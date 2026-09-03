"""Graph helpers: calendar -> online meeting -> transcript content.

Security invariant: meeting identity is ALWAYS derived from the authenticated
caller's calendarView. In app-only mode the mailbox ID comes only from the
validated token's oid claim. No caller-supplied user IDs, join URLs, meeting
IDs, thread IDs or transcript tokens ever reach Graph.
"""

from __future__ import annotations

import datetime as dt
import re
import time
from urllib.parse import quote

import requests

GRAPH = "https://graph.microsoft.com/v1.0"


class AccessDeniedError(Exception):
    """HTTP 403 — access denied. Distinct from 'no transcript exists'."""


class AuthExpiredError(Exception):
    """HTTP 401 — token rejected; interaction required."""


class ThrottledError(Exception):
    """HTTP 429/5xx persisted after bounded retries."""


class OnlineMeetingNotFoundError(Exception):
    """The calendar event's join URL resolved to no online meeting."""


class TranscriptContentMissingError(Exception):
    """A listed transcript's content 404'd (deleted between list and fetch)."""


class _NoTranscriptCollection(Exception):
    """Internal: the transcripts collection is empty/absent for the meeting."""


_SAFE_GRAPH_403_CODES = {
    "GraphAccessToTranscriptsDisabled",
    "SpeakerAttributionNotAllowed",
}


def _safe_graph_error(response, path: str) -> tuple[str, str]:
    """Return finite-allowlisted inner code and endpoint class, never IDs/body."""
    try:
        error = response.json().get("error") or {}
        inner = error.get("innerError") or {}
        candidate = inner.get("code")
    except (AttributeError, TypeError, ValueError):
        candidate = None
    code = (
        candidate
        if isinstance(candidate, str) and candidate in _SAFE_GRAPH_403_CODES
        else "Forbidden"
    )
    if "/transcripts/" in path and path.rstrip("/").endswith("content"):
        operation = "transcript_content"
    elif path.rstrip("/").endswith("transcripts"):
        operation = "transcript_list"
    elif "attendanceReports" in path:
        operation = "attendance_report"
    else:
        operation = "graph_request"
    return code, operation


def _get(token: str, path: str, params: dict | None = None, raw: bool = False,
         max_retries: int = 3):
    """GET with status-specific errors, Retry-After on 429, backoff on 5xx."""
    delay = 1.0
    for attempt in range(max_retries + 1):
        response = requests.get(
            f"{GRAPH}{path}",
            headers={
                "Authorization": f"Bearer {token}",
                "Prefer": 'outlook.timezone="UTC"',  # pin event times to UTC
            },
            params=params,
            timeout=60,
        )
        status = response.status_code
        if status < 400:
            return response.text if raw else response.json()
        if status == 401:
            raise AuthExpiredError(f"401 from Graph on {path} — sign in again.")
        if status == 403:
            code, operation = _safe_graph_error(response, path)
            raise AccessDeniedError(f"Graph 403 code={code} operation={operation}")
        if status == 404:
            if "transcripts" in path:
                if path.rstrip("/").endswith("transcripts"):
                    raise _NoTranscriptCollection(path)
                raise TranscriptContentMissingError(path)
            raise OnlineMeetingNotFoundError(path)
        if status == 429 or 500 <= status < 600:
            if attempt < max_retries:
                retry_after = response.headers.get("Retry-After")
                wait = float(retry_after) if retry_after and retry_after.isdigit() else delay
                time.sleep(min(wait, 30))
                delay *= 2
                continue
            raise ThrottledError(
                f"Graph {path} still failing after {max_retries} retries "
                f"(last status {status})."
            )
        response.raise_for_status()
    raise ThrottledError(f"Graph {path}: exhausted retries.")


def _parse_dt(value: str | None) -> dt.datetime | None:
    if not value:
        return None
    if value.endswith("Z"):
        return dt.datetime.fromisoformat(value.removesuffix("Z")).replace(tzinfo=dt.UTC)
    parsed = dt.datetime.fromisoformat(value)
    # calendarView event times are pinned to UTC via the Prefer header; treat
    # any residual offset-naive value as UTC.
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=dt.UTC)


def _segment(value: str) -> str:
    return quote(value, safe="")


def _user_path(user_id: str | None, suffix: str) -> str:
    return f"/users/{_segment(user_id)}{suffix}" if user_id else f"/me{suffix}"


def list_teams_meetings(token: str, days: int = 7, limit: int = 100,
                        user_id: str | None = None) -> list[dict]:
    """Teams meetings on the user's calendar in the window, newest first.

    Paginates @odata.nextLink (calendarView cannot $filter on isOnlineMeeting,
    so online-meeting filtering is client-side). Subject lookups pass a larger
    limit; the user-facing list passes its own.
    """
    now = dt.datetime.now(dt.UTC)
    start = now - dt.timedelta(days=days)
    params = {
        "startDateTime": start.isoformat(),
        "endDateTime": now.isoformat(),
        "$select": "id,subject,start,end,organizer,attendees,onlineMeeting,"
                   "isOrganizer,responseStatus,seriesMasterId,type,isCancelled",
        "$orderby": "start/dateTime desc",
        "$top": "50",
    }
    meetings: list[dict] = []
    url: str | None = _user_path(user_id, "/calendarView")
    pages = 0
    while url and pages < 20 and len(meetings) < limit:
        pages += 1
        if url.startswith("http"):
            data = _get(token, url.replace(GRAPH, ""))
        else:
            data = _get(token, url, params=params)
        for event in data.get("value", []):
            # calendarView can't $filter on isOnlineMeeting; onlineMeeting
            # presence is the reliable client-side signal.
            join_url = (event.get("onlineMeeting") or {}).get("joinUrl")
            if not join_url:
                continue
            meetings.append(
                {
                    "event_id": event.get("id"),
                    "subject": event.get("subject") or "(no subject)",
                    "start": (event.get("start") or {}).get("dateTime"),
                    "end": (event.get("end") or {}).get("dateTime"),
                    "organizer": ((event.get("organizer") or {}).get("emailAddress") or {}).get("name"),
                    "organizer_id": ((event.get("organizer") or {}).get("emailAddress") or {}).get("address"),
                    "is_organizer": bool(event.get("isOrganizer")),
                    "is_cancelled": bool(event.get("isCancelled")),
                    # event.type is authoritative (singleInstance/occurrence/
                    # seriesMaster); seriesMasterId proves unreliable.
                    "is_recurring": event.get("type") in ("occurrence", "seriesMaster"),
                    "response": ((event.get("responseStatus") or {}).get("response")),
                    "attendees": [
                        ((a.get("emailAddress") or {}).get("address") or "").lower()
                        for a in event.get("attendees", [])
                    ],
                    "join_url": join_url,
                }
            )
        url = data.get("@odata.nextLink")
        params = None  # nextLink already carries the query
    return meetings[:limit]


ACCESS_GATES = frozenset({"invited", "accepted", "attended"})


def check_calendar_gate(
    meeting: dict, user_email: str, access_gate: str
) -> tuple[bool, str]:
    """Apply the selected calendar-backed gate before transcript access.

    ``attended`` first requires a real invitation, then separately requires a
    positive app-only attendance record after the online meeting is resolved.
    """
    if access_gate not in ACCESS_GATES:
        raise ValueError(f"invalid access gate: {access_gate}")
    if meeting.get("is_cancelled"):
        return False, "the meeting was cancelled"
    if meeting["is_organizer"]:
        return True, "organizer"
    email = user_email.lower()
    if email not in meeting["attendees"]:
        return False, "your address is not in the attendee list (forwarded invite?)"
    if access_gate == "accepted" and meeting.get("response") != "accepted":
        return False, "the invitation was not accepted"
    return True, "listed invitee"


def check_attended_gate(attendance_seconds: int) -> tuple[bool, str]:
    """Require positive joined time for the attended gate."""
    if attendance_seconds <= 0:
        return False, "the attendance report does not show that you joined"
    return True, "attendance verified"


def revalidate_event(token: str, meeting: dict, user_email: str,
                     user_id: str | None = None,
                     access_gate: str = "invited") -> tuple[bool, str]:
    """Re-read the calendar event immediately before fetching.

    Catches stale cached selections: event deleted/cancelled or attendee list
    changed after listing.
    """
    event_id = meeting.get("event_id")
    if not event_id:
        return False, "no calendar event ID recorded for this meeting"
    try:
        event = _get(
            token, _user_path(user_id, f"/events/{_segment(event_id)}"),
            params={
                "$select": "id,attendees,isOrganizer,responseStatus,isCancelled,"
                "organizer,onlineMeeting"
            },
        )
    except OnlineMeetingNotFoundError:
        return False, "the calendar event no longer exists"
    if event.get("isCancelled"):
        return False, "the meeting was cancelled"
    response = ((event.get("responseStatus") or {}).get("response"))
    attendees = [
        ((a.get("emailAddress") or {}).get("address") or "").lower()
        for a in event.get("attendees", [])
    ]
    is_organizer = bool(event.get("isOrganizer"))
    fresh = {
        "is_organizer": is_organizer,
        "response": response,
        "attendees": attendees,
    }
    allowed, reason = check_calendar_gate(fresh, user_email, access_gate)
    if not allowed:
        return False, reason
    join_url = (event.get("onlineMeeting") or {}).get("joinUrl")
    if not join_url:
        return False, "online meeting details are no longer available"
    organizer = (event.get("organizer") or {}).get("emailAddress") or {}
    meeting["organizer"] = organizer.get("name")
    meeting["organizer_id"] = organizer.get("address")
    meeting["join_url"] = join_url
    meeting["response"] = fresh["response"]
    meeting["attendees"] = fresh["attendees"]
    meeting["is_organizer"] = fresh["is_organizer"]
    return True, reason


ATTENDANCE_OCCURRENCE_LEAD = dt.timedelta(minutes=30)


def _attendance_reports_for_occurrence(
    reports: list[dict], occurrence_start: str | None, occurrence_end: str | None
) -> list[dict]:
    """Select reports whose actual start matches the chosen occurrence."""
    start = _parse_dt(occurrence_start)
    end = _parse_dt(occurrence_end)
    if not start or (end and end < start):
        return []
    selected = []
    for report in reports:
        report_start = _parse_dt(report.get("meetingStartDateTime"))
        if report_start and abs(report_start - start) <= ATTENDANCE_OCCURRENCE_LEAD:
            selected.append(report)
    return selected


def get_my_attendance_seconds(
    token: str,
    meeting_id: str,
    user_email: str,
    user_id: str | None = None,
    occurrence_start: str | None = None,
    occurrence_end: str | None = None,
) -> int:
    """Hard attendance proof from Teams attendance reports (`attended` gate).

    Returns total seconds joined. Raises AttendanceReportUnavailableError when
    reports are inaccessible/absent — never conflate that with zero seconds.
    """
    try:
        reports = _get(
            token,
            _user_path(
                user_id,
                f"/onlineMeetings/{_segment(meeting_id)}/attendanceReports",
            ),
            params={
                "$select": "id,meetingStartDateTime,meetingEndDateTime"
            },
        ).get("value", [])
    except AccessDeniedError as exc:
        raise AttendanceReportUnavailableError(
            "attendance reports are organizer/co-organizer only under the current "
            "Teams meeting policy"
        ) from exc
    except OnlineMeetingNotFoundError as exc:
        raise AttendanceReportUnavailableError(
            "no attendance report exists for this meeting (expired or never generated)"
        ) from exc
    if not reports:
        raise AttendanceReportUnavailableError("no attendance report exists for this meeting")
    reports = _attendance_reports_for_occurrence(
        reports, occurrence_start, occurrence_end
    )
    if not reports:
        raise AttendanceReportUnavailableError(
            "no attendance report exists for this occurrence"
        )
    email = user_email.lower()
    total = 0
    for report in reports:
        records = _get(
            token,
            _user_path(
                user_id,
                f"/onlineMeetings/{_segment(meeting_id)}/attendanceReports/"
                f"{_segment(report['id'])}"
                f"/attendanceRecords",
            ),
            params={"$filter": f"emailAddress eq '{user_email}'"},
        ).get("value", [])
        for record in records:
            if (record.get("emailAddress") or "").lower() == email:
                total += int(record.get("totalAttendanceInSeconds") or 0)
    return total


class AttendanceReportUnavailableError(Exception):
    """Attendance reports inaccessible (403) or absent.

    Usually the Teams meeting policy restricts reports to the organizer, or
    the report expired. A tenant policy boundary — never code around it.
    """


def get_event(token: str, event_id: str, user_id: str | None = None) -> dict:
    """Fetch one event from the authenticated caller's calendar by id.

    Delegated mode uses `/me/events`; application mode uses only the validated
    caller oid supplied by the server. A meeting handle cannot select a mailbox.
    """
    event = _get(
        token, _user_path(user_id, f"/events/{_segment(event_id)}"),
        params={"$select": "id,subject,start,end,organizer,attendees,"
                           "onlineMeeting,isOrganizer,responseStatus,"
                           "seriesMasterId,isCancelled,type"},
    )
    return event


def shape_event(event: dict) -> dict | None:
    """Normalise a raw Graph event into the meeting dict; None if not an
    online meeting with a join URL."""
    join_url = (event.get("onlineMeeting") or {}).get("joinUrl")
    if not join_url:
        return None
    return {
        "event_id": event.get("id"),
        "subject": event.get("subject") or "(no subject)",
        "start": (event.get("start") or {}).get("dateTime"),
        "end": (event.get("end") or {}).get("dateTime"),
        "organizer": ((event.get("organizer") or {}).get("emailAddress") or {}).get("name"),
        "organizer_id": ((event.get("organizer") or {}).get("emailAddress") or {}).get("address"),
        "is_organizer": bool(event.get("isOrganizer")),
        "is_cancelled": bool(event.get("isCancelled")),
        "is_recurring": event.get("type") in ("occurrence", "seriesMaster"),
        "response": ((event.get("responseStatus") or {}).get("response")),
        "attendees": [
            ((a.get("emailAddress") or {}).get("address") or "").lower()
            for a in event.get("attendees", [])
        ],
        "join_url": join_url,
    }


def resolve_online_meeting(token: str, join_url: str,
                           user_id: str | None = None) -> dict:
    escaped_join_url = join_url.replace("'", "''")
    filt = f"joinWebUrl eq '{escaped_join_url}'"
    data = _get(
        token,
        _user_path(user_id, "/onlineMeetings"),
        params={"$filter": filt},
    )
    meetings = data.get("value", [])
    if not meetings:
        raise OnlineMeetingNotFoundError(
            "No online meeting found for this calendar event."
        )
    return meetings[0]


def list_transcripts(token: str, meeting_id: str,
                     user_id: str | None = None) -> list[dict]:
    try:
        data = _get(
            token,
            _user_path(
                user_id,
                f"/onlineMeetings/{_segment(meeting_id)}/transcripts",
            ),
        )
    except _NoTranscriptCollection:
        return []
    return [
        {"id": t["id"], "created": t.get("createdDateTime")}
        for t in data.get("value", [])
    ]


# Occurrence correlation window: transcripts can be created slightly before
# the scheduled start (early recording) and processing can lag after the end.
OCCURRENCE_LEAD = dt.timedelta(minutes=30)
OCCURRENCE_TAIL = dt.timedelta(hours=4)


def transcripts_for_occurrence(
    transcripts: list[dict], occ_start: str | None, occ_end: str | None
) -> list[dict]:
    """All transcript segments belonging to ONE occurrence, in created order.

    Recurring series share one onlineMeeting object whose transcripts
    accumulate across occurrences. A single occurrence can also produce
    multiple segments (transcription stopped/restarted). Returns every segment
    in the window; the caller concatenates. Empty list = no transcript for
    this occurrence.
    """
    start = _parse_dt(occ_start)
    end = _parse_dt(occ_end)
    if not start:
        return transcripts[:1]  # no occurrence info — single-instance meeting
    window_start = start - OCCURRENCE_LEAD
    window_end = (end or start) + OCCURRENCE_TAIL
    in_window = [
        t for t in transcripts
        if (created := _parse_dt(t["created"])) and window_start <= created <= window_end
    ]
    return sorted(in_window, key=lambda t: t["created"] or "")


def get_transcript_vtt(token: str, meeting_id: str, transcript_id: str,
                       user_id: str | None = None) -> str:
    return _get(
        token,
        _user_path(
            user_id,
            f"/onlineMeetings/{_segment(meeting_id)}/transcripts/"
            f"{_segment(transcript_id)}/content",
        ),
        params={"$format": "text/vtt"},
        raw=True,
    )


_VTT_CUE = re.compile(
    r"^(?P<start>\d{2}:\d{2}:\d{2}\.\d+) --> .*?\n(?P<body>(?:.*\n)+?)(?=\n|\Z)",
    re.MULTILINE,
)
_VTT_SPEAKER = re.compile(r"<v\s+([^>]+)>(.*?)</v>", re.DOTALL)


def vtt_to_dialogue(vtt: str) -> str | None:
    """Convert Teams VTT to 'Speaker: text' lines, merging same-speaker cues.

    Returns None when no cues parse — the caller must NOT silently fall back
    to raw VTT (unstructured, larger injection surface).
    """
    lines: list[str] = []
    last_speaker = None
    for cue in _VTT_CUE.finditer(vtt):
        body = cue.group("body").strip()
        for match in _VTT_SPEAKER.finditer(body):
            speaker = match.group(1).strip()
            text = re.sub(r"<[^>]+>", "", match.group(2)).strip()
            if not text:
                continue
            if speaker == last_speaker and lines:
                lines[-1] += " " + text
            else:
                lines.append(f"{speaker}: {text}")
                last_speaker = speaker
    return "\n".join(lines) if lines else None

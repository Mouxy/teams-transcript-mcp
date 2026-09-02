"""Graph helpers: calendar -> online meeting -> transcript content.

Security invariant: meeting identity is ALWAYS derived from the signed-in
user's calendarView. No caller-supplied join URLs, meeting IDs, thread IDs
or transcript tokens ever reach Graph.
"""

from __future__ import annotations

import datetime as dt
import re
import time

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
            raise AccessDeniedError(
                f"403 from Graph on {path}. This is a permissions denial, NOT a "
                "missing artifact. Response: " + response.text[:300]
            )
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


def list_teams_meetings(token: str, days: int = 7, limit: int = 100) -> list[dict]:
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
                   "isOrganizer,responseStatus,seriesMasterId,type",
        "$orderby": "start/dateTime desc",
        "$top": "50",
    }
    meetings: list[dict] = []
    url: str | None = "/me/calendarView"
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
                    "is_organizer": bool(event.get("isOrganizer")),
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


def check_attendance(meeting: dict, user_email: str) -> tuple[bool, str]:
    """Invite gate: calendar presence + accepted invite (or organizer).

    Code-enforced before any transcript call. Matches Teams' own trust model:
    anyone holding the invite may request to join. The optional strict mode
    (get_my_attendance_seconds) additionally proves actual join time, but is
    currently blocked tenant-wide for delegated attendees (403).
    """
    if meeting["is_organizer"]:
        return True, "organizer"
    email = user_email.lower()
    if email not in meeting["attendees"]:
        return False, "your address is not in the attendee list (forwarded invite?)"
    if meeting["response"] != "accepted":
        return False, f"your response status is '{meeting['response']}', not 'accepted'"
    return True, "accepted attendee"


def revalidate_event(token: str, meeting: dict, user_email: str) -> tuple[bool, str]:
    """Re-read the calendar event immediately before fetching.

    Catches stale cached selections: event deleted, response changed to
    declined, attendee list changed after listing.
    """
    event_id = meeting.get("event_id")
    if not event_id:
        return False, "no calendar event ID recorded for this meeting"
    try:
        event = _get(
            token, f"/me/events/{event_id}",
            params={"$select": "id,attendees,isOrganizer,responseStatus,isCancelled"},
        )
    except OnlineMeetingNotFoundError:
        return False, "the calendar event no longer exists"
    if event.get("isCancelled"):
        return False, "the meeting was cancelled"
    fresh = {
        "is_organizer": bool(event.get("isOrganizer")),
        "response": ((event.get("responseStatus") or {}).get("response")),
        "attendees": [
            ((a.get("emailAddress") or {}).get("address") or "").lower()
            for a in event.get("attendees", [])
        ],
    }
    return check_attendance(fresh, user_email)


def get_my_attendance_seconds(token: str, meeting_id: str, user_email: str) -> int:
    """Hard attendance proof from Teams attendance reports (strict mode).

    Returns total seconds joined. Raises AttendanceReportUnavailableError when
    reports are inaccessible/absent — never conflate that with zero seconds.
    """
    try:
        reports = _get(
            token, f"/me/onlineMeetings/{meeting_id}/attendanceReports"
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
    email = user_email.lower()
    total = 0
    for report in reports:
        records = _get(
            token,
            f"/me/onlineMeetings/{meeting_id}/attendanceReports/{report['id']}"
            f"/attendanceRecords",
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


def get_event(token: str, event_id: str) -> dict:
    """Fetch one event from the USER'S OWN calendar by id (for id-keyed lookups).

    Raises OnlineMeetingNotFoundError if absent. Only /me/events — a supplied
    id can never escape the caller's own calendar.
    """
    event = _get(
        token, f"/me/events/{event_id}",
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
        "is_organizer": bool(event.get("isOrganizer")),
        "is_recurring": event.get("type") in ("occurrence", "seriesMaster"),
        "response": ((event.get("responseStatus") or {}).get("response")),
        "attendees": [
            ((a.get("emailAddress") or {}).get("address") or "").lower()
            for a in event.get("attendees", [])
        ],
        "join_url": join_url,
    }


def resolve_online_meeting(token: str, join_url: str) -> dict:
    filt = f"joinWebUrl eq '{join_url}'"
    data = _get(token, "/me/onlineMeetings", params={"$filter": filt})
    meetings = data.get("value", [])
    if not meetings:
        raise OnlineMeetingNotFoundError(
            "No online meeting found for this calendar event."
        )
    return meetings[0]


def list_transcripts(token: str, meeting_id: str) -> list[dict]:
    try:
        data = _get(token, f"/me/onlineMeetings/{meeting_id}/transcripts")
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


def get_transcript_vtt(token: str, meeting_id: str, transcript_id: str) -> str:
    return _get(
        token,
        f"/me/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content",
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

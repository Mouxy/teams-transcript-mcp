"""Acceptance-criteria tests with a mocked Graph. No live tenant calls."""

from __future__ import annotations

import json

import pytest

from transcript_sync import graph, server

# ---------- fixtures ----------

def make_event(subject="Standup", organizer="Alice", is_org=False, response="accepted",
               attendees=("user@example.com",), join="https://teams.microsoft.com/l/meetup-join/x",
               start="2026-08-11T12:00:00.0000000", end="2026-08-11T13:00:00.0000000",
               event_id="evt-1"):
    return {
        "event_id": event_id,
        "subject": subject,
        "start": start,
        "end": end,
        "organizer": organizer,
        "is_organizer": is_org,
        "response": response,
        "attendees": [a.lower() for a in attendees],
        "join_url": join,
    }


SAMPLE_VTT = """WEBVTT

00:00:01.000 --> 00:00:04.000
<v Alice Smith>Good morning everyone.</v>

00:00:05.000 --> 00:00:08.000
<v Alice Smith>Let's start.</v>

00:00:09.000 --> 00:00:12.000
<v Bob Jones>Deploy is green.</v>
"""


# ---------- VTT parsing ----------

def test_vtt_merges_same_speaker_cues():
    out = graph.vtt_to_dialogue(SAMPLE_VTT)
    assert out == (
        "Alice Smith: Good morning everyone. Let's start.\n"
        "Bob Jones: Deploy is green."
    )


def test_vtt_unparsable_returns_none_not_raw():
    assert graph.vtt_to_dialogue("WEBVTT\n\nno cues here") is None


# ---------- calendar access gate matrix ----------

@pytest.mark.parametrize("gate,is_org,response,attendees,expected", [
    ("invited", True, "organizer", [], True),
    ("invited", False, "accepted", ["user@example.com"], True),
    ("invited", False, "tentative", ["user@example.com"], True),
    ("invited", False, "declined", ["user@example.com"], True),
    ("invited", False, "notResponded", ["user@example.com"], True),
    ("accepted", True, "organizer", [], True),
    ("accepted", False, "accepted", ["user@example.com"], True),
    ("accepted", False, "tentative", ["user@example.com"], False),
    ("accepted", False, "declined", ["user@example.com"], False),
    ("accepted", False, "notResponded", ["user@example.com"], False),
    # Attended mode first requires a real invitation. Positive joined time is
    # evaluated separately from an app-only attendance report.
    ("attended", True, "organizer", [], True),
    ("attended", False, "declined", ["user@example.com"], True),
    ("attended", False, "accepted", ["other@example.com"], False),
])
def test_calendar_access_gate(gate, is_org, response, attendees, expected):
    meeting = make_event(is_org=is_org, response=response, attendees=attendees)
    allowed, _ = graph.check_calendar_gate(meeting, "user@example.com", gate)
    assert allowed is expected


@pytest.mark.parametrize("seconds,expected", [(1, True), (30, True), (0, False)])
def test_attended_gate_requires_positive_joined_time(seconds, expected):
    allowed, _ = graph.check_attended_gate(seconds)
    assert allowed is expected


# ---------- occurrence matching ----------

def _ts(created):
    return {"id": f"t-{created}", "created": created}


def test_occurrence_window_selects_correct_segment():
    transcripts = [
        _ts("2026-07-14T12:05:00Z"),   # July occurrence
        _ts("2026-08-11T11:57:40Z"),   # Aug occurrence (early recording start)
        _ts("2026-08-11T12:40:00Z"),   # second segment, same occurrence
    ]
    got = graph.transcripts_for_occurrence(
        transcripts, "2026-08-11T12:00:00.0000000", "2026-08-11T13:00:00.0000000"
    )
    assert [g["id"] for g in got] == ["t-2026-08-11T11:57:40Z", "t-2026-08-11T12:40:00Z"]


def test_occurrence_outside_window_returns_empty():
    got = graph.transcripts_for_occurrence(
        [_ts("2026-07-14T12:05:00Z")],
        "2026-08-11T12:00:00.0000000", "2026-08-11T13:00:00.0000000",
    )
    assert got == []


def test_occurrence_handles_naive_event_times():
    # calendarView times are offset-naive; transcript times are Z-suffixed.
    got = graph.transcripts_for_occurrence(
        [_ts("2026-08-11T12:10:00Z")],
        "2026-08-11T12:00:00.0000000", "2026-08-11T13:00:00.0000000",
    )
    assert len(got) == 1


def test_single_instance_no_occurrence_info_returns_first():
    got = graph.transcripts_for_occurrence([_ts("2026-08-11T12:10:00Z")], None, None)
    assert len(got) == 1


# ---------- error taxonomy (mocked HTTP) ----------

class FakeResponse:
    def __init__(self, status, body=None, headers=None):
        self.status_code = status
        self._body = body or {}
        self.text = json.dumps(self._body) if isinstance(body, (dict, list)) else str(body)
        self.headers = headers or {}
        self.content = b"x"

    def json(self):
        return self._body

    def raise_for_status(self):
        import requests
        raise requests.HTTPError(f"{self.status_code}", response=self)


def test_403_raises_access_denied_with_safe_graph_code(monkeypatch):
    body = {
        "error": {
            "code": "Forbidden",
            "message": "sensitive provider detail",
            "innerError": {"code": "SpeakerAttributionNotAllowed"},
        }
    }
    monkeypatch.setattr(graph.requests, "get",
                        lambda *a, **k: FakeResponse(403, body))
    path = "/me/onlineMeetings/secret-meeting/transcripts/secret-transcript/content"
    with pytest.raises(graph.AccessDeniedError) as caught:
        graph._get("tok", path)
    detail = str(caught.value)
    assert "SpeakerAttributionNotAllowed" in detail
    assert "transcript_content" in detail
    assert "secret-meeting" not in detail
    assert "secret-transcript" not in detail
    assert "sensitive provider detail" not in detail


def test_403_unknown_codes_are_not_logged(monkeypatch):
    body = {
        "error": {
            "code": "TOP_LEVEL_SECRET_456",
            "innerError": {"code": "SECRET_TRANSCRIPT_ID_123"},
        }
    }
    monkeypatch.setattr(graph.requests, "get",
                        lambda *a, **k: FakeResponse(403, body))
    with pytest.raises(graph.AccessDeniedError) as caught:
        graph._get("tok", "/me/onlineMeetings/x/transcripts/y/content")
    detail = str(caught.value)
    assert "code=Forbidden" in detail
    assert "SECRET_TRANSCRIPT_ID_123" not in detail
    assert "TOP_LEVEL_SECRET_456" not in detail


def test_403_non_string_inner_code_maps_to_forbidden(monkeypatch):
    body = {"error": {"innerError": {"code": ["not", "hashable"]}}}
    monkeypatch.setattr(graph.requests, "get",
                        lambda *a, **k: FakeResponse(403, body))
    with pytest.raises(graph.AccessDeniedError, match="code=Forbidden"):
        graph._get("tok", "/me/onlineMeetings/x/transcripts/y/content")


def test_401_raises_auth_expired(monkeypatch):
    monkeypatch.setattr(graph.requests, "get", lambda *a, **k: FakeResponse(401))
    with pytest.raises(graph.AuthExpiredError):
        graph._get("tok", "/me/calendarView")


def test_404_on_transcript_collection_is_internal_no_transcript(monkeypatch):
    monkeypatch.setattr(graph.requests, "get", lambda *a, **k: FakeResponse(404))
    with pytest.raises(graph._NoTranscriptCollection):
        graph._get("tok", "/me/onlineMeetings/x/transcripts")


def test_404_on_content_is_content_missing(monkeypatch):
    monkeypatch.setattr(graph.requests, "get", lambda *a, **k: FakeResponse(404))
    with pytest.raises(graph.TranscriptContentMissingError):
        graph._get("tok", "/me/onlineMeetings/x/transcripts/y/content")


def test_429_honours_retry_after_then_succeeds(monkeypatch):
    calls = []
    responses = [
        FakeResponse(429, headers={"Retry-After": "0"}),
        FakeResponse(200, {"value": []}),
    ]

    def fake_get(*a, **k):
        calls.append(1)
        return responses.pop(0)

    monkeypatch.setattr(graph.requests, "get", fake_get)
    monkeypatch.setattr(graph.time, "sleep", lambda s: None)
    assert graph._get("tok", "/me/calendarView") == {"value": []}
    assert len(calls) == 2


def test_persistent_429_raises_throttled(monkeypatch):
    monkeypatch.setattr(graph.requests, "get",
                        lambda *a, **k: FakeResponse(429, headers={"Retry-After": "0"}))
    monkeypatch.setattr(graph.time, "sleep", lambda s: None)
    with pytest.raises(graph.ThrottledError):
        graph._get("tok", "/me/calendarView", max_retries=2)


def test_prefer_utc_header_sent(monkeypatch):
    seen = {}

    def fake_get(url, headers=None, **k):
        seen.update(headers or {})
        return FakeResponse(200, {"value": []})

    monkeypatch.setattr(graph.requests, "get", fake_get)
    graph._get("tok", "/me/calendarView")
    assert seen.get("Prefer") == 'outlook.timezone="UTC"'


# ---------- pagination ----------

def test_pagination_follows_nextlink(monkeypatch):
    pages = {
        "/me/calendarView": {
            "value": [{"id": "1", "subject": "A", "isOnlineMeeting": True,
                       "onlineMeeting": {"joinUrl": "u1"}, "start": {}, "end": {},
                       "organizer": {"emailAddress": {"name": "X"}},
                       "attendees": [], "responseStatus": {"response": "accepted"}}],
            "@odata.nextLink": graph.GRAPH + "/me/calendarView?$skip=1",
        },
        "/me/calendarView?$skip=1": {
            "value": [{"id": "2", "subject": "B", "isOnlineMeeting": True,
                       "onlineMeeting": {"joinUrl": "u2"}, "start": {}, "end": {},
                       "organizer": {"emailAddress": {"name": "Y"}},
                       "attendees": [], "responseStatus": {"response": "accepted"}}],
        },
    }
    monkeypatch.setattr(graph, "_get",
                        lambda token, path, params=None, **k: pages[path])
    meetings = graph.list_teams_meetings("tok", days=7)
    assert [m["event_id"] for m in meetings] == ["1", "2"]


def test_non_online_events_filtered(monkeypatch):
    page = {"value": [
        {"id": "1", "subject": "Offline", "onlineMeeting": None,
         "start": {}, "end": {}, "organizer": {"emailAddress": {}},
         "attendees": [], "responseStatus": {}},
    ]}
    monkeypatch.setattr(graph, "_get", lambda *a, **k: page)
    assert graph.list_teams_meetings("tok") == []


# ---------- stale selection revalidation ----------

def test_revalidation_allows_declined_invitee(monkeypatch):
    m = make_event(response="accepted")
    declined_event = {"id": "evt-1", "isOrganizer": False,
                      "responseStatus": {"response": "declined"},
                      "organizer": {"emailAddress": {
                          "name": "Alice", "address": "alice@example.com"}},
                      "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/x"},
                      "attendees": [{"emailAddress": {"address": "user@example.com"}}]}
    monkeypatch.setattr(graph, "_get", lambda *a, **k: declined_event)
    ok, reason = graph.revalidate_event("tok", m, "user@example.com")
    assert ok is True and reason == "listed invitee"


def test_revalidation_rejects_deleted_event(monkeypatch):
    m = make_event()
    monkeypatch.setattr(
        graph, "_get",
        lambda *a, **k: (_ for _ in ()).throw(graph.OnlineMeetingNotFoundError("gone")))
    ok, reason = graph.revalidate_event("tok", m, "user@example.com")
    assert ok is False and "no longer exists" in reason


# ---------- audit log ----------

def test_audit_writes_record_with_permissions(tmp_path, monkeypatch):
    monkeypatch.setattr(server, "AUDIT_DIR", tmp_path / "ts")
    monkeypatch.setattr(server, "AUDIT_LOG", tmp_path / "ts" / "audit.log")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    server._audit(make_event(), "mid", "ok", attendance_seconds=None)
    line = (tmp_path / "ts" / "audit.log").read_text().strip()
    record = json.loads(line)
    assert record["result"] == "ok"
    assert record["user"] == "user@example.com"
    assert record["access_gate"] == "invited"
    assert (tmp_path / "ts").stat().st_mode & 0o777 == 0o700
    assert (tmp_path / "ts" / "audit.log").stat().st_mode & 0o777 == 0o600


# ---------- subject resolution ----------

def test_subject_resolution_ambiguous_returns_candidates(monkeypatch):
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server.core.graph, "list_teams_meetings",
                        lambda *a, **k: [make_event(subject="Townhall", event_id="e1"),
                                         make_event(subject="Townhall", event_id="e2",
                                                    start="2026-07-14T12:00:00.0000000")])
    server._last_meetings = []
    out = server.get_transcript("Townhall")
    assert out["status"] == "error"
    assert out["error_code"] == "ambiguous_match"
    assert len(out["candidates"]) == 2
    assert all(c["id"].startswith("mtg_") for c in out["candidates"])


def test_subject_resolution_no_match_names_window(monkeypatch):
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server.core.graph, "list_teams_meetings", lambda *a, **k: [])
    out = server.get_transcript("Nonexistent")
    assert out["error_code"] == "not_found"
    assert "60 days" in out["message"]


def test_never_invited_meeting_is_unreachable_by_design(monkeypatch):
    """'Meeting between X and Y' where user is not invited never resolves."""
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server.core.graph, "list_teams_meetings",
                        lambda *a, **k: [make_event(subject="X and Y sync")])
    # User IS invited here (fixture) — but rejection happens at the gate:
    monkeypatch.setattr(server.core.graph, "check_calendar_gate",
                        lambda m, u, g: (False, "not a confirmed participant"))
    monkeypatch.setattr(server.core.graph, "revalidate_event", lambda *a: (True, ""))
    monkeypatch.setattr(server, "_audit", lambda *a, **k: None)
    out = server.get_transcript("X and Y")
    assert out["error_code"] == "not_attendee"


def test_is_recurring_uses_event_type_not_series_master_id():
    """Claude's live finding: seriesMasterId misclassified a one-off."""
    single = {"type": "singleInstance", "seriesMasterId": None,
              "onlineMeeting": {"joinUrl": "u"}, "start": {}, "end": {},
              "organizer": {"emailAddress": {}}, "attendees": [],
              "responseStatus": {}}
    occurrence = dict(single, type="occurrence", seriesMasterId="master-1")
    assert graph.shape_event(single)["is_recurring"] is False
    assert graph.shape_event(occurrence)["is_recurring"] is True
    # The live false positive: non-null seriesMasterId on a singleInstance
    weird = dict(single, seriesMasterId="unexpected-value")
    assert graph.shape_event(weird)["is_recurring"] is False


# ---------- opaque ids ----------

def test_meeting_id_roundtrip():
    m = make_event(event_id="AAMkADAGAADDdm4NAAA=")
    mid = server.core.make_meeting_id(m)
    assert mid.startswith("mtg_")
    assert server.core._decode_meeting_id(mid) == "AAMkADAGAADDdm4NAAA="


def test_malformed_id_rejected_without_graph_call(monkeypatch):
    """Forged/malformed handles fail fast: no Graph call, generic message,
    internals only in the audit record."""
    calls = []
    audits = []
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server.core.graph, "get_event",
                        lambda *a: calls.append(1))
    monkeypatch.setattr(server, "_audit",
                        lambda *a, **k: audits.append(a))
    out = server.get_transcript("mtg_!!!notb64!!!")
    assert out["error_code"] == "invalid_id"
    assert calls == []                       # fail-fast: zero Graph calls
    assert "AAMk" not in out["message"]      # no internals leaked
    assert audits and audits[0][2] == "invalid_id"


def test_forged_id_never_reaches_graph_even_when_wellformed(monkeypatch):
    """A base64 handle decoding to a non-Exchange id is rejected pre-flight."""
    import base64 as b64
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server, "_audit", lambda *a, **k: None)
    calls = []
    monkeypatch.setattr(server.core.graph, "get_event", lambda *a: calls.append(1))
    fake = "mtg_" + b64.urlsafe_b64encode(b"someone-elses-event").decode()
    out = server.get_transcript(fake)
    assert out["error_code"] == "invalid_id" and calls == []


def test_id_lookup_scoped_to_own_calendar(monkeypatch):
    """A well-formed id only resolves via /me/events/{id} — own calendar."""
    seen = {}

    def fake_get_event(token, event_id):
        seen["id"] = event_id
        return {"id": event_id, "subject": "Townhall", "onlineMeeting": None}

    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server.core.graph, "get_event", fake_get_event)
    monkeypatch.setattr(server, "_audit", lambda *a, **k: None)
    mid = server.core.make_meeting_id({"event_id": "AAMkAGenuine42="})
    out = server.get_transcript(mid)
    assert seen["id"] == "AAMkAGenuine42="
    assert out["error_code"] == "not_found"  # not an online meeting


def test_graph_500_lands_in_taxonomy_not_raw(monkeypatch):
    """Catch-all: unexpected Graph errors map to a coded error, never raw."""
    monkeypatch.setattr(server.auth, "get_token", lambda: "tok")
    monkeypatch.setattr(server.auth, "auth_status",
                        lambda: {"account": "user@example.com"})
    monkeypatch.setattr(server, "_audit", lambda *a, **k: None)

    def boom(*a, **k):
        raise RuntimeError("Graph exploded: /me/events/AAMk internal detail")
    monkeypatch.setattr(server.core.graph, "get_event", boom)
    mid = server.core.make_meeting_id({"event_id": "AAMkAGenuine42="})
    out = server.get_transcript(mid)
    assert out["status"] == "error"
    assert "internal detail" not in out["message"]
    assert out["error_code"] in ("invalid_id", "error")


# ---------- tri-state availability ----------

def _ctx_for_probe():
    return server.core.FetchContext(token="t", user_email="user@example.com",
                                    access_gate="invited",
                                    audit=lambda *a, **k: None)


def test_probe_available(monkeypatch):
    monkeypatch.setattr(server.core.graph, "resolve_online_meeting",
                        lambda *a: {"id": "om1"})
    monkeypatch.setattr(server.core.graph, "list_transcripts",
                        lambda *a: [{"id": "t1", "created": "2026-08-11T12:10:00Z"}])
    m = make_event()
    assert server.core.probe_transcript_status(_ctx_for_probe(), m) == "available"


def test_probe_unavailable(monkeypatch):
    monkeypatch.setattr(server.core.graph, "resolve_online_meeting",
                        lambda *a: {"id": "om1"})
    monkeypatch.setattr(server.core.graph, "list_transcripts", lambda *a: [])
    assert server.core.probe_transcript_status(
        _ctx_for_probe(), make_event()) == "unavailable"


def test_probe_unknown_on_error(monkeypatch):
    def boom(*a):
        raise server.core.graph.AccessDeniedError("403")
    monkeypatch.setattr(server.core.graph, "resolve_online_meeting", boom)
    assert server.core.probe_transcript_status(
        _ctx_for_probe(), make_event()) == "unknown"


# ---------- list payload ----------

def test_list_payload_structure_and_rendering(monkeypatch):
    monkeypatch.setattr(server.core.graph, "list_teams_meetings",
                        lambda *a, **k: [make_event(subject="Townhall"),
                                         make_event(subject="1:1", event_id="e2")])
    monkeypatch.setattr(server.core, "probe_transcript_status",
                        lambda ctx, m: "available" if m["subject"] == "Townhall"
                        else "unavailable")
    _, payload = server.core.list_meetings(_ctx_for_probe(), days=7)
    assert payload["truncated"] is False
    assert payload["range"]["days"] == 7
    first = payload["meetings"][0]
    assert first["id"].startswith("mtg_")
    assert first["transcript_status"] == "available"
    assert first["start"].endswith("Z")           # explicit UTC, not naive
    assert "transcript: available" in payload["rendered"]
    assert "1 of 2 have transcripts" in payload["rendered"]


def test_list_only_with_transcripts_filter(monkeypatch):
    monkeypatch.setattr(server.core.graph, "list_teams_meetings",
                        lambda *a, **k: [make_event(subject="A", event_id="e1"),
                                         make_event(subject="B", event_id="e2")])
    monkeypatch.setattr(server.core, "probe_transcript_status",
                        lambda ctx, m: "unavailable")
    _, payload = server.core.list_meetings(_ctx_for_probe(), days=7,
                                           only_with_transcripts=True)
    assert payload["meetings"] == []
    assert "No meetings in this window have fetchable transcripts" in payload["rendered"]

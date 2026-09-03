"""End-to-end policy tests for the three explicit access gates."""

from __future__ import annotations

import pytest

from transcript_sync import core, graph


def _meeting(response: str = "accepted") -> dict:
    return {
        "event_id": "AAMkGateEvent",
        "subject": "Project call",
        "start": "2026-09-03T09:00:00Z",
        "end": "2026-09-03T10:00:00Z",
        "organizer": "Organiser",
        "organizer_id": "organiser@example.com",
        "is_organizer": False,
        "is_cancelled": False,
        "is_recurring": False,
        "response": response,
        "attendees": ["caller@example.com"],
        "join_url": "https://teams.microsoft.com/l/meetup-join/gate",
    }


def _ctx(access_gate: str) -> core.FetchContext:
    return core.FetchContext(
        token="app-token",
        user_email="caller@example.com",
        access_gate=access_gate,
        audit=lambda *args, **kwargs: None,
        graph_user_id="caller-object-id",
    )


def test_accepted_list_filters_declined_before_transcript_probe(monkeypatch):
    probes = []
    monkeypatch.setattr(graph, "list_teams_meetings", lambda *a, **k: [_meeting("declined")])
    monkeypatch.setattr(
        core,
        "probe_transcript_status",
        lambda *args: probes.append("transcript") or "available",
    )

    meetings, payload = core.list_meetings(_ctx("accepted"), days=7)

    assert meetings == []
    assert payload["meetings"] == []
    assert payload["access_gate"] == "accepted"
    assert probes == []


def test_attended_list_checks_joined_time_before_transcript_metadata(monkeypatch):
    calls = []
    monkeypatch.setattr(graph, "list_teams_meetings", lambda *a, **k: [_meeting("declined")])
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda *a, **k: calls.append("resolve") or {"id": "online-id"},
    )
    monkeypatch.setattr(
        graph,
        "get_my_attendance_seconds",
        lambda *a, **k: calls.append("attendance") or 42,
    )
    monkeypatch.setattr(
        graph,
        "list_transcripts",
        lambda *a, **k: calls.append("transcripts")
        or [{"id": "transcript-id", "created": "2026-09-03T09:15:00Z"}],
    )

    meetings, payload = core.list_meetings(_ctx("attended"), days=7)

    assert len(meetings) == 1
    assert payload["access_gate"] == "attended"
    assert payload["meetings"][0]["transcript_status"] == "available"
    assert calls == ["resolve", "attendance", "transcripts"]


def test_attended_list_hides_non_attendee_before_transcript_metadata(monkeypatch):
    calls = []
    monkeypatch.setattr(graph, "list_teams_meetings", lambda *a, **k: [_meeting()])
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda *a, **k: calls.append("resolve") or {"id": "online-id"},
    )
    monkeypatch.setattr(
        graph,
        "get_my_attendance_seconds",
        lambda *a, **k: calls.append("attendance") or 0,
    )
    monkeypatch.setattr(
        graph,
        "list_transcripts",
        lambda *a, **k: calls.append("transcripts") or [],
    )

    meetings, payload = core.list_meetings(_ctx("attended"), days=7)

    assert meetings == []
    assert payload["meetings"] == []
    assert calls == ["resolve", "attendance"]


def test_accepted_fetch_denies_declined_before_artifact_calls(monkeypatch):
    calls = []
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda *a, **k: calls.append("artifact"),
    )

    result = core.fetch_transcript(_ctx("accepted"), _meeting("declined"))

    assert result["error_code"] == "not_accepted"
    assert calls == []


def test_attended_fetch_denies_zero_joined_time_before_transcript_list(monkeypatch):
    calls = []
    monkeypatch.setattr(graph, "revalidate_event", lambda *a, **k: (True, "listed invitee"))
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda *a, **k: calls.append("resolve") or {"id": "online-id"},
    )
    monkeypatch.setattr(
        graph,
        "get_my_attendance_seconds",
        lambda *a, **k: calls.append("attendance") or 0,
    )
    monkeypatch.setattr(
        graph,
        "list_transcripts",
        lambda *a, **k: calls.append("transcripts") or [],
    )

    result = core.fetch_transcript(_ctx("attended"), _meeting())

    assert result["error_code"] == "not_attended"
    assert calls == ["resolve", "attendance"]


def test_accepted_revalidation_fails_when_response_changed(monkeypatch):
    raw_event = {
        "id": "AAMkGateEvent",
        "attendees": [{"emailAddress": {"address": "caller@example.com"}}],
        "isOrganizer": False,
        "isCancelled": False,
        "responseStatus": {"response": "declined"},
        "organizer": {"emailAddress": {"name": "Organiser", "address": "organiser@example.com"}},
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/gate"},
    }
    monkeypatch.setattr(graph, "_get", lambda *a, **k: raw_event)

    allowed, reason = graph.revalidate_event(
        "app-token",
        _meeting("accepted"),
        "caller@example.com",
        access_gate="accepted",
        user_id="caller-object-id",
    )

    assert allowed is False
    assert reason == "the invitation was not accepted"


def test_attended_gate_uses_only_selected_recurring_occurrence(monkeypatch):
    records_requested = []
    reports = {
        "value": [
            {
                "id": "old-report",
                "meetingStartDateTime": "2026-08-27T09:00:00Z",
                "meetingEndDateTime": "2026-08-27T10:00:00Z",
            },
            {
                "id": "current-report",
                "meetingStartDateTime": "2026-09-03T09:02:00Z",
                "meetingEndDateTime": "2026-09-03T09:58:00Z",
            },
        ]
    }

    def fake_get(token, path, params=None, **kwargs):
        if path.endswith("/attendanceReports"):
            return reports
        records_requested.append(path)
        if "old-report" in path:
            return {"value": [{
                "emailAddress": "caller@example.com",
                "totalAttendanceInSeconds": 600,
            }]}
        return {"value": []}

    monkeypatch.setattr(graph, "_get", fake_get)

    seconds = graph.get_my_attendance_seconds(
        "app-token",
        "online-id",
        "caller@example.com",
        user_id="caller-object-id",
        occurrence_start="2026-09-03T09:00:00Z",
        occurrence_end="2026-09-03T10:00:00Z",
    )

    assert seconds == 0
    assert len(records_requested) == 1
    assert "current-report" in records_requested[0]


def test_attended_gate_fails_closed_without_matching_occurrence_report(monkeypatch):
    monkeypatch.setattr(
        graph,
        "_get",
        lambda *args, **kwargs: {
            "value": [{
                "id": "old-report",
                "meetingStartDateTime": "2026-08-27T09:00:00Z",
                "meetingEndDateTime": "2026-08-27T10:00:00Z",
            }]
        },
    )

    with pytest.raises(graph.AttendanceReportUnavailableError):
        graph.get_my_attendance_seconds(
            "app-token",
            "online-id",
            "caller@example.com",
            occurrence_start="2026-09-03T09:00:00Z",
            occurrence_end="2026-09-03T10:00:00Z",
        )


def test_attended_gate_fails_closed_when_selected_occurrence_has_no_start(monkeypatch):
    monkeypatch.setattr(
        graph,
        "_get",
        lambda *args, **kwargs: {
            "value": [{
                "id": "unscoped-report",
                "meetingStartDateTime": "2026-09-03T09:00:00Z",
                "meetingEndDateTime": "2026-09-03T10:00:00Z",
            }]
        },
    )

    with pytest.raises(graph.AttendanceReportUnavailableError):
        graph.get_my_attendance_seconds(
            "app-token",
            "online-id",
            "caller@example.com",
            occurrence_start=None,
            occurrence_end=None,
        )


def test_attended_gate_rejects_report_starting_hours_after_occurrence(monkeypatch):
    monkeypatch.setattr(
        graph,
        "_get",
        lambda *args, **kwargs: {
            "value": [{
                "id": "later-report",
                "meetingStartDateTime": "2026-09-03T13:00:00Z",
                "meetingEndDateTime": "2026-09-03T14:00:00Z",
            }]
        },
    )

    with pytest.raises(graph.AttendanceReportUnavailableError):
        graph.get_my_attendance_seconds(
            "app-token",
            "online-id",
            "caller@example.com",
            occurrence_start="2026-09-03T09:00:00Z",
            occurrence_end="2026-09-03T10:00:00Z",
        )

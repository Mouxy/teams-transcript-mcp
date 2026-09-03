"""Application-only Graph routing and guard-boundary tests."""

from __future__ import annotations

import base64

from transcript_sync import core, graph


def _calendar_event() -> dict:
    return {
        "id": "AAMkCallerEvent",
        "subject": "All-Hands Huddle",
        "start": {"dateTime": "2026-09-01T09:00:00"},
        "end": {"dateTime": "2026-09-01T10:00:00"},
        "organizer": {
            "emailAddress": {
                "name": "Tenant News",
                "address": "news@example.com",
            }
        },
        "attendees": [
            {"emailAddress": {"address": "caller@example.com"}}
        ],
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/x"},
        "isOrganizer": False,
        "responseStatus": {"response": "notResponded"},
        "type": "singleInstance",
    }


def test_app_calendar_listing_uses_validated_caller_object_id(monkeypatch):
    seen = []

    def fake_get(token, path, params=None, **kwargs):
        seen.append((token, path))
        return {"value": [_calendar_event()]}

    monkeypatch.setattr(graph, "_get", fake_get)

    meetings = graph.list_teams_meetings(
        "app-token", days=7, user_id="caller-object-id"
    )

    assert seen == [("app-token", "/users/caller-object-id/calendarView")]
    assert meetings[0]["organizer_id"] == "news@example.com"


def test_app_calendar_revalidation_uses_validated_caller_object_id(monkeypatch):
    seen = []

    def fake_get(token, path, params=None, **kwargs):
        seen.append((token, path))
        return _calendar_event()

    monkeypatch.setattr(graph, "_get", fake_get)

    allowed, reason = graph.revalidate_event(
        "app-token",
        {
            "event_id": "AAMkCallerEvent",
            "is_organizer": False,
            "attendees": ["caller@example.com"],
        },
        "caller@example.com",
        user_id="caller-object-id",
    )

    assert (allowed, reason) == (True, "listed invitee")
    assert seen == [
        ("app-token", "/users/caller-object-id/events/AAMkCallerEvent")
    ]


def test_caller_supplied_event_id_is_one_encoded_path_segment(monkeypatch):
    seen = []
    monkeypatch.setattr(
        graph,
        "_get",
        lambda token, path, **kwargs: seen.append(path) or _calendar_event(),
    )

    graph.get_event(
        "app-token",
        "AAMk/../../users/another-user",
        user_id="caller-object-id",
    )

    assert seen == [
        "/users/caller-object-id/events/AAMk%2F..%2F..%2Fusers%2Fanother-user"
    ]


def test_forged_path_shaped_handle_is_rejected_before_graph(monkeypatch):
    calls = []
    monkeypatch.setattr(
        graph,
        "get_event",
        lambda *args, **kwargs: calls.append("calendar"),
    )
    forged = "mtg_" + base64.urlsafe_b64encode(
        b"AAMk/../../users/another-user"
    ).decode().rstrip("=")

    target, error, _ = core.resolve_target(_fetch_context(), forged, [])

    assert target is None
    assert error["error_code"] == "invalid_id"
    assert calls == []


def test_revalidation_refreshes_organizer_and_join_url_used_for_artifacts(monkeypatch):
    target = graph.shape_event(_calendar_event())
    target["organizer_id"] = "stale@example.com"
    target["join_url"] = "https://teams.microsoft.com/stale"
    fresh = _calendar_event()
    fresh["organizer"]["emailAddress"]["address"] = "fresh@example.com"
    fresh["onlineMeeting"]["joinUrl"] = "https://teams.microsoft.com/fresh"
    monkeypatch.setattr(graph, "_get", lambda *args, **kwargs: fresh)

    allowed, _ = graph.revalidate_event(
        "app-token",
        target,
        "caller@example.com",
        user_id="caller-object-id",
    )

    assert allowed is True
    assert target["organizer_id"] == "fresh@example.com"
    assert target["join_url"] == "https://teams.microsoft.com/fresh"


def test_app_artifact_paths_use_validated_caller_object_id(monkeypatch):
    seen = []

    def fake_get(token, path, params=None, raw=False, **kwargs):
        seen.append((token, path, raw))
        if path.endswith("/onlineMeetings"):
            return {"value": [{"id": "online-meeting-id"}]}
        if path.endswith("/transcripts"):
            return {"value": [{"id": "transcript-id", "createdDateTime": "now"}]}
        return "WEBVTT"

    monkeypatch.setattr(graph, "_get", fake_get)

    meeting = graph.resolve_online_meeting(
        "app-token",
        "https://teams.microsoft.com/l/meetup-join/x",
        user_id="caller-object-id",
    )
    transcripts = graph.list_transcripts(
        "app-token", meeting["id"], user_id="caller-object-id"
    )
    content = graph.get_transcript_vtt(
        "app-token",
        meeting["id"],
        transcripts[0]["id"],
        user_id="caller-object-id",
    )

    root = "/users/caller-object-id/onlineMeetings/online-meeting-id"
    assert seen == [
        ("app-token", "/users/caller-object-id/onlineMeetings", False),
        ("app-token", f"{root}/transcripts", False),
        ("app-token", f"{root}/transcripts/transcript-id/content", True),
    ]
    assert content == "WEBVTT"


def _fetch_context() -> core.FetchContext:
    return core.FetchContext(
        token="app-token",
        user_email="caller@example.com",
        access_gate="invited",
        audit=lambda *args, **kwargs: None,
        graph_user_id="caller-object-id",
    )


def test_app_core_routes_calendar_and_artifacts_to_validated_caller(monkeypatch):
    target = graph.shape_event(_calendar_event())
    calls = []

    monkeypatch.setattr(
        graph,
        "revalidate_event",
        lambda token, meeting, email, user_id=None, access_gate="invited": (
            calls.append(("revalidate", token, user_id)) or (True, "listed invitee")
        ),
    )
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda token, join_url, user_id=None: (
            calls.append(("resolve", token, user_id)) or {"id": "online-id"}
        ),
    )
    monkeypatch.setattr(
        graph,
        "list_transcripts",
        lambda token, meeting_id, user_id=None: (
            calls.append(("list", token, user_id))
            or [{"id": "transcript-id", "created": "2026-09-01T09:05:00Z"}]
        ),
    )
    monkeypatch.setattr(
        graph,
        "get_transcript_vtt",
        lambda token, meeting_id, transcript_id, user_id=None: (
            calls.append(("content", token, user_id))
            or "WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Speaker>Hello</v>\n"
        ),
    )

    result = core.fetch_transcript(_fetch_context(), target)

    assert result["status"] == "ok"
    assert calls == [
        ("revalidate", "app-token", "caller-object-id"),
        ("resolve", "app-token", "caller-object-id"),
        ("list", "app-token", "caller-object-id"),
        ("content", "app-token", "caller-object-id"),
    ]


def test_app_core_denial_stops_before_any_app_artifact_call(monkeypatch):
    target = graph.shape_event(_calendar_event())
    target["attendees"] = ["someone-else@example.com"]
    calls = []
    monkeypatch.setattr(
        graph,
        "resolve_online_meeting",
        lambda *args, **kwargs: calls.append("artifact"),
    )

    result = core.fetch_transcript(_fetch_context(), target)

    assert result["error_code"] == "not_attendee"
    assert calls == []


def test_list_filters_uninvited_events_before_app_artifact_probe(monkeypatch):
    invited = graph.shape_event(_calendar_event())
    uninvited = graph.shape_event(_calendar_event())
    uninvited["event_id"] = "AAMkForwardedEvent"
    uninvited["attendees"] = ["someone-else@example.com"]
    probes = []

    monkeypatch.setattr(
        graph,
        "list_teams_meetings",
        lambda *args, **kwargs: [invited, uninvited],
    )
    monkeypatch.setattr(
        core,
        "probe_transcript_status",
        lambda ctx, meeting: probes.append(meeting["event_id"]) or "available",
    )

    meetings, payload = core.list_meetings(_fetch_context(), days=7)

    assert [meeting["event_id"] for meeting in meetings] == ["AAMkCallerEvent"]
    assert probes == ["AAMkCallerEvent"]
    assert len(payload["meetings"]) == 1


def test_list_filters_cancelled_event_before_app_artifact_probe(monkeypatch):
    cancelled = _calendar_event()
    cancelled["isCancelled"] = True
    probes = []
    monkeypatch.setattr(
        graph,
        "_get",
        lambda *args, **kwargs: {"value": [cancelled]},
    )
    monkeypatch.setattr(
        core,
        "probe_transcript_status",
        lambda ctx, meeting: probes.append(meeting["event_id"]) or "available",
    )

    meetings, payload = core.list_meetings(_fetch_context(), days=7)

    assert meetings == []
    assert payload["meetings"] == []
    assert probes == []

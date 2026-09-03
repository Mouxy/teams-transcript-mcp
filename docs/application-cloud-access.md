# Application-only cloud transcript access

## Goal

An assigned connector user may retrieve a transcript when the user organised the
meeting or appears in its calendar attendee list. The user's RSVP state does not
matter: `accepted`, `tentative`, `declined` and `notResponded` are all invitations.
Transcript retrieval does not depend on the organiser sharing the transcript with
the caller.

The cloud connector authenticates the human and Microsoft Graph authorises the
backend. Deterministic server code joins those two identities and enforces the
per-user rule before any transcript request.

## Authentication and Graph context

The incoming Claude connector token must contain:

- the expected tenant-specific issuer;
- the API application's exact client-ID audience;
- a valid signature and lifetime;
- the exposed delegated `access_as_user` scope;
- an Entra object ID (`oid`) formatted as a GUID.

`access_as_user` authenticates the caller to this API. It is not used for a
Microsoft Graph OBO exchange.

The backend uses its certificate to acquire a Microsoft Graph client-credentials
token for `https://graph.microsoft.com/.default`. All Graph calls use that
application token.

## Required Microsoft Graph application permissions

Grant tenant admin consent for:

- `Calendars.Read` — read the authenticated caller's calendar through
  `/users/{validated-caller-oid}/...`.
- `OnlineMeetings.Read.All` — resolve an online meeting through the validated
  caller's user context. Microsoft supports an invited attendee's user ID for
  this app-only path.[2]
- `OnlineMeetingTranscript.Read.All` — list and retrieve transcript metadata and
  content.[1]
- `OnlineMeetingArtifact.Read.All` — retrieve attendance reports and records
  when the `attended` access gate is enabled.[3]

`OnlineMeetingRecording.Read.All` is not required for transcript retrieval.
Neither `Files.Read.All`, `Sites.Read.All` nor `Group.Read.All` is required for
this direct Teams artifact path.

The local stdio transport remains separate. It uses delegated Graph permissions
and `/me` paths because it runs directly as its interactive user.

## Teams application access policy

Microsoft requires a Teams application access policy for app-only online-meeting,
transcript and attendance operations.[1][2][3] Microsoft documents tenant-wide
and per-user organiser assignments separately.[4]

To cover meetings organised by any user in the tenant:

```powershell
Connect-MicrosoftTeams

New-CsApplicationAccessPolicy `
  -Identity "TranscriptSync-AppOnly" `
  -AppIds "<TRANSCRIPT_SYNC_CLOUD_CLIENT_ID>" `
  -Description "Allow Transcript Sync application access after its caller guard passes"

Grant-CsApplicationAccessPolicy `
  -PolicyName "TranscriptSync-AppOnly" `
  -Global
```

The policy authorises the application to use the user ID supplied in the Graph
path. This implementation always supplies the validated connector caller's
`oid`; a global policy covers every assigned caller. Enterprise Application
assignment continues to restrict who can authenticate to the connector.[2][4]

The tenant must also have Microsoft Graph transcript access enabled. Speaker
attribution must be enabled when the service requests speaker-labelled
WebVTT.[5]

## Request flow

1. Validate the connector bearer token and extract its immutable `oid` and user
   address.
2. Acquire an application Graph token. Never expose it to the MCP client.
3. List events only from `/users/{validated-caller-oid}/calendarView`.
4. Retain an event only when the caller is its organiser or their address appears
   in the event attendee list. Ignore RSVP response status.
5. Apply that guard before probing transcript availability.
6. Return an opaque `mtg_` handle derived from the caller's calendar event.
7. On fetch, resolve the handle only through
   `/users/{validated-caller-oid}/events/{event-id}` and revalidate that the event
   still exists, is not cancelled, and still lists the caller.
8. Resolve and retrieve meeting artifacts through
   `/users/{validated-caller-oid}/onlineMeetings/...` with the application
   token. Microsoft permits the invited attendee's user ID for app-only online
   meeting retrieval.[2]

## Guard invariants

- Tools never accept caller-supplied mailbox IDs, join URLs, online-meeting IDs,
  transcript IDs or organiser IDs.
- The calendar mailbox comes only from the validated bearer token's `oid`.
- Transcript probing and retrieval occur only after the configured access gate
  passes. In `attended` mode, the attendance report is checked first.
- Deleted, cancelled and no-longer-listed events fail closed.
- Meeting-list caches are keyed by caller `oid` and are not shared between users.
- Every fetch revalidates the calendar event, even when the caller uses a cached
  list position or opaque handle.
- Transcript text remains untrusted, bounded and excluded from logs.
- Application tokens, connector bearer tokens and certificate material never
  appear in tool output or audit records.

## Access gates

Set one canonical `TRANSCRIPT_SYNC_ACCESS_GATE` value:

- `invited` (default): organiser or attendee listed on the calendar event,
  regardless of RSVP state;
- `accepted`: organiser or listed attendee whose current RSVP is `accepted`;
- `attended`: organiser or listed attendee with positive joined time in an
  app-only attendance report. The attendance proof runs before transcript
  metadata or content is returned. Recurring-series reports are matched to the
  selected calendar occurrence by requiring the report's actual meeting start
  to be within 30 minutes of the scheduled start. Missing timestamps fail closed.

All gates deny deleted, cancelled and forwarded events where the caller is
absent from the attendee list. Attendance reports do not support channel
meetings, so `attended` fails closed for those meetings.[3]

A user who attended through a link but has no calendar event is not discoverable
through `list_recent_meetings`. Supporting attendance-only discovery would need a
separate server-owned meeting index and is outside the current tool contract.

## Deployment verification

Before widening use:

1. Verify the app registration requests exactly the four required Graph
   application roles.
2. Verify service-principal app-role assignments after admin consent.
3. Read back Enterprise Application assignment restriction and user/group
   assignments.
4. Read back the Teams application access policy.
5. Confirm tenant Graph transcript access and speaker attribution settings.
6. Run an allowed `notResponded` invitation and retrieve content without exposing
   transcript text in diagnostics.
7. Run a never-invited control and prove zero meeting-artifact calls occur.
8. Run deleted, cancelled, missing-organiser and malformed-handle controls.
9. Verify audit records contain decisions but no tokens or transcript content.

## Sources

[1] https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-transcripts?view=graph-rest-1.0 — List transcripts - Microsoft Graph
[2] https://learn.microsoft.com/en-us/graph/api/onlinemeeting-get?view=graph-rest-1.0 — Get onlineMeeting - Microsoft Graph
[3] https://learn.microsoft.com/en-us/graph/api/meetingattendancereport-get?view=graph-rest-1.0 — Get meetingAttendanceReport - Microsoft Graph
[4] https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy — Configure application access to online meetings
[5] https://learn.microsoft.com/en-us/microsoftteams/meeting-transcript-api-access — Manage transcript API access for Teams meetings

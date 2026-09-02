# Hybrid cloud transcript access

**Status:** Required design for sharing-independent cloud retrieval. It is **not
implemented** by the current release. The current cloud transport uses delegated
OBO for both eligibility and transcript retrieval.

## Goal

An assigned connector user may retrieve a transcript when the user was invited
to the meeting or has a positive attendance record. Transcript retrieval must
not depend on the organiser manually sharing the transcript with that user.

The caller-facing policy and the backend retrieval capability are separate:

- delegated user context establishes who is calling and what appears on that
  user's calendar;
- app-only context retrieves online-meeting artifacts independently of the
  caller's native Teams transcript entitlement;
- server-side policy must authorise the caller before any app-only transcript
  request is made.

## Current release limitation

The current cloud implementation exchanges the connector token through OBO and
uses that delegated Graph token for calendar, online-meeting and transcript
calls. This means Microsoft can list transcript metadata yet deny transcript
content when the signed-in user lacks native content entitlement.

A certificate on the Entra app does not change that behaviour by itself. The
runtime must acquire a client-credentials token and use the token's `roles`
claim for app-only Graph requests.

## Required permissions

### Delegated permissions

Keep:

- `Calendars.Read` — read and revalidate the signed-in caller's own calendar.
- The connector's exposed `access_as_user` scope — authenticate Claude to this
  API. This is not a Microsoft Graph permission.

The existing delegated `OnlineMeetings.Read`,
`OnlineMeetingTranscript.Read.All`, and `OnlineMeetingArtifact.Read.All` scopes
may remain while the current delegated and local transports still use them.
They are not the authority for sharing-independent cloud transcript retrieval.

### Microsoft Graph application permissions

Add and grant tenant admin consent for:

- `OnlineMeetings.Read.All` — resolve an online meeting in the organiser's user
  context.[2]
- `OnlineMeetingTranscript.Read.All` — list and retrieve transcript metadata and
  content.[1]
- `OnlineMeetingArtifact.Read.All` — retrieve attendance reports and records for
  attendance-based eligibility.[3]

The direct online-meeting transcript path does not require
`Calendars.Read.All`, `Files.Read.All`, `Sites.Read.All`, or `Group.Read.All`.
Those permissions belong to different architectures, such as a SharePoint
collector, and must not be added for this mode without a separate requirement.

## Teams application access policy

Application permissions alone are insufficient. Microsoft requires a Teams
application access policy for app-only online-meeting, transcript and attendance
operations.[1][2][3] Microsoft documents the policy creation and tenant-wide or
per-user assignment flow separately.[4]

To cover meetings organised by any user in the tenant:

```powershell
Connect-MicrosoftTeams

New-CsApplicationAccessPolicy `
  -Identity "TranscriptSync-AppOnly" `
  -AppIds "<TRANSCRIPT_SYNC_CLOUD_CLIENT_ID>" `
  -Description "Allow Transcript Sync to retrieve meeting artifacts after its caller policy passes"

Grant-CsApplicationAccessPolicy `
  -PolicyName "TranscriptSync-AppOnly" `
  -Global
```

A scoped rollout can grant the same policy to individual organiser object IDs
instead of `-Global`. The policy applies to the organiser whose user ID appears
in `/users/{organiser-id}/onlineMeetings/...`; it is not a caller assignment
policy.[4]

The tenant must also have Microsoft Graph transcript access enabled. Speaker
attribution must be enabled when the service requests WebVTT with speaker
labels.[5]

## Runtime token separation

The cloud service needs two Graph token paths:

1. **Delegated token**
   - acquired through OBO;
   - reads `/me/calendarView` and `/me/events/{id}`;
   - proves caller identity and calendar invitation state.

2. **App-only token**
   - acquired with the app certificate through client credentials;
   - resolves the meeting through
     `/users/{organiser-id}/onlineMeetings/...`;
   - lists and retrieves transcripts through the organiser-scoped path;
   - reads attendance reports when attendance evidence is required.

The app-only token must never be returned to the MCP client or accepted as a
caller credential.

## Target authorisation rule

Before using the app-only transcript path, the service must establish at least
one of:

- **Invited:** the meeting is derived from the caller's own calendar and the
  caller is the organiser or an attendee on the event.
- **Attended:** an app-only attendance report contains a matching caller object
  ID or verified address with positive attendance duration.

Deleted or cancelled meetings remain denied. Exact treatment of declined
invitations must be explicit in tests and operator documentation.

The existing `list_recent_meetings` tool discovers calendar-backed meetings. A
user who attended through a link but has no corresponding calendar event needs
a separate discovery design before the attendance-only branch can expose that
meeting. Attendance reports also do not support channel meetings.[3]

## Existing controls that remain required

- Entra Enterprise Application assignment remains required for connector users.
- Caller bearer tokens still require tenant, issuer, signature, audience,
  expiry, object ID, and `access_as_user` validation.
- Meeting handles remain calendar-derived and opaque.
- Events are revalidated before retrieval.
- Transcript content remains untrusted, bounded, and excluded from logs.
- Every allow and deny decision is audited without transcript text or tokens.

Entra assignment restricts who can call the connector. It does not reduce the
backend app-only token's Graph capability; the Teams application access policy
and server-side eligibility gate provide those boundaries.

## Implementation and rollout checklist

Do not grant the application roles and claim completion until a release supports
this flow.

1. Extend the Entra bootstrap to request the three application permissions by
   immutable Graph app-role ID and verify their service-principal assignments.
2. Add a certificate-based client-credentials token provider separate from OBO.
3. Preserve organiser address or object ID in the calendar-derived meeting
   record.
4. Split Graph calls so delegated context handles eligibility and app-only
   context handles meeting artifacts.
5. Add tests proving no app-only artifact call occurs before eligibility passes.
6. Add invited, attended, declined, cancelled, foreign-ID, channel-meeting and
   organiser-not-covered cases.
7. Configure and read back the Teams application access policy.
8. Pilot with one allowed meeting and one denied control meeting without
   returning transcript contents in diagnostic output.
9. Verify audit records, then widen organiser policy coverage deliberately.

## Sources

[1] https://learn.microsoft.com/en-us/graph/api/onlinemeeting-list-transcripts?view=graph-rest-1.0 — List transcripts - Microsoft Graph
[2] https://learn.microsoft.com/en-us/graph/api/onlinemeeting-get?view=graph-rest-1.0 — Get onlineMeeting - Microsoft Graph
[3] https://learn.microsoft.com/en-us/graph/api/meetingattendancereport-get?view=graph-rest-1.0 — Get meetingAttendanceReport - Microsoft Graph
[4] https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy — Configure application access to online meetings
[5] https://learn.microsoft.com/en-us/microsoftteams/meeting-transcript-api-access — Manage transcript API access for Teams meetings

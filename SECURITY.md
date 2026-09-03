# Security Policy

## Reporting a vulnerability

Please do not open a public issue for a suspected vulnerability involving
authentication, authorisation, transcript exposure or secret handling.

Use GitHub's private vulnerability reporting feature for this repository. If
that feature is unavailable, contact the repository owner through their GitHub
profile and ask for a private reporting channel. Do not include access tokens,
private keys, client secrets, transcript content or tenant data in the first
message.

## Supported versions

Only the latest revision on the `main` branch receives security fixes until the
project begins publishing versioned releases.

## Deployment responsibility

Operators are responsible for Entra consent, user assignment, Azure region,
log retention, MCP-host data handling and local secret storage. Never report a
live credential in an issue, log excerpt or test fixture.

The cloud release uses application-only Graph access so transcript retrieval does
not depend on the caller's native Teams sharing rights. In this mode:

- Entra connector assignment limits who can authenticate but does not limit the
  backend application token.
- The authenticated caller's validated `oid` is the only mailbox identifier used
  for calendar calls. The MCP client cannot supply a mailbox, organiser, join URL
  or Graph meeting ID.
- The server applies exactly one configured access gate before transcript
  availability probes and final artifact requests: `invited`, `accepted` or
  `attended`. The default `invited` gate ignores RSVP status. `accepted`
  requires an accepted RSVP. `attended` requires positive joined time from an
  app-only attendance report before transcript metadata is exposed.
- A Teams application access policy authorises the app-only user context used
  for meeting access. This implementation always uses the validated caller
  `oid`.
- Tests and logs must never include transcript content, bearer tokens,
  certificate private material or raw Graph response bodies.

See [Application-only cloud transcript access](docs/application-cloud-access.md)
for the complete permission and guard model.

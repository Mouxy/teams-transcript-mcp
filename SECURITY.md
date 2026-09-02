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

The current release uses delegated Graph access. The planned hybrid cloud mode
adds application permissions that let the backend retrieve meeting artifacts
without relying on the caller's native transcript sharing rights. In that mode:

- Entra connector assignment limits callers but does not limit the backend
  app-only token.
- A Teams application access policy limits which organisers' meetings the
  backend can access.
- The server's calendar/invitation/attendance gate is an authorisation boundary
  and must run before any app-only artifact request.
- Tests and logs must never include transcript content, bearer tokens,
  certificate private material or raw Graph response bodies.

See [Hybrid cloud transcript access](docs/hybrid-cloud-access.md) before granting
application roles or a global Teams application access policy.

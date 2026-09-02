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

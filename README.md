# Teams Transcript MCP

A Microsoft Model Context Protocol (MCP) server for listing a signed-in user's
Microsoft Teams meetings and fetching native Teams transcripts through
Microsoft Graph.

It supports two transports:

- **Cloud:** streamable HTTP, Microsoft Entra OAuth and an on-behalf-of (OBO)
  exchange. This is suitable for a managed Claude custom connector.
- **Local:** stdio, interactive Microsoft sign-in and an OS-keyring token cache.

This project is not affiliated with Microsoft or Anthropic.

## What it does

The server exposes two tools:

- `list_recent_meetings(days, only_with_transcripts)` lists recent online
  meetings from the user's own calendar. Each result has a stable `mtg_` ID
  and an `available`, `unavailable` or `unknown` transcript status.
- `get_transcript(meeting, raw_vtt)` fetches the transcript for a listed
  meeting. It returns speaker-attributed text by default or raw WebVTT.

The server never accepts a Teams join URL or Graph online-meeting ID directly.
It derives meeting identity from the signed-in user's calendar, revalidates the
calendar event before each fetch, and records every attempt in an audit log.

## Architecture

```text
Cloud MCP client ──HTTPS──> Transcript Sync
                             │ validates Entra JWT
                             │ exchanges token through OBO
                             ▼
                        Microsoft Graph
                 delegated access as the signed-in user
```

The cloud process keeps caches per Entra object ID. It does not share meeting
state between users. Scale-to-zero cold starts only clear those caches.

## Security model

- Cloud requests validate the Entra issuer, exact API client-ID audience, expiry,
  object ID and delegated `access_as_user` scope before OBO exchange.
- A caller must be the meeting organiser or an accepted attendee.
- The server resolves opaque meeting handles only through `/me/events/{id}`.
- It rechecks deleted, cancelled, declined and changed events before fetching.
- Transcript text is returned between untrusted-content delimiters and capped
  at 200,000 characters. The MCP host must still treat it as untrusted data.
- Cloud audit records go to stdout. Local audit records go to
  `~/.transcript-sync/audit.log` with user-only permissions.
- `TRANSCRIPT_SYNC_ATTENDANCE_MODE=invite` is the default. `strict` additionally
  requests an attendance report, but Microsoft Graph normally restricts that
  delegated endpoint to organisers and co-organisers.

## Prerequisites

For either mode:

- Python 3.11 or newer
- [`uv`](https://docs.astral.sh/uv/)
- A Microsoft 365 tenant with Teams transcription enabled
- An Entra administrator who can create app registrations and grant consent

For cloud mode:

- An Azure subscription
- Azure CLI with the Container Apps extension
- Contributor access to the target subscription, or Contributor on an existing
  resource group after a subscription administrator has registered the required
  resource providers
- A public HTTPS endpoint
- A Claude plan that supports custom connectors, if Claude is the MCP client

## Required Microsoft Graph permissions

Add these as **delegated** permissions to the Entra app:

- `Calendars.Read`
- `OnlineMeetings.Read`
- `OnlineMeetingTranscript.Read.All`
- `OnlineMeetingArtifact.Read.All`

The last permission supports strict attendance checks. The default invite mode
still requests the same fixed scope set so switching modes does not silently
change consent.

## Cloud setup

### 1. Create the Entra app registration

In the Microsoft Entra admin centre:

1. Open **Identity → Applications → App registrations → New registration**.
2. Name it `Transcript Sync Cloud` or choose your own name.
3. Select **Accounts in this organisational directory only**.
4. Leave the redirect URI empty for now and create the registration.
5. Record the **Application (client) ID** and **Directory (tenant) ID**.
6. Under **Authentication**, add these **Web** redirect URIs:
   - `https://claude.ai/api/mcp/auth_callback`
   - `https://claude.com/api/mcp/auth_callback`
7. Under **API permissions**, add the four delegated Microsoft Graph
   permissions listed above.
8. Under **Expose an API**, set an initial Application ID URI of
   `api://<client-id>` and add a delegated scope named `access_as_user`.
9. Under **Certificates & secrets**, create a client secret for the MCP
   connector. Copy it immediately and store it in a secret manager.

Do not commit tenant IDs, client secrets, private certificates or `.env` files.
A tenant ID and client ID are identifiers rather than passwords, but keeping
instance-specific values out of the repository makes the deployment portable.

### 2. Create the OBO certificate

Generate a private key and self-signed certificate locally:

```bash
install -d -m 700 "$HOME/.transcript-sync"
openssl req -x509 -newkey rsa:2048 -sha256 -nodes -days 730 \
  -subj "/CN=Transcript Sync Cloud" \
  -keyout "$HOME/.transcript-sync/cloud-key.pem" \
  -out "$HOME/.transcript-sync/cloud-cert.crt"
cat "$HOME/.transcript-sync/cloud-key.pem" \
    "$HOME/.transcript-sync/cloud-cert.crt" \
  > "$HOME/.transcript-sync/cloud-cert.pem"
chmod 600 "$HOME/.transcript-sync/"*
```

Upload `cloud-cert.crt`, which contains only the public certificate, under the
app registration's **Certificates & secrets → Certificates** page. Keep
`cloud-key.pem` and `cloud-cert.pem` private. The deployment script uploads the
combined PEM to Container Apps as a secret.

### 3. Grant consent and restrict access

1. Open **Enterprise applications → Transcript Sync Cloud → Permissions**.
2. Select **Grant admin consent** and accept the four delegated Graph scopes.
3. Read the permissions page back and confirm every permission shows
   **Granted for** your tenant. Clicking the consent button alone is not proof.
4. Open **Properties**, set **Assignment required?** to **Yes**, and save.
5. Assign only the intended pilot users or group under **Users and groups**.

Use the Entra portal for consent. Direct `adminconsent` URLs are prone to
browser-session failures and are not part of this setup.

### 4. Deploy to Azure Container Apps

```bash
uv sync --group dev

export SUBSCRIPTION='<subscription name or ID>'
export TENANT_ID='<directory tenant ID>'
export CLOUD_CLIENT_ID='<application client ID>'
export CERT_PEM="$HOME/.transcript-sync/cloud-cert.pem"

# Optional names and region:
export LOCATION='uksouth'
export RG='rg-transcript-sync'
export ENV_NAME='cae-transcript-sync'
export APP_NAME='transcript-sync'

scripts/deploy_azure.sh
```

The script registers required Azure providers, creates an explicitly located
Container Apps environment and Log Analytics workspace, builds the image,
stores the OBO certificate as a Container Apps secret, and deploys with zero
minimum replicas. It stops if the existing environment is in a different
region.

Record the public origin printed by the script, for example:
`https://transcript-sync.example-region.azurecontainerapps.io`.

### 5. Finalise the Entra resource URI

Claude derives the OAuth resource from the MCP server URL. Microsoft Entra
requires that resource to match the resource part of the requested scope.
After deployment:

1. Return to **App registrations → Transcript Sync Cloud → Expose an API**.
2. Replace the Application ID URI with the public HTTPS origin, without `/mcp`.
3. Confirm the exposed scope is now
   `https://<public-origin>/access_as_user`.

The server must receive matching values:

```text
TRANSCRIPT_SYNC_TENANT_ID=<tenant ID>
TRANSCRIPT_SYNC_CLOUD_CLIENT_ID=<client ID>
TRANSCRIPT_SYNC_SERVER_URL=https://<public-origin>
TRANSCRIPT_SYNC_CLOUD_CERT_PEM=<combined PEM content or file path>
TRANSCRIPT_SYNC_APP_NAME=Transcript Sync Cloud
```

`deploy_azure.sh` sets these values. Its `secretref:` value resolves to the PEM
content inside the container, not to a filesystem path.

### 6. Add the Claude custom connector

In the Claude organisation settings, add a custom connector with:

- **Server URL:** `https://<public-origin>/mcp`
- **Client ID:** the Entra Application (client) ID
- **Client secret:** the secret created for the connector

OAuth discovery is available at:

```text
https://<public-origin>/.well-known/oauth-protected-resource
```

It advertises the tenant authorisation server and the
`https://<public-origin>/access_as_user` scope. Each assigned user must complete
an individual Microsoft sign-in.

### Optional automated Entra bootstrap

The scripts can create and converge the app if you already have a
certificate-authenticated Graph management app with
`Application.ReadWrite.All`:

```bash
uv run python scripts/create_cloud_app.py \
  --tenant '<tenant ID>' \
  --caller-client-id '<management app client ID>' \
  --caller-pem '/secure/path/management-app.pem' \
  --create-client-secret

# Deploy, then converge the resource URI without creating another secret:
uv run python scripts/create_cloud_app.py \
  --tenant '<tenant ID>' \
  --app-client-id '<application client ID from the first run>' \
  --caller-client-id '<management app client ID>' \
  --caller-pem '/secure/path/management-app.pem' \
  --server-url 'https://<public-origin>'
```

The first command prints the connector secret once. Store it securely. The
second command registers both Claude callbacks and changes the identifier URI
and exposed scope to the deployed server origin. The scripts never mutate an
existing app from a display-name match alone. A follow-up run must identify the
app by its immutable client ID, and the cloud script verifies that the local OBO
certificate matches a key credential registered on that app.

## Local stdio setup

### 1. Create a local public-client app

In **App registrations**:

1. Create a single-tenant app named `Transcript Sync Local`.
2. Under **Authentication**, add `http://localhost` as a **Mobile and desktop
   applications** redirect URI.
3. Enable public client flows.
4. Add the four delegated Graph permissions and grant admin consent through
   **Enterprise applications → Transcript Sync Local → Permissions**.
5. Record the tenant ID and client ID.

The optional automation script requires the same pre-existing Graph management
app described above:

```bash
uv run python scripts/create_entra_app.py \
  --tenant '<tenant ID>' \
  --caller-client-id '<management app client ID>' \
  --caller-pem '/secure/path/management-app.pem'
```

### 2. Configure the MCP client

Use an absolute path to `uv`; desktop applications often have a restricted
`PATH`.

```json
{
  "mcpServers": {
    "teams-transcripts": {
      "command": "/absolute/path/to/uv",
      "args": [
        "--directory",
        "/absolute/path/to/teams-transcript-mcp",
        "run",
        "transcript-sync"
      ],
      "env": {
        "TRANSCRIPT_SYNC_TENANT_ID": "<tenant ID>",
        "TRANSCRIPT_SYNC_CLIENT_ID": "<client ID>",
        "TRANSCRIPT_SYNC_ATTENDANCE_MODE": "invite"
      }
    }
  }
}
```

The first `sign_in` call opens the browser and forces account selection. The
refresh-token cache is stored through the operating system keyring. `sign_out`
removes only the local cache; it does not revoke Microsoft sessions.

## Verification

Run the local checks:

```bash
uv sync --group dev
uv run pytest tests/ -q
uvx ruff check src scripts tests
bash -n scripts/deploy_azure.sh
```

For a cloud deployment:

1. Fetch `/.well-known/oauth-protected-resource` and confirm the resource and
   scope use the exact public origin.
2. Call `/mcp` without a bearer token and confirm it returns `401` with a
   `WWW-Authenticate` resource-metadata link.
3. Sign in as an assigned pilot user.
4. Run `list_recent_meetings` and fetch one known transcript.
5. Run a negative test with a user who is not assigned or an event outside the
   caller's calendar.
6. Confirm the audit entry appears in Container Apps logs.

The automated test suite mocks Microsoft Graph and makes no live tenant calls.

## Data boundary and privacy

A transcript returned by this server leaves Microsoft 365 and enters the MCP
host's processing and retention boundary. Before production use, assess the
host's DPA, retention, region, model-training and access-control terms.

Audit records can contain user addresses, meeting subjects, meeting handles and
occurrence times. Treat them as sensitive business data, set an appropriate Log
Analytics retention period, restrict log access, and avoid logging transcript
content. This project does not log transcript bodies.

## Rollback

- Remove or disable the MCP connector in the host.
- Disable or delete the Azure Container App.
- Remove user/group assignments from the Enterprise application.
- Delete the client secret or the whole Entra app registration.
- Delete the Azure resource group if it is dedicated to this service.
- Review audit-retention obligations before deleting Log Analytics data or
  `~/.transcript-sync/audit.log`.

## Development

```bash
uv sync --group dev
uv run pytest tests/ -q
uvx ruff check src scripts tests
```

The core Graph and policy logic lives in `src/transcript_sync/core.py` and is
shared by both transports. `mcp>=1.10,<2` is intentional because this code uses
the MCP 1.x `mcp.server.fastmcp` import path.

## Licence

MIT. See [LICENSE](LICENSE).

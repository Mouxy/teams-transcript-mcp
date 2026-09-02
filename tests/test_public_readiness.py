"""Public-release guardrails for tenant-neutral defaults and documentation."""

from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PUBLIC_FILES = [
    ROOT / "README.md",
    ROOT / "pyproject.toml",
    *sorted((ROOT / "scripts").glob("*.py")),
    *sorted((ROOT / "scripts").glob("*.sh")),
    *sorted((ROOT / "src").rglob("*.py")),
    *[
        path
        for path in sorted((ROOT / "tests").glob("*.py"))
        if path.name != "test_public_readiness.py"
    ],
]

def test_runtime_tenant_configuration_has_no_uuid_default():
    source = "\n".join(path.read_text() for path in PUBLIC_FILES)
    tenant_default = re.compile(
        r'TRANSCRIPT_SYNC_TENANT_ID["\']\s*,\s*["\']'
        r'[0-9a-fA-F]{8}-[0-9a-fA-F-]{27,}'
    )
    assert tenant_default.search(source) is None


def test_deployment_requires_subscription_and_tenant():
    deploy = (ROOT / "scripts" / "deploy_azure.sh").read_text()
    assert '${SUBSCRIPTION:?' in deploy
    assert '${TENANT_ID:?' in deploy
    assert "az provider show" in deploy
    assert "registrationState" in deploy


def test_deployment_uploads_certificate_content_not_a_local_path():
    deploy = (ROOT / "scripts" / "deploy_azure.sh").read_text()
    assert 'cloud-cert-pem=@"$CERT_PEM"' not in deploy
    assert 'CERT_PEM_CONTENT=' in deploy
    assert 'cloud-cert-pem=$CERT_PEM_CONTENT' in deploy
    assert '--app-client-id $CLOUD_CLIENT_ID' in deploy


def test_docker_context_excludes_local_secrets_and_state():
    patterns = set((ROOT / ".dockerignore").read_text().splitlines())
    assert {".git", ".env", "*.pem", "*.key", "*.crt", ".venv", ".transcript-sync", "audit.log"} <= patterns


def test_local_auth_requires_explicit_tenant_configuration():
    env = os.environ.copy()
    env.pop("TRANSCRIPT_SYNC_TENANT_ID", None)
    env.pop("TRANSCRIPT_SYNC_CLIENT_ID", None)
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            (
                "from transcript_sync import auth; "
                "auth._app()"
            ),
        ],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode != 0
    output = result.stdout + result.stderr
    assert "TRANSCRIPT_SYNC_TENANT_ID" in output
    assert "TRANSCRIPT_SYNC_CLIENT_ID" in output


def test_setup_documentation_covers_required_public_configuration():
    readme = (ROOT / "README.md").read_text().casefold()
    required = (
        "tenant id",
        "client id",
        "enterprise applications",
        "grant admin consent",
        "claude.ai/api/mcp/auth_callback",
        "claude.com/api/mcp/auth_callback",
        "assignment required",
        "transcript_sync_server_url",
        "onlineMeetingTranscript.Read.All".casefold(),
        "data boundary",
    )
    missing = [item for item in required if item not in readme]
    assert missing == []

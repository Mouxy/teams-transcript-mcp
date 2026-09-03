#!/usr/bin/env bash
# Deploy Transcript Sync to Azure Container Apps with scale-to-zero.
# Prerequisites: Azure CLI login with subscription Contributor rights and an
# Entra app created by scripts/create_cloud_app.py.
set -euo pipefail

: "${SUBSCRIPTION:?Set SUBSCRIPTION to an Azure subscription name or ID}"
: "${CLOUD_CLIENT_ID:?Set CLOUD_CLIENT_ID from create_cloud_app.py output}"
: "${TENANT_ID:?Set TENANT_ID to the Microsoft Entra tenant ID}"

RG="${RG:-rg-transcript-sync}"
LOCATION="${LOCATION:-uksouth}"
ENV_NAME="${ENV_NAME:-cae-transcript-sync}"
APP_NAME="${APP_NAME:-transcript-sync}"
LAW_NAME="${LAW_NAME:-law-transcript-sync}"
DISPLAY_NAME="${DISPLAY_NAME:-Transcript Sync Cloud}"
ACCESS_GATE="${ACCESS_GATE:-invited}"
CERT_PEM="${CERT_PEM:-$HOME/.transcript-sync/cloud-cert.pem}"
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"

case "$ACCESS_GATE" in
  invited|accepted|attended) ;;
  *) echo "ACCESS_GATE must be 'invited', 'accepted' or 'attended'." >&2; exit 2 ;;
esac

if [[ ! -f "$CERT_PEM" ]]; then
  printf 'Certificate not found: %s\n' "$CERT_PEM" >&2
  exit 1
fi

az account set --subscription "$SUBSCRIPTION"

# Fresh subscriptions need these providers. Registration is idempotent.
for namespace in \
  Microsoft.App \
  Microsoft.ContainerRegistry \
  Microsoft.OperationalInsights \
  Microsoft.Insights
do
  PROVIDER_STATE=$(az provider show \
    --namespace "$namespace" --query registrationState -o tsv)
  if [[ "$PROVIDER_STATE" != "Registered" ]]; then
    if ! az provider register --namespace "$namespace" --wait -o none; then
      printf 'Provider %s is not registered. Subscription-level permission is required to register it.\n' \
        "$namespace" >&2
      exit 1
    fi
  fi
done

if ! az group show --name "$RG" -o none 2>/dev/null; then
  az group create --name "$RG" --location "$LOCATION" -o none
fi

LAW_ID=$(az monitor log-analytics workspace show \
  --resource-group "$RG" --workspace-name "$LAW_NAME" \
  --query id -o tsv 2>/dev/null || true)
if [[ -z "$LAW_ID" ]]; then
  az monitor log-analytics workspace create \
    --resource-group "$RG" --workspace-name "$LAW_NAME" \
    --location "$LOCATION" -o none
  LAW_ID=$(az monitor log-analytics workspace show \
    --resource-group "$RG" --workspace-name "$LAW_NAME" --query id -o tsv)
fi
printf 'Log Analytics: %s\n' "$LAW_ID"

# Container Apps expects the workspace customer ID and shared key, not its ARM ID.
LAW_GUID=$(az monitor log-analytics workspace show \
  --resource-group "$RG" --workspace-name "$LAW_NAME" \
  --query customerId -o tsv)
LAW_KEY=$(az monitor log-analytics workspace get-shared-keys \
  --resource-group "$RG" --workspace-name "$LAW_NAME" \
  --query primarySharedKey -o tsv)

if ! az containerapp env show \
  --name "$ENV_NAME" --resource-group "$RG" -o none 2>/dev/null
then
  az containerapp env create \
    --name "$ENV_NAME" --resource-group "$RG" --location "$LOCATION" \
    --logs-destination log-analytics \
    --logs-workspace-id "$LAW_GUID" --logs-workspace-key "$LAW_KEY" -o none
fi

ACTUAL_LOCATION=$(az containerapp env show \
  --name "$ENV_NAME" --resource-group "$RG" --query location -o tsv)
NORMALISED_EXPECTED=$(printf '%s' "$LOCATION" | tr '[:upper:]' '[:lower:]' | tr -d ' ')
NORMALISED_ACTUAL=$(printf '%s' "$ACTUAL_LOCATION" | tr '[:upper:]' '[:lower:]' | tr -d ' ')
if [[ "$NORMALISED_ACTUAL" != "$NORMALISED_EXPECTED" ]]; then
  printf 'Refusing deployment: Container Apps environment is in %s, expected %s.\n' \
    "$ACTUAL_LOCATION" "$LOCATION" >&2
  exit 1
fi

# Build from source and deploy into the explicitly created environment.
az containerapp up \
  --name "$APP_NAME" --resource-group "$RG" --environment "$ENV_NAME" \
  --source "$PROJECT_DIR" --ingress external --target-port 8000

CERT_PEM_CONTENT=$(<"$CERT_PEM")
az containerapp secret set \
  --name "$APP_NAME" --resource-group "$RG" \
  --secrets "cloud-cert-pem=$CERT_PEM_CONTENT" -o none
unset CERT_PEM_CONTENT

FQDN=$(az containerapp show \
  --name "$APP_NAME" --resource-group "$RG" \
  --query properties.configuration.ingress.fqdn -o tsv)
SERVER_URL="https://$FQDN"

az containerapp update \
  --name "$APP_NAME" --resource-group "$RG" \
  --min-replicas 0 --max-replicas 2 \
  --remove-env-vars TRANSCRIPT_SYNC_ATTENDANCE_MODE \
  --set-env-vars \
    "TRANSCRIPT_SYNC_TENANT_ID=$TENANT_ID" \
    "TRANSCRIPT_SYNC_CLOUD_CLIENT_ID=$CLOUD_CLIENT_ID" \
    "TRANSCRIPT_SYNC_SERVER_URL=$SERVER_URL" \
    "TRANSCRIPT_SYNC_APP_NAME=$DISPLAY_NAME" \
    "TRANSCRIPT_SYNC_ACCESS_GATE=$ACCESS_GATE" \
    "TRANSCRIPT_SYNC_CLOUD_CERT_PEM=secretref:cloud-cert-pem" \
  -o none

printf '\nDeployed: %s/mcp\n' "$SERVER_URL"
printf 'Metadata: %s/.well-known/oauth-protected-resource\n' "$SERVER_URL"
printf '\nRequired next step:\n'
echo "  Rerun scripts/create_cloud_app.py with --app-client-id $CLOUD_CLIENT_ID --server-url $SERVER_URL"
printf '  Then grant and verify admin consent before adding the connector.\n'

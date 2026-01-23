set -e

gcloud config set project boda8-6879
gcloud config set run/region asia-southeast1

SERVICE="boda8-bot"
REGION="asia-southeast1"

RUN_URL="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')"
echo "RUN_URL=$RUN_URL"

echo "---- health checks ----"
curl -sS -o /dev/null -w "healthz/ HTTP=%{http_code}\n" "$RUN_URL/healthz/" || true
curl -sS -o /dev/null -w "healthz  HTTP=%{http_code}\n" "$RUN_URL/healthz" || true

echo "---- revisions ----"
gcloud run services describe "$SERVICE" --region "$REGION" \
  --format="value(status.latestReadyRevisionName,status.latestCreatedRevisionName)"

REV_CREATED="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.latestCreatedRevisionName)')"
echo "LATEST_CREATED=$REV_CREATED"

echo "---- recent errors (created revision) ----"
gcloud logging read \
'resource.type="cloud_run_revision"
 AND resource.labels.service_name="boda8-bot"
 AND resource.labels.revision_name="'"$REV_CREATED"'"
 AND (severity>=ERROR OR textPayload:"Traceback" OR textPayload:"Exception" OR textPayload:"ModuleNotFoundError" OR textPayload:"RuntimeError")' \
--limit 80 \
--format="value(textPayload)" || true

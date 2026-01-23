set -e

SERVICE="boda8-bot"
REGION="asia-southeast1"

RUN_URL="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')"
READY_REV="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.latestReadyRevisionName)')"
CREATED_REV="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.latestCreatedRevisionName)')"

echo "RUN_URL=$RUN_URL"
echo "READY_REV=$READY_REV"
echo "CREATED_REV=$CREATED_REV"

echo "---- HTTP checks ----"
echo "GET /healthz/ =>"
curl -sS -i "$RUN_URL/healthz/" | head -n 12 || true
echo
echo "GET /telegram => (PATUT 405 sebab GET)"
curl -sS -i "$RUN_URL/telegram" | head -n 8 || true
echo

echo "---- last 10 min request logs (telegram) ----"
gcloud logging read \
'resource.type="cloud_run_revision"
 AND resource.labels.service_name="boda8-bot"
 AND httpRequest.requestUrl:"/telegram"
 AND timestamp>="'"$(date -u -d '10 minutes ago' +%Y-%m-%dT%H:%M:%SZ)"'"' \
--limit 30 \
--format="table(timestamp,httpRequest.requestMethod,httpRequest.status,httpRequest.requestUrl)" || true

echo
echo "---- last 30 errors/exceptions ----"
gcloud logging read \
'resource.type="cloud_run_revision"
 AND resource.labels.service_name="boda8-bot"
 AND (severity>=ERROR OR textPayload:"Bad Request" OR textPayload:"Traceback" OR textPayload:"Exception")' \
--limit 30 \
--format="value(textPayload)" || true


#!/bin/bash
set -e

# Determine service type from ZEABUR_SERVICE_ID or SERVICE_TYPE
SERVICE_TYPE="${SERVICE_TYPE:-api}"

# Check if ZEABUR_SERVICE_ID contains "worker"
if [[ "${ZEABUR_SERVICE_ID}" == *"worker"* ]]; then
    SERVICE_TYPE="worker"
fi

echo "Starting service: ${SERVICE_TYPE}"

if [ "${SERVICE_TYPE}" = "worker" ]; then
    exec python -m apps.worker.run
else
    exec python -m uvicorn apps.api.main:app --host 0.0.0.0 --port ${PORT:-8000}
fi

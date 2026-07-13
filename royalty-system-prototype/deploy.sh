#!/bin/bash
echo "=== Starting Cloud Run Deployment Workflow ==="

# Check if Docker is running
if ! docker info >/dev/null 2>&1; then
    echo "❌ Error: Docker is not running. Please start Docker Desktop first!"
    exit 1
fi

echo "🟢 Step 1: Building Docker Image (AMD64)..."
docker build --platform linux/amd64 -t gcr.io/royalty-502107/royalty-app:latest .
if [ $? -ne 0 ]; then
    echo "❌ Error: Docker build failed!"
    exit 1
fi

echo "🟢 Step 2: Pushing Image to Google Container Registry..."
docker push gcr.io/royalty-502107/royalty-app:latest
if [ $? -ne 0 ]; then
    echo "❌ Error: Docker push failed!"
    exit 1
fi

echo "🟢 Step 3: Deploying Service to GCP Cloud Run..."
gcloud run deploy royalty-app \
  --image=gcr.io/royalty-502107/royalty-app:latest \
  --region=asia-east1 \
  --project=royalty-502107 \
  --memory=2Gi

if [ $? -eq 0 ]; then
    echo "✅ Success: Cloud Run Deployment completed successfully!"
    echo "Service URL: https://royalty-app-674727095219.asia-east1.run.app"
else
    echo "❌ Error: gcloud deployment failed!"
    exit 1
fi

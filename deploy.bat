@echo off
echo Deploying to Azure Functions...
az login
az functionapp deployment source config-zip ^
  --resource-group ticket-processor-rg ^
  --name mjhughes-ticket-processor ^
  --src release.zip
echo Done.
pause

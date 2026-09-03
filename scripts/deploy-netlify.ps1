$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "Team Workload - Netlify Deploy" -ForegroundColor Cyan
Write-Host "--------------------------------" -ForegroundColor Cyan

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "ไม่พบ Node.js กรุณาติดตั้ง Node.js ก่อน"
}

Write-Host "1) เปิด Netlify login/link ถ้าจำเป็น..." -ForegroundColor Yellow
npx netlify-cli status 2>$null
if ($LASTEXITCODE -ne 0) {
  npx netlify-cli login
}

Write-Host ""
Write-Host "2) Deploy production..." -ForegroundColor Yellow
npx netlify-cli deploy --prod --dir public --functions netlify/functions

Write-Host ""
Write-Host "เสร็จแล้ว หากเป็นการ deploy ครั้งแรก ให้ตั้ง Environment Variables:" -ForegroundColor Green
Write-Host "  APPS_SCRIPT_API_URL"
Write-Host "  APPS_SCRIPT_API_SECRET"

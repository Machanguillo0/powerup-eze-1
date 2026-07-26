param(
    [string]$Repo = "Machanguillo0/powerup-eze-1"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$gitExe = "C:\Program Files\Git\cmd\git.exe"
if (-not (Test-Path $gitExe)) { $gitExe = "git" }

Write-Host "🚀 Subiendo cambios a GitHub..." -ForegroundColor Cyan
& $gitExe add .
$commitMsg = "Actualización automática - " + (Get-Date -Format "HH:mm:ss")
& $gitExe commit -m $commitMsg
& $gitExe push origin main

Write-Host "`n⏳ Monitoreando el despliegue en GitHub Pages..." -ForegroundColor Yellow
Start-Sleep -Seconds 4

$startTime = Get-Date
$initialRunId = $null

try {
    $latest = (Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/actions/runs?per_page=1").workflow_runs[0]
    $initialRunId = $latest.id
} catch {}

while ($true) {
    try {
        $run = (Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/actions/runs?per_page=1").workflow_runs[0]
        $status = $run.status
        $conclusion = $run.conclusion

        if ($status -eq "completed") {
            if ($conclusion -eq "success") {
                Write-Host "`n"
                Write-Host "==================================================" -ForegroundColor Green
                Write-Host "  ✅ ¡DESPLIEGUE COMPLETADO CON ÉXITO!" -ForegroundColor Green
                Write-Host "  🌐 Tu web ya está lista en Trello." -ForegroundColor Green
                Write-Host "==================================================" -ForegroundColor Green
            } else {
                Write-Host "`n❌ Error en el despliegue de GitHub: $conclusion" -ForegroundColor Red
            }
            break
        } else {
            Write-Host -NoNewline "⏳ Desplegando en GitHub Pages... "
            Write-Host "[$status]" -ForegroundColor Yellow
            Start-Sleep -Seconds 4
        }
    } catch {
        Start-Sleep -Seconds 4
    }

    if ((Get-Date) - $startTime -gt [timespan]::FromMinutes(3)) {
        Write-Host "`n⚠️ Tiempo de espera agotado. Comprueba GitHub Actions." -ForegroundColor Yellow
        break
    }
}

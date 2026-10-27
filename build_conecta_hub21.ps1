Stop-Process -Name "N.E.X.U.S" -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

Write-Host "Compilando N.E.X.U.S via PyInstaller..." -ForegroundColor Cyan
py -m PyInstaller "N.E.X.U.S.spec" --clean --noconfirm

Stop-Process -Name "N.E.X.U.S" -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800

# Copia para Desktop e pasta NEXUS local
if (Test-Path "dist\N.E.X.U.S.exe") {
    $desktops = @([Environment]::GetFolderPath('Desktop'), "$HOME\Desktop", "$env:USERPROFILE\OneDrive\Área de Trabalho", "$env:USERPROFILE\OneDrive\Desktop") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
    foreach ($d in $desktops) {
        Copy-Item -Path "dist\N.E.X.U.S.exe" -Destination (Join-Path $d "N.E.X.U.S.exe") -Force
    }
    if (Test-Path "$HOME\Desktop\NEXUS") {
        Copy-Item -Path "dist\N.E.X.U.S.exe" -Destination "$HOME\Desktop\NEXUS\N.E.X.U.S.exe" -Force
    }
    Write-Host "N.E.X.U.S.exe gerado e copiado para a Área de Trabalho com sucesso!" -ForegroundColor Green
} else {
    Write-Host "Erro: dist\N.E.X.U.S.exe não foi encontrado após o build." -ForegroundColor Red
}

# Publicação automática e compactação na rede via Python (suporte nativo a acentos e caminhos de rede)
py scripts\publicar_rede.py


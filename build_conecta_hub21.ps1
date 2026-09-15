py -m PyInstaller "N.E.X.U.S.spec" --clean --noconfirm
Copy-Item -Path "dist\N.E.X.U.S.exe" -Destination "$HOME\Desktop\N.E.X.U.S.exe" -Force
Write-Host "N.E.X.U.S.exe gerado com sucesso no Desktop e na pasta dist!" -ForegroundColor Green

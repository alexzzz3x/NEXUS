Stop-Process -Name "N.E.X.U.S" -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

Write-Host "Compilando N.E.X.U.S via PyInstaller..." -ForegroundColor Cyan
py -m PyInstaller "N.E.X.U.S.spec" --clean --noconfirm

Stop-Process -Name "N.E.X.U.S" -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800

# Copia para Desktop e pasta NEXUS local
Copy-Item -Path "dist\N.E.X.U.S.exe" -Destination "$HOME\Desktop\N.E.X.U.S.exe" -Force
if (Test-Path "$HOME\Desktop\NEXUS") {
    Copy-Item -Path "dist\N.E.X.U.S.exe" -Destination "$HOME\Desktop\NEXUS\N.E.X.U.S.exe" -Force
}
Write-Host "N.E.X.U.S.exe gerado localmente com sucesso!" -ForegroundColor Green

# Publicacao automatica e compactacao na rede via Python (suporte nativo a acentos e caminhos de rede)
py -c "
import os, zipfile, json, datetime

exe_path = r'dist\N.E.X.U.S.exe'
pasta_rede = r'X:\obras\Elpa - Almox CR Mauá\ANDERSON'

if os.path.exists(pasta_rede) and os.path.exists(exe_path):
    print('Compactando NEXUS.zip para publicacao na rede...')
    zip_dest = os.path.join(pasta_rede, 'NEXUS.zip')
    with zipfile.ZipFile(zip_dest, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.write(exe_path, arcname='N.E.X.U.S.exe')
    
    versao_dest = os.path.join(pasta_rede, 'versao.json')
    dados_versao = {
        'versao': '1.0.0',
        'data': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'),
        'arquivo_zip': 'NEXUS.zip',
        'notas': 'N.E.X.U.S v1.0 — Versao Oficial de Lancamento.'
    }
    if os.path.exists('versao.json'):
        try:
            with open('versao.json', 'r', encoding='utf-8') as vf:
                dados_versao = json.load(vf)
                dados_versao['data'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        except Exception:
            pass
    with open(versao_dest, 'w', encoding='utf-8') as f:
        json.dump(dados_versao, f, indent=2, ensure_ascii=False)
    print('Publicado com sucesso em: ' + pasta_rede)
else:
    print('Pasta da rede nao encontrada ou exe ausente.')
"

# -*- coding: utf-8 -*-
import os
import sys
import zipfile
import json
import datetime

def publicar():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    exe_path = os.path.join(base_dir, "dist", "N.E.X.U.S.exe")
    pasta_rede = r"X:\obras\Elpa - Almox CR Mauá\ANDERSON"
    versao_local_path = os.path.join(base_dir, "versao.json")

    print(f"Verificando pasta de rede: {pasta_rede}")
    if not os.path.exists(pasta_rede):
        print("Pasta de rede não acessível. Publicação ignorada.")
        return False

    if not os.path.exists(exe_path):
        print(f"Executável não encontrado em: {exe_path}")
        return False

    try:
        print("Compactando NEXUS.zip para publicação na rede...")
        zip_dest = os.path.join(pasta_rede, "NEXUS.zip")
        with zipfile.ZipFile(zip_dest, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            zf.write(exe_path, arcname="N.E.X.U.S.exe")
        print(f"NEXUS.zip criado com sucesso em: {zip_dest}")

        dados_versao = {
            "versao": "2.3.2",
            "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "arquivo_zip": "NEXUS.zip",
            "notas": "Atualização automática N.E.X.U.S."
        }
        if os.path.exists(versao_local_path):
            with open(versao_local_path, "r", encoding="utf-8") as vf:
                dados_versao = json.load(vf)
                dados_versao["versao"] = "2.3.2"
                dados_versao["data"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        versao_dest = os.path.join(pasta_rede, "versao.json")
        with open(versao_dest, "w", encoding="utf-8") as f:
            json.dump(dados_versao, f, indent=2, ensure_ascii=False)
        print(f"versao.json publicado com sucesso em: {versao_dest}")

        print("Publicação na rede concluída com 100% de sucesso!")
        return True
    except Exception as e:
        print(f"Erro ao publicar na rede: {e}")
        return False

if __name__ == "__main__":
    publicar()

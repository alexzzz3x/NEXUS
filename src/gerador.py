import os
import base64
import qrcode
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.graphics.barcode import code128
from utils_os import obter_caminho_recurso
from utils_metrics import incrementar_metrica


def gerar_html_para_impressao(deposito, material, descricao, qtd, unidade, lote="", tamanho="Grande", qr_topo=""):
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    pasta_destino = os.path.join(diretorio_atual, "etiquetas_geradas")
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    mapa_medidas = {"Pequena": (100, 63), "Média": (150, 88), "Grande": (210, 148)}
    largura_mm, altura_mm = mapa_medidas.get(tamanho, (210, 148))

    qr_base64 = ""
    if qr_topo:
        img = qrcode.make(str(qr_topo))
        buf = BytesIO()
        img.save(buf, format="PNG")
        qr_base64 = base64.b64encode(buf.getvalue()).decode()

    nome_arquivo = str(material).strip() or "identificacao_livre"
    caminho_html = os.path.join(pasta_destino, f"etiqueta_{nome_arquivo}.html")

    qr_html = ""
    if qr_base64:
        qr_html = f'<img src="data:image/png;base64,{qr_base64}" style="width:35mm;height:35mm;position:absolute;top:3mm;right:3mm;">'

    logo_base64 = ""
    caminho_logo = obter_caminho_recurso("assets", "logo_cntc.png")
    if os.path.exists(caminho_logo):
        with open(caminho_logo, "rb") as img_f:
            logo_base64 = base64.b64encode(img_f.read()).decode()

    logo_html = ""
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="height:12mm; margin-right: 10mm;">'
    else:
        logo_html = '<span class="logo-text">CONECTA LOG</span>'

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Etiqueta {nome_arquivo}</title>
<style>
  @page {{ size: {largura_mm}mm {altura_mm}mm; margin: 0; }}
  @media print {{ body {{ margin: 0; }} .no-print {{ display: none; }} }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{
    width: {largura_mm}mm;
    height: {altura_mm}mm;
    font-family: Arial, Helvetica, sans-serif;
    position: relative;
    border: 1.5px solid black;
    padding: 4mm;
    overflow: hidden;
  }}
  .header {{
    border-bottom: 1px solid black;
    padding-bottom: 2mm;
    margin-bottom: 2mm;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .logo-text {{ font-size: 11pt; font-weight: bold; }}
  .dep-text {{ font-size: 12pt; font-weight: bold; text-align: right; }}
  .deposito {{ font-size: 42pt; font-weight: bold; text-align: center; margin: 4mm 0 2mm 0; }}
  .material {{ font-size: 20pt; font-weight: bold; text-align: center; margin-bottom: 2mm; }}
  .descricao {{ font-size: 10pt; font-weight: bold; text-align: center; margin-bottom: 2mm; }}
  .quantidade {{ font-size: 14pt; font-weight: bold; text-align: center; margin-bottom: 1mm; }}
  .lote {{ font-size: 12pt; font-weight: bold; text-align: center; margin-bottom: 2mm; }}
</style>
</head>
<body onload="setTimeout(function(){{window.print();}}, 300);">
  {qr_html}
  <div class="header">
    {logo_html}
    <span class="dep-text">DEP: {deposito}</span>
  </div>
  <div class="deposito">{deposito}</div>
  <div class="material">{material}</div>
  <div class="descricao">{descricao[:70]}</div>
  {"<div class='quantidade'>QUANTIDADE: " + str(qtd) + " " + str(unidade) + "</div>" if qtd else ""}
  {"<div class='lote'>LOTE: " + str(lote) + "</div>" if lote else ""}
</body>
</html>"""

    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html)

    return caminho_html


def gerar_pdf_placa_tradicional(deposito, material, descricao, qtd, unidade, lote="", tamanho="Grande", qr_topo=""):
    # Configuração de diretórios
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    pasta_destino = os.path.join(diretorio_atual, "etiquetas_geradas")
    if not os.path.exists(pasta_destino): os.makedirs(pasta_destino)
    
    # 1. Definição do tamanho e escala
    if tamanho == "Pequena":
        page_size = (400, 250)
        escala = 0.45 
    elif tamanho == "Média":
        page_size = (600, 350)
        escala = 0.62
    else: # Grande
        page_size = landscape(A4)
        escala = 1.0 
        
    nome_arquivo = str(material).strip() or "identificacao_livre"
    caminho_salvamento = os.path.join(pasta_destino, f"etiqueta_{nome_arquivo}.pdf")
    c = canvas.Canvas(caminho_salvamento, pagesize=page_size)
    l, a = page_size

    # 2. Desenho da Estrutura
    c.setStrokeColor(colors.black); c.setLineWidth(2)
    c.rect(20 * escala, 20 * escala, l - (40 * escala), a - (40 * escala))
    c.line(20 * escala, a - (100 * escala), l - (20 * escala), a - (100 * escala))
    
    # 3. Logo
    caminho_logo = obter_caminho_recurso("assets", "logo_cntc.png")
    if os.path.exists(caminho_logo):
        c.drawImage(caminho_logo, 30 * escala, a - (85 * escala), width=150 * escala, height=50 * escala, preserveAspectRatio=True, anchor='sw')

    # 4. QR Code do Canto Superior Direito (Se preenchido)
    c.setFont("Helvetica-Bold", 20 * escala)
    if qr_topo:
        img_qr_topo = qrcode.make(str(qr_topo))
        caminho_qr_topo_temp = os.path.join(pasta_destino, f"temp_qr_topo_{nome_arquivo}.png")
        img_qr_topo.save(caminho_qr_topo_temp)
        
        tamanho_qr_topo = 65 * escala
        pos_x_topo = l - (30 * escala) - tamanho_qr_topo
        pos_y_topo = a - (25 * escala) - tamanho_qr_topo
        
        c.drawImage(caminho_qr_topo_temp, pos_x_topo, pos_y_topo, width=tamanho_qr_topo, height=tamanho_qr_topo)
        
        if os.path.exists(caminho_qr_topo_temp):
            os.remove(caminho_qr_topo_temp)
            
        # Desloca o texto do depósito para a esquerda do QR Code para não sobrepor
        c.drawRightString(pos_x_topo - (15 * escala), a - (65 * escala), f"DEPÓSITO: {deposito}")
    else:
        c.drawRightString(l - (30 * escala), a - (65 * escala), f"DEPÓSITO: {deposito}")
    
    # 5. Textos Principais
    c.setFont("Helvetica-Bold", 110 * escala)
    c.drawCentredString(l/2, 370 * escala, deposito)
    
    c.setFont("Helvetica-Bold", 55 * escala)
    c.drawCentredString(l/2, 290 * escala, str(material))
    
    c.setFont("Helvetica-Bold", 24 * escala)
    c.drawCentredString(l/2, 230 * escala, descricao[:70])
    
    if qtd:
        c.setFont("Helvetica-Bold", 38 * escala)
        c.drawCentredString(l/2, 160 * escala, f"QUANTIDADE: {qtd} {unidade}")
    
    if lote:
        c.setFont("Helvetica-Bold", 24 * escala)
        c.drawCentredString(l/2, 110 * escala, f"LOTE: {lote}")

    # 6. Código de Barras Tradicional (Embaixo)
    if str(material).strip():
        codigo_barras = code128.Code128(str(material), barHeight=65 * escala, barWidth=2.2 * escala)
        codigo_barras.drawOn(c, (l - codigo_barras.width) / 2, 30 * escala)

    c.showPage()
    c.save()
    incrementar_metrica("etiquetas_geradas", 1)
    return caminho_salvamento

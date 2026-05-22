import os
import json
import base64
import requests
import tempfile
import uuid
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from pdf2docx import Converter
import fitz  # PyMuPDF

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = Path("uploads")
OUTPUT_FOLDER = Path("outputs")
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# ── DETECÇÃO DE PDF ESCANEADO ─────────────────────────────────────
def is_scanned(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        total_text = ""
        for page in doc:
            total_text += page.get_text()
        doc.close()
        return len(total_text.strip()) < 100
    except:
        return True

# ── CONVERSÃO DIRETA PDF → DOCX (pdf2docx) ───────────────────────
def convert_direct(pdf_path, output_path):
    cv = Converter(str(pdf_path))
    cv.convert(str(output_path))
    cv.close()

# ── CONVERSÃO VIA API CLAUDE ──────────────────────────────────────
PROMPT_SIMPLES = """Extraia o conteúdo da página e retorne SOMENTE JSON válido.
{"blocks":[{"type":"heading"|"paragraph"|"table","runs":[{"text":"...","bold":false,"italic":false,"superscript":false,"subscript":false,"color":null}],"align":"left"|"center"|"right"|"justify","table":{"header_bg":"D9D9D9","border_color":"000000","rows":[{"cells":[{"runs":[{"text":"...","bold":false}],"bg":null}]}]}}]}
REGRAS:
- runs: trechos com formatação diferente viram runs separados. Negrito parcial DEVE ser separado.
- Tabelas: detecte cor de fundo de cada célula visualmente
- Títulos numerados (1., 4.1., 4.1.1.): type "heading"
- Sobrescrito: superscript:true. Subscrito: subscript:true
- Símbolos especiais (α, μ, Ø, ≤, ≥): preservar exatamente
- Texto colorido: color:"RRGGBB" (hex sem #, null se preto)
- Preserve exatamente o texto com acentos e pontuação
- Ignore logo, cabeçalho repetitivo de página, assinaturas, carimbos"""

PROMPT_JURAMENTADO = """Você é especialista em formatação de traduções juramentadas brasileiras.
Analise a imagem e retorne SOMENTE JSON válido.
{"idioma":"pt"|"en"|"es"|"de"|"fr"|"it"|"zh","blocks":[{"type":"paragraph"|"table","runs":[{"text":"...","bold":false,"italic":false,"superscript":false,"subscript":false,"color":null}],"align":"left"|"center"|"right"|"justify","table":{"header_bg":"D9D9D9","border_color":"000000","rows":[{"cells":[{"runs":[{"text":"...","bold":false}],"bg":null}]}]}}]}

SUBSTITUIÇÕES VISUAIS OBRIGATÓRIAS:
- Brasão/armas de país → [Consta brasão de (nome do país/estado)]
- Selo oficial → [Consta selo de (nome do país/estado/instituição)]
- Logo legível → [Consta logotipo com o seguinte teor] e transcreva o texto visível no logo abaixo
- Foto de pessoa(s) → [Consta fotografia]
- Imagem genérica única → [Consta imagem]
- Múltiplas imagens → [Constam imagens]
- Desenho técnico único → [Consta desenho técnico]
- Múltiplos desenhos → [Constam desenhos técnicos]
- Gráfico → [Consta gráfico]
- Diagrama → [Consta diagrama]
- QR Code → [Consta código QR]
- Assinatura manuscrita única → [Consta assinatura]
- Múltiplas assinaturas → [Constam assinaturas]
- Carimbo legível → [Consta carimbo com o seguinte teor] e transcreva o texto do carimbo abaixo
- Carimbo ilegível → [Consta carimbo ilegível]
- Texto ilegível → [Ilegível]
- Palavras riscadas → NÃO incluir
REGRAS: Texto corrido, ordem lógica. Símbolos (α,μ,Ø,≤,≥) preservados. Superscript:true para sobrescritos."""

PROMPT_CETRA_EXTRA = """
MODO CETRA ATIVADO:
- Documentos bilíngues: manter APENAS o texto em inglês
- EXCEÇÕES: cabeçalho e nomes de cargos/referências institucionais
- Palavras riscadas: NUNCA incluir
- Rodapé repetitivo: incluir APENAS na última página, precedido de [consta nota de rodapé]
- Imagens com título "Drawing No. X": incluir título + [Constam desenhos técnicos]
- Numeração: corrigir pontos faltantes (1.11 → 1.1.1)"""

def extract_page_claude(image_bytes, page_num, api_key, mode, cetra):
    b64 = base64.b64encode(image_bytes).decode()
    system = PROMPT_JURAMENTADO + (PROMPT_CETRA_EXTRA if cetra else "") if mode == "juramentado" else PROMPT_SIMPLES
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"Content-Type": "application/json", "x-api-key": api_key, "anthropic-version": "2023-06-01"},
        json={
            "model": "claude-opus-4-5",
            "max_tokens": 4000,
            "system": system,
            "messages": [{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": f"Página {page_num}. Retorne JSON."}
            ]}]
        },
        timeout=90
    )
    data = resp.json()
    if not resp.ok:
        raise Exception(data.get("error", {}).get("message", "Erro na API"))
    raw = data["content"][0]["text"].strip().replace("```json","").replace("```","").strip()
    return json.loads(raw)

# ── CONSTRUTOR DE DOCX A PARTIR DE JSON ───────────────────────────
def build_docx_from_json(all_pages, output_path, mode):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    INDENT = {1: 0, 2: 504, 3: 1080, 4: 1656}
    HANG   = {1: 360, 2: 432, 3: 504, 4: 576}

    import re
    def detect_level(runs):
        text = "".join(r.get("text","") for r in (runs or []))
        m = re.match(r'^(\d+(?:\.\d+)*)\.?\s', text)
        return len(m.group(1).split(".")) if m else None

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    def add_runs(para, runs, size_pt=10):
        for r in (runs or []):
            run = para.add_run(r.get("text",""))
            run.bold = r.get("bold", False)
            run.italic = r.get("italic", False)
            run.font.name = "Arial"
            run.font.size = Pt(size_pt)
            if r.get("superscript"):
                run.font.superscript = True
            if r.get("subscript"):
                run.font.subscript = True
            if r.get("color"):
                try:
                    c = r["color"].lstrip("#")
                    run.font.color.rgb = RGBColor(int(c[0:2],16), int(c[2:4],16), int(c[4:6],16))
                except:
                    pass

    def add_table(doc, table_data):
        rows = table_data.get("rows", [])
        if not rows: return
        num_cols = max(len(r.get("cells",[])) for r in rows)
        if num_cols == 0: return
        table = doc.add_table(rows=len(rows), cols=num_cols)
        table.style = "Table Grid"
        for ri, row in enumerate(rows):
            cells = row.get("cells", [])
            for ci in range(num_cols):
                cell = table.cell(ri, ci)
                cell.text = ""
                if ci < len(cells):
                    cell_data = cells[ci]
                    bg = cell_data.get("bg")
                    if not bg and ri == 0:
                        bg = table_data.get("header_bg", "D9D9D9")
                    if bg:
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        shd = OxmlElement("w:shd")
                        shd.set(qn("w:val"), "clear")
                        shd.set(qn("w:color"), "auto")
                        shd.set(qn("w:fill"), bg.lstrip("#"))
                        tcPr.append(shd)
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if ri == 0 else WD_ALIGN_PARAGRAPH.LEFT
                    add_runs(para, cell_data.get("runs", []), size_pt=9)
        doc.add_paragraph()

    for pi, page_data in enumerate(all_pages):
        if pi > 0:
            doc.add_page_break()
        blocks = page_data.get("blocks", [])
        cur_lvl = 1

        for block in blocks:
            btype = block.get("type", "paragraph")
            runs = block.get("runs", [])
            align = block.get("align", "justify")

            if btype == "table":
                add_table(doc, block.get("table", {}))
                continue

            lvl = detect_level(runs) if mode == "simples" else None

            p = doc.add_paragraph()
            align_map = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}

            if lvl is not None:
                cur_lvl = lvl
                left = (INDENT.get(lvl, 0) + HANG.get(lvl, 360))
                p.paragraph_format.left_indent = Pt(left / 20)
                p.paragraph_format.first_line_indent = Pt(-HANG.get(lvl, 360) / 20)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            else:
                left = (INDENT.get(cur_lvl, 0) + HANG.get(cur_lvl, 0)) if mode == "simples" else 0
                if left > 0:
                    p.paragraph_format.left_indent = Pt(left / 20)
                p.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)

            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            add_runs(p, runs)

    doc.save(str(output_path))

# ── ROTAS ─────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]
    mode = request.form.get("mode", "simples")
    use_api = request.form.get("use_api", "false") == "true"
    cetra = request.form.get("cetra", "false") == "true"
    api_key = request.form.get("api_key", "")

    # Salvar PDF
    pdf_id = str(uuid.uuid4())
    pdf_path = UPLOAD_FOLDER / f"{pdf_id}.pdf"
    output_path = OUTPUT_FOLDER / f"{pdf_id}.docx"
    file.save(str(pdf_path))

    try:
        scanned = is_scanned(pdf_path)

        if not use_api and not scanned and mode == "simples":
            # PDF digital + simples + sem IA → conversão direta
            convert_direct(pdf_path, output_path)

        elif use_api or scanned:
            # Usa API Claude
            if not api_key:
                return jsonify({"error": "Chave de API necessária para este tipo de conversão"}), 400

            doc = fitz.open(str(pdf_path))
            all_pages = []
            detected_langs = set()

            for page_num in range(len(doc)):
                page = doc[page_num]
                mat = fitz.Matrix(1.5, 1.5)
                pix = page.get_pixmap(matrix=mat)
                img_bytes = pix.tobytes("jpeg")
                try:
                    data = extract_page_claude(img_bytes, page_num + 1, api_key, mode, cetra)
                    all_pages.append(data)
                    if data.get("idioma"):
                        detected_langs.add(data["idioma"])
                except Exception as e:
                    all_pages.append({"blocks": [{"type": "paragraph", "runs": [{"text": f"[Erro na página {page_num+1}: {str(e)}]"}], "align": "left"}]})
            doc.close()
            build_docx_from_json(all_pages, output_path, mode)

        else:
            # Fallback: conversão direta
            convert_direct(pdf_path, output_path)

        original_name = Path(file.filename).stem
        return send_file(
            str(output_path),
            as_attachment=True,
            download_name=f"{original_name}_convertido.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if pdf_path.exists():
            pdf_path.unlink()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

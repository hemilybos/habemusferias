"""
Gerador Automático de Alertas de Vencimento de Férias
Empresa: MOBS2 - BWA Global Departamento Pessoal
"""

import os
import re
import io
import copy
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import pdfplumber
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor
import lxml.etree as etree

from flask import (
    Flask, request, render_template_string,
    send_file, session, redirect, url_for, flash
)

app = Flask(__name__)
app.secret_key = "ferias_bwa_2026"

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "template_alerta.docx"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# HTML TEMPLATE
# ─────────────────────────────────────────────
HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Alertas de Férias – BWA Global</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    background: #f0f2f5;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 16px;
    color: #1a1a2e;
  }

  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 4px 24px rgba(0,0,0,.10);
    padding: 40px 44px;
    width: 100%;
    max-width: 680px;
  }

  .logo-bar {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 28px;
    padding-bottom: 20px;
    border-bottom: 2px solid #f0f0f0;
  }

  .logo-circle {
    width: 48px; height: 48px;
    border-radius: 50%;
    background: linear-gradient(135deg, #d4a017, #a0522d);
    display: flex; align-items: center; justify-content: center;
    font-weight: 900; color: #fff; font-size: 20px;
  }

  .logo-bar h1 { font-size: 1.15rem; color: #1a1a2e; font-weight: 700; }
  .logo-bar p  { font-size: 0.78rem; color: #888; margin-top: 2px; }

  h2 { font-size: 1.05rem; color: #333; margin-bottom: 18px; font-weight: 600; }

  .upload-zone {
    border: 2px dashed #c8c8d8;
    border-radius: 10px;
    padding: 32px 24px;
    text-align: center;
    background: #fafafa;
    transition: border-color .2s, background .2s;
    cursor: pointer;
    position: relative;
  }
  .upload-zone:hover, .upload-zone.drag-over {
    border-color: #d4a017;
    background: #fffdf0;
  }
  .upload-zone input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
  }
  .upload-zone .icon { font-size: 2.2rem; margin-bottom: 8px; }
  .upload-zone p { color: #666; font-size: .9rem; }
  .upload-zone strong { color: #333; }
  .file-name {
    margin-top: 10px;
    font-size: .85rem;
    color: #d4a017;
    font-weight: 600;
    min-height: 20px;
  }

  .info-box {
    background: #fff8e1;
    border-left: 4px solid #d4a017;
    border-radius: 6px;
    padding: 12px 16px;
    margin: 20px 0;
    font-size: .87rem;
    color: #5a4000;
  }
  .info-box b { color: #a0522d; }

  .btn {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 12px 28px;
    border-radius: 8px;
    font-size: .95rem;
    font-weight: 600;
    border: none;
    cursor: pointer;
    text-decoration: none;
    transition: opacity .2s, transform .1s;
  }
  .btn:hover { opacity: .88; transform: translateY(-1px); }
  .btn:active { transform: translateY(0); }

  .btn-primary { background: linear-gradient(135deg, #d4a017, #a0522d); color: #fff; width: 100%; justify-content: center; margin-top: 20px; font-size: 1rem; padding: 14px; }
  .btn-docx    { background: #2b5797; color: #fff; }
  .btn-pdf     { background: #c0392b; color: #fff; }

  .result-section { margin-top: 28px; }
  .result-section h3 { font-size: .95rem; color: #444; margin-bottom: 14px; font-weight: 600; }

  .preview-table {
    width: 100%;
    border-collapse: collapse;
    font-size: .82rem;
    margin-bottom: 20px;
  }
  .preview-table th {
    background: #2b5797;
    color: #fff;
    padding: 9px 12px;
    text-align: left;
  }
  .preview-table td {
    padding: 8px 12px;
    border-bottom: 1px solid #eee;
  }
  .preview-table tr:last-child td { border-bottom: none; }
  .preview-table tr:nth-child(even) td { background: #f7f7fb; }
  .badge-urgent {
    background: #fdecea;
    color: #c0392b;
    border-radius: 12px;
    padding: 2px 10px;
    font-size: .76rem;
    font-weight: 700;
  }

  .download-bar {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
  }

  .alert-msg {
    padding: 12px 16px;
    border-radius: 8px;
    font-size: .88rem;
    margin-bottom: 16px;
  }
  .alert-error   { background: #fdecea; color: #b71c1c; border-left: 4px solid #e53935; }
  .alert-warning { background: #fff8e1; color: #795800; border-left: 4px solid #f9a825; }
  .alert-success { background: #e8f5e9; color: #1b5e20; border-left: 4px solid #43a047; }

  .spinner {
    display: none;
    width: 18px; height: 18px;
    border: 3px solid rgba(255,255,255,.4);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin .7s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  footer {
    margin-top: 32px;
    font-size: .78rem;
    color: #aaa;
    text-align: center;
  }
</style>
</head>
<body>

<div class="card">
  <div class="logo-bar">
    <div class="logo-circle">B</div>
    <div>
      <h1>BWA Global – Departamento Pessoal</h1>
      <p>Gerador de Alertas de Vencimento de Férias</p>
    </div>
  </div>

  {% with msgs = get_flashed_messages(with_categories=true) %}
    {% for cat, msg in msgs %}
      <div class="alert-msg alert-{{ cat }}">{{ msg }}</div>
    {% endfor %}
  {% endwith %}

  {% if not result %}
  {# ── UPLOAD FORM ── #}
  <h2>📂 Enviar Relatório de Férias (PDF)</h2>

  <form method="POST" enctype="multipart/form-data" id="upload-form">
    <div class="upload-zone" id="drop-zone">
      <input type="file" name="pdf_file" id="pdf-input" accept=".pdf" required>
      <div class="icon">📄</div>
      <p><strong>Clique para selecionar</strong> ou arraste o arquivo aqui</p>
      <p>Formato aceito: <strong>PDF</strong></p>
      <div class="file-name" id="file-label">Nenhum arquivo selecionado</div>
    </div>

    <div class="info-box">
      O sistema irá filtrar automaticamente os colaboradores com prazo de gozo de férias
      vencendo nos <b>próximos 3 meses</b> a partir de hoje (<b>{{ today }}</b>),
      gerar o documento Word preenchido e convertê-lo para PDF.
    </div>

    <button type="submit" class="btn btn-primary" id="submit-btn">
      <div class="spinner" id="spinner"></div>
      ⚙️ Processar e Gerar Alertas
    </button>
  </form>

  {% else %}
  {# ── RESULT ── #}
  <div class="result-section">
    <div class="alert-msg alert-success">
      ✅ <strong>{{ result.count }} colaborador(es)</strong> encontrado(s) com férias vencendo
      entre <strong>{{ today }}</strong> e <strong>{{ result.limit_date }}</strong>.
    </div>

    <h3>📋 Colaboradores incluídos no relatório</h3>
    <table class="preview-table">
      <thead>
        <tr>
          <th>Empregado</th>
          <th>Período Aquisitivo</th>
          <th>Data Limite</th>
        </tr>
      </thead>
      <tbody>
        {% for r in result.rows %}
        <tr>
          <td>{{ r.nome }}</td>
          <td>{{ r.periodo }}</td>
          <td>
            {{ r.data_limite }}
            {% if r.urgente %}
            <span class="badge-urgent">⚠ Urgente</span>
            {% endif %}
          </td>
        </tr>
        {% endfor %}
      </tbody>
    </table>

    <h3>⬇️ Downloads</h3>
    <div class="download-bar">
      <a href="{{ url_for('download', fmt='docx') }}" class="btn btn-docx">
        📝 Baixar DOCX
      </a>
      <a href="{{ url_for('download', fmt='pdf') }}" class="btn btn-pdf">
        📄 Baixar PDF
      </a>
    </div>

    <form method="GET" action="/" style="margin-top:24px">
      <button type="submit" class="btn" style="background:#eee;color:#333;width:100%;justify-content:center;">
        🔄 Processar outro arquivo
      </button>
    </form>
  </div>
  {% endif %}
</div>

<footer>BWA Global Consultoria Contábil Ltda &nbsp;·&nbsp; Sistema de Alertas de Férias &nbsp;·&nbsp; {{ today }}</footer>

<script>
const input = document.getElementById('pdf-input');
const label = document.getElementById('file-label');
const form  = document.getElementById('upload-form');
const btn   = document.getElementById('submit-btn');
const spin  = document.getElementById('spinner');

if (input) {
  input.addEventListener('change', () => {
    label.textContent = input.files[0] ? input.files[0].name : 'Nenhum arquivo selecionado';
  });
}

const zone = document.getElementById('drop-zone');
if (zone) {
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      label.textContent = e.dataTransfer.files[0].name;
    }
  });
}

if (form) {
  form.addEventListener('submit', () => {
    btn.disabled = true;
    spin.style.display = 'inline-block';
    btn.querySelector('span') && (btn.querySelector('span').textContent = ' Processando...');
  });
}
</script>
</body>
</html>
"""

# ─────────────────────────────────────────────
# PDF PARSING
# ─────────────────────────────────────────────
DATE_RE = re.compile(r'\d{2}/\d{2}/\d{4}')

def parse_date(s):
    """Parse dd/mm/yyyy string to date object."""
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y").date()
    except Exception:
        return None

def extract_records(pdf_bytes):
    """
    Parse the 'RELATÓRIO DE FÉRIAS' PDF and return a list of dicts:
      {nome, periodo_aquisitivo, prazo_final}
    Only rows that have a 'Prazo final p/ iniciar as férias sem gerar dobro' are returned.
    """
    records = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")

            for line in lines:
                line = line.strip()
                # Data lines start with a numeric code followed by the name
                # e.g. "21 DAVI VELOSO DA SILVA 30,00 01/08/2024 a 31/07/2025 01/08/2025 a 31/07/2026 02/07/2026"
                # We look for lines that start with digits then a name
                match = re.match(r'^(\d+)\s+([A-ZÁÉÍÓÚÃÕÂÊÔÇÀÜ][\w\s]+?)\s+([\d,]+)\s+(.+)$', line)
                if not match:
                    continue

                cod   = match.group(1)
                nome  = match.group(2).strip()
                resto = match.group(4).strip()

                # Extract all dates from the remainder
                dates = DATE_RE.findall(resto)
                if not dates:
                    continue

                # "Período Aquisitivo" is always first two dates (start a end)
                # "Deverá gozar entre" is next two dates
                # "Prazo final" (optional) is the last date if count is odd or >= 5

                # Layout possibilities:
                #  4 dates: periodo_aq(2) + gozar(2) -> no prazo_final
                #  5 dates: periodo_aq(2) + gozar(2) + prazo_final(1)
                # Actually from PDF text the order is:
                # periodo_aq_start periodo_aq_end  gozar_start gozar_end [prazo_final]

                # The "Referente Período Aquisitivo" columns in the PDF are rendered
                # AFTER "Deverá gozar" in the visual layout but pdfplumber reads
                # left-to-right reordered. Let's figure from PDF text:
                # "30,00 01/08/2024 a 31/07/2025 01/08/2025 a 31/07/2026 02/07/2026"
                #        ^-- ref periodo aquisitivo           ^-- deverá gozar  ^-- prazo

                # Check how many date groups separated by ' a '
                date_groups = re.findall(r'(\d{2}/\d{2}/\d{4})\s+a\s+(\d{2}/\d{2}/\d{4})', resto)

                # Trailing date (prazo_final) = last standalone date not part of a pair
                all_dates_in_order = dates
                paired_dates = [d for grp in date_groups for d in grp]

                prazo_final = None
                if len(all_dates_in_order) >= 5:
                    # Last date is prazo final
                    prazo_final = all_dates_in_order[-1]
                elif len(all_dates_in_order) == 5:
                    prazo_final = all_dates_in_order[4]

                # If no prazo_final, skip (no alert needed for this row)
                if not prazo_final:
                    continue

                # Referente período aquisitivo: the LAST date pair (rightmost columns in PDF)
                # From text sample: dates[0] a dates[1] = ref periodo aq,
                # dates[2] a dates[3] = deverá gozar, dates[4] = prazo final
                # Confirmed from PDF document column order.
                if len(date_groups) >= 2:
                    periodo = f"{date_groups[0][0]} a {date_groups[0][1]}"
                elif len(all_dates_in_order) >= 2:
                    periodo = f"{all_dates_in_order[0]} a {all_dates_in_order[1]}"
                else:
                    periodo = ""

                records.append({
                    "nome":             nome,
                    "periodo_aquisitivo": periodo,
                    "prazo_final":      prazo_final,
                    "prazo_date":       parse_date(prazo_final),
                })

    return records

def filter_records(records, today=None, months=3):
    """Keep only records whose prazo_final falls within [today, today+3months]."""
    if today is None:
        today = date.today()
    # Approx 3 months ahead
    limit = date(today.year + (today.month + months - 1) // 12,
                 (today.month + months - 1) % 12 + 1,
                 today.day)
    filtered = []
    for r in records:
        d = r["prazo_date"]
        if d and today <= d <= limit:
            filtered.append(r)
    return filtered, limit

# ─────────────────────────────────────────────
# DOCX GENERATION
# ─────────────────────────────────────────────

def copy_cell_format(src_cell, dst_cell):
    """Copy XML properties from one cell to another."""
    src_tc = src_cell._tc
    dst_tc = dst_cell._tc
    # Copy tcPr (cell properties: borders, shading, width)
    src_tcPr = src_tc.find(qn('w:tcPr'))
    if src_tcPr is not None:
        dst_tcPr = dst_tc.find(qn('w:tcPr'))
        if dst_tcPr is not None:
            dst_tc.remove(dst_tcPr)
        dst_tc.insert(0, copy.deepcopy(src_tcPr))

def copy_run_format(src_run, dst_run):
    """Copy run properties from source to destination run."""
    src_rPr = src_run._r.find(qn('w:rPr'))
    dst_rPr = dst_run._r.find(qn('w:rPr'))
    if src_rPr is not None:
        if dst_rPr is not None:
            dst_run._r.remove(dst_rPr)
        dst_run._r.insert(0, copy.deepcopy(src_rPr))

def set_cell_text(cell, text, bold=False, center=False):
    """Clear cell content and set new text, preserving existing style."""
    # Clear existing paragraphs
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    para = cell.paragraphs[0]
    para.clear()
    if center:
        para.alignment = 1  # CENTER
    run = para.add_run(text)
    run.bold = bold

def add_table_row(table, template_row, nome, periodo, data_limite):
    """Add a new data row to the table, cloned from template_row styling."""
    from docx.oxml import OxmlElement

    # Create a new row by deep-copying the template row XML
    new_tr = copy.deepcopy(template_row._tr)
    table._tbl.append(new_tr)

    # Access cells in the new row
    # We need to get cells from the newly appended row
    new_row = table.rows[-1]
    cells = new_row.cells

    if len(cells) >= 3:
        # Clear and set text
        for cell in cells:
            for p in cell.paragraphs[1:]:
                p._element.getparent().remove(p._element)
            cell.paragraphs[0].clear()

        # Set cell texts
        cells[0].paragraphs[0].add_run(nome)
        cells[1].paragraphs[0].add_run(periodo)
        cells[2].paragraphs[0].add_run(data_limite)

def generate_docx(records, today, limit_date):
    """
    Open the template, clear existing data rows (keep header),
    insert new rows from records, return bytes.
    """
    doc = Document(str(TEMPLATE_PATH))
    table = doc.tables[0]

    # Row 0 = header; rows 1..N = data rows to remove
    # Remove existing data rows (keep only header row)
    while len(table.rows) > 1:
        tr = table.rows[-1]._tr
        tr.getparent().remove(tr)

    # Template data row format (clone from header style adapted)
    # We'll build rows manually matching the original style
    header_row = table.rows[0]

    for rec in records:
        add_table_row(
            table,
            header_row,
            rec["nome"],
            rec["periodo_aquisitivo"],
            rec["prazo_final"],
        )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────
# PDF CONVERSION via LibreOffice
# ─────────────────────────────────────────────

def docx_to_pdf(docx_bytes):
    """Convert DOCX bytes to PDF bytes using docx2pdf."""
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "report.docx")
        pdf_path  = os.path.join(tmpdir, "report.pdf")

        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
        except Exception as e:
            raise RuntimeError(f"Erro na conversão para PDF: {e}")

        if not os.path.exists(pdf_path):
            raise RuntimeError("Arquivo PDF não foi gerado.")

        with open(pdf_path, "rb") as f:
            return f.read()

# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    session.pop("docx_path", None)
    session.pop("pdf_path", None)
    today_str = date.today().strftime("%d/%m/%Y")
    return render_template_string(HTML, today=today_str, result=None)

@app.route("/", methods=["POST"])
def process():
    today_str = date.today().strftime("%d/%m/%Y")

    if "pdf_file" not in request.files or request.files["pdf_file"].filename == "":
        flash("Nenhum arquivo enviado. Por favor, selecione o PDF do relatório de férias.", "error")
        return render_template_string(HTML, today=today_str, result=None)

    file = request.files["pdf_file"]
    pdf_bytes = file.read()

    try:
        # 1. Parse PDF
        records = extract_records(pdf_bytes)
    except Exception as e:
        flash(f"Erro ao ler o PDF: {e}", "error")
        return render_template_string(HTML, today=today_str, result=None)

    if not records:
        flash("Nenhum dado válido encontrado no PDF. Verifique se o arquivo é o 'Relatório de Férias' correto.", "warning")
        return render_template_string(HTML, today=today_str, result=None)

    # 2. Filter by next 3 months
    today = date.today()
    filtered, limit_date = filter_records(records, today=today, months=3)

    if not filtered:
        limit_str = limit_date.strftime("%d/%m/%Y")
        flash(
            f"Nenhum colaborador possui prazo de férias vencendo nos próximos 3 meses "
            f"(até {limit_str}). Nenhum alerta gerado.",
            "warning"
        )
        return render_template_string(HTML, today=today_str, result=None)

    # 3. Generate DOCX
    try:
        docx_bytes = generate_docx(filtered, today, limit_date)
    except Exception as e:
        flash(f"Erro ao gerar o documento Word: {e}", "error")
        return render_template_string(HTML, today=today_str, result=None)

    # 4. Convert to PDF
    try:
        pdf_out_bytes = docx_to_pdf(docx_bytes)
    except Exception as e:
        flash(f"Erro ao converter para PDF: {e}", "error")
        return render_template_string(HTML, today=today_str, result=None)

    # 5. Save outputs
    month_str  = today.strftime("%m%Y")
    docx_fname = f"Alerta_Ferias_{month_str}.docx"
    pdf_fname  = f"Alerta_Ferias_{month_str}.pdf"
    docx_path  = OUTPUT_DIR / docx_fname
    pdf_path   = OUTPUT_DIR / pdf_fname

    with open(docx_path, "wb") as f:
        f.write(docx_bytes)
    with open(pdf_path, "wb") as f:
        f.write(pdf_out_bytes)

    session["docx_path"] = str(docx_path)
    session["pdf_path"]  = str(pdf_path)

    # Prepare preview data
    urgent_threshold = today + timedelta(days=45)
    rows_preview = []
    for r in filtered:
        rows_preview.append({
            "nome":       r["nome"],
            "periodo":    r["periodo_aquisitivo"],
            "data_limite": r["prazo_final"],
            "urgente":    r["prazo_date"] and r["prazo_date"] <= urgent_threshold,
        })

    result = {
        "count":      len(filtered),
        "rows":       rows_preview,
        "limit_date": limit_date.strftime("%d/%m/%Y"),
    }

    return render_template_string(HTML, today=today_str, result=result)

@app.route("/download/<fmt>")
def download(fmt):
    if fmt == "docx":
        path = session.get("docx_path")
        mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        fname = "Alerta_Ferias.docx"
    elif fmt == "pdf":
        path = session.get("pdf_path")
        mimetype = "application/pdf"
        fname = "Alerta_Ferias.pdf"
    else:
        return "Formato inválido", 400

    if not path or not os.path.exists(path):
        flash("Arquivo não encontrado. Por favor, processe um PDF primeiro.", "error")
        return redirect(url_for("index"))

    return send_file(path, mimetype=mimetype, as_attachment=True, download_name=fname)

if __name__ == "__main__":
    print("=" * 55)
    print("  Alertas de Férias – BWA Global")
    print("  Acesse: http://localhost:5000")
    print("=" * 55)
    app.run(debug=False, host="0.0.0.0", port=5000)

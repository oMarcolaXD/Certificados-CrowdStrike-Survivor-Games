#!/usr/bin/env python3
"""Certificate generator — pure Python stdlib + Microsoft PowerPoint via AppleScript."""

import cgi
import io
import json
import os
import re
import subprocess
import threading
import unicodedata
import uuid
import zipfile
from http.server import BaseHTTPRequestHandler, HTTPServer, ThreadingHTTPServer
from pathlib import Path

# ── Lock to prevent concurrent PowerPoint operations ─────────────────────────
_pptx_lock = threading.Lock()


# ── Utilities ─────────────────────────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """Remove accents, replace spaces/special chars → safe filename segment."""
    nfd = unicodedata.normalize('NFD', name)
    ascii_only = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    safe = re.sub(r'[^A-Za-z0-9_-]', '_', ascii_only)
    safe = re.sub(r'_+', '_', safe).strip('_')
    return safe


def build_filename(first: str, last: str, colocacao: str) -> str:
    parts = [p for p in (sanitize_name(first), sanitize_name(last)) if p]
    base = '_'.join(parts) if parts else 'Participante'
    if colocacao in ('1', '2', '3'):
        return f'{base}_{colocacao}_Lugar'
    return f'{base}_Certificado'


def xml_escape(s: str) -> str:
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;')
             .replace("'", '&apos;'))


def colocacao_label(c: str) -> str:
    return {'1': '1st Place', '2': '2nd Place', '3': '3rd Place'}.get(c, '')


# ── PPTX placeholder replacement ──────────────────────────────────────────────

_PARA_RE = re.compile(r'<a:p\b[^>]*>.*?</a:p>', re.DOTALL)
_RUN_RE  = re.compile(r'<a:r\b[^>]*>.*?</a:r>', re.DOTALL)
_T_RE    = re.compile(r'(<a:t\b[^>]*>)(.*?)(</a:t>)', re.DOTALL)


def _process_paragraph(para_xml: str, data: dict) -> str:
    """Replace placeholders in a single <a:p> element, handling split text runs."""

    # Combined text of all <a:t> elements in this paragraph
    all_texts = _T_RE.findall(para_xml)
    combined = ''.join(t for _, t, _ in all_texts)

    if '{{' not in combined:
        return para_xml

    # Build merged_text: combined with placeholders replaced (values properly escaped).
    # combined is raw XML text and may already contain entities like &amp; — only
    # escape the substituted values, not the surrounding text.
    merged_text = combined
    changed = False
    for key, value in data.items():
        placeholder = '{{' + key + '}}'
        if placeholder in merged_text:
            merged_text = merged_text.replace(placeholder, xml_escape(value))
            changed = True

    if not changed:
        return para_xml

    # Fast path: placeholder is fully contained in a single <a:t> — substitute in-place
    def sub_t(m):
        open_t, text, close_t = m.groups()
        r = text
        for key, value in data.items():
            r = r.replace('{{' + key + '}}', xml_escape(value))
        return open_t + r + close_t

    result = _T_RE.sub(sub_t, para_xml)

    # Check if placeholders remain (means they were split across multiple runs)
    after_texts = _T_RE.findall(result)
    still_combined = ''.join(t for _, t, _ in after_texts)
    if '{{' not in still_combined:
        return result

    # Split-run case: the placeholder spans multiple <a:r> elements.
    # Keep the FIRST run entirely intact (preserving its <a:rPr> formatting),
    # just replace its <a:t> content. Remove all subsequent runs.
    runs = list(_RUN_RE.finditer(para_xml))
    if not runs:
        return para_xml

    first_run_xml = runs[0].group(0)
    updated = [False]

    def _replace_t(m):
        if updated[0]:
            return ''  # drop extra <a:t> elements inside the first run (rare)
        updated[0] = True
        open_t, _, close_t = m.groups()
        return open_t + merged_text + close_t

    new_first_run = _T_RE.sub(_replace_t, first_run_xml)

    if not updated[0]:
        # First run has no <a:t> at all — insert one before </a:r>
        close_pos = first_run_xml.rfind('</a:r>')
        if close_pos != -1:
            new_first_run = first_run_xml[:close_pos] + f'<a:t>{merged_text}</a:t></a:r>'

    # Rebuild paragraph: before first run + updated first run + after last run
    return para_xml[:runs[0].start()] + new_first_run + para_xml[runs[-1].end():]


_SP_RE_FULL  = re.compile(r'(<p:sp\b.*?</p:sp>)', re.DOTALL)
_SZ_ATTR_RE  = re.compile(r'(\bsz=")(\d+)(")')
_EXT_CX_RE   = re.compile(r'<a:ext\s+cx="(\d+)"')


def _scale_font_to_fit(xml: str) -> str:
    """For large-font single-line boxes (title), scale sz down if text is long."""
    def _fix_sp(m):
        block = m.group(0)
        texts = re.findall(r'<a:t[^>]*>(.*?)</a:t>', block, re.DOTALL)
        combined = ''.join(texts)
        if not combined.strip():
            return block
        sizes = [int(s) for s in _SZ_ATTR_RE.findall(block) and re.findall(r'\bsz="(\d+)"', block)]
        if not sizes or max(sizes) < 3600:
            return block
        cx_m = _EXT_CX_RE.search(block)
        if not cx_m:
            return block
        cx_pt = int(cx_m.group(1)) / 12700  # EMU → points
        orig_sz = max(sizes)
        orig_pt = orig_sz / 100
        # Approximate chars that fit in one line at orig_pt (0.6 × pt = avg char width in pts)
        max_chars = cx_pt / (orig_pt * 0.58)
        if len(combined) <= max_chars:
            return block
        # Scale font so text fits in one line
        new_pt = cx_pt / (len(combined) * 0.58)
        new_pt = max(new_pt, 18)  # floor at 18pt
        new_sz = int(new_pt * 100)
        def _replace_sz(sm):
            if int(sm.group(2)) >= 3600:
                return sm.group(1) + str(new_sz) + sm.group(3)
            return sm.group(0)
        return _SZ_ATTR_RE.sub(_replace_sz, block)
    return _SP_RE_FULL.sub(_fix_sp, xml)


def fill_pptx(template_bytes: bytes, participant: dict) -> bytes:
    """Fill all placeholders in a PPTX buffer, return modified PPTX bytes."""
    data = {
        'NOME_COMPLETO': participant.get('nomeCompleto', ''),
        'PRIMEIRO_NOME': participant.get('primeiroNome', ''),
        'SOBRENOME':     participant.get('sobrenome', ''),
        'EVENTO':        participant.get('evento', ''),
        'COLOCACAO':     colocacao_label(participant.get('colocacao', '')),
        'DATA':          participant.get('data', ''),
    }

    in_zip  = zipfile.ZipFile(io.BytesIO(template_bytes))
    out_buf = io.BytesIO()
    out_zip = zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED)

    slide_re = re.compile(r'^ppt/slides/slide\d+\.xml$')

    for item in in_zip.infolist():
        raw = in_zip.read(item.filename)
        if slide_re.match(item.filename):
            try:
                xml = raw.decode('utf-8')
                xml = _PARA_RE.sub(lambda m: _process_paragraph(m.group(0), data), xml)
                xml = _scale_font_to_fit(xml)  # must run AFTER placeholder substitution
                raw = xml.encode('utf-8')
            except Exception:
                pass  # keep original on any error
        out_zip.writestr(item, raw)

    out_zip.close()
    return out_buf.getvalue()


def find_placeholders(template_bytes: bytes) -> list[str]:
    """Return list of placeholder keys found in the PPTX slides."""
    z = zipfile.ZipFile(io.BytesIO(template_bytes))
    found = set()
    slide_re = re.compile(r'^ppt/slides/slide\d+\.xml$')
    ph_re = re.compile(r'\{\{([A-Z_]+)\}\}')
    for item in z.infolist():
        if slide_re.match(item.filename):
            try:
                xml = z.read(item.filename).decode('utf-8')
                # Also check combined paragraph text to catch split tags
                for para in _PARA_RE.findall(xml):
                    texts = _T_RE.findall(para)
                    combined = ''.join(t for _, t, _ in texts)
                    for k in ph_re.findall(combined):
                        found.add(k)
            except Exception:
                pass
    return sorted(found)


# ── PDF conversion via Microsoft PowerPoint (AppleScript) ─────────────────────

def convert_batch_to_pdf(pptx_paths: list, output_dir: str) -> list:
    """Convert a list of PPTX files to PDF in a single PowerPoint session.
    Returns list of PDF paths in the same order as pptx_paths."""
    if not pptx_paths:
        return []

    out_abs = os.path.abspath(output_dir)

    # Build list of (pptx_abs, pdf_abs) pairs
    pairs = []
    for p in pptx_paths:
        pptx_abs = os.path.abspath(p)
        stem = os.path.splitext(os.path.basename(p))[0]
        pdf_abs = os.path.join(out_abs, f'{stem}.pdf')
        pairs.append((pptx_abs, pdf_abs))

    # Build a single AppleScript that processes the entire batch sequentially.
    # PowerPoint is kept hidden throughout to avoid interfering with the user.
    conversions = '\n'.join(
        f'    open POSIX file "{pptx}"\n'
        f'    delay 2\n'
        f'    save active presentation in POSIX file "{pdf}" as save as PDF\n'
        f'    close active presentation saving no\n'
        f'    delay 0.5'
        for pptx, pdf in pairs
    )

    script = f'''tell application "Microsoft PowerPoint"
    if not running then
        launch
        delay 3
    end if
end tell
tell application "System Events"
    set visible of process "Microsoft PowerPoint" to false
end tell
tell application "Microsoft PowerPoint"
{conversions}
end tell
tell application "System Events"
    set visible of process "Microsoft PowerPoint" to false
end tell
'''
    with _pptx_lock:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True, text=True,
            timeout=30 + 10 * len(pairs)
        )

    # Verify all PDFs were created
    missing = [pdf for _, pdf in pairs if not os.path.exists(pdf)]
    if missing:
        raise RuntimeError(
            f'Falha ao converter {len(missing)} arquivo(s) para PDF.\n'
            f'Erro AppleScript: {result.stderr.strip() or "(sem mensagem)"}'
        )

    return [pdf for _, pdf in pairs]


def check_powerpoint():
    """Check if Microsoft PowerPoint is available via AppleScript."""
    result = subprocess.run(
        ['osascript', '-e', 'tell application "Microsoft PowerPoint" to version'],
        capture_output=True, text=True, timeout=15
    )
    if result.returncode == 0 and result.stdout.strip():
        return True, result.stdout.strip()
    return False, result.stderr.strip()


# Cache populated once at startup in a background thread
_ppt_status = {'ok': False, 'version': '', 'ready': False}

def _probe_powerpoint():
    ok, version = check_powerpoint()
    _ppt_status['ok'] = ok
    _ppt_status['version'] = version
    _ppt_status['ready'] = True
    print(f'PowerPoint: {"ok " + version if ok else "não encontrado"}')
    if ok:
        # Pre-warm PowerPoint so the first generation is faster
        subprocess.run(
            ['osascript', '-e',
             'tell application "Microsoft PowerPoint"\n'
             '  if not running then launch\n'
             '  delay 2\n'
             'end tell\n'
             'tell application "System Events"\n'
             '  set visible of process "Microsoft PowerPoint" to false\n'
             'end tell'],
            capture_output=True, text=True, timeout=20
        )


# ── Certificate generation pipeline ───────────────────────────────────────────

def derive_names(nome_completo: str) -> tuple[str, str]:
    parts = nome_completo.strip().split()
    return parts[0] if parts else '', ' '.join(parts[1:])


def generate_certificates(colocados_bytes, participantes_bytes, participants):
    """Run full pipeline: fill PPTX → convert to PDF → zip. Returns zip bytes."""
    # Fixed paths so macOS only asks PowerPoint permission once, ever.
    tmp      = '/tmp/cert_work'
    pptx_dir = os.path.join(tmp, 'pptx')
    pdf_dir  = os.path.join(tmp, 'pdf')

    # Clean previous run then recreate
    import shutil
    if os.path.exists(tmp):
        shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(pptx_dir)
    os.makedirs(pdf_dir)

    try:
        pptx_paths = []
        stems = []

        # Step 1: fill all PPTX templates
        for p in participants:
            nome = p.get('nomeCompleto', '').strip()
            if not nome:
                continue

            primeiro, sobrenome = derive_names(nome)
            colocacao = str(p.get('colocacao', '')).strip()

            participant = {
                'nomeCompleto': nome,
                'primeiroNome': primeiro,
                'sobrenome':    sobrenome,
                'evento':       p.get('evento', ''),
                'colocacao':    colocacao,
                'data':         p.get('data', ''),
            }

            if colocacao in ('1', '2', '3'):
                if colocados_bytes is None:
                    raise ValueError(
                        f'Participante "{nome}" tem colocação {colocacao} mas o '
                        'modelo_colocados.pptx não foi enviado.'
                    )
                template = colocados_bytes
            else:
                if participantes_bytes is None:
                    raise ValueError(
                        f'Participante "{nome}" não tem colocação mas o '
                        'modelo_participantes.pptx não foi enviado.'
                    )
                template = participantes_bytes

            stem = build_filename(primeiro, sobrenome, colocacao)
            pptx_path = os.path.join(pptx_dir, f'{stem}.pptx')

            filled = fill_pptx(template, participant)
            with open(pptx_path, 'wb') as f:
                f.write(filled)

            pptx_paths.append(pptx_path)
            stems.append(stem)

        # Step 2: convert entire batch to PDF in a single PowerPoint session
        pdf_paths = convert_batch_to_pdf(pptx_paths, pdf_dir)

        # Step 3: build zip
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pdf_path, stem in zip(pdf_paths, stems):
                zf.write(pdf_path, f'{stem}.pdf')

        return zip_buf.getvalue()

    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


# ── HTTP Server ───────────────────────────────────────────────────────────────

STATIC_DIR = Path(__file__).parent / 'static'


class CertificateHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f'[{self.address_string()}] {format % args}')

    def send_json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def send_error_json(self, code: int, message: str):
        self.send_json(code, {'ok': False, 'error': message})

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        from urllib.parse import unquote
        path = unquote(self.path.split('?')[0])
        if path == '/' or path == '':
            path = '/index.html'

        file_path = STATIC_DIR / path.lstrip('/')
        if file_path.is_file():
            content = file_path.read_bytes()
            ext = path.rsplit('.', 1)[-1].lower() if '.' in path else ''
            ct = {'html': 'text/html', 'svg': 'image/svg+xml', 'png': 'image/png',
                  'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'css': 'text/css',
                  'js': 'application/javascript'}.get(ext, 'application/octet-stream')
            self.send_response(200)
            self.send_header('Content-Type', ct)
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        else:
            self.send_error_json(404, 'Not found')

    def do_POST(self):
        if self.path == '/api/validate':
            self._handle_validate()
        elif self.path == '/api/generate':
            self._handle_generate()
        else:
            self.send_error_json(404, 'Not found')

    def _parse_form(self):
        return cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': self.headers.get('Content-Type', ''),
                'CONTENT_LENGTH': self.headers.get('Content-Length', '0'),
            }
        )

    def _handle_validate(self):
        try:
            ok      = _ppt_status['ok']
            version = _ppt_status['version']
            ready   = _ppt_status['ready']

            result: dict = {
                'ok': True,
                'powerpointReady': ready,
                'powerpointAvailable': ok,
                'powerpointVersion': version if ok else None,
                'colocadosPlaceholders': [],
                'participantesPlaceholders': [],
                'warnings': [],
            }

            if ready and not ok:
                result['warnings'].append(
                    'Microsoft PowerPoint não encontrado. A conversão para PDF não funcionará.'
                )

            # Only parse uploaded files if there is a multipart body
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', '0') or '0')
            if content_length > 0 and 'multipart' in content_type:
                form = self._parse_form()
                colocados_item    = form['templateColocados']    if 'templateColocados'    in form else None
                participantes_item = form['templateParticipantes'] if 'templateParticipantes' in form else None

                if colocados_item and hasattr(colocados_item, 'file'):
                    result['colocadosPlaceholders'] = find_placeholders(colocados_item.file.read())

                if participantes_item and hasattr(participantes_item, 'file'):
                    result['participantesPlaceholders'] = find_placeholders(participantes_item.file.read())

            self.send_json(200, result)

        except Exception as e:
            self.send_error_json(500, str(e))

    def _handle_generate(self):
        try:
            form = self._parse_form()

            # Read template files
            colocados_bytes = None
            participantes_bytes = None

            if 'templateColocados' in form:
                item = form['templateColocados']
                if hasattr(item, 'file'):
                    colocados_bytes = item.file.read()

            if 'templateParticipantes' in form:
                item = form['templateParticipantes']
                if hasattr(item, 'file'):
                    participantes_bytes = item.file.read()

            if colocados_bytes is None and participantes_bytes is None:
                self.send_error_json(400, 'Nenhum template foi enviado.')
                return

            # Validate file types (check PPTX magic: PK zip header)
            for name, data in [('modelo_colocados', colocados_bytes),
                                ('modelo_participantes', participantes_bytes)]:
                if data and not data.startswith(b'PK'):
                    self.send_error_json(400, f'{name}.pptx não parece ser um arquivo válido.')
                    return

            # Parse participants
            participants_raw = form.getvalue('participants', '[]')
            try:
                participants = json.loads(participants_raw)
            except json.JSONDecodeError:
                self.send_error_json(400, 'Lista de participantes inválida (JSON malformado).')
                return

            if not participants:
                self.send_error_json(400, 'Lista de participantes está vazia.')
                return

            # Validate colocacao values
            for p in participants:
                c = str(p.get('colocacao', '')).strip()
                if c not in ('', '1', '2', '3'):
                    self.send_error_json(
                        400,
                        f'Colocação inválida "{c}" para "{p.get("nomeCompleto", "?")}". '
                        'Use 1, 2, 3 ou vazio.'
                    )
                    return

            # Check PowerPoint (use startup cache — avoids re-launching PPT)
            if not _ppt_status['ok']:
                self.send_error_json(
                    503,
                    'Microsoft PowerPoint não encontrado. '
                    'Instale o Office para macOS para habilitar a conversão para PDF.'
                )
                return

            zip_bytes = generate_certificates(colocados_bytes, participantes_bytes, participants)

            self.send_response(200)
            self.send_header('Content-Type', 'application/zip')
            self.send_header('Content-Disposition', 'attachment; filename="certificados.zip"')
            self.send_header('Content-Length', str(len(zip_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(zip_bytes)

        except ValueError as e:
            self.send_error_json(400, str(e))
        except RuntimeError as e:
            self.send_error_json(500, str(e))
        except Exception as e:
            self.send_error_json(500, f'Erro inesperado: {e}')


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    port = int(os.environ.get('PORT', 8080))
    # Probe PowerPoint in background so the server starts instantly
    threading.Thread(target=_probe_powerpoint, daemon=True).start()
    server = ThreadingHTTPServer(('localhost', port), CertificateHandler)
    print(f'Servidor rodando em http://localhost:{port}')
    print('Pressione Ctrl+C para parar.')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServidor encerrado.')


if __name__ == '__main__':
    main()

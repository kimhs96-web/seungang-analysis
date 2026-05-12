"""
Vercel Python Serverless - 승강시설 장애분석 API
2개 파일(장애신고 + 장애분류) 업로드 처리
"""
import sys, os, io, json, zipfile, traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from http.server import BaseHTTPRequestHandler

class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        path = self.path.split('?')[0]
        if path == '/api/health':
            # 패키지 import 가능 여부 진단
            diag = {}
            try:
                import openpyxl
                diag['openpyxl'] = openpyxl.__version__
            except Exception as e:
                diag['openpyxl'] = f'ERROR: {e}'
            try:
                import pptx
                diag['python-pptx'] = pptx.__version__
            except Exception as e:
                diag['python-pptx'] = f'ERROR: {e}'
            try:
                from processor import run_pipeline
                diag['processor'] = 'OK'
            except Exception as e:
                diag['processor'] = f'ERROR: {e}'
            try:
                from el_report import build_el_report
                diag['el_report'] = 'OK'
            except Exception as e:
                diag['el_report'] = f'ERROR: {e}'
            try:
                from es_report import build_es_report
                diag['es_report'] = 'OK'
            except Exception as e:
                diag['es_report'] = f'ERROR: {e}'

            self._json(200, {'status': 'ok', 'version': '3.1', 'diag': diag})
        else:
            self._json(404, {'error': 'not found'})

    def do_POST(self):
        path = self.path.split('?')[0]
        if path == '/api/process':
            self._handle()
        else:
            self._json(404, {'error': 'not found'})

    def _handle(self):
        try:
            ct = self.headers.get('Content-Type', '')
            cl = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(cl)

            files = self._parse_multipart(body, ct)
            fault_bytes = files.get('fault')
            cls_bytes   = files.get('cls')

            if not fault_bytes:
                self._json(400, {'error': '장애신고 파일이 없습니다'}); return
            if not cls_bytes:
                self._json(400, {'error': '장애분류 파일이 없습니다'}); return

            from processor import run_pipeline
            from el_report  import build_el_report
            from es_report  import build_es_report

            excel_bytes, stats = run_pipeline(fault_bytes, cls_bytes)
            el_bytes = build_el_report(stats)
            es_bytes = build_es_report(stats)

            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.writestr('장애신고_처리완료.xlsx',  excel_bytes)
                zf.writestr('EL_장애분석_보고서.pptx', el_bytes)
                zf.writestr('ES_장애분석_보고서.pptx', es_bytes)
            buf.seek(0)
            zip_bytes = buf.read()

            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/zip')
            self.send_header('Content-Disposition', 'attachment; filename="result.zip"')
            self.send_header('Content-Length', str(len(zip_bytes)))
            self.send_header('X-Stats', json.dumps(stats, ensure_ascii=False))
            self.end_headers()
            self.wfile.write(zip_bytes)

        except Exception as e:
            tb = traceback.format_exc()
            print(f"[ERROR]\n{tb}", flush=True)
            self._json(500, {'error': str(e), 'detail': tb})

    def _parse_multipart(self, body, ct):
        files = {}
        boundary = None
        for part in ct.split(';'):
            part = part.strip()
            if part.startswith('boundary='):
                boundary = part[9:].strip('"\'')
                break
        if not boundary:
            return files
        delimiter = ('--' + boundary).encode()
        parts = body.split(delimiter)
        for part in parts[1:]:
            if b'\r\n\r\n' not in part:
                continue
            header_raw, data = part.split(b'\r\n\r\n', 1)
            data = data.rstrip(b'\r\n--')
            header_str = header_raw.decode('utf-8', errors='replace')
            name = None
            for line in header_str.split('\r\n'):
                if 'Content-Disposition' in line:
                    for seg in line.split(';'):
                        seg = seg.strip()
                        if seg.startswith('name='):
                            name = seg[5:].strip('"\'')
            if name:
                files[name] = data
        return files

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self._cors()
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_response(self, code, message=None):
        super().send_response(code, message)

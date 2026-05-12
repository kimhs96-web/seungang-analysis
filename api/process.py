"""
Vercel Python Serverless Function
단일 엔드포인트로 업로드 → 처리 → ZIP 반환 (동기 처리)
"""
import os, sys, uuid, zipfile, io, json, traceback, base64

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from http.server import BaseHTTPRequestHandler
import cgi

class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors()
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        self._cors()
        path = self.path.split('?')[0]

        if path == '/api/health':
            self._json(200, {
                'status': 'ok',
                'runtime': 'vercel-python',
                'python': sys.version.split()[0],
            })
        else:
            self._json(404, {'error': 'Not found'})

    def do_POST(self):
        self._cors()
        path = self.path.split('?')[0]

        if path == '/api/process':
            self._handle_process()
        else:
            self._json(404, {'error': 'Not found'})

    def _handle_process(self):
        try:
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            # multipart/form-data 파싱
            if 'multipart/form-data' in content_type:
                # boundary 추출
                boundary = None
                for part in content_type.split(';'):
                    part = part.strip()
                    if part.startswith('boundary='):
                        boundary = part[9:].strip('"')
                        break

                if not boundary:
                    self._json(400, {'error': 'boundary 없음'}); return

                file_bytes = self._parse_multipart(body, boundary.encode())
            else:
                # raw body (base64 JSON)
                try:
                    data = json.loads(body)
                    file_bytes = base64.b64decode(data.get('file', ''))
                except:
                    self._json(400, {'error': '잘못된 요청 형식'}); return

            if not file_bytes:
                self._json(400, {'error': '파일이 없습니다'}); return

            # 처리
            from processor import run_pipeline
            from el_report import build_el_report
            from es_report import build_es_report

            excel_bytes, stats = run_pipeline(file_bytes)
            el_bytes  = build_el_report(stats)
            es_bytes  = build_es_report(stats)

            # ZIP 생성
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.writestr('장애신고_처리완료.xlsx', excel_bytes)
                zf.writestr('EL_장애분석_보고서.pptx', el_bytes)
                zf.writestr('ES_장애분석_보고서.pptx', es_bytes)
            zip_buf.seek(0)
            zip_bytes = zip_buf.read()

            # 응답
            self.send_response(200)
            self.send_header('Content-Type', 'application/zip')
            self.send_header('Content-Disposition', 'attachment; filename="승강기_장애분석_결과.zip"')
            self.send_header('Content-Length', str(len(zip_bytes)))
            self.send_header('X-Stats', json.dumps(stats, ensure_ascii=False))
            self.end_headers()
            self.wfile.write(zip_bytes)

        except Exception as e:
            detail = traceback.format_exc()
            print(f"[ERROR] {detail}")
            self._json(500, {'error': str(e), 'detail': detail})

    def _parse_multipart(self, body, boundary):
        """단순 multipart 파서 - 첫 번째 파일 필드 추출"""
        delimiter = b'--' + boundary
        parts = body.split(delimiter)
        for part in parts:
            if b'Content-Disposition' in part and b'filename' in part:
                # 헤더와 바디 분리
                if b'\r\n\r\n' in part:
                    _, file_data = part.split(b'\r\n\r\n', 1)
                    # 끝의 \r\n 제거
                    return file_data.rstrip(b'\r\n--')
        return None

    def _cors(self):
        pass  # end_headers 전에 호출되어야 하므로 별도 처리

    def _json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(body)

    def send_response(self, code, message=None):
        super().send_response(code, message)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

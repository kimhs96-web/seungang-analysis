"""
Vercel Python Serverless - Flask 기반
파일 전송: base64 JSON 방식 (multipart 대신)
"""
import sys, os, io, json, zipfile, traceback, base64
from flask import Flask, request, jsonify, Response

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__)

@app.after_request
def cors(resp):
    resp.headers['Access-Control-Allow-Origin']  = '*'
    resp.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp

@app.route('/api/<path:p>', methods=['OPTIONS'])
@app.route('/', methods=['OPTIONS'])
def options(p=''):
    return '', 204

@app.route('/api/health')
def health():
    diag = {}
    for name, mod in [('openpyxl','openpyxl'),('python-pptx','pptx')]:
        try:
            m = __import__(mod)
            diag[name] = getattr(m, '__version__', 'OK')
        except Exception as e:
            diag[name] = f'ERROR: {e}'
    for name in ['processor','el_report','es_report']:
        try:
            __import__(name)
            diag[name] = 'OK'
        except Exception as e:
            diag[name] = f'ERROR: {e}'
    return jsonify({'status': 'ok', 'version': '3.3', 'diag': diag})

@app.route('/api/process', methods=['POST'])
def process():
    try:
        # base64 JSON 방식으로 수신
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({'error': '요청 데이터가 없습니다'}), 400

        fault_b64 = data.get('fault')
        cls_b64   = data.get('cls')

        if not fault_b64:
            return jsonify({'error': '장애신고 파일이 없습니다'}), 400
        if not cls_b64:
            return jsonify({'error': '장애분류 파일이 없습니다'}), 400

        fault_bytes = base64.b64decode(fault_b64)
        cls_bytes   = base64.b64decode(cls_b64)

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

        # ZIP을 base64로 반환
        zip_b64 = base64.b64encode(zip_bytes).decode('utf-8')
        return jsonify({
            'zip':   zip_b64,
            'stats': stats
        })

    except Exception as e:
        tb = traceback.format_exc()
        print(f"[ERROR]\n{tb}", flush=True)
        return jsonify({'error': str(e), 'detail': tb}), 500

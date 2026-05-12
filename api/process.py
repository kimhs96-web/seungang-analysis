"""
Vercel Python Serverless - Flask 기반
"""
import sys, os, io, json, zipfile, traceback
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
    return jsonify({'status': 'ok', 'version': '3.2', 'diag': diag})

@app.route('/api/process', methods=['POST'])
def process():
    try:
        fault_file = request.files.get('fault')
        cls_file   = request.files.get('cls')

        if not fault_file:
            return jsonify({'error': '장애신고 파일이 없습니다'}), 400
        if not cls_file:
            return jsonify({'error': '장애분류 파일이 없습니다'}), 400

        fault_bytes = fault_file.read()
        cls_bytes   = cls_file.read()

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

        resp = Response(zip_bytes, status=200, mimetype='application/zip')
        resp.headers['Content-Disposition'] = 'attachment; filename="result.zip"'
        resp.headers['Content-Length']      = str(len(zip_bytes))
        resp.headers['X-Stats']             = json.dumps(stats, ensure_ascii=False)
        return resp

    except Exception as e:
        tb = traceback.format_exc()
        print(f"[ERROR]\n{tb}", flush=True)
        return jsonify({'error': str(e), 'detail': tb}), 500

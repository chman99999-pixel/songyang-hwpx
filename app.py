#!/usr/bin/env python3
"""송영서비스 제공현황 HWPX 웹앱"""
import os
import shutil
import tempfile
from flask import Flask, request, send_file, render_template_string
from hwpx_replace import parse_excel, parse_trip_info, generate
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>송영서비스 HWPX 생성기</title>
<style>
  body { font-family: 'Malgun Gothic', sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
  .container { max-width: 900px; margin: 0 auto; background: #fff; border-radius: 12px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
  h1 { margin: 0 0 8px; font-size: 22px; color: #333; }
  p.sub { margin: 0 0 24px; color: #888; font-size: 13px; }
  .row { display: flex; gap: 12px; align-items: center; margin-bottom: 16px; }
  .row label { min-width: 180px; font-weight: bold; font-size: 14px; }
  .row input { flex: 1; }
  .btn { display: block; width: 100%; padding: 14px; background: #1976d2; color: #fff; border: none; border-radius: 8px; font-size: 16px; cursor: pointer; margin-top: 20px; }
  .btn:hover { background: #1565c0; }
  .msg { margin-top: 16px; padding: 12px; border-radius: 8px; font-size: 14px; }
  .msg.ok { background: #e8f5e9; color: #2e7d32; }
  .msg.err { background: #ffebee; color: #c62828; }
  .hint { font-size: 11px; color: #999; margin-top: 2px; }
  .result { margin-top: 24px; }
  .result h2 { font-size: 17px; margin: 0 0 4px; color: #333; }
  .result .info { font-size: 13px; color: #666; margin-bottom: 12px; }
  .dl-btn { display: inline-block; padding: 10px 24px; background: #1976d2; color: #fff; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold; margin-bottom: 16px; }
  .dl-btn:hover { background: #1565c0; }
  table.data { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 8px; }
  table.data th { background: #f0f0f0; padding: 8px 6px; border: 1px solid #ddd; text-align: center; font-weight: bold; }
  table.data td { padding: 7px 6px; border: 1px solid #ddd; text-align: center; }
  table.data tr:hover { background: #f9f9f9; }
  .warn-row { background: #fff3e0 !important; }
  .warn-list { margin: 12px 0; padding: 10px 16px; background: #fff3e0; border-radius: 6px; font-size: 13px; color: #e65100; }
</style>
</head>
<body>
<div class="container">
  <h1>송영서비스 HWPX 생성기</h1>
  <p class="sub">원본 한글 파일, 엑셀 바우처 데이터, 송영정보를 업로드하면<br>월/날짜/시간/비용이 자동 교체된 한글 문서를 생성합니다.</p>
  <form method="POST" enctype="multipart/form-data" action="/generate">
    <div class="row">
      <label>① 원본 HWPX (.hwpx)</label>
      <input type="file" name="hwpx" accept=".hwpx" required>
    </div>
    <div class="row">
      <label>② 송영서비스 엑셀 (.xlsx)</label>
      <input type="file" name="excel" accept=".xlsx,.xls" required>
    </div>
    <div class="row">
      <label>③ 송영정보 엑셀 (.xlsx)</label>
      <input type="file" name="trip_info" accept=".xlsx,.xls" required>
    </div>
    <p class="hint">송영정보: 이용자별 오전/오후 송영 시간이 기재된 파일</p>
    <button type="submit" class="btn">HWPX 문서 생성</button>
  </form>

  {% if message %}
  <div class="msg {{ msg_type }}">{{ message }}</div>
  {% endif %}

  {% if result %}
  <div class="result">
    <h2>{{ result.month }}월 송영서비스 처리 결과</h2>
    <div class="info">총 {{ result.total }}명 처리 완료</div>
    <a class="dl-btn" href="/download_warn?f={{ result.filename }}">문서 다운로드 ({{ result.filename }})</a>

    {% if result.warnings %}
    <div class="warn-list">
      {% for w in result.warnings %}{{ w }}<br>{% endfor %}
    </div>
    {% endif %}

    <table class="data">
      <thead>
        <tr>
          <th>No</th>
          <th>이용자명</th>
          <th>그룹</th>
          <th>오전</th>
          <th>오후</th>
          <th>일수</th>
          <th>날짜범위</th>
          <th>총시간</th>
          <th>단가</th>
          <th>비용</th>
        </tr>
      </thead>
      <tbody>
        {% for u in result.users %}
        <tr>
          <td>{{ loop.index }}</td>
          <td>{{ u.name }}</td>
          <td>{{ u.group }}</td>
          <td>{{ u.am or '-' }}</td>
          <td>{{ u.pm or '-' }}</td>
          <td>{{ u.days }}일</td>
          <td>{{ u.date_range }}</td>
          <td>{{ u.total_hours }}시간</td>
          <td>{{ u.unit_price }}</td>
          <td>{{ u.cost }}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
  {% endif %}
</div>
</body>
</html>'''


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/generate', methods=['POST'])
def gen():
    hwpx_file = request.files.get('hwpx')
    excel_file = request.files.get('excel')
    trip_file = request.files.get('trip_info')

    if not hwpx_file or not excel_file or not trip_file:
        return render_template_string(HTML_TEMPLATE, message='세 파일 모두 업로드해 주세요.', msg_type='err')

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            hwpx_path = os.path.join(tmpdir, 'input.hwpx')
            excel_path = os.path.join(tmpdir, 'input.xlsx')
            trip_path = os.path.join(tmpdir, 'trip_info.xlsx')
            hwpx_file.save(hwpx_path)
            excel_file.save(excel_path)
            trip_file.save(trip_path)

            trip_info = parse_trip_info(trip_path)
            users_data, month = parse_excel(excel_path, trip_info)
            year = datetime.now().year
            output_name = f'주간활동송영서비스_{year}.{month:02d}.hwpx'
            output_path = os.path.join(tmpdir, output_name)

            warnings = generate(hwpx_path, excel_path, output_path, trip_path) or []

            # 다운로드용 파일 복사
            dl_path = os.path.join(tempfile.gettempdir(), output_name)
            shutil.copy2(output_path, dl_path)

            # 결과 테이블 데이터
            user_rows = []
            for name in sorted(users_data.keys()):
                u = users_data[name]
                user_rows.append({
                    'name': name,
                    'group': u.get('group_str', ''),
                    'am': f"{u['am_min']}분" if u.get('am_min') else None,
                    'pm': f"{u['pm_min']}분" if u.get('pm_min') else None,
                    'days': u.get('day_count', 0),
                    'date_range': u.get('date_range', ''),
                    'total_hours': u.get('total_hours_int', 0),
                    'unit_price': f"{u.get('unit_price', 0):,}원",
                    'cost': f"{u.get('final_cost', 0):,}원",
                })

            result = {
                'month': month,
                'total': len(users_data),
                'filename': output_name,
                'users': user_rows,
                'warnings': warnings,
            }

            return render_template_string(HTML_TEMPLATE, result=result)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, message=f'오류: {str(e)}', msg_type='err')


@app.route('/download_warn')
def download_warn():
    fname = request.args.get('f', '')
    fpath = os.path.join(tempfile.gettempdir(), fname)
    if os.path.exists(fpath):
        return send_file(fpath, as_attachment=True, download_name=fname, mimetype='application/hwp+zip')
    return 'File not found', 404


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)

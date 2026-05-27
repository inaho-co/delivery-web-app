import os
import io
import tempfile
import smtplib
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import core_transfer

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-me')

# ── BOX トークン（インメモリ保持） ──────────────────────────────
_tokens = {
    'access': os.environ.get('BOX_ACCESS_TOKEN', ''),
    'refresh': os.environ.get('BOX_REFRESH_TOKEN', ''),
}
_token_lock = threading.Lock()


def _store_tokens(access_token, refresh_token):
    with _token_lock:
        _tokens['access'] = access_token
        _tokens['refresh'] = refresh_token


def _get_box_client():
    from boxsdk import OAuth2, Client
    with _token_lock:
        a = _tokens['access']
        r = _tokens['refresh']
    oauth = OAuth2(
        client_id=os.environ['BOX_CLIENT_ID'],
        client_secret=os.environ['BOX_CLIENT_SECRET'],
        access_token=a,
        refresh_token=r,
        store_tokens=_store_tokens,
    )
    return Client(oauth)


def _list_excel_files(folder_id):
    client = _get_box_client()
    items = client.folder(folder_id).get_items(limit=200)
    files = [
        {'id': item.id, 'name': item.name}
        for item in items
        if item.type == 'file' and item.name.lower().endswith(('.xlsx', '.xlsm'))
    ]
    return sorted(files, key=lambda x: x['name'])


def _download_to_temp(file_id, suffix='.xlsx'):
    client = _get_box_client()
    content = client.file(file_id).content()
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, 'wb') as f:
        f.write(content)
    return path


def _update_box_file(file_id, file_bytes):
    client = _get_box_client()
    client.file(file_id).update_contents_with_stream(io.BytesIO(file_bytes))


def _send_email(file_bytes, filename, store_name, date_str):
    gmail_user = 'tomo.dream69@gmail.com'
    gmail_password = os.environ['GMAIL_APP_PASSWORD']
    gmail_to = os.environ['GMAIL_TO']

    subject = f'納品書 {date_str} {store_name}'
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = gmail_to
    msg['Subject'] = subject
    msg.attach(MIMEText(f'納品書を送付します。\n\n日付: {date_str}\n店舗: {store_name}', 'plain', 'utf-8'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(gmail_user, gmail_password)
        smtp.send_message(msg)


def _cleanup(*paths):
    for p in paths:
        try:
            if p and os.path.exists(p):
                os.unlink(p)
        except Exception:
            pass


# ── 認証 ───────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        if request.form.get('password') == os.environ.get('APP_PASSWORD', 'inaho'):
            session['logged_in'] = True
            return redirect(url_for('index'))
        error = 'パスワードが違います'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── トップ ─────────────────────────────────────────────────────
@app.route('/')
@login_required
def index():
    import subprocess
    try:
        version = subprocess.check_output(
            ['git', 'rev-parse', '--short', 'HEAD'],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            text=True
        ).strip()
    except Exception:
        version = 'unknown'
    return render_template('index.html', version=version)


# ── 通常モード ─────────────────────────────────────────────────
@app.route('/normal')
@login_required
def normal():
    order_folder_id = os.environ.get('BOX_ORDER_FOLDER_ID', '')
    delivery_folder_id = os.environ.get('BOX_DELIVERY_FOLDER_ID', '')
    preselected_order_id = session.pop('fax_order_file_id', None)
    try:
        order_files = _list_excel_files(order_folder_id) if order_folder_id else []
        delivery_files = _list_excel_files(delivery_folder_id) if delivery_folder_id else []
    except Exception as e:
        return render_template('normal.html', error=f'BOX接続エラー: {e}',
                               order_files=[], delivery_files=[],
                               preselected_order_id=None)
    return render_template('normal.html', order_files=order_files,
                           delivery_files=delivery_files, error=None,
                           preselected_order_id=preselected_order_id)


@app.route('/api/stores', methods=['POST'])
@login_required
def api_stores():
    order_file_id = request.json.get('order_file_id')
    if not order_file_id:
        return jsonify({'error': '発注書が未選択'}), 400
    tmp = None
    try:
        tmp = _download_to_temp(order_file_id, suffix='.xlsx')
        stores = core_transfer.get_store_names(tmp)
        return jsonify({'stores': stores})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        _cleanup(tmp)


@app.route('/normal/create', methods=['POST'])
@login_required
def normal_create():
    order_file_id = request.form.get('order_file_id')
    delivery_file_id = request.form.get('delivery_file_id')
    store_name = request.form.get('store_name', '').strip()
    date_str = request.form.get('date_str', '').replace('/', '').replace('-', '')

    if not all([order_file_id, delivery_file_id, store_name, date_str]):
        return render_template('normal.html', error='入力が不足しています',
                               order_files=[], delivery_files=[])

    order_tmp = delivery_tmp = output_tmp = None
    try:
        order_tmp = _download_to_temp(order_file_id, suffix='.xlsx')
        delivery_tmp = _download_to_temp(delivery_file_id, suffix='.xlsx')

        price_map = core_transfer.build_kubun_price_map(delivery_tmp)
        items = core_transfer.extract_order_items(order_tmp, store_name, price_map)

        if not items:
            return render_template('normal.html',
                                   error=f'「{store_name}」の商品データが見つかりませんでした',
                                   order_files=[], delivery_files=[])

        fd, output_tmp = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)

        sheet_name, wb = core_transfer.add_delivery_sheet(
            delivery_tmp, items, date_str, output_tmp, price_map)
        wb.save(output_tmp)

        with open(output_tmp, 'rb') as f:
            file_bytes = f.read()

        _update_box_file(delivery_file_id, file_bytes)

        client = _get_box_client()
        delivery_filename = client.file(delivery_file_id).get().name
        _send_email(file_bytes, delivery_filename, store_name, date_str)

        return render_template('success.html',
                               filename=delivery_filename,
                               store=store_name,
                               date_str=date_str,
                               item_count=len(items))

    except Exception as e:
        return render_template('normal.html', error=f'エラー: {e}',
                               order_files=[], delivery_files=[])
    finally:
        _cleanup(order_tmp, delivery_tmp, output_tmp)


# ── 手入力モード ───────────────────────────────────────────────
@app.route('/manual')
@login_required
def manual():
    delivery_folder_id = os.environ.get('BOX_DELIVERY_FOLDER_ID', '')
    try:
        delivery_files = _list_excel_files(delivery_folder_id) if delivery_folder_id else []
    except Exception as e:
        return render_template('manual.html', error=f'BOX接続エラー: {e}',
                               delivery_files=[])
    return render_template('manual.html', delivery_files=delivery_files, error=None)


@app.route('/manual/create', methods=['POST'])
@login_required
def manual_create():
    delivery_file_id = request.form.get('delivery_file_id')
    date_str = request.form.get('date_str', '').replace('/', '').replace('-', '')

    kubuns  = request.form.getlist('kubun[]')
    codes   = request.form.getlist('code[]')
    names   = request.form.getlist('name[]')
    hinbans = request.form.getlist('hinban[]')
    qtys    = request.form.getlist('qty[]')
    prices  = request.form.getlist('price[]')

    items = []
    for i in range(len(codes)):
        try:
            qty = int(qtys[i]) if qtys[i] else 0
            price = int(prices[i]) if prices[i] else 0
            if qty > 0 and codes[i]:
                items.append({
                    'kubun':  kubuns[i] if i < len(kubuns) else '',
                    'code':   codes[i],
                    'name':   names[i] if i < len(names) else '',
                    'hinban': hinbans[i] if i < len(hinbans) else '',
                    'qty':    qty,
                    'price':  price,
                    'amount': qty * price,
                })
        except (ValueError, IndexError):
            continue

    if not all([delivery_file_id, date_str]) or not items:
        return render_template('manual.html', error='入力が不足しています（商品は1行以上必要です）',
                               delivery_files=[])

    delivery_tmp = output_tmp = None
    try:
        delivery_tmp = _download_to_temp(delivery_file_id, suffix='.xlsx')
        fd, output_tmp = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)

        sheet_name, wb = core_transfer.add_delivery_sheet(
            delivery_tmp, items, date_str, output_tmp)
        wb.save(output_tmp)

        with open(output_tmp, 'rb') as f:
            file_bytes = f.read()

        _update_box_file(delivery_file_id, file_bytes)

        client = _get_box_client()
        delivery_filename = client.file(delivery_file_id).get().name
        _send_email(file_bytes, delivery_filename, '手入力', date_str)

        return render_template('success.html',
                               filename=delivery_filename,
                               store='手入力',
                               date_str=date_str,
                               item_count=len(items))

    except Exception as e:
        return render_template('manual.html', error=f'エラー: {e}',
                               delivery_files=[])
    finally:
        _cleanup(delivery_tmp, output_tmp)


# ── 倉庫FAX確認 ────────────────────────────────────────────────

def _parse_fax_order(path):
    """発注書Excelの「稲穂」シートを解析して店舗・商品リストを返す。
    tab_fax.py の _fax_load_sheet() 相当。"""
    from openpyxl import load_workbook

    exclude = {'', 'None', '区分', '商品ｺｰﾄﾞ', '商品コード', '品名', '品番',
               'JANコード', 'JAN', '数量', '単価', '金額', '備考',
               'BB', '工具館', '店舗名', '引取'}
    PAGE_MARKS = {'①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'}

    wb = load_workbook(path, data_only=True)
    result_stores = []
    result_items = []

    for ws in wb.worksheets:
        if '稲穂' not in ws.title:
            continue
        data = list(ws.iter_rows(values_only=True))

        store_col_map = {}
        store_row_vals = None
        for ri, row in enumerate(data[:15]):
            row_vals = [str(v or '').strip() for v in row]
            if '店舗名' in row_vals:
                candidates = [v for v in row_vals[4:] if v not in exclude]
                if candidates:
                    store_row_vals = row_vals
                else:
                    if ri + 1 < len(data):
                        store_row_vals = [str(v or '').strip() for v in data[ri + 1]]
                break

        if store_row_vals is None:
            continue

        for ci, v in enumerate(store_row_vals):
            if ci >= 4 and v not in exclude:
                store_col_map[v] = ci

        if not store_col_map:
            continue

        stores = list(store_col_map.keys())
        result_stores = stores
        items_map = {}

        for row in data:
            kubun = str(row[0] or '').strip()
            code = str(row[1] or '').strip()
            name = str(row[2] or '').replace('\n', '').strip()
            hinban = str(row[3] or '').strip()

            if kubun in PAGE_MARKS:
                continue
            if not code or code in ['商品ｺｰﾄﾞ', 'None', '']:
                continue
            if not kubun or kubun in ['区分', '企業名', '店舗名']:
                continue
            if not code.isdigit():
                continue

            if code not in items_map:
                items_map[code] = {
                    'code': code, 'name': name,
                    'kubun': kubun, 'hinban': hinban,
                    'qty_by_store': {}
                }

            for sname, ci in store_col_map.items():
                val = row[ci] if ci < len(row) else None
                try:
                    qty = int(val)
                except (TypeError, ValueError):
                    qty = 0
                existing = items_map[code]['qty_by_store'].get(sname, 0)
                items_map[code]['qty_by_store'][sname] = max(existing, qty)

        result_items = list(items_map.values())
        break  # 最初の「稲穂」シートのみ使用

    wb.close()
    return {'stores': result_stores, 'items': result_items}


def _write_fax_quantities(path, stores, quantities):
    """発注書Excelの数量を quantities で上書きして保存。
    quantities: {store_name: {code: qty}}"""
    from openpyxl import load_workbook

    exclude = {'', 'None', '区分', '商品ｺｰﾄﾞ', '商品コード', '品名', '品番',
               'JANコード', 'JAN', '数量', '単価', '金額', '備考',
               'BB', '工具館', '店舗名', '引取'}
    PAGE_MARKS = {'①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'}

    wb = load_workbook(path)
    for ws in wb.worksheets:
        if '稲穂' not in ws.title:
            continue
        data = list(ws.iter_rows(values_only=False))

        store_col_map = {}
        store_row_vals = None
        for ri, row in enumerate(data[:15]):
            row_vals = [str(row[ci].value or '').strip() for ci in range(len(row))]
            if '店舗名' in row_vals:
                candidates = [v for v in row_vals[4:] if v not in exclude]
                if candidates:
                    store_row_vals = row_vals
                else:
                    if ri + 1 < len(data):
                        store_row_vals = [str(data[ri + 1][ci].value or '').strip()
                                          for ci in range(len(data[ri + 1]))]
                break

        if store_row_vals is None:
            continue

        for ci, v in enumerate(store_row_vals):
            if ci >= 4 and v not in exclude:
                store_col_map[v] = ci

        for row in data:
            kubun = str(row[0].value or '').strip()
            code = str(row[1].value or '').strip()
            if not code or code in ['商品ｺｰﾄﾞ', 'None']:
                continue
            if not kubun or kubun in ['区分', '企業名', '店舗名'] or kubun in PAGE_MARKS:
                continue
            for sname, ci in store_col_map.items():
                if sname in quantities and code in quantities[sname]:
                    qty = quantities[sname][code]
                    row[ci].value = qty if qty > 0 else None

    wb.save(path)
    wb.close()


@app.route('/fax')
@login_required
def fax():
    order_folder_id = os.environ.get('BOX_ORDER_FOLDER_ID', '')
    try:
        order_files = _list_excel_files(order_folder_id) if order_folder_id else []
    except Exception as e:
        return render_template('fax.html', error=f'BOX接続エラー: {e}', order_files=[])
    return render_template('fax.html', order_files=order_files, error=None)


@app.route('/api/fax/table', methods=['POST'])
@login_required
def api_fax_table():
    order_file_id = request.json.get('order_file_id')
    if not order_file_id:
        return jsonify({'error': '発注書が未選択'}), 400
    tmp = None
    try:
        tmp = _download_to_temp(order_file_id, suffix='.xlsx')
        result = _parse_fax_order(tmp)
        if not result['stores']:
            return jsonify({'error': '「稲穂」シートまたは店舗情報が見つかりませんでした'}), 400
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        _cleanup(tmp)


@app.route('/fax/save', methods=['POST'])
@login_required
def fax_save():
    import json as _json, shutil, tempfile as _tempfile
    order_file_id = request.form.get('order_file_id')
    stores_json = request.form.get('stores', '[]')
    try:
        stores = _json.loads(stores_json)
    except Exception:
        stores = []

    quantities = {}
    for key, val in request.form.items():
        if not key.startswith('qty_'):
            continue
        parts = key[4:].split('_', 1)
        if len(parts) != 2:
            continue
        code, store_idx_str = parts
        try:
            store_idx = int(store_idx_str)
            store_name = stores[store_idx]
            qty = int(val) if val else 0
            quantities.setdefault(store_name, {})[code] = qty
        except (ValueError, IndexError):
            continue

    tmp = output_tmp = None
    try:
        tmp = _download_to_temp(order_file_id, suffix='.xlsx')
        fd, output_tmp = _tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        shutil.copy2(tmp, output_tmp)
        _write_fax_quantities(output_tmp, stores, quantities)
        with open(output_tmp, 'rb') as f:
            file_bytes = f.read()
        _update_box_file(order_file_id, file_bytes)
        session['fax_order_file_id'] = order_file_id
        return redirect(url_for('normal'))
    except Exception as e:
        order_folder_id = os.environ.get('BOX_ORDER_FOLDER_ID', '')
        try:
            order_files = _list_excel_files(order_folder_id)
        except Exception:
            order_files = []
        return render_template('fax.html', error=f'保存エラー: {e}', order_files=order_files)
    finally:
        _cleanup(tmp, output_tmp)


if __name__ == '__main__':
    app.run(debug=True, port=5010)

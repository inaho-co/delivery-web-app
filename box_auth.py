"""
BOX OAuth2 初回認証スクリプト
実行: python box_auth.py
→ ブラウザが開く → Boxにログイン → 認証後にリダイレクト
→ access_token と refresh_token が表示される
→ Render の環境変数にコピーする
"""
import os
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from boxsdk import OAuth2

CLIENT_ID     = input('Box Client ID を入力: ').strip()
CLIENT_SECRET = input('Box Client Secret を入力: ').strip()

oauth = OAuth2(client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
auth_url, csrf_token = oauth.get_authorization_url('http://localhost:8765')

print(f'\nブラウザで認証ページを開きます...')
webbrowser.open(auth_url)

# 認証コードをローカルサーバーで受け取る
auth_code = None

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        qs = parse_qs(urlparse(self.path).query)
        auth_code = qs.get('code', [None])[0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write('認証完了！このタブは閉じてください。'.encode('utf-8'))

    def log_message(self, *args):
        pass  # ログ抑制

server = HTTPServer(('localhost', 8765), Handler)
print('認証待機中...')
server.handle_request()

if not auth_code:
    print('認証コードが取得できませんでした')
    exit(1)

access_token, refresh_token = oauth.authenticate(auth_code)

print('\n' + '='*60)
print('【Render に設定する環境変数】')
print('='*60)
print(f'BOX_CLIENT_ID     = {CLIENT_ID}')
print(f'BOX_CLIENT_SECRET = {CLIENT_SECRET}')
print(f'BOX_ACCESS_TOKEN  = {access_token}')
print(f'BOX_REFRESH_TOKEN = {refresh_token}')
print('='*60)
print('\nこの値を Render の Environment Variables にコピーしてください。')

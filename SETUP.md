# delivery-web-app セットアップガイド

## ステップ① BOX Developer アプリ作成

### 1. Box Developer Console にアクセス
- ブラウザで https://developer.box.com を開く
- 稲穂のBOXアカウント（tomo.dream69@gmail.com）でログイン

### 2. アプリを作成
1. 左メニュー「My Apps」をクリック
2. 「Create New App」ボタン
3. 「Custom App」を選択
4. アプリ名: `delivery-web-app` など（任意）
5. 「Create App」

### 3. OAuth 2.0 認証を設定
1. アプリの設定ページで「Configuration」タブを開く
2. **Application type** → 「User Authentication (OAuth 2.0)」を選択
3. **Redirect URIs** の欄に以下を追加：
   ```
   http://localhost:8765
   ```
   → 「Add」ボタンで追加 → 「Save Changes」

### 4. Client ID と Secret を確認
- 同じ Configuration ページの上部に以下が表示されている：
  - **Client ID**（20文字程度）
  - **Client Secret**（32文字程度）
  
この2つをメモしておく

---

## ステップ② BOX初回認証（トークン取得）

### 前提
- delivery-web-app フォルダが `C:\Users\tomod\Box\稲穂フォルダ\自作ツール\delivery-web-app\` にある

### 実行手順

1. **PowerShell またはコマンドプロンプトで実行：**
   ```
   cd "C:\Users\tomod\Box\稲穂フォルダ\自作ツール\delivery-web-app"
   pip install boxsdk
   python box_auth.py
   ```

2. **プロンプトが出たら Client ID を入力：**
   ```
   Box Client ID を入力: [ステップ①で確認したClient IDをペースト]
   ```

3. **Client Secret を入力：**
   ```
   Box Client Secret を入力: [ステップ①で確認したClient Secretをペースト]
   ```

4. **ブラウザが自動で開く：**
   - Box ログイン画面が表示される
   - 稲穂アカウントでログイン
   - 「Authorize」をクリック

5. **トークンが表示される：**
   ```
   ============================================================
   【Render に設定する環境変数】
   ============================================================
   BOX_CLIENT_ID     = [表示される値]
   BOX_CLIENT_SECRET = [表示される値]
   BOX_ACCESS_TOKEN  = [表示される値]
   BOX_REFRESH_TOKEN = [表示される値]
   ============================================================
   ```
   → これらの4つの値をメモする

---

## ステップ③ GitHub + Render デプロイ

（次のステップで説明）

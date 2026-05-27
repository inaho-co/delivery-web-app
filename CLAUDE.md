# CLAUDE.md — delivery-web-app（スマホ向け納品書作成Webアプリ）

外出先のスマホから納品書を作成・BOX保存・メール送信するWebアプリ（Flask + Render デプロイ）

## セッション開始時のルール

- 仕事 TODO ページを確認（自作ツール全体の ページID: `36670bbc-e793-81bc-89a4-c64ad3c8e1ad`）
- 本プロジェクト固有の進捗は Notion ログに記録

## 共通ルール

【最重要・絶対遵守ルール】

1. この会話は**常に日本語のみ**で行う。絶対に英語を使用しない。
2. どんなに会話が長くなっても、コンテキストが圧迫されても、**日本語を維持**する。

## Git 操作ルール

コミット作成後、自動的に `git push origin main` を実行する。

## プロジェクト情報

| 項目 | 値 |
|------|-----|
| **デプロイ URL** | https://delivery-web-app-74r1.onrender.com |
| **GitHub リポジトリ** | https://github.com/inaho-co/delivery-web-app |
| **デプロイ先** | Render （サーバーレス、無料枠内） |
| **フレームワーク** | Flask 3.0.3 + Gunicorn |
| **認証方式** | BOX OAuth2 + セッション認証（パスワード） |
| **メール送信** | Gmail SMTP（アプリパスワード） |

## ファイル構成

```
delivery-web-app/
├── app.py                      # Flask メインアプリケーション
├── box_auth.py                 # BOX 初回認証スクリプト（ローカル実行）
├── core_transfer.py            # 稲穂事務作業ツールからコピー（変更なし）
├── requirements.txt            # Python 依存パッケージ
├── render.yaml                 # Render デプロイ設定
├── CLAUDE.md                   # このファイル
├── SETUP.md                    # セットアップガイド
└── templates/
    ├── login.html              # ログインページ（パスワード認証）
    ├── index.html              # モード選択トップ
    ├── normal.html             # 通常モード（発注書から自動抽出）
    ├── manual.html             # 手入力モード（商品を手動入力）
    └── success.html            # 完了画面
```

## 環境変数（Render 設定）

| 変数名 | 用途 | 例 |
|--------|------|-----|
| `BOX_CLIENT_ID` | BOX OAuth2 認証 | （20文字） |
| `BOX_CLIENT_SECRET` | BOX OAuth2 認証 | （32文字） |
| `BOX_REFRESH_TOKEN` | BOX トークン更新 | 初回認証で取得 |
| `BOX_FOLDER_ID` | BOX 保存先フォルダID | 359331951814 |
| `GMAIL_APP_PASSWORD` | Gmail SMTP 認証 | gptm xpme zkzb opgt |
| `GMAIL_TO` | メール送信先（固定） | inaho.co@kke.biglobe.ne.jp |
| `APP_PASSWORD` | アプリログイン用PW | 6969 |
| `SECRET_KEY` | Flask セッション用 | delivery-2026 |

## 実装の主な特徴

- **モバイル最適化**：全 HTML がスマホ対応（48px+ タッチターゲット）
- **2つの入力モード**：
  - 通常モード：発注書を選択すると自動で商品を抽出
  - 手入力モード：商品を手動で入力
- **BOX との連携**：
  - 発注書ダウンロード → Excel 解析 → 納品書作成
  - 完成した納品書を BOX にアップロード
- **メール送信**：完成した納品書を Gmail で固定宛先に送信
- **トークン自動更新**：BOX Refresh Token を自動管理

## 開発・テスト

### ローカルで実行

```bash
cd delivery-web-app
pip install -r requirements.txt
python app.py
```

ブラウザで `http://127.0.0.1:5010` にアクセス

### Render へのデプロイ

GitHub にプッシュすると自動デプロイ（`git push origin main`）

## トラブルシューティング

- **404 エラー**：Render のデプロイ状況を確認（`Deploy latest commit` でリデプロイ）
- **BOX 認証エラー**：`box_auth.py` で Refresh Token を再取得
- **メール送信失敗**：Gmail のアプリパスワード確認、2段階認証有効か確認
- **ログインできない**：`APP_PASSWORD` 環境変数が 6969 に設定されているか確認

---

## 記録

作業ログは Notion ログページに日付付きで記録する。


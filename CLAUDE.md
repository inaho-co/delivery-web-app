# CLAUDE.md — delivery-web-app（スマホ向け納品書作成Webアプリ）

外出先のスマホから納品書を作成・BOX保存・メール送信するWebアプリ（Flask）

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
| **稼働環境** | 自宅 NAS（Synology DS118） |
| **アプリURL** | http://192.168.0.220:5010 |
| **アクセス方法** | 自宅Wi-Fi直接 / 外出先はOpenVPN経由 |
| **VPN DDNS** | inaho11.synology.me（UDP 1194） |
| **GitHub リポジトリ** | https://github.com/inaho-co/delivery-web-app |
| **フレームワーク** | Flask 3.0.3 |
| **認証方式** | BOX OAuth2 + トークン自動保存（box_tokens.json） |
| **メール送信** | Gmail SMTP（アプリパスワード） |

## NAS ファイル構成

```
/volume1/webapp/
├── app.py
├── core_transfer.py
├── config.json          # BOX JWT設定（使用していないが保持）
├── .env                 # 環境変数
├── start.sh             # 起動スクリプト
├── box_tokens.json      # BOXトークン自動保存（再起動後も維持）
└── templates/
```

## ローカル（PC）ファイル構成

```
delivery-web-app/
├── app.py
├── box_auth.py          # BOX初回認証スクリプト（ローカル実行）
├── core_transfer.py
├── requirements.txt
├── config.json          # .gitignore で除外
├── CLAUDE.md
└── templates/
```

## NAS 環境変数（/volume1/webapp/.env）

| 変数名 | 用途 |
|--------|------|
| `BOX_CLIENT_ID` | BOX OAuth2 認証 |
| `BOX_CLIENT_SECRET` | BOX OAuth2 認証 |
| `BOX_ACCESS_TOKEN` | BOX アクセストークン（初回用） |
| `BOX_REFRESH_TOKEN` | BOX リフレッシュトークン（初回用） |
| `BOX_ORDER_FOLDER_ID` | 発注書フォルダID：163709833837 |
| `BOX_DELIVERY_FOLDER_ID` | 納品書フォルダID：359331951814 |
| `GMAIL_APP_PASSWORD` | Gmail SMTP 認証 |
| `GMAIL_TO` | inaho.co@kke.biglobe.ne.jp |
| `APP_PASSWORD` | 6969 |
| `SECRET_KEY` | delivery-2026 |

## NAS 起動方法

```bash
# 手動起動
ssh nastom@192.168.0.220
sh /volume1/webapp/start.sh

# 自動起動
タスクスケジューラ（ブートアップ時）に設定済み
```

## NASへのファイル転送

```powershell
scp -O app.py nastom@192.168.0.220:/volume1/webapp/
```

## BOX 認証について

- OAuth2 + `store_tokens` コールバックで `box_tokens.json` に自動保存
- 再起動後もトークンが維持されるため手動更新不要
- JWT は個人BOXアカウントでは使用不可（エンタープライズ専用）
- トークン失効は60日以上未使用の場合のみ → `box_auth.py` で再取得

## トラブルシューティング

- **BOX 認証エラー**：`box_auth.py` でトークン再取得 → `.env` を更新 → 再起動
- **アプリが起動しない**：SSH で `sh /volume1/webapp/start.sh` を手動実行
- **外出先からアクセスできない**：OpenVPN Connect がオンになっているか確認

---

## 記録

作業ログは Notion ログページに日付付きで記録する。

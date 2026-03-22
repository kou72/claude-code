# nginx 設定統一チェックリスト（.88 → .188 相当）

設定変更は nginx-ui（`http://192.168.21.88:9000`）から行う。

---

## Stage 1 — HTTP リバースプロキシの疎通確認（証明書なし）

### Step 1: proxy_pass 先の疎通確認

nginx-ui のターミナル、または SSH で以下を実行。

- [ ] `.88` から `192.168.11.207` に到達できることを確認

  ```bash
  curl http://192.168.11.207
  ```

### Step 2: nginx-ui で設定変更

- [ ] nginx-ui → **Configuration** を開く
- [ ] 既存の設定を以下に書き換える

  ```nginx
  events {
      worker_connections 1024;
  }

  http {
      access_log /var/log/nginx/access.log;
      error_log  /var/log/nginx/error.log;

      server {
          listen 7273;
          server_name _;

          location / {
              proxy_pass http://192.168.11.207;
              proxy_set_header Host              $host;
              proxy_set_header X-Real-IP         $remote_addr;
              proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
              proxy_set_header X-Forwarded-Proto $scheme;

              proxy_connect_timeout 60s;
              proxy_send_timeout    60s;
              proxy_read_timeout    60s;
          }
      }
  }
  ```

- [ ] nginx-ui の **Test** でエラーがないことを確認
- [ ] nginx-ui の **Save & Reload** を実行

### Step 3: 動作確認

- [ ] nginx-ui でステータスが `running` のままであることを確認
- [ ] nginx-ui のエラーログに異常がないことを確認
- [ ] ポート 7273 のリスン確認

  ```bash
  ss -tlnp | grep 7273
  ```

- [ ] リバースプロキシ経由でレスポンスが返ることを確認

  ```bash
  curl http://192.168.21.88:7273
  # → 192.168.11.207 のレスポンスが返ること
  ```

---

## Stage 2 — SSL / mTLS の追加（証明書取得後）

### Step 4: 証明書の取得・配置

> CAサーバで以下3ファイルを発行してもらい、.88 に転送する。

| ファイル | 用途 |
| -------- | ---- |
| `server.crt` | サーバ証明書 |
| `server.key` | サーバ秘密鍵 |
| `ca.crt` | CA 証明書（クライアント認証用） |

- [ ] CAサーバで `server.crt` / `server.key` / `ca.crt` を発行
- [ ] 動作確認用クライアント証明書（`client.crt` / `client.key`）も発行
- [ ] 証明書を .88 に転送して配置

  ```bash
  sudo mkdir -p /etc/nginx/ssl/ca
  sudo cp reproxy.crt /etc/nginx/ssl/
  sudo cp reproxy.key /etc/nginx/ssl/
  sudo cp root.crt    /etc/nginx/ssl/ca/
  sudo chmod 600 /etc/nginx/ssl/reproxy.key
  ```

- [ ] 証明書の検証

  ```bash
  openssl verify -CAfile /etc/nginx/ssl/ca/root.crt /etc/nginx/ssl/reproxy.crt
  # → reproxy.crt: OK
  ```

### Step 5: nginx-ui で SSL / mTLS 設定を追加

- [ ] nginx-ui → **Manage Configs** → `conf.d` → `proxy.conf` の **Modify** を開く
- [ ] 内容を以下に書き換える

  ```nginx
  server {
      listen 7273 ssl;
      server_name _;

      ssl_certificate     /etc/nginx/ssl/reproxy.crt;
      ssl_certificate_key /etc/nginx/ssl/reproxy.key;

      ssl_client_certificate /etc/nginx/ssl/ca/root.crt;
      ssl_verify_client on;
      ssl_verify_depth  2;

      ssl_protocols TLSv1.2 TLSv1.3;
      ssl_ciphers HIGH:!aNULL:!MD5;
      ssl_prefer_server_ciphers on;

      location / {
          proxy_pass http://192.168.11.207;
          proxy_set_header Host              $host;
          proxy_set_header X-Real-IP         $remote_addr;
          proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
          proxy_set_header X-Forwarded-Proto $scheme;

          proxy_connect_timeout 60s;
          proxy_send_timeout    60s;
          proxy_read_timeout    60s;
      }
  }
  ```

- [ ] nginx-ui の **Test** でエラーがないことを確認
- [ ] nginx-ui の **Save & Reload** を実行

### Step 6: 動作確認

- [ ] nginx-ui でステータスが `running` のままであることを確認
- [ ] nginx-ui のエラーログに異常がないことを確認
- [ ] クライアント証明書なし → 拒否されることを確認

  ```bash
  curl -k https://192.168.21.88:7273
  # → 400 No required SSL certificate was sent
  ```

- [ ] クライアント証明書あり → 転送されることを確認

  ```bash
  curl -k \
    --cert client.crt \
    --key  client.key \
    https://192.168.21.88:7273
  # → 192.168.11.207 のレスポンスが返ること
  ```

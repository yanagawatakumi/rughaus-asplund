# ASPLUND Order Generator (Cloudflare Pages Static)

Serverless static app.  
CSV parsing, validation, phone normalization, XLSX generation, and ZIP download all run in the browser.

## Local test

```bash
cd "/Users/yana/Desktop/RUGHAUS/作業用0414/アスプルンド発注書/pages_app"
python3 -m http.server 8080
```

Open:

`http://127.0.0.1:8080`

## Cloudflare Pages deploy

1. Push this folder to GitHub.
2. In Cloudflare Dashboard:
   - `Workers & Pages` -> `Create` -> `Pages` -> `Connect to Git`
3. Repository settings:
   - `Root directory`: `アスプルンド発注書/pages_app`
   - `Build command`: (empty)
   - `Build output directory`: `.`
4. Deploy.

## Notes

- Template file is bundled at `assets/asplund_template.xlsx`.
- This app does not upload CSV/order data to a backend server.
- If template layout changes, replace `assets/asplund_template.xlsx`.

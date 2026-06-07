# Debt Tracker

A simple Python web app to track outstanding debt balances locally.

## Setup

1. Create and activate a Python virtual environment (recommended):

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Run

```bash
python app.py
```

Then open `http://127.0.0.1:5000` in your browser.

If port `5000` is already in use, the app will automatically try the next available port and print the URL in the terminal.

## Access from iPhone / iPad

If your Mac and your iPhone/iPad are on the same Wi-Fi network, use your Mac's local IP address instead of `127.0.0.1`. For example:

```bash
http://192.168.1.42:5000
```

To find your Mac's IP address, run:

```bash
ipconfig getifaddr en0
```

If you want remote access from outside your network, use a tunnel service such as `ngrok`.

## Notes

- Debt data is saved locally to `debts.json`.
- Use the main form to edit balances and delete accounts.
- Use the add-account form to add a new debt entry.

# LitReady — Talking River Review Production Tool

Drop a messy .docx. Get a clean one back. Ready for InDesign.

## How to Deploy (15 minutes, no coding required)

### Step 1: Get a GitHub account
If you don't have one: go to [github.com](https://github.com) and sign up. Free.

### Step 2: Create a new repository
1. Go to [github.com/new](https://github.com/new)
2. Name it `litready` (or whatever you want)
3. Set it to **Public**
4. Click **Create repository**

### Step 3: Upload these files
1. On your new repo page, click **"uploading an existing file"**
2. Drag ALL the files from this folder into the upload area:
   - `server.py`
   - `litready_engine.py`
   - `requirements.txt`
   - `Procfile`
   - `.gitignore`
   - The `static/` folder (with `index.html` inside it)
3. Click **Commit changes**

### Step 4: Deploy to Railway
1. Go to [railway.app](https://railway.app) and sign in with your GitHub account
2. Click **New Project**
3. Click **Deploy from GitHub Repo**
4. Select your `litready` repo
5. Railway will automatically detect Python, install dependencies, and start the server
6. Wait ~2 minutes for it to build
7. Click **Settings** → **Networking** → **Generate Domain**
8. That's your live URL. Done.

### Step 5: Use it
Go to your Railway URL. Drop a .docx. Pick the genre. Click "Clean & Map Styles." Download the cleaned file. Place it in InDesign.

## Local Development (optional)

```bash
pip install -r requirements.txt
python server.py
```

Then open http://localhost:8000 in your browser.

## What's Inside

- `server.py` — The web server (FastAPI)
- `litready_engine.py` — The formatting engine
- `static/index.html` — The frontend
- `requirements.txt` — Python dependencies
- `Procfile` — Tells Railway how to start the app

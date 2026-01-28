# Timetable API — Railway Deployment

This repository contains a FastAPI app (`api/index.py`). These notes show exactly how to deploy it to Railway.

## Files already present

- `requirements.txt` — Python dependencies (includes `gunicorn`, `uvicorn[standard]`).
- `Procfile` — start command for production (uses Gunicorn + Uvicorn worker).
- `runtime.txt` — Python version.
- `render.yaml`, `.gitignore`, etc.

## Local run (quick test)

Install deps and run locally:

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
# Run with Gunicorn (use same command Railway will use):
gunicorn -k uvicorn.workers.UvicornWorker api.index:app --bind 0.0.0.0:8000
# or with uvicorn for dev
uvicorn api.index:app --host 0.0.0.0 --port 8000
```

Open http://localhost:8000/ to verify.

## Git steps

If you haven't already:

```bash
cd c:\timetable-api
git init
git add .
git commit -m "Prepare project for Railway deployment"
git branch -M main
# create a repo on GitHub and then:
git remote add origin https://github.com/YOUR_USERNAME/timetable-api.git
git push -u origin main
```

## Deploy to Railway (UI)

1. Create a Railway account (https://railway.app) and sign in with GitHub.
2. Click **New Project** → **Deploy from GitHub repo**.
3. Select your `timetable-api` repository.
4. Set **Root Directory** to blank (`.`) — leave it empty if asked.
5. For **Start Command** use (Railway may auto-detect, but set explicitly):

```bash
gunicorn -k uvicorn.workers.UvicornWorker api.index:app --bind 0.0.0.0:$PORT
```

Railway automatically provides the `PORT` environment variable. Do not hardcode the port.

6. Add any environment variables under Settings → Variables (if your app needs them).
7. Click **Deploy**. Railway will build using `requirements.txt` and run the start command.

## Notes & troubleshooting

- If Railway build fails due to dependencies, check `requirements.txt` versions.
- If the service crashes, open the Railway logs to see the error (missing env vars, import errors, etc.).
- Ensure `api/index.py` exposes `app` at top-level (it does).
- Root Directory: only set it if your code is inside a subfolder (not the case here).

## After deploy

- Railway will provide a public URL. Visit `/` to validate JSON response.
- Push new commits to `main` and Railway will redeploy automatically if you enabled the GitHub integration.

---

If you want, I can now:
- Commit these changes for you and show the exact `git` commands to run locally, or
- Attempt to run a local container/test to verify the `gunicorn` start command works.

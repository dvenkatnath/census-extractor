# ✅ Streamlit Cloud Deployment Checklist

## Before Deploying

- [ ] Code is pushed to GitHub (public repository)
- [ ] `nf9_streamlit.py` is in the root directory
- [ ] `nf9.py` is in the root directory
- [ ] `requirements.txt` is up to date
- [ ] `.streamlit/config.toml` exists (created ✅)
- [ ] `.gitignore` excludes sensitive files (already done ✅)

## Deployment Steps

1. [ ] Sign in to [share.streamlit.io](https://share.streamlit.io)
2. [ ] Click "New app"
3. [ ] Select repository: `YOUR_USERNAME/census-extractor`
4. [ ] Main file: `nf9_streamlit.py`
5. [ ] Click "Deploy!"

## Configure Secrets

1. [ ] Go to app Settings → Secrets
2. [ ] Add `GROQ_API_KEY = "your_key_here"`
3. [ ] (Optional) Add `OPENAI_API_KEY = "your_key_here"`
4. [ ] Save and wait for app to restart

## Test

- [ ] App loads at `https://YOUR_APP_NAME.streamlit.app`
- [ ] Warning page displays correctly
- [ ] File upload works
- [ ] Extraction works with test file
- [ ] Share URL with colleague

## Quick Commands

```bash
# Check if git is initialized
git status

# If not, initialize and push
git init
git add .
git commit -m "Ready for Streamlit Cloud"
git remote add origin https://github.com/YOUR_USERNAME/census-extractor.git
git push -u origin main
```

---

**Need help?** See `STREAMLIT_CLOUD_SETUP.md` for detailed instructions.


# Streamlit Community Cloud Deployment Guide

## Step-by-Step Instructions

### Step 1: Create GitHub Repository

1. Go to https://github.com and sign in (or create account)
2. Click the **"+"** icon → **"New repository"**
3. Repository name: `census-extractor` (or any name you prefer)
4. Description: "Census Data Mapper & Extractor"
5. Make it **Public** (or Private - both work)
6. **DO NOT** initialize with README, .gitignore, or license
7. Click **"Create repository"**

### Step 2: Push Your Code to GitHub

Open terminal in this directory and run:

```bash
# Initialize git (if not already done)
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit - Census Extractor App"

# Add GitHub remote (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/census-extractor.git

# Push to GitHub
git branch -M main
git push -u origin main
```

**Note:** If you get asked for credentials:
- Use a **Personal Access Token** (not password)
- Generate one: GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)

### Step 3: Deploy to Streamlit Cloud

1. Go to https://share.streamlit.io
2. Click **"Sign in"** → Sign in with **GitHub**
3. Click **"New app"**
4. Fill in:
   - **Repository**: Select `YOUR_USERNAME/census-extractor`
   - **Branch**: `main`
   - **Main file path**: `app.py`
   - **App URL** (optional): Choose a custom name like `census-extractor`
5. Click **"Deploy"**

### Step 4: Add API Key Secret

1. In Streamlit Cloud, go to your app
2. Click **"Settings"** (⚙️ icon) → **"Secrets"**
3. Click **"Edit secrets"**
4. Add this:

```toml
GROQ_API_KEY = "your-actual-api-key-here"
```

5. Click **"Save"**
6. The app will automatically restart with the new secret

### Step 5: Access Your Public App

Once deployed, you'll get a URL like:
- `https://census-extractor.streamlit.app`

Share this URL with your team!

---

## Troubleshooting

### App won't start?
- Check **"Logs"** in Streamlit Cloud dashboard
- Verify `requirements.txt` has all dependencies
- Make sure `app.py` is the main file

### API key not working?
- Double-check the secret name is exactly `GROQ_API_KEY`
- Restart the app after adding secrets

### Import errors?
- Check `requirements.txt` includes all packages
- Streamlit Cloud installs packages from requirements.txt automatically

---

## Updating Your App

After making changes locally:

```bash
git add .
git commit -m "Description of changes"
git push
```

Streamlit Cloud will automatically redeploy!


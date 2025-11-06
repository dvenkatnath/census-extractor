# 🚀 Streamlit Cloud Deployment Guide

This guide will help you deploy `nf9_streamlit.py` to Streamlit Cloud so your colleague can access it from any network.

## 📋 Prerequisites

1. **GitHub Account**: You need a GitHub account (free)
2. **Streamlit Cloud Account**: Sign up at [share.streamlit.io](https://share.streamlit.io) (free)
3. **API Keys**: You'll need your API keys ready:
   - `GROQ_API_KEY` (for Groq models)
   - `OPENAI_API_KEY` (optional, if you want to use OpenAI)

## 🔧 Step-by-Step Setup

### Step 1: Push Your Code to GitHub

1. **Initialize Git** (if not already done):
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Streamlit Cloud deployment"
   ```

2. **Create a GitHub Repository**:
   - Go to [github.com](https://github.com)
   - Click "New repository"
   - Name it (e.g., `census-extractor`)
   - Make it **Public** (required for free Streamlit Cloud)
   - Click "Create repository"

3. **Push Your Code**:
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/census-extractor.git
   git branch -M main
   git push -u origin main
   ```
   Replace `YOUR_USERNAME` with your GitHub username.

### Step 2: Deploy to Streamlit Cloud

1. **Sign in to Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click "Sign in" and authorize with GitHub

2. **Deploy Your App**:
   - Click "New app"
   - Select your repository: `YOUR_USERNAME/census-extractor`
   - Select branch: `main`
   - **Main file path**: `nf9_streamlit.py`
   - Click "Deploy!"

### Step 3: Configure Secrets (API Keys)

1. **Add Secrets**:
   - In your Streamlit Cloud app dashboard, click "⚙️ Settings" (or "Manage app")
   - Click "Secrets" tab
   - Add your API keys in this format:

   ```toml
   GROQ_API_KEY = "your_groq_api_key_here"
   OPENAI_API_KEY = "your_openai_api_key_here"  # Optional
   ```

2. **Save and Restart**:
   - Click "Save"
   - The app will automatically restart with the new secrets

### Step 4: Access Your App

- Your app will be available at: `https://YOUR_APP_NAME.streamlit.app`
- Share this URL with your colleague - they can access it from any network!

## 📝 Important Notes

### Security
- ⚠️ **Never commit API keys** to GitHub
- ✅ Use Streamlit Cloud's "Secrets" feature for API keys
- ✅ The `.gitignore` file already excludes `.env` files

### File Structure
Your repository should have:
```
census-extractor/
├── nf9_streamlit.py      # Main app file
├── nf9.py                # Core extraction logic
├── requirements.txt      # Python dependencies
├── .streamlit/
│   └── config.toml       # Streamlit configuration
└── .gitignore           # Git ignore rules
```

### Requirements.txt
Make sure `requirements.txt` includes all dependencies:
```
streamlit>=1.28
pandas>=1.5
openpyxl>=3.1
groq>=0.4
openai>=1.0
python-dotenv>=1.0
```

## 🔄 Updating Your App

1. **Make changes** to your code locally
2. **Commit and push** to GitHub:
   ```bash
   git add .
   git commit -m "Update app"
   git push
   ```
3. **Streamlit Cloud automatically redeploys** when you push to the main branch

## 🐛 Troubleshooting

### App Won't Deploy
- Check that `nf9_streamlit.py` is in the root directory
- Verify `requirements.txt` has all dependencies
- Check the "Logs" tab in Streamlit Cloud for errors

### API Key Errors
- Verify secrets are set correctly in Streamlit Cloud
- Check the secret names match exactly: `GROQ_API_KEY`, `OPENAI_API_KEY`
- Restart the app after adding secrets

### Import Errors
- Ensure all Python files (`nf9.py`, etc.) are in the repository
- Check that all dependencies are in `requirements.txt`

## 📞 Support

- Streamlit Cloud Docs: [docs.streamlit.io/streamlit-community-cloud](https://docs.streamlit.io/streamlit-community-cloud)
- Streamlit Community: [discuss.streamlit.io](https://discuss.streamlit.io)

---

**🎉 Once deployed, your colleague can access the app from anywhere!**


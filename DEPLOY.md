# How to Host DocMorph Online as a Live Web Converter

This guide explains how to host this web application online so anyone can open your website in their browser, drag & drop files, and download converted documents.

---

## ⚡ Option 1: Host Free on Render.com (Recommended)

[Render.com](https://render.com) offers a free tier that supports custom `Dockerfile` deployments.

### Step 1: Push Your Code to GitHub
1. Create a new repository on [GitHub](https://github.com/new) (e.g., `word-pdf-converter`).
2. In your terminal inside this folder, run:
   ```bash
   git init
   git add .
   git commit -m "Initial commit of DocMorph Web App"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/word-pdf-converter.git
   git push -u origin main
   ```

### Step 2: Deploy on Render
1. Go to [dashboard.render.com](https://dashboard.render.com/) and sign in with GitHub.
2. Click **New +** → **Web Service**.
3. Select your `word-pdf-converter` repository.
4. Fill in the settings:
   - **Name:** `my-docmorph-converter` (or any name you choose)
   - **Region:** Pick the closest region to you (e.g., Oregon or Frankfurt)
   - **Language / Environment:** Select **Docker** (Render will automatically detect your `Dockerfile`)
   - **Instance Type:** Select **Free**
5. Click **Create Web Service**.

Render will automatically build the container (installing Linux, LibreOffice, and Python) and give you a free live public URL (e.g., `https://my-docmorph-converter.onrender.com`).

---

## 🚂 Option 2: Host on Railway.app

1. Sign in to [railway.app](https://railway.app/).
2. Click **New Project** → **Deploy from GitHub repo**.
3. Select this repository.
4. Railway automatically detects `Dockerfile` and deploys it immediately.
5. Under service settings, click **Generate Domain** to get your public web link.

---

## 🤗 Option 3: Host 100% Free on Hugging Face Spaces

Hugging Face provides free persistent Docker spaces:
1. Go to [huggingface.co/spaces](https://huggingface.co/spaces) and click **Create new Space**.
2. Space SDK: Select **Docker** → **Blank**.
3. Push your files to the Hugging Face git repository or upload them directly.
4. Your web converter will be live with an embedded public URL.

---

## 💻 Option 4: Run the Web App Locally on Your PC

You can also run the web interface directly on your own machine to use it in your web browser:

1. Open PowerShell or Command Prompt in this folder.
2. Run:
   ```powershell
   python web_app.py
   ```
3. Open your browser and go to:
   [http://127.0.0.1:5000](http://127.0.0.1:5000)

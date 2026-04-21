# Preperly — AI Lesson Planner & PowerPoint Generator

## Deploy to Vercel (follow these steps exactly)

### Step 1 — Put the files on GitHub

1. Go to **github.com** and sign in
2. Click the **+** button (top right) → **New repository**
3. Name it `preperly`
4. Leave everything else as default → click **Create repository**
5. On the next page, click **uploading an existing file**
6. Drag ALL the files from the `preperly-app` folder into the browser window:
   - `vercel.json`
   - `requirements.txt`
   - `api/index.py`
   - `public/index.html`
7. Click **Commit changes**

### Step 2 — Deploy on Vercel

1. Go to **vercel.com** and sign in
2. Click **Add New → Project**
3. Click **Import** next to your `preperly` repository
4. Under **Environment Variables**, add:
   - Name: `ANTHROPIC_API_KEY`
   - Value: *(leave blank for now — teachers use their own key)*
5. Click **Deploy**
6. Wait 2 minutes — Vercel builds and deploys automatically
7. You get a URL like `preperly.vercel.app` — that's your live app!

### Step 3 — Test it

1. Open the Vercel URL in any browser
2. Enter your Anthropic API key when prompted
3. Generate a lesson and download the PowerPoint

### Connect your domain (later)

1. In Vercel → your project → **Settings → Domains**
2. Add `preperly.uk`
3. Follow the DNS instructions Vercel gives you

## How it works

- Teachers visit the URL, enter their API key once (saved in browser)
- They fill in lesson details → Claude generates the lesson plan
- They review and edit → click Build PowerPoint
- The server generates a real `.pptx` using python-pptx and sends it back
- Teacher downloads it and opens in PowerPoint, Keynote, Google Slides, or Canva

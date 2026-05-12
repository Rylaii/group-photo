# Group Photo Uploader

Upload group photos with a group label → stored in the cloud → export all as a PowerPoint.

## Deploy to Vercel (step-by-step)

### 1. Push to GitHub

1. Create a new repository on [github.com](https://github.com/new)
2. In your terminal:
   ```bash
   cd group-photo-app
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
   git push -u origin main
   ```

### 2. Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) → **Add New Project**
2. Import your GitHub repository
3. Click **Deploy** (default settings work fine)

### 3. Add Vercel Blob Storage (for photos to persist)

1. In your Vercel project dashboard → **Storage** tab
2. Click **Create Database** → choose **Blob**
3. Name it anything (e.g. `group-photos`) → **Create**
4. Click **Connect to Project** → select your project → **Connect**
5. Vercel will automatically add `BLOB_READ_WRITE_TOKEN` to your environment variables

### 4. Redeploy

After connecting Blob, go to **Deployments** → click the three dots on the latest deployment → **Redeploy**.

Your app is now live at `https://your-project.vercel.app` 🎉

---

## How it works

| Tab | What it does |
|-----|-------------|
| **Upload** | Pick an image + enter a group name → saves to Vercel Blob |
| **Gallery** | Shows all uploaded entries (shared across everyone) — delete individual entries |
| **Generate** | Downloads a `.pptx` file with one slide per group entry |

## Local development

```bash
npm install
# Copy .env.example to .env.local and fill in your token
cp .env.example .env.local
npm run dev
```

Then open [http://localhost:3000](http://localhost:3000)

## Tech stack

- **Next.js 14** (App Router)
- **Vercel Blob** — cloud image + metadata storage
- **pptxgenjs** — PowerPoint generation (runs in the browser)

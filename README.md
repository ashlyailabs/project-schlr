# Project Scheduler
Corporate Department · Construction & Furnishing

## Stack
- **Next.js 14** (App Router)
- **Supabase** (database)
- **Tailwind CSS** (styling)
- **Vercel** (hosting)

---

## Setup — 3 steps

### 1. Supabase database

1. Go to [supabase.com](https://supabase.com) → sign in → create a new project
2. Go to **Database → SQL Editor → New query**
3. Paste the entire contents of `supabase/schema.sql` and click **Run**
4. Go to **Project Settings → API** and copy:
   - **Project URL** → `NEXT_PUBLIC_SUPABASE_URL`
   - **anon / public key** → `NEXT_PUBLIC_SUPABASE_ANON_KEY`

### 2. Local development

```bash
npm install
cp .env.local.example .env.local
# Paste your Supabase values into .env.local
npm run dev
```

Open [http://localhost:3000](http://localhost:3000)

### 3. Deploy to Vercel

```bash
git init
git add .
git commit -m "initial commit"
```

Push to GitHub, then:

1. Go to [vercel.com](https://vercel.com) → **Add New Project** → import your repo
2. Before clicking Deploy, open **Environment Variables** and add:
   - `NEXT_PUBLIC_SUPABASE_URL`
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY`
3. Click **Deploy**

---

## How to use

| Action | How |
|---|---|
| Add a project | Home page → "Add new project" card |
| Open a project | Click any project card |
| Add a task | Inside a project → "+ Add Task" button |
| Edit a task | Click any activity row |
| Highlight a day | Click any day on the calendar |
| Un-highlight | Click the same day again |
| Delete a project | Hover the card → ✕ button |

---

## Data model

```
projects       — id, name, created_at
tasks          — id, project_id, sl, name, description, start_date, end_date, status, progress
highlights     — id, project_id, date
```

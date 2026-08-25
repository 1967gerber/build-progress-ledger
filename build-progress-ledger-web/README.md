# BUILD Progress Ledger — Web App

This is a real, deployable version of the Progress Ledger, built with React. It uses
[Supabase](https://supabase.com) for real teacher accounts (email + password) and a real
database, and is meant to be hosted on [Render](https://render.com) as a Static Site.

You do **not** need to know how to code to get this running — just follow the steps below
in order. It'll take about 20–30 minutes the first time.

---

## Overview of what you're setting up

1. **Supabase** — a free account that gives you a real database and login system
2. **GitHub** — where the code lives so Render can find it
3. **Render** — hosts the actual website, for free

---

## Part 1: Set up Supabase

1. Go to [supabase.com](https://supabase.com) and sign up for a free account.
2. Click **New Project**. Give it any name (e.g. "build-progress-ledger"), set a database
   password (save it somewhere — you likely won't need it again, but keep it just in case),
   and pick the region closest to you. Click **Create new project** and wait a minute or two
   while it sets up.
3. Once it's ready, click **SQL Editor** in the left sidebar.
4. Open the file `supabase/schema.sql` from this project (in a text editor, or on GitHub
   once you've uploaded it — see Part 2), copy its entire contents, paste them into the SQL
   Editor, and click **Run**. This creates the table and security rules the app needs.
   You should see "Success. No rows returned."
5. Click **Authentication** in the left sidebar → **Providers**, and confirm **Email** is
   enabled (it is by default).
6. **Turn off email confirmation for now** (recommended while testing): go to
   **Authentication → Sign In / Providers → Email**, and toggle off "Confirm email". This
   lets teachers start using their account immediately after signing up instead of waiting
   for a confirmation email. You can turn this back on later if you want extra security.
7. Click **Settings → API** in the left sidebar. You'll need two values from this page in
   Part 3:
   - **Project URL** (looks like `https://xxxxx.supabase.co`)
   - **anon public** key (a long string under "Project API keys")

Keep this tab open — you'll copy these two values into Render shortly.

---

## Part 2: Put the code on GitHub

Render deploys from a GitHub repository, so the code needs to live there first.

1. Go to [github.com](https://github.com) and sign up for a free account if you don't have one.
2. Click the **+** icon (top right) → **New repository**. Name it anything (e.g.
   `build-progress-ledger`), leave it Public or Private (either works), and click
   **Create repository**.
3. On the new repository's page, click **uploading an existing file**.
4. Drag in every file and folder from this project (everything you downloaded/extracted),
   keeping the folder structure intact (the `src` folder and `supabase` folder need to stay
   as folders, not get flattened).
5. Scroll down and click **Commit changes**.

If you're comfortable with git on the command line instead, the usual steps work too
(`git init`, `git add .`, `git commit`, `git remote add origin ...`, `git push`) — either
way gets you to the same place.

---

## Part 3: Deploy on Render

1. Go to [render.com](https://render.com) and sign up (you can sign up with your GitHub
   account, which makes the next step easier).
2. Click **New +** → **Static Site**.
3. Connect your GitHub account if prompted, then select the repository you just created.
4. Fill in these settings:
   - **Name**: anything, e.g. `build-progress-ledger`
   - **Build Command**: `npm install && npm run build`
   - **Publish Directory**: `dist`
5. Before clicking Create, scroll to **Environment Variables** and add two:
   | Key | Value |
   |---|---|
   | `VITE_SUPABASE_URL` | the Project URL from Supabase (Part 1, step 7) |
   | `VITE_SUPABASE_ANON_KEY` | the anon public key from Supabase (Part 1, step 7) |
6. Click **Create Static Site**. Render will build and deploy — this takes a few minutes.
   You'll get a live URL like `https://build-progress-ledger.onrender.com` when it's done.

That URL is your real, permanent website. Bookmark it, share it with other teachers — it's
the same app every time, unlike the Claude preview links.

---

## Part 4: First use

1. Open your new Render URL.
2. Click **Create an account** — enter your name, a real email, and a password (6+
   characters). Since you turned off email confirmation, you'll be signed in right away.
3. Everything works exactly like the version you tried in Claude: add students, log
   checkpoints, track mastery, view growth charts, export to Excel.
4. Other teachers can go to the same URL and create their own accounts. Anyone signed in
   can browse **All Teachers** to view (but not edit) anyone else's data — that's by design,
   matching what you asked for earlier.

---

## Making changes later

If you ever want a feature added or a bug fixed, the easiest path is to come back here,
ask for the change, get the updated `App.jsx`, and re-upload just that one file to your
GitHub repository (in the `src` folder) — Render will automatically rebuild and redeploy
within a couple of minutes.

## Troubleshooting

- **Blank white page after deploying**: almost always means the two environment variables
  in Render are missing or mistyped. Double check them against Supabase's Settings → API.
- **"Invalid API key" errors**: same as above — the anon key was probably copied
  incorrectly (make sure there's no extra space at the start or end).
- **Sign-up says "check your email" and you don't want that**: go back to Supabase →
  Authentication → Providers → Email and turn off "Confirm email".
- **Local testing before deploying** (optional, if you want to try it on your own computer
  first): install [Node.js](https://nodejs.org), then in this folder run:
  ```
  npm install
  cp .env.example .env
  ```
  Edit `.env` with your real Supabase values, then run `npm run dev` and open the URL it
  gives you.

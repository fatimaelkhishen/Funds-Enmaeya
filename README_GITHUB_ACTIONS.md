# Running the weekly scrape without your computer (GitHub Actions)

This lets the scraper run every Sunday on GitHub's own servers, for free -
your computer can be off. One-time setup, all done in a browser, no coding
or terminal needed.

## 1. Create a GitHub account (skip if you already have one)

Go to https://github.com/signup and create a free account.

## 2. Create a new repository

- Click the "+" in the top-right corner -> "New repository".
- Name it something like `enmaeya-news-scraper`.
- Set it to **Private** (recommended, since it's internal monitoring).
- Do NOT add a README/gitignore/license on this screen (leave everything
  unchecked) - just click "Create repository".

## 3. Upload this whole `News` folder

- On the new (empty) repo page, click "uploading an existing file".
- Open this `News` folder in File Explorer, select everything inside it
  (including the `scraper` folder and the hidden `.github` folder - in File
  Explorer, turn on View -> Show -> Hidden items so you can see `.github`),
  and drag it all into the GitHub upload box.
- Scroll down and click "Commit changes".

## 4. Turn on push permissions for the workflow (one-time)

- In your new repo: **Settings** tab -> **Actions** -> **General** (left
  sidebar) -> scroll to **Workflow permissions**.
- Select **"Read and write permissions"**.
- Click **Save**.

(Without this step, the weekly run will scrape fine but fail to save the
new articles back to the repo.)

## 5. Test it once

- Go to the **Actions** tab -> click **"Weekly Lebanon News Scrape"** on the
  left -> click **"Run workflow"** (dropdown on the right) -> **Run
  workflow**.
- It takes a while to finish (scraping ~45 sites) - refresh the page to
  watch progress. When it's green/done, check that the `.xlsx` files in the
  repo have newer commit dates.

After this, it runs automatically every Sunday (~8am Beirut time) with no
further action needed from you.

## Getting the new articles back into this OneDrive folder

The scraper now lives in two places: this OneDrive folder (what you're
looking at) and the GitHub repo (where it actually runs weekly). GitHub
does not automatically write back into OneDrive. Easiest options:

- **Manual**: once in a while, open the repo on github.com and download the
  updated `.xlsx` files (each file's page has a "Download raw file" button),
  replacing the ones in this folder.
- **Automatic**: tell me your repo's URL and I can set up a weekly step on
  my side that pulls the latest files from GitHub straight into this
  OneDrive folder, so you never have to download anything by hand.

## Note on the Windows Task Scheduler files

`run_weekly.bat` and `setup_weekly_task.bat` in the `scraper` folder are the
alternative "run on your own computer" option from before. You don't need
them if you go with GitHub Actions - safe to ignore or delete.

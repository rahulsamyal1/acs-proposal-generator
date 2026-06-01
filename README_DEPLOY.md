# ACS Proposal Generator — Web App

A password-protected web version of the proposal generator. Fill the form in
any browser (laptop or phone), click **Generate**, and download the branded
Word doc + PDF. Hosted free on **Streamlit Community Cloud**.

---

## What's in this folder

| File / folder | Purpose |
|---|---|
| `app.py` | The web app (the form + generate logic) |
| `proposal.py` | Builds the Word doc and the PDF |
| `template/proposal_template.docx` | The branded master template |
| `presets.json` | Job-type presets + default service terms (editable) |
| `assets/acs_logo.png` | Logo shown at the top of the app |
| `requirements.txt` | Python packages the host installs |
| `packages.txt` | System packages the host installs (LibreOffice, for PDF) |
| `.streamlit/config.toml` | Green ACS theme |
| `secrets.toml.example` | Shows the password setting to add on the host |

---

## Deploy it (one time, ~10 minutes)

You need a **GitHub account** and a **Streamlit Community Cloud account**
(both free; you sign in to Streamlit with your GitHub login).

### 1. Put this folder on GitHub
- Create a new repository (a **private** repo is fine — Streamlit can read it).
- Upload the entire contents of this `Proposal Generator Web` folder.
  - Easiest no-command option: github.com → New repository → "uploading an
    existing file" → drag everything in.
  - Do **not** upload `.venv/` or `.streamlit/secrets.toml` (the `.gitignore`
    already excludes them).

### 2. Create the app on Streamlit
- Go to **https://share.streamlit.io** and sign in with GitHub.
- Click **Create app → Deploy a public app from GitHub** (private repos work too).
- Repository: your repo. Branch: `main`. Main file path: `app.py`.
- Click **Advanced settings → Secrets** and paste:
  ```toml
  app_password = "choose-a-strong-password"
  ```
- Click **Deploy**.

The first build takes a few minutes (it installs LibreOffice for the PDFs).
When it finishes you get a link like `https://your-app-name.streamlit.app`.

### 3. Use it
- Open the link, enter the password, fill the form, click **Generate**,
  download the Word doc and/or PDF.
- Share the link + password with Hannah. Works on phone and laptop.

> **Free-tier note:** the app goes to sleep after a period of no use. The next
> visit takes ~30–60 seconds to wake up, then it's fast again.

---

## Run it locally (optional, for testing)

From this folder:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

(The local password is `changeme` unless you edit `.streamlit/secrets.toml`.)
PDF generation locally needs LibreOffice installed; if it isn't, you still get
the Word document and a note. On the deployed app, PDF always works.

---

## Making changes

- **Password:** edit the `app_password` secret in Streamlit (App → Settings → Secrets).
- **Presets / wording:** edit `presets.json`, commit to GitHub — the app redeploys automatically.
- **Logo, contact details, layout:** these live in `template/proposal_template.docx`.
  Update that file, commit it, and the app uses the new template.

---

## Privacy

The app runs on Streamlit's servers. Proposals are generated on demand and sent
straight to your browser as downloads; they are not stored anywhere. The app is
protected by the password you set.

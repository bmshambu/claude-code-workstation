# Run Pitchcraft on the office laptop

Instructions for a coding agent (Cursor) on the Windows office laptop. Goal: run
**Pitchcraft**, the PowerPoint add-in for authoring proposal templates, against
its backend (`ppt_gen_v3p1`), and keep developing it there.

Pitchcraft started as the hello-world add-in in this folder; the folder name
`ppt-addin-hello` is kept so existing clones keep working.

## What it is

A task pane in PowerPoint with five tabs, all talking to the `ppt_gen_v3p1`
FastAPI backend through the dev server's proxy:

| Tab | What it does | Backend calls |
|---|---|---|
| Library | Import a library `.pptx` (or the open presentation); after import the library's slides are shown in PowerPoint, each tagged `BLOCK=<id>` | `POST /api/libraries`, `GET /api/libraries/{lib}` |
| Questions | Upload / delete JSON answer sets; list the questions (fields) they produce | `GET .../fields`, `POST .../payloads`, `DELETE .../payloads/{name}` |
| Rules | Deck order (drag), include/exclude, conditions (all/any of field is / is not / is one of / answered); try the rules on an answer set before saving | `GET`/`PUT .../rules` |
| Placeholders | Mark text in PowerPoint: select words, pick an existing placeholder or **+ New placeholder…**, Insert; set where each value comes from; **Update library from PowerPoint** | `PUT .../rules`, `POST /api/libraries` (replace) |
| Build | Pick an answer set, build on the backend, insert the slides into the open deck tagged by block | `POST .../select`, `POST .../build`, `GET /download/{token}` |

Decisions already made, do not reopen them:

- **No PDFs.** Stakeholders do not want the PDF preview. Placeholders come from
  the `.pptx` itself and PowerPoint is the only renderer. The backend must run
  **without Graph** (see step 4).
- **XML manifest on this laptop.** The JSON manifest's sideload is blocked by the
  firewall (`titles.prod.mos.microsoft.com`). Use `npm run start:xml`.
- **Port 3002**, because AMAgent holds 3000 on this laptop.
- **Do not change the backend's behaviour.** `ppt_gen_v3p1/CLAUDE.md` says so. The
  add-in uses only its existing API.
- Building decks will later move to Gemini Enterprise chat; the add-in is the
  authoring tool.

## Ground rules for the agent

- Run commands in a terminal and show the user their output. Report errors verbatim.
- Do **not** change Office Trust Center settings, Windows sharing, group policy,
  or the registry by hand. If a step needs that, stop and ask the user.
- Never put credentials in files, never create `ppt_gen_v3p1/.env`, and never
  commit real (client) templates, decks or answer sets. `templates/demo/` and
  `fixtures/` are synthetic and safe.
- Commit and push only when the user asks.
- PowerPoint steps (clicking the ribbon and the pane) are for the user to do;
  ask them what they see.

## 1. Prerequisites

```
node -v          # v18 or later
python --version # 3.10 or later, for the backend
git --version
```

If Node or Python is missing, stop and tell the user to install it (Node.js LTS,
Python 3.11+) from the company software portal.

## 2. Get or update the code

Fresh machine:

```
git clone --depth 1 https://github.com/bmshambu/claude-code-workstation.git
cd claude-code-workstation
```

Existing clone from the hello-world test: it has local edits (port 3002 in four
files, placeholder icons). Throw those away first, then pull. The port is set
again in step 3 with one command.

```
git status
git checkout -- ppt-addin-hello
git clean -n ppt-addin-hello      # shows what would be removed; e.g. make-icons.js
git pull
```

Remove the files `git clean -n` listed only after the user agrees
(`git clean -f ppt-addin-hello`). Keep `results.md` if it is there.

Check the icons came across (`dir ppt-addin-hello\assets` lists
`icon-16/32/64/80.png`). If not, re-download them from the repo; do not generate
placeholders.

**Backend location.** The backend is **not** in this repo; it is already on
this laptop (typically `...\Templfy\proposal_builder\ppt_gen_v3p1`). Ask the
user for the path if unsure. Do not copy it into this repo. Before relying on
it, check it has the routes the pane calls (step 4 lists the checks). An older
copy (`ppt_gen_v3`) has the same API apart from Graph rendering, which is not
used.

## 3. Install the add-in and set its port

```
cd ppt-addin-hello
npm install --ignore-scripts
npm run set-port -- 3002
```

`--ignore-scripts` skips `keytar`'s native build, which fails here because
nodejs.org is unreachable; nothing Pitchcraft uses needs it. `set-port` rewrites
the port in `package.json` and in every URL of both manifests (webpack reads it
from `package.json`). Do not commit the port change unless the user asks.

Network errors on install (ECONNRESET, ETIMEDOUT, self-signed certificate): ask
the user for the company proxy and registry, then `npm config set proxy …`,
`https-proxy …`, `registry …`. Do not set `strict-ssl false` without the user's OK.

## 4. Run the backend (no Graph, no PDF)

In a second terminal, in the backend folder already on this laptop (if it
already has a working `.venv`, just activate it and skip the `venv`/`pip` lines):

```
cd <path-to>\ppt_gen_v3p1
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
set GRAPH_DRIVE_ID=
uvicorn app:app --host 127.0.0.1 --port 8000
```

- `pywin32` in requirements is optional (only `tools/make_library_pdf.py` uses
  it, and PDFs are out). If it fails to install, remove that line locally and
  carry on.
- **Graph must stay off.** With Graph configured the backend uploads every
  imported library to SharePoint to make a PDF. Without a `.env` it is off; the
  `set GRAPH_DRIVE_ID=` line guards against one appearing. Check:
  `curl http://127.0.0.1:8000/api/renderers` must show `graph` with
  `"available": false`.
- Check it serves the demo library: `curl http://127.0.0.1:8000/api/libraries`
  lists `demo`, and `curl http://127.0.0.1:8000/api/libraries/demo/fields`
  returns `fields` and `payloads`. If either route 404s, the copy is too old:
  stop and tell the user.
- If the office copy has a `.env` with Graph settings, leave the file alone; the
  `set GRAPH_DRIVE_ID=` line above switches Graph off for this run, since a
  variable already set wins over `.env`.
- If port 8000 is taken, pick another (e.g. 8010) for uvicorn, and before
  starting the add-in run `set PB_BACKEND=http://127.0.0.1:8010` in the add-in
  terminal. The dev server proxies `/api` and `/download` to `PB_BACKEND`.

## 5. Start Pitchcraft in PowerPoint

Close all PowerPoint windows, then in the add-in terminal:

```
npm run start:xml
```

- First run may ask to install a localhost development certificate; the user
  clicks **Yes**.
- Output ends with `Debugging started.` and PowerPoint opens a temporary deck.
  Leave the terminal running.
- If a **WebView Stop On Load** dialog appears, the user clicks OK.
- **Home** tab › **Pitchcraft** group › **Open pane**. The header pill should
  say **Connected**. "Offline" means the backend is not reachable: check step 4
  and `PB_BACKEND`.
- With the XML manifest the add-in only lives in the window this command opens.
  Reopening PowerPoint drops it; rerun `npm run start:xml`.
- The pane's HTML hot-reloads: after editing `taskpane.html`, close and reopen
  the pane. Changes to a manifest need `npm run stop:xml` then `npm run start:xml`.

## 6. Smoke test (user does this, agent records results)

Use the synthetic `demo` library.

1. **Build**: answer set `p2_expansion` › Build & insert slides. Expect 8 slides
   with *Globex International Holdings* on the cover; the block list shows
   `cover … contacts`.
2. **Questions**: 8 questions from 4 answer sets; `Quality` flagged "in 3 of 4".
3. **Rules**: open the condition on `expansion_detail`; *Try the rules* with
   `p1_new_client` vs `p2_expansion` shows it dropping in and out. Do not save.
4. **Placeholders**: the list shows `{{ClientName}}`, `{{DueDate}}`, fees, etc.
   Select a word on a slide, pick **+ New placeholder…**, name it, Insert: the word
   becomes `{{Name}}` and an orange "not in the library yet" note appears. Do
   **not** press Update library on `demo`.
5. **Library**: to test import without touching `demo`, import
   `<backend>\fixtures\master.pptx` as a new id (e.g. `office_test`) with
   "Show the library's slides" ticked. Its slides replace the open deck's.
6. **Update library from PowerPoint** on `office_test`: mark one placeholder,
   press it, confirm. The report lists the new placeholder.

Writes land in `ppt_gen_v3p1/templates/<id>/` (gitignored except `demo`);
saving rules keeps a `rules.json.<timestamp>.bak` beside it.

## 7. Clean up

Close PowerPoint, then `npm run stop:xml` in the add-in terminal and Ctrl+C in
the backend terminal.

## 8. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `EADDRINUSE` on 3000/3002 | Port held (AMAgent holds 3000) | `npm run set-port -- <free port>`, then `npm run stop:xml` / `npm run start:xml` |
| Pane pill "Offline", "Backend not reachable" | Backend not running or on another port | Step 4; set `PB_BACKEND` |
| Pane blank or "Script error" overlay | JavaScript error in `taskpane.html` | Runtime overlay is off; check the pane's status box, and `node -e` syntax-check the inline script |
| `Unable to sideload`, `401`, `Get ServiceUrl failed`, ETIMEDOUT to `titles.prod.mos.microsoft.com` | JSON manifest route blocked | Use `npm run start:xml` only |
| `keytar` / node-gyp build error on install | Native module can't fetch Node headers | `npm install --ignore-scripts` |
| Import or Update library takes long and mentions Graph or SharePoint | Graph is configured | Stop the backend, make sure no `.env`, `set GRAPH_DRIVE_ID=`, restart |
| "add-ins inserted during development are only available during debugging" | Opened the temp deck outside a debug session | Rerun `npm run start:xml` |

## 9. Where things are

- `ppt-addin-hello/taskpane.html` holds the whole pane: HTML, CSS and script in
  one file, sections marked `// ---- library`, `questions`, `rules`,
  `placeholders`, `build`.
- `ppt-addin-hello/webpack.config.js` has the dev server, the HTTPS cert and the
  proxy to the backend (`PB_BACKEND`).
- `ppt-addin-hello/manifest.xml` is the ribbon group "Pitchcraft" (used here);
  `manifest.json` is the Microsoft 365 route.
- The backend (outside this repo, on this laptop): `app.py` holds every API route
  the pane calls (thin; logic in `engine/`). Read its `CLAUDE.md` before touching
  it.

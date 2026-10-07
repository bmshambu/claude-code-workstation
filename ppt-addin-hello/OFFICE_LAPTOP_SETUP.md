# Run the Proposal Builder hello-world add-in on the office laptop

Instructions for a coding agent (Cursor) on a Windows office laptop. Goal: load
the hello-world PowerPoint add-in in this folder and report whether its four
buttons work on this machine's PowerPoint.

Already proven on a personal laptop (Office Home & Student 2021): the XML
manifest loads and buttons 1–3 work. This run checks **Microsoft 365
PowerPoint**, which should accept the newer JSON manifest (`manifest.json`).

## Ground rules for the agent

- Run commands in a terminal and show the user their output. Report errors verbatim.
- Do **not** change Office Trust Center settings, Windows sharing, group policy,
  or the registry by hand. If a step needs that, stop and ask the user.
- Do **not** commit or push anything. This is a local test only.
- PowerPoint steps (clicking the ribbon and buttons) are for the user to do; ask
  them what they see.

## 1. Check prerequisites

```
node -v        # need v18 or later
git --version
```

If Node is missing or older than 18, stop and tell the user to install Node.js
LTS from https://nodejs.org (or via the company software portal).

Ask the user to check PowerPoint: **File > Account > About PowerPoint**. Note:
- whether it says **Microsoft 365** (subscription) or Office 2021/2024 (perpetual),
- the **Version** number (for example 2508). JSON manifest needs **2501 or later**,
- that they are signed in with their **work account**.

## 2. Get the code

If this folder isn't already on the machine:

```
git clone --depth 1 https://github.com/bmshambu/claude-code-workstation.git
cd claude-code-workstation\ppt-addin-hello
```

If `git clone` fails (proxy/GitHub blocked), ask the user to download the repo
as a ZIP from GitHub, or copy the `ppt-addin-hello` folder over from the
personal laptop.

## 3. Install

```
npm install
```

If it fails with network/proxy errors (ECONNRESET, ETIMEDOUT, self-signed
certificate in chain), ask the user for the company proxy and registry, then:

```
npm config set proxy http://<proxy-host>:<port>
npm config set https-proxy http://<proxy-host>:<port>
npm config set registry <company-npm-registry-if-any>
```

Do not set `strict-ssl false` without the user's explicit OK.

## 4. Start — Microsoft 365 path (JSON manifest)

Close all PowerPoint windows first, then:

```
npm start
```

- First run asks to install a **localhost development certificate**. The user
  must click **Yes**.
- Expected output ends with `Debugging started.` and PowerPoint opens.
- Leave this terminal running; it hosts the pane at https://localhost:3000.

### If `npm start` fails

| Error | Meaning | Next step |
|---|---|---|
| `401`, `Unable to sideload`, `atk install` error, "custom app upload disabled" | Tenant blocks sideloading custom apps, or not signed in | Run `npm stop`, then use step 5 (XML) |
| PowerPoint version below 2501 | Too old for JSON manifest | Use step 5 (XML) |
| Certificate install blocked / access denied | IT policy blocks the dev cert | Stop; user must ask IT |
| `EADDRINUSE` port 3000 | Something already on 3000 | Find and close it, or ask the user |

## 5. Fallback — XML manifest

```
npm stop
npm run start:xml
```

Same expected behaviour. Note: with the XML manifest, the add-in only appears in
the PowerPoint window this command opens; reopening PowerPoint drops it (that's
normal for development mode). Rerun `npm run start:xml` to get it back.

## 6. Test in PowerPoint (user does this)

1. If a **"WebView Stop On Load"** dialog appears, click OK.
2. **Home** tab > **Proposal Builder** group > **Open pane**.
   - Not there? Try Home > Add-ins, or Insert > My Add-ins.
3. The pane first shows **PowerPointApi support**, e.g. `1.3: yes`. Note all four lines
   (1.1, 1.3, 1.5, 1.8) before clicking anything — the text gets replaced.
4. Button 1: click inside a text box on a slide, then the button. Text becomes `Hello world!`.
5. Button 2: click a slide thumbnail on the left, then the button. Expect `BLOCK = cover`.
6. Button 3: expect a list of slides; the tagged one shows `BLOCK=cover`.
   Drag that slide to another position, press button 3 again: the tag should move with it.
7. Button 4: a new blank PowerPoint window should open.
8. Close PowerPoint, reopen it with a blank deck, and check whether **Proposal Builder**
   is still on the Home tab (with the JSON manifest on Microsoft 365 it may persist).

## 7. Clean up

Close PowerPoint, then in the terminal:

```
npm stop            # or: npm run stop:xml   if step 5 was used
```

This unregisters the add-in and stops the local server.

## 8. Report back

Write the results in this shape so the user can paste them into the project thread:

```
Office laptop result
- PowerPoint: <Microsoft 365 / Office 2021 / ...>, Version <....>
- Manifest used: <JSON (npm start) / XML (npm run start:xml)>
- Errors (verbatim): <none / ...>
- PowerPointApi support: 1.1 <yes/no>, 1.3 <yes/no>, 1.5 <yes/no>, 1.8 <yes/no>
- Button 1 (write text): <worked / error>
- Button 2 (tag slide): <worked / error>
- Button 3 (list slides + tags, tag moves with slide): <worked / error>
- Button 4 (new presentation): <opened / error>
- Still on ribbon after restarting PowerPoint: <yes / no>
```

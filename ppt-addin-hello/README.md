# ppt-addin-hello

A hello-world PowerPoint add-in using the unified JSON manifest. It checks
the four things the proposal builder add-in depends on:

1. writing text into the open slide,
2. writing a block id as a slide tag and reading it back,
3. reading every slide in order with its tags,
4. opening a new presentation (how the built deck will be previewed).

Based on Microsoft's
[PowerPoint hello world sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world),
unified-manifest configuration.

## Run it (Windows, PowerPoint for Microsoft 365, version 2501 or later)

```
npm install
npm start
```

The first run asks to install a localhost development certificate. Accept it.
PowerPoint then opens with a **Proposal Builder** group on the **Home** tab.
Choose **Open pane**.

When finished, close PowerPoint and run `npm stop`.

### Perpetual Office (2021, 2024): use the XML manifest

Perpetual Office cannot load the JSON manifest (sideloading fails with a 401,
because it needs a Microsoft 365 account). `manifest.xml` describes the same
pane in the older format:

```
npm run start:xml
```

and `npm run stop:xml` when finished. The pane reports which PowerPoint API
versions this copy of PowerPoint supports, which matters here: perpetual
Office may lack the versions the tag buttons need.

## What to check

- The pane opens and shows `PowerPointApi support` with 1.3 and 1.5 as `yes`.
- Click into a text box, then button 1: the text becomes `Hello world!`.
- Select a slide in the thumbnail pane, then button 2: it shows `BLOCK = cover`.
- Button 3 lists the slides; the tagged one shows `BLOCK=cover`. Drag that
  slide somewhere else and press button 3 again: the tag moves with it.
- Button 4 opens a new blank presentation.

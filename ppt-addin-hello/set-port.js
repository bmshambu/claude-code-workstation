// Move the dev server to another port: `npm run set-port -- 3002`.
//
// The port lives in package.json (webpack reads it from there) and is written
// out in full in every URL of both manifests, which Office reads as-is. On the
// office laptop port 3000 is held by a corporate agent, and editing four files
// by hand is how one URL gets missed and the pane loads blank.
const fs = require("fs");
const path = require("path");

const port = process.argv[2];
if (!/^\d{2,5}$/.test(port || "")) {
  console.error("usage: npm run set-port -- <port>   e.g. npm run set-port -- 3002");
  process.exit(1);
}

const here = __dirname;
const pkgPath = path.join(here, "package.json");
const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf8"));
const was = pkg.config.dev_server_port;
pkg.config.dev_server_port = Number(port);
fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n");

for (const name of ["manifest.xml", "manifest.json"]) {
  const file = path.join(here, name);
  const text = fs.readFileSync(file, "utf8");
  const out = text.replace(/https:\/\/localhost:\d+/g, "https://localhost:" + port);
  fs.writeFileSync(file, out);
  const n = (text.match(/https:\/\/localhost:\d+/g) || []).length;
  console.log(name + ": " + n + " URL(s) now use port " + port);
}
console.log("package.json: dev_server_port " + was + " -> " + port);
console.log("If the add-in was registered on the old port, run the matching stop script, then start again.");

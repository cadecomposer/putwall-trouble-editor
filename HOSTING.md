# Hosting options

You asked how to put this online without actually doing it yet. Because the
whole tool is static, client-side files (no server code, no database), almost
any way of publishing a web page works. This file compares the realistic
options, with an eye on two constraints: **work may block unfamiliar domains**,
and **you probably should not send order/item data to a third party**.

The single self-contained file is `dist/putwall-trouble-editor.html`. Every
option below can serve that one file, or the whole folder.

A quick note that applies everywhere: **the tool never uploads the spreadsheet**.
It reads and writes the file in the browser. Hosting only decides where the
*page* comes from, not where your data goes.

---

## Option 0 - No hosting at all (recommended first)

Keep `dist/putwall-trouble-editor.html` and open it directly (double-click,
USB stick, or email it to your work address and save it). It needs no internet
and cannot be blocked, because it is not a website at all.

- **Pros:** zero setup, cannot be blocked, works offline, data never leaves the PC.
- **Cons:** you have to move the file around; to update it you replace the file.
  Browsers disagree about local storage for `file://` pages, so the tool's
  "remembered between visits" session may not survive a browser restart. A
  normal `http(s)` host below is more reliable if you rely on that memory.
- **Best for:** getting through the holiday season today.

Many people stop here. The rest are about making it accessible from any device.

---

## Option 1 - Free static hosts (fastest "real URL")

Drag-and-drop the folder or the single HTML file and get an HTTPS URL in under
a minute. No command line needed for some of them.

| Host | How | Free URL looks like |
| --- | --- | --- |
| **Netlify Drop** | Go to `app.netlify.com/drop`, drag the folder in | `something.netlify.app` |
| **Cloudflare Pages** | `pages.cloudflare.com`, connect a Git repo or upload | `something.pages.dev` |
| **Vercel** | `vercel.com`, import a project or use the CLI | `something.vercel.app` |
| **GitHub Pages** | Push the folder to a repo, enable Pages | `you.github.io/...` |
| **Surge** | `npm i -g surge && surge dist` | `something.surge.sh` |

- **Pros:** minutes to set up, free, HTTPS included.
- **Cons:** the domain could be blocked by work, or filters may treat
  `*.netlify.app` / `*.pages.dev` as "unknown". `github.io` is sometimes allowed,
  sometimes not - it depends on your building's policy.
- **Tip:** most of these let you attach a **custom domain** you own. Using a
  custom domain (e.g. `troubles.yourname.com`) plus **Cloudflare** in front
  usually looks like ordinary business web traffic and is less likely to be
  flagged than a random `*.vercel.app` hostname.

```bash
# Example: Cloudflare Pages with the Wrangler CLI
npm i -g wrangler
wrangler pages deploy dist --project-name putwall-tools
```

---

## Option 2 - Host it yourself and expose it with a tunnel

Run a tiny static server on a machine you control (home PC, NAS, Raspberry Pi)
and create an outbound tunnel. No inbound firewall changes, no port forwarding.

| Tool | Command sketch | Notes |
| --- | --- | --- |
| **Cloudflare Tunnel** | `cloudflared tunnel --url http://localhost:8137` | Free, can give a stable hostname on your own domain, good uptime |
| **Tailscale Funnel** | `tailscale funnel 8137` | Free, serves on a `*.ts.net` name; great if the work PC is on Tailscale |
| **ngrok** | `ngrok http 8137` | Fastest to try; free URLs rotate and often show an interstitial warning page |

Serve the folder with Python's built-in server:

```bash
cd "putwall trouble spreadsheet editor"
python3 -m http.server 8137
# in another terminal, for example:
cloudflared tunnel --url http://localhost:8137
```

- **Pros:** you own everything; can add basic-auth; a custom domain looks legit.
- **Cons:** your home machine has to stay on; you are now running a service
  (keep it patched, put a password in front of it).
- **Privacy:** still fully client-side, so only the page is served, not your data.

---

## Option 3 - Google Apps Script web app (`script.google.com`)

Google domains are very often allowed on corporate networks, which makes this
attractive if everything else is blocked. Google Apps Script can serve a static
HTML page with `HtmlService`, and you can inline the same libraries, so the tool
still runs entirely in the browser.

- **Pros:** served from a Google-owned domain; likely reachable; no server to run.
- **Cons:** some setup/porting; Apps Script HTML sandboxing has a few quirks
  (inline everything, avoid external `<script src>`); still needs a Google account.
- **How:** create a new Apps Script project, add an `index.html` with the
  contents of `dist/putwall-trouble-editor.html`, add a `doGet()` that returns
  `HtmlService.createHtmlOutputFromFile('index')`, then **Deploy → Web app**.

This is the option to reach for if Options 1 and 2 are blocked.

---

## Option 4 - Same Wi-Fi only (LAN)

Run the server on a device connected to the same network and open the device's
LAN IP (e.g. `http://192.168.1.50:8137`).

- **Pros:** nothing is exposed to the internet.
- **Cons:** only works on that network; work may isolate guest/personal devices;
  the IP can change.

---

## A useful upgrade once hosted: make it work offline (PWA)

Add a tiny service worker to the hosted copy. Then a browser can **install** it
and keep working even if the domain later gets blocked, or the network drops.
The single-file version already has this property; a hosted version gets it via
the service worker. Worth doing before the busy season.

---

## About "Open in Google Sheets"

The tool intentionally does **not** include an "Open in Google Sheets" button - you
chose to download the `.xlsx` and import it manually, which is the simplest and
most reliable option. For reference, a genuinely one-click "upload and open"
would require either a registered Google OAuth app or a small backend (Apps
Script) with Drive access, because a web page cannot silently push a local file
into your Drive. Both approaches remain available if you ever want them; the
manual **File → Import** flow in Google Sheets handles the `.xlsx` (including the
barcode images) without any extra setup.

---

## Recommendation

1. Start with **Option 0** today - it is immediate and unblockable.
2. When you want a shareable URL, try **Netlify Drop / Cloudflare Pages**
   (Option 1) on a **custom domain**.
3. If work blocks those, move to **Option 2** (Cloudflare Tunnel) or
   **Option 3** (Apps Script).
4. Add the PWA service worker once a host is chosen.

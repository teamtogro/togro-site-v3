# Togro — website (v3)

Marketing site for **Togro**, the real-time countryside intelligence app. Static single-page site, no build step.

## Structure
- `index.html` — the whole site (HTML + CSS + JS inline)
- `images/` — countryside photos, team photos, logo
- `apps-script/` — reference copy of the Google Apps Script that powers the waitlist + launch email automation (deployed separately at script.google.com under team@togro.co)

## Deploy to Netlify
1. Netlify → **Add new site → Import an existing project → GitHub → `teamtogro/togro-site`**
2. Build command: *(none)* · Publish directory: `.`
3. Deploy. The site is served from `index.html` at the root.

Or drag-and-drop this folder into the Netlify dashboard.

## Waitlist
The "Join the waitlist" form POSTs (name, email, user type) to a Google Apps Script web app that writes to the **Togro Waitlist** Google Sheet and sends a welcome email + a daily launch countdown. Endpoint is set in `index.html` as `WAITLIST_ENDPOINT`.

## Launch
App launches **22 June 2026** — the site shows a live countdown and flips all CTAs to "Get the app" automatically on launch day. Set the real App Store / Play link in the `#wlLive` button and the launch-day email's Button URL.

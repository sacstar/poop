# Handoff: put the Attendance Bonus app on 3plview.com

Paste everything below into Claude Code running on **3pl-printsvr01**.

---

You're continuing work started in another Claude Code session. Read all of this before doing anything.

## Goal
Host the **Perfect Attendance Bonus** web app on **3plview.com**. Only users created in the existing site's **admin page** may see it, because it shows private pay information.

## What exists already
- **App code:** GitHub repo `sacstar/poop`, branch `claude/new-session-tdh8vt`, file `attendance-bonus/index.html` (one HTML file, no build step). There's also a `README.md`. The file was edited outside the first session after it was published, so use the current version on that branch.
- **Current version:** published as a private claude.ai page (https://claude.ai/artifact/X9cC1wABCoxJXy86WNZ5ff). Its login check uses claude.ai. Outside claude.ai it stays on the locked screen, so it must be adapted.
- **Hosting:** 3plview.com resolves to Cloudflare IPs. It's probably a Cloudflare Tunnel or proxy in front of this server. Confirm.

## What to do
1. **Look before changing anything.** Find the web server (IIS, nginx, Apache, Node, etc.) and the other sites served under 3plview.com: their folders, bindings, and how traffic reaches them (Cloudflare tunnel config, IIS bindings, reverse proxy). Find the **admin page**: how users are created and stored (database, table, password hashing) and how a logged-in session is checked (cookie, session store, framework auth). Report what you find to the user before making changes.
2. **Plan with the user.** Propose where the bonus app will live (e.g. `3plview.com/bonus`) and how it reuses the admin page's logins. Confirm before touching existing sites, server config, or the Cloudflare setup. Ask before restarting anything.
3. **Adapt the app's gate.** In `attendance-bonus/index.html`, the `start()` function at the bottom unlocks the page only when `window.claude.use('user')` returns a signed-in viewer. Replace that check with the existing site's login:
   - Preferred: serve the page from the site's backend behind its existing auth middleware, so the server never sends the HTML to anyone not logged in. Unauthenticated requests redirect to the existing login page.
   - Don't rely on a client-side check alone. The HTML must not be served without a valid session.
   - Show the logged-in user's name where it currently says "Signed in as …".
   - `downloads` (from `claude.use('downloads')`) won't exist off claude.ai. Replace "Download CSV" with a normal Blob download link. "Copy for Excel" already works anywhere.
   - If users have roles, ask the user which roles may see the bonus page. Consider adding a "Bonus report" permission in the admin page.
4. **Keep data private.** The app reads uploaded timeclock files in the browser only. Don't add server-side storage of punch data unless the user asks. Serve it over HTTPS only, through Cloudflare like the other sites. Add `Cache-Control: no-store` to the page response.
5. **Test.** Logged out, the page redirects to login and the HTML isn't readable. Logged in as an admin-page user, it works. Test with the app's "Load example data" button (fictional employees), not real pay data.
6. Commit the changes to `sacstar/poop` (or wherever the user wants the site code kept) and explain to the user how to add or remove access.

## How the app works (so you don't break it)
- **Inputs:**
  - *Punch_Report export.* CSV or Excel. Header row has `EMP L NAME, EMP F NAME, EMP##, DATE, IN, OUT, TOTAL, DEPT CODE, …`. Each shift is one row spanning the whole shift plus a separate row for lunch. "Manual Edit" in a row means a hand-entered punch.
  - *Optional schedule CSV.* `Last Name, First Name, Date, Start, End, Shift`, with mandatory overtime included in End.
- **Rules (from the flyer, "NOW APPROVED Perfect Attendance Bonus", Northern Nevada 3PL × Bearpaw):**
  - Drivers: $160 weekdays, $144 weekends, $180 nights per week.
  - Warehouse: $80 weekdays, $72 weekends, $100 nights per week.
  - Night amounts include the $0.50/hr shift differential.
  - To qualify: work every scheduled shift and all mandatory overtime in full. No lates, no early outs, no call-outs. Sent home for performance or conduct means no bonus. No exceptions.
- **Pay week:** Sunday through Saturday. Grace is 0 minutes by default, adjustable in Settings.
- **Departments:** `WARE` maps to Warehouse, driver-like codes to Driver, everything else (e.g. `SURGE`, the staffing agency) to Not eligible. Editable in Settings.
- **Statuses:** Qualifies, Does not qualify, Needs review (hand-edited punches, missing clock-out, no schedule, unfinished week, mixed day/night), Not eligible.
- **Overrides:** the reviewer can override role, bonus rate, and decision per employee with a note, e.g. "Sent home for conduct". Settings, roster overrides, and decisions are saved in the browser's localStorage (`apb.*` keys).
- **Meal-break notes:** under 30 min, or no meal within 8 hours (Nevada NRS 608.019). These are compliance notes only and don't affect the bonus.

## Open questions from the first session (raise with the user)
- The flyer counts the $0.50/hr night differential as part of the bonus. If the differential is normally earned regardless of attendance, making it depend on perfect attendance could be a wage problem.
- A promised attendance bonus generally has to be included in the overtime "regular rate" under federal wage law (FLSA). Payroll should account for it.
- Without a schedule, the app can't detect lates, early outs, or call-outs. Find out whether the timeclock system can export schedules.
- Data review of the 9/27–10/3 week flagged:
  - Chuck Burns and Alejandro Castillo: shift punches entered by hand (Manual Edit).
  - Tina Phelps: meals taken more than 8 hours in.
  - Short lunches: Phelps, Sara Ashworth, Burns.

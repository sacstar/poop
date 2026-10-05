# Attendance Bonus Builder

`attendance-bonus/index.html` is a single-page app that reads the timeclock
**Punch_Report** export (CSV or Excel) plus an optional schedule file. It works
out who earns the Perfect Attendance Bonus each Sunday–Saturday pay week, using
the rules on the flyer:

- Drivers: $160 weekdays / $144 weekends / $180 nights per week
- Warehouse: $80 weekdays / $72 weekends / $100 nights per week
- Qualify only by working every scheduled shift and all mandatory overtime in full:
  no lates, no early outs, no call-outs.

The page is published as a private claude.ai artifact. It shows nothing unless
the viewer is signed in to claude.ai with access to it. Opened anywhere else,
including straight from this repo, it stays locked. Uploaded files are read in
the browser only and are never sent anywhere.

Schedule CSV columns: `Last Name, First Name, Date, Start, End, Shift`
(put mandatory overtime in the End time; Shift is Weekday, Weekend or Nights).

# NookBird sign-ups (Google Apps Script)

`_workspace/` is not published on the website (GitHub Pages skips folders that start with `_`).

`signup.gs` receives the website's boxes: the waitlist (home page) and the testers' applications and answers
(`pilot.html`). It keeps one row per person in the private Sheet's `People` tab, sends the approved thank-you email once
from hello@nookbird.com, emails the owner at each tester's answers and every morning at 08:00 Melbourne time, and keeps a
separate `Counts` Sheet file with numbers only, for operations. A tester's answers are accepted only with the one-time
code made at Apply (the Sheet keeps its SHA-256 fingerprint, in a hidden column). The steps to set it up are in the
planning thread's message to the owner. Its web app address is the `action=` of `#waitlist-form` (`index.html`) and of
`#apply-form` and `#questions` (`pilot.html`).

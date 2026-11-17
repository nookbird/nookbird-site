# NookBird sign-ups (Google Apps Script)

`_workspace/` is not published on the website (GitHub Pages skips folders that start with `_`).

`signup.gs` receives the website's email box (the waitlist now, the testers later), keeps each address in a private Sheet
and sends the approved thank-you email once, from hello@nookbird.com. The steps to set it up are in the planning
thread's message to the owner. Its web app address goes into `index.html` (`action=` of `#waitlist-form`).

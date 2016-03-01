A simple tray icon app which will show a notification when there are new gmails.

Run it
------

- Run via GmailNotifier.exe
- If this is the first time you run it a page should come up in your browser, asking for authorization. This will only be needed once.
- You need to select the account to check (only one supported)
- Double-click on the tray icon as a short-cut to open https://mail.google.com/ in the default browser
- Right-click on tray icon for menu options, like Re-authentication and to Exit the app
- A grey icon means no unread message. A read icon means there are 1 or more unread messages.
- It checks every 1 to 5 minutes if there are new messages.

Build
-----

Just grab the pre-build one from https://github.com/PLYoung/GMailNotifier/releases or download the source and build it yourself.

You will need to add credentials to /GmailNotifier/client_id.json. This can be created via the Google Developers Console.
Check this link for a guide: https://developers.google.com/gmail/api/quickstart/dotnet
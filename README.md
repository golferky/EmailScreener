# EmailScreener

EmailScreener is a VB.NET Windows Forms application that screens unread Gmail and Yahoo messages, stores message metadata in SQLite, and can automatically forward messages from configured senders or domains to Gmail.

The running version is displayed in the form title bar using the `yyyyMMdd.N` build format. Version 20260811.1 adds SQLite-backed mail-account management while keeping app passwords in per-user environment settings.

## Secure configuration

Credentials are never stored in the repository. Configure them as Windows user environment variables before starting the application:

```powershell
[Environment]::SetEnvironmentVariable("EMAILSCREENER_GMAIL_USER", "your-address@gmail.com", "User")
[Environment]::SetEnvironmentVariable("EMAILSCREENER_GMAIL_APP_PASSWORD", "your-new-app-password", "User")
[Environment]::SetEnvironmentVariable("EMAILSCREENER_YAHOO_USER", "your-yahoo-user", "User")
[Environment]::SetEnvironmentVariable("EMAILSCREENER_YAHOO_APP_PASSWORD", "your-new-app-password", "User")
[Environment]::SetEnvironmentVariable("EMAILSCREENER_SQLITE_PATH", "C:\EmailScreener\GarysEmails.db", "User")
```

Restart EmailScreener after changing environment variables. The optional legacy SQL Server migration connection can be supplied as `EMAILSCREENER_SQL_CONNECTION`.

Use provider-specific app passwords rather than primary account passwords. Revoke every credential that was previously committed before using the application again.

## Mail accounts

Open the **Mail accounts** tab to add, edit, enable, disable, or delete Gmail, Yahoo, and custom IMAP accounts. For a second Gmail login:

1. Choose **Gmail** and give the account a unique name such as `Gmail 2`.
2. Enter the complete Gmail address and that Google account's app password.
3. Click **Add account**.
4. Return to **Email screening**, select the new account, and click **Connect to Email**.

Account names, usernames, IMAP/SMTP servers, and enabled status are stored in the `MailAccounts` SQLite table. App passwords are not stored in SQLite: the screen writes each one to a uniquely named Windows user environment variable and the grid shows only `Configured` or `Missing`. Leaving the app-password box blank while editing preserves the existing password. Deleting an account does not delete its password environment variable.

## Auto forwarding

Open the **Auto forwarding** tab to:

1. Add an exact sender address, a name-or-email-contains rule, or a sender-domain rule.
2. Enter the destination Gmail address.
3. Enable the rule and automatic forwarding.

The destination is prefilled from the first enabled Gmail account but remains editable for a one-off destination. Editing it does not change the saved mail account or persist the override.

Forwarding runs while unread mail is screened. The complete original message is attached to the forwarded message, preserving its content and attachments. Successful sends are recorded by account, folder, IMAP UID, and destination so a message is not forwarded twice.

To search older mail safely, choose a **Search since** date and click **Preview matches**. Previewing sends nothing; Yahoo/Gmail first filters candidates by sender and date on the server, then EmailScreener verifies the sender locally. The preview groups matches by sender display name and email address and shows the message count plus oldest/newest dates. After reviewing the summary, enter the Gmail destination and click **Forward previewed**, which requires a final confirmation. The dated preview checks both read and unread inbox messages and uses the same duplicate protection when forwarding. Original plain-text content is included inline so downstream Gmail and calendar workflows can read it, while the complete original remains attached. Use the Cancel button on either tab to stop the active screening, preview, or forwarding pass after its current message finishes.

The selected source account's SMTP server and app password are used to send forwarded messages.

## Removing leaked credentials from Git history

Changing the current files does not remove credentials from old commits. After revoking the old credentials, rewrite the repository history with `git filter-repo` or BFG, force-push the rewritten branches and tags, and have every existing clone re-clone or carefully reset to the rewritten history.

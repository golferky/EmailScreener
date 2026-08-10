# EmailScreener

EmailScreener is a VB.NET Windows Forms application that screens unread Gmail and Yahoo messages, stores message metadata in SQLite, and can automatically forward messages from configured senders or domains to Gmail.

The running version is displayed in the form title bar. Version 1.1.0 adds secure configuration, forwarding rules, sender/date previews, cancellation, and confirmed forwarding.

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

## Auto forwarding

Open the **Auto forwarding** tab to:

1. Add an exact sender address, sender display name, or sender-domain rule.
2. Enter the destination Gmail address.
3. Enable the rule and automatic forwarding.

The destination is prefilled from `EMAILSCREENER_GMAIL_USER` but remains editable for a one-off destination. Editing it does not change the environment variable or persist the override.

Forwarding runs while unread mail is screened. The complete original message is attached to the forwarded message, preserving its content and attachments. Successful sends are recorded by account, folder, IMAP UID, and destination so a message is not forwarded twice.

To search older mail safely, choose a **Search since** date and click **Preview matches**. Previewing sends nothing; Yahoo/Gmail first filters candidates by sender and date on the server, then EmailScreener verifies the sender locally. The preview groups matches by sender display name and email address and shows the message count plus oldest/newest dates. After reviewing the summary, enter the Gmail destination and click **Forward previewed**, which requires a final confirmation. The dated preview checks both read and unread inbox messages and uses the same duplicate protection when forwarding. Original plain-text content is included inline so downstream Gmail and calendar workflows can read it, while the complete original remains attached. Use the Cancel button on either tab to stop the active screening, preview, or forwarding pass after its current message finishes.

The selected source account's SMTP server and app password are used to send forwarded messages.

## Removing leaked credentials from Git history

Changing the current files does not remove credentials from old commits. After revoking the old credentials, rewrite the repository history with `git filter-repo` or BFG, force-push the rewritten branches and tags, and have every existing clone re-clone or carefully reset to the rewritten history.

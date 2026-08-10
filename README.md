# EmailScreener

EmailScreener is a VB.NET Windows Forms application that screens unread Gmail and Yahoo messages, stores message metadata in SQLite, and can automatically forward messages from configured senders or domains to Gmail.

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

1. Add an exact sender or sender-domain rule.
2. Enter the destination Gmail address.
3. Enable the rule and automatic forwarding.

Forwarding runs while unread mail is screened. The complete original message is attached to the forwarded message, preserving its content and attachments. Successful sends are recorded by account, folder, IMAP UID, and destination so a message is not forwarded twice.

The selected source account's SMTP server and app password are used to send forwarded messages.

## Removing leaked credentials from Git history

Changing the current files does not remove credentials from old commits. After revoking the old credentials, rewrite the repository history with `git filter-repo` or BFG, force-push the rewritten branches and tags, and have every existing clone re-clone or carefully reset to the rewritten history.

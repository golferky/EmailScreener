Imports MailKit.Net.Smtp
Imports MailKit.Security
Imports MimeKit

Public Class ForwardingService
    Public Function ForwardMatching(message As MimeMessage,
                                    uid As Integer,
                                    accountName As String,
                                    folder As String,
                                    smtpHost As String,
                                    smtpUser As String,
                                    smtpPassword As String,
                                    rules As IEnumerable(Of ForwardingRule),
                                    repository As ForwardingRepository) As Integer
        Dim matchingRules = rules.Where(Function(rule) RuleMatches(rule, message)).ToList()
        If matchingRules.Count = 0 Then Return 0
        If String.IsNullOrWhiteSpace(smtpHost) OrElse String.IsNullOrWhiteSpace(smtpUser) OrElse String.IsNullOrWhiteSpace(smtpPassword) Then
            Throw New InvalidOperationException("SMTP settings are missing for the selected email account.")
        End If

        Dim forwardedCount = 0
        Using smtp As New SmtpClient()
            smtp.Connect(smtpHost, 587, SecureSocketOptions.StartTls)
            smtp.Authenticate(smtpUser, smtpPassword)

            For Each rule In matchingRules
                If repository.IsForwarded(accountName, folder, uid, rule.Destination) Then Continue For
                smtp.Send(BuildForwardedMessage(message, smtpUser, rule.Destination))
                repository.RecordForwarded(accountName, folder, uid, rule.Destination)
                forwardedCount += 1
            Next

            smtp.Disconnect(True)
        End Using
        Return forwardedCount
    End Function

    Private Shared Function RuleMatches(rule As ForwardingRule, message As MimeMessage) As Boolean
        If Not rule.Enabled OrElse message.From Is Nothing Then Return False
        Dim sender = message.From.Mailboxes.FirstOrDefault()
        If sender Is Nothing OrElse String.IsNullOrWhiteSpace(sender.Address) Then Return False

        If String.Equals(rule.MatchType, "Sender", StringComparison.OrdinalIgnoreCase) Then
            Return String.Equals(sender.Address, rule.MatchValue.Trim(), StringComparison.OrdinalIgnoreCase)
        End If

        If String.Equals(rule.MatchType, "Domain", StringComparison.OrdinalIgnoreCase) Then
            Dim separator = sender.Address.LastIndexOf("@"c)
            If separator < 0 Then Return False
            Return String.Equals(sender.Address.Substring(separator + 1), rule.MatchValue.Trim().TrimStart("@"c), StringComparison.OrdinalIgnoreCase)
        End If

        Return False
    End Function

    Private Shared Function BuildForwardedMessage(original As MimeMessage, fromAddress As String, destination As String) As MimeMessage
        Dim forwarded As New MimeMessage()
        forwarded.From.Add(New MailboxAddress("EmailScreener", fromAddress))
        forwarded.To.Add(MailboxAddress.Parse(destination))
        forwarded.Subject = $"Fwd: {If(original.Subject, "(no subject)")}"

        Dim originalSender = original.From.Mailboxes.FirstOrDefault()
        If originalSender IsNot Nothing Then forwarded.ReplyTo.Add(originalSender)

        Dim intro As New TextPart("plain") With {
            .Text = $"Forwarded automatically by EmailScreener.{Environment.NewLine}" &
                    $"Original sender: {original.From}{Environment.NewLine}" &
                    $"Original date: {original.Date}{Environment.NewLine}{Environment.NewLine}" &
                    "The complete original message is attached."
        }
        Dim originalPart As New MessagePart With {.Message = original}
        Dim body As New Multipart("mixed") From {intro, originalPart}
        forwarded.Body = body
        Return forwarded
    End Function
End Class

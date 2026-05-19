Imports System.Net.Mail
Public Class GGSMTP_GMAIL
    Dim Temp_GmailAccount As String
    Dim Temp_GmailPassword As String
    Dim Temp_SMTPSERVER As String
    Dim Temp_ServerPort As Int32
    Dim Temp_ErrorText As String = ""
    Dim Temp_EnableSSl As Boolean = True
    Public ReadOnly Property ErrorText() As String
        Get
            Return Temp_ErrorText
        End Get
    End Property
    Public Property EnableSSL() As Boolean
        Get
            Return Temp_EnableSSl
        End Get
        Set(ByVal value As Boolean)
            Temp_EnableSSl = value
        End Set
    End Property
    Public Property GmailAccount() As String
        Get
            Return Temp_GmailAccount
        End Get
        Set(ByVal value As String)
            Temp_GmailAccount = value
        End Set
    End Property
    Public Property GmailPassword() As String
        Get
            Return Temp_GmailPassword
        End Get
        Set(ByVal value As String)
            Temp_GmailPassword = value
        End Set
    End Property
    Public Property SMTPSERVER() As String
        Get
            Return Temp_SMTPSERVER
        End Get
        Set(ByVal value As String)
            Temp_SMTPSERVER = value
        End Set
    End Property
    Public Property ServerPort() As Int32
        Get
            Return Temp_ServerPort
        End Get
        Set(ByVal value As Int32)
            Temp_ServerPort = value
        End Set
    End Property
    Public Sub New(ByVal GmailAccount As String, ByVal GmailPassword As String, Optional ByVal SMTPSERVER As String = "smtp.gmail.com", Optional ByVal ServerPort As Int32 = 587, Optional ByVal EnableSSl As Boolean = True)
        Temp_GmailAccount = GmailAccount
        Temp_GmailPassword = GmailPassword
        Temp_SMTPSERVER = SMTPSERVER
        Temp_ServerPort = ServerPort
        Temp_EnableSSl = EnableSSl
        'https://myaccount.google.com/u/1/lesssecureapps?pli=1&rapt=AEjHL4MbakEKQccI8V0HBVpniQ4GotT_5ld_-dDbKcYXcuRldAOinDF60SeFsfmsKA7g2oFBQaonOmZRhXSz1C3a6As-sZPIoA
    End Sub
    'Public Function SendMail(ByVal ToAddressies As String(), ByVal Subject As String, ByVal BodyText As String, Optional ByVal AttachedFiles As String() = Nothing) As Boolean
    Public Function SendMail(ByVal ToAddressies As List(Of String), ByVal Subject As String, ByVal BodyText As String, Optional ByVal AttachedFiles As String() = Nothing) As Boolean
        Temp_ErrorText = ""
        Dim Mail As New MailMessage
        Dim SMTP As New SmtpClient(Temp_SMTPSERVER)
        Mail.Subject = Subject
        Mail.From = New MailAddress(Temp_GmailAccount)
        SMTP.Credentials = New System.Net.NetworkCredential(Temp_GmailAccount, Temp_GmailPassword) '<-- Password Here
        Mail.To.Clear()
        'For i As Int16 = 0 To ToAddressies.Length - 1
        '    Mail.To.Add(ToAddressies(i))
        'Next i
        For Each addr In ToAddressies
            Mail.To.Add(addr)
        Next
        Mail.Body = BodyText
        Mail.Attachments.Clear()

        If AttachedFiles IsNot Nothing Then
            For i As Int16 = 0 To AttachedFiles.Length - 1
                Mail.Attachments.Add(New Attachment(AttachedFiles(i)))
            Next
        End If

        SMTP.EnableSsl = Temp_EnableSSl
        SMTP.Port = Temp_ServerPort

        Try
            SMTP.Send(Mail)
            Return True
        Catch ex As Exception
            Me.Temp_ErrorText = ex.Message.ToString
            Return False
        End Try

    End Function
End Class
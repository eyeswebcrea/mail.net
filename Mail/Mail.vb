Imports System.Net
Imports System.Net.Mail
Imports System.Configuration

Public Class Mail

    Private objMessage As MailMessage = New MailMessage()
    Private Shared objSMPTClient As New SmtpClient()

    Public Sub New()

        setSmtpAdress(ConfigurationSettings.AppSettings("mail.smtpServer"), ConfigurationSettings.AppSettings("mail.smtpPort"))

        If (Not String.IsNullOrEmpty(ConfigurationSettings.AppSettings("mail.smtpUsername"))) Then

            Console.WriteLine("Utilisatation de l'authentitication smtp")

            Dim username As New String(ConfigurationSettings.AppSettings("mail.smtpUsername"))
            Dim password As New String(ConfigurationSettings.AppSettings("mail.smtpPassword"))

            setCreditentialSmtp(username, password)

        End If


    End Sub

    Public Function setSender(ByVal mail As String) As Mail
        Me.objMessage.From = New MailAddress(mail)

        Return Me
    End Function

    Public Shared Function getSmtpClient() As SmtpClient
        Return objSMPTClient
    End Function

    Public Function getObjMessage() As MailMessage
        Return Me.objMessage
    End Function

    Public Sub isHtml(ByVal value As Boolean)
        Me.objMessage.IsBodyHtml = value
    End Sub

    Public Shared Sub setSmtpAdress(ByVal address As String, ByVal port As Integer)
        objSMPTClient.Host = address
        objSMPTClient.Port = port
    End Sub

    Public Shared Sub activeSSL()
        objSMPTClient.EnableSsl = True
    End Sub

    Public Shared Sub setCreditentialSmtp(ByVal user As String, ByVal password As String)
        objSMPTClient.Credentials = New NetworkCredential(user, password)
    End Sub

    Public Function setRecipient(ByVal mail As String) As Mail
        Me.objMessage.To.Clear()
        Me.addRecipient(mail)

        Return Me
    End Function

    Public Function addRecipient(ByVal mail As String) As Mail
        Me.objMessage.To.Add(New MailAddress(mail))
        Return Me
    End Function

    Public Function setBody(ByVal message As String) As Mail
        Me.objMessage.Body = message

        Return Me
    End Function

    Public Function setSubject(ByVal value As String) As Mail
        Me.objMessage.Subject = value

        Return Me
    End Function

    Public Sub send()
        objSMPTClient.Send(Me.objMessage)
    End Sub

    Public Sub send(ByVal message As String)
        Me.setBody(message).send()
    End Sub

    Public Sub send(ByVal message As String, ByVal recipient As String)
        Me.setRecipient(recipient).setBody(message).send()
    End Sub

End Class

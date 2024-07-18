Imports System.ComponentModel
Imports System.IO
Imports System.Net.Mail
Imports System.Text

Public Class Form1
    Dim EmailSubject As String
    Dim em As New EmailHandler
    Dim dsEmail As DataSet = em.GetMailRecipients(141)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SendShortMessage()
    End Sub

    Private Sub SendShortMessage()
        EmailSubject = "Test Message"

        Dim message As String = "This is a test email from Form1."

        ' Send email
        em.SendEmail(EmailSubject, message, "", dsEmail)
    End Sub
End Class

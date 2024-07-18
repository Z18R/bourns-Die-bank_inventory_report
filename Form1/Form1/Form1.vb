Imports System.ComponentModel
Imports System.Data
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text

Public Class Form1
    Dim EmailSubject As String
    Dim em As New EmailHandler
    Dim dsEmail As DataSet
    Dim filename, filePath, filesave, Sqlemail, backup, Datenow As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Optionally, you can call Nexgen or Bourns methods here based on your requirement
            'Nexgen()
            Bourns()
            SendEmail(EmailSubject, CreateMsgBody(), filesave)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Me.Close()
    End Sub

    Private Sub Bourns()
        Datenow = Date.Now().ToString("yyyyMMdd")
        filename = "\\192.168.6.16\10.PC_common\DBI\Bourns DBI\BOURNS DIEBANK-INVENTORY.xlsx"
        filesave = Application.StartupPath & "\backup\Bourns\BOURNS DIEBANK-INVENTORY " & Datenow & ".xlsx"
        System.IO.File.Move(filename, filesave)

        EmailSubject = "Bourns Die Bank Inventory " & Datenow
        Sqlemail = "usp_SPT_AutoEmail_GetRecipients 146"
        'dsEmail = em.GetMailRecipients(127)
        dsEmail = em.GetMailRecipients(146)
    End Sub

    Private Sub SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal file As String)
        Try
            em.SendEmail(strSubject, strMessage, file, dsEmail)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function CreateMsgBody() As String
        Dim msgBody As String = ""
        msgBody &= "<html><body><pre>"
        msgBody &= "<H3>" & EmailSubject & " "
        msgBody &= "<pre>Hello Team </pre>"
        msgBody &= "<br><pre>See attached file for " & EmailSubject & "</pre>"
        msgBody &= "<br><br><pre>This auto mail is still in Beta Test"
        msgBody &= "<br><pre>DO NOT REPLY to this Mail"
        msgBody &= "</pre></body></html>"
        Return msgBody
    End Function
End Class
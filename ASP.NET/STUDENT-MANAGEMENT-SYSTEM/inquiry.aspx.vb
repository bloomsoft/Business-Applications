Imports System.Web.Mail
Public Class inquiry
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub
    Private Sub ShowSentConfirmation()
        Button1.Visible = False
        Button2.Visible = False
        Label5.Text = "We Have Received Your Query! We will Contact you Soon"

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'Dim NM As New System.Web.Mail.MailMessage
            'Dim MBody As String

            'NM.To = "zahid@whiteroseschool.com"
            'NM.Subject = "Online Inquiry Recvd. On Web Site http://www.WiteRoseSchool.com"
            'NM.From = TextBox1.Text
            'NM.Body = "Name : " & TextBox2.Text & vbCrLf & "Phone : " & TextBox3.Text & vbCrLf & " Message : " & TextBox4.Text

            ''System.Web.Mail.SmtpMail.SmtpServer = "smtp.naxfashion.com"
            'System.Web.Mail.SmtpMail.SmtpServer = "relay-hosting.secureserver.net"
            'System.Web.Mail.SmtpMail.Send(NM)
            Const SERVER As String = "relay-hosting.secureserver.net"
            Dim OMail As New MailMessage
            OMail.From = TextBox1.Text
            OMail.To = "zahid@whiteroseschool.com"
            'OMail.Cc = "zahid@whiteroseschool.com"
            OMail.Subject = "Online Inquiry Recvd. On Web Site http://www.WiteRoseSchool.com"
            OMail.BodyFormat = MailFormat.Html ' // enumeration
            OMail.Priority = MailPriority.High ' // enumeration
            OMail.Body = "Name : " & TextBox2.Text & vbCrLf & "Phone : " & TextBox3.Text & vbCrLf & " Message : " & TextBox4.Text

            SmtpMail.SmtpServer = SERVER
            SmtpMail.Send(OMail)
            OMail = Nothing

            ShowSentConfirmation()
        Catch Ex As Exception
            Label5.Text = (Ex.Message & " " & Ex.Source.ToString)
        End Try
    End Sub
End Class

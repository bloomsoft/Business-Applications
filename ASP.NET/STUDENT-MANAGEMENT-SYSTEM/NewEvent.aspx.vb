Imports System.IO
Public Class IbrarEvent
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents txtimgNo As System.Web.UI.WebControls.TextBox

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region
    Dim ImgNo As Int16
    Private Sub ShowImage()
        Dim S As String
        Dim I As Int16
        Dim Files() As String

        Files = Directory.GetFiles(Server.MapPath("/EventSnaps"))
        For Each S In Files

            If I = ImgNo Then
                Dim FI As New FileInfo(S)
                txtimgNo.Text = I
                Image1.ImageUrl = "EventSnaps/" & FI.Name
                Exit For
            End If
            I += 1
        Next
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        ImgNo = Val(txtimgNo.Text)
        'If ImgNo = 0 Then ImgNo = -1
        ShowImage()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtimgNo.Text = Val(txtimgNo.Text) - 1
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtimgNo.Text = Val(txtimgNo.Text) + 1
    End Sub
End Class

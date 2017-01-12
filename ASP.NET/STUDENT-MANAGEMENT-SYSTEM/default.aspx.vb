Imports System.IO
Public Class _default
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Cn1 = New System.Data.SqlClient.SqlConnection
        '
        'Cn1
        '
        Me.Cn1.ConnectionString = "workstation id=HRBLM;packet size=4096;integrated security=SSPI;data source=HRBLM;" & _
        "persist security info=False;initial catalog=zakria"

    End Sub
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents txtRegNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtpin As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblError As System.Web.UI.WebControls.Label
    Protected WithEvents lblWelcome As System.Web.UI.WebControls.Label
    Protected WithEvents linkremarks As System.Web.UI.WebControls.HyperLink
    Protected WithEvents linkResults As System.Web.UI.WebControls.HyperLink
    Protected WithEvents lblNews1 As System.Web.UI.WebControls.Label
    Protected WithEvents lblNews1Body As System.Web.UI.WebControls.Label
    Protected WithEvents lblNews2 As System.Web.UI.WebControls.Label
    Protected WithEvents lblNews3 As System.Web.UI.WebControls.Label
    Protected WithEvents lblNews2Body As System.Web.UI.WebControls.Label
    Protected WithEvents lblNews3Body As System.Web.UI.WebControls.Label
    Protected WithEvents logintable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents tblimage As System.Web.UI.WebControls.Table
    Protected WithEvents Cn1 As System.Data.SqlClient.SqlConnection

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region
    Private Sub Readnews1()
        Try
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim S As String

            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If

            S = "Select * from News where NewsID=1"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("NewsTitle")) Then
                    lblNews1.Text = DR.Item("NewsTitle")
                    lblNews1Body.Text = DR.Item("NewsBody")
                End If
            End If
            DR.Close()
            Cn1.Close()
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try
    End Sub
    Private Sub Readnews2()
        Try
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim S As String

            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If

            S = "Select * from News where NewsID=2"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("NewsTitle")) Then
                    lblNews2.Text = DR.Item("NewsTitle")
                    lblNews2Body.Text = DR.Item("NewsBody")
                End If
            End If
            DR.Close()
            Cn1.Close()
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try
    End Sub
    Private Sub Readnews3()
        Try
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim S As String

            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If

            S = "Select * from News where NewsID=3"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("NewsTitle")) Then
                    lblNews3.Text = DR.Item("NewsTitle")
                    lblNews3Body.Text = DR.Item("NewsBody")
                End If
            End If
            DR.Close()
            Cn1.Close()
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try
    End Sub
    Private Sub SignoutProc(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Response.Redirect("Signout.aspx")
    End Sub
    Private Sub LoginUser()
        Try
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim S As String

            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If

            S = "Select a.*,b.Name as ClassName,c.SectionName from Students a,ClassMain b,[Section] c where a.ClassCode=b.Code and a.SectionCode=c.SectionCode and a.AdmNo=" & txtRegNo.Text
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If DR.Item("Password") = txtpin.Text Then
                    lblError.Text = ""
                    lblWelcome.Text = "Welcome! Dear Parents of " & DR.Item("Name") & " ...."
                    linkResults.Visible = True
                    linkResults.NavigateUrl = "ResultCard.aspx?AdmNo=" & txtRegNo.Text
                    linkremarks.Visible = True
                    linkremarks.NavigateUrl = "TeacherRemarks.aspx?AdmNo=" & txtRegNo.Text
                    logintable.Visible = False

                    Dim R As New TableRow
                    Dim C As New TableCell
                    Dim C2 As New TableCell
                    Dim P As String
                    Dim I As Int16
                    Dim SB As New Button
                    If Not IsDBNull(DR.Item("Photo")) Then
                        I = InStrRev(DR.Item("Photo"), "\")
                        If I > 0 Then
                            P = Right(DR.Item("Photo"), Len(DR.Item("Photo")) - I)
                        End If
                        'C.Controls.Add(New LiteralControl("<Img Src=stdimages/" & P & " Height=110 Width=100>"))
                        C2.Controls.Add(New LiteralControl("<Font Face=verdana Size=3 Color=white><B>" & DR.Item("Name") & "<BR>" & DR.Item("ClassName") & "<BR>" & DR.Item("SectionName") & "</B></Font><BR><BR><BR>"))
                        SB.Text = "Signout"

                        AddHandler SB.Click, AddressOf SignoutProc

                        C2.Controls.Add(SB)


                    End If
                    R.Cells.Add(C)
                    R.Cells.Add(C2)
                    tblimage.BorderStyle = BorderStyle.None
                    tblimage.Rows.Add(R)
                    Session.RemoveAll()
                    Session.Add("StudentID", txtRegNo.Text)
                Else
                    lblError.Text = "Invalid Registration Number or Pin"
                End If
            Else
                lblError.Text = "Invalid Registration Number"
            End If
            DR.Close()
            Cn1.Close()
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        Try
            'If Page.IsPostBack = False Then
            If Session.Item("StudentID") <> "" Then
                txtRegNo.Text = Session.Item("StudentID")
                txtpin.Text = "zakria"
                LoginUser()
            End If
            'End If
            Readnews1()
            Readnews2()
            Readnews3()
        Catch Ex As Exception
            lblError.Text = Ex.Message
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If txtRegNo.Text.Length <= 0 Then
            lblError.Text = "Please Enter Registration No."
            Exit Sub
        End If
        If txtpin.Text.Length <= 0 Then
            lblError.Text = "Please Enter Pin"
            Exit Sub
        End If
        LoginUser()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtRegNo.Text = ""
        txtpin.Text = ""

    End Sub
End Class

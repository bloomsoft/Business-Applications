Public Class Resultcard
    Inherits System.Web.UI.Page
    Dim OutOfTotal As Double
    Protected WithEvents Cn1 As System.Data.SqlClient.SqlConnection
    Protected WithEvents DaResults As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents DsResults1 As Whiterose.DsResults
    Dim ObtainedTotal As Double
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Cn1 = New System.Data.SqlClient.SqlConnection
        Me.DaResults = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.DsResults1 = New Whiterose.DsResults
        CType(Me.DsResults1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Cn1
        '
        Me.Cn1.ConnectionString = "workstation id=HRBLM;packet size=4096;integrated security=SSPI;data source=HRBLM;" & _
        "persist security info=False;initial catalog=zakria"
        '
        'DaResults
        '
        Me.DaResults.SelectCommand = Me.SqlSelectCommand1
        Me.DaResults.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "ExamResultVW", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("RegNo", "RegNo"), New System.Data.Common.DataColumnMapping("ExamYear", "ExamYear"), New System.Data.Common.DataColumnMapping("ExamTerm", "ExamTerm"), New System.Data.Common.DataColumnMapping("Subject", "Subject"), New System.Data.Common.DataColumnMapping("Obtained", "Obtained"), New System.Data.Common.DataColumnMapping("Outof", "Outof"), New System.Data.Common.DataColumnMapping("Grade", "Grade"), New System.Data.Common.DataColumnMapping("ExamDate", "ExamDate"), New System.Data.Common.DataColumnMapping("STUDENTS.NAME", "STUDENTS.NAME"), New System.Data.Common.DataColumnMapping("FRNAME", "FRNAME"), New System.Data.Common.DataColumnMapping("SECTIONNAME", "SECTIONNAME"), New System.Data.Common.DataColumnMapping("CLASSMAIN.NAME", "CLASSMAIN.NAME")})})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT RegNo, ExamYear, ExamTerm, Subject, Obtained, Outof, Grade, ExamDate, [STU" & _
        "DENTS.NAME], FRNAME, SECTIONNAME, [CLASSMAIN.NAME] FROM ExamResultVW"
        Me.SqlSelectCommand1.Connection = Me.Cn1
        '
        'DsResults1
        '
        Me.DsResults1.DataSetName = "DsResults"
        Me.DsResults1.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.DsResults1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents lblStdName As System.Web.UI.WebControls.Label
    Protected WithEvents txtAdmNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents lstYears As System.Web.UI.WebControls.DropDownList
    Protected WithEvents lstTerms As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub ShowStudentName()
        Dim Ssql As String
        Dim DR As SqlClient.SqlDataReader
        Dim Cmd As New SqlClient.SqlCommand

        If Cn1.State = ConnectionState.Open Then
            Cn1.Close()
            Cn1.ConnectionString = Bloom1.ConnectionString
            Cn1.Open()
        Else
            Cn1.ConnectionString = Bloom1.ConnectionString
            Cn1.Open()
        End If

        Ssql = "Select * from Students where admNo=" & Val(txtAdmNo.Text)
        Cmd.CommandText = Ssql
        Cmd.CommandType = CommandType.Text
        Cmd.Connection = Cn1

        DR = Cmd.ExecuteReader
        If DR.Read Then
            lblStdName.Text = DR.Item("Name") & " | Father Name : " & DR.Item("FRname")
        End If
        DR.Close()
    End Sub
    Private Sub BindGrid()

        Dim Ssql As String
        Ssql = "SELECT * FROM ExamResultVW "
        Ssql &= " where RegNo=" & txtAdmNo.Text
        Ssql &= " and ExamYear=" & lstYears.SelectedItem.Value
        Ssql &= " and ExamTerm=" & lstTerms.SelectedItem.Value
        Ssql &= " Order by Subject"
        'Response.Write(Ssql)
        'Response.End()
        ' Exit Sub
        DaResults.SelectCommand.CommandText = Ssql
        'lblError.Text = Ssql
        'DaInbox.SelectCommand.Parameters.Item(0).Value = Session.Item("CurrentUsername")
        'DaInbox.SelectCommand.Parameters.Item(1).Value = txtFolderType.Text
        DaResults.Fill(DsResults1) ', "Inbox")
        DataGrid1.DataBind()
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        If Page.IsPostBack = False Then
            Dim R As Integer
            lstYears.Items.Clear()
            For R = 2007 To 2008
                lstYears.Items.Add(New ListItem(CStr(R), R))
            Next
            lstTerms.Items.Clear()
            For R = 1 To 6
                lstTerms.Items.Add(New ListItem(CStr(R), R))
            Next
        End If
        If Request.QueryString("AdmNo") <> "" Then
            txtAdmNo.Text = Request.QueryString("AdmNo")
            ShowStudentName()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        BindGrid()
    End Sub

    Private Sub DataGrid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            OutOfTotal += Val(e.Item.Cells(1).Text)
            ObtainedTotal += Val(e.Item.Cells(2).Text)
        ElseIf e.Item.ItemType = ListItemType.Footer Then
            e.Item.Cells(0).Text = "Totals"
            e.Item.Cells(1).Text = OutOfTotal
            e.Item.Cells(2).Text = ObtainedTotal
        End If
    End Sub
End Class

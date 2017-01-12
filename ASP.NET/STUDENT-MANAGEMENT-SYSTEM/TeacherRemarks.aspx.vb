Public Class TeacherRemarks
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.Cn1 = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.DaRemarks = New System.Data.SqlClient.SqlDataAdapter
        Me.DsRemarks1 = New Whiterose.DsRemarks
        CType(Me.DsRemarks1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT ADMNO, REMARKS, TEACHERNAME, RDate, [STUDENTS.NAME], FRNAME, RollNo, [CLAS" & _
        "SMAIN.NAME], SECTIONNAME FROM RemrksVW"
        Me.SqlSelectCommand1.Connection = Me.Cn1
        '
        'Cn1
        '
        Me.Cn1.ConnectionString = "workstation id=HRBLM;packet size=4096;integrated security=SSPI;data source=HRBLM;" & _
        "persist security info=False;initial catalog=zakria"
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO RemrksVW(ADMNO, REMARKS, TEACHERNAME, RDate, [STUDENTS.NAME], FRNAME," & _
        " RollNo, [CLASSMAIN.NAME], SECTIONNAME) VALUES (@ADMNO, @REMARKS, @TEACHERNAME, " & _
        "@RDate, @Param1, @FRNAME, @RollNo, @Param2, @SECTIONNAME); SELECT ADMNO, REMARKS" & _
        ", TEACHERNAME, RDate, [STUDENTS.NAME], FRNAME, RollNo, [CLASSMAIN.NAME], SECTION" & _
        "NAME FROM RemrksVW"
        Me.SqlInsertCommand1.Connection = Me.Cn1
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ADMNO", System.Data.SqlDbType.Int, 4, "ADMNO"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@REMARKS", System.Data.SqlDbType.NVarChar, 250, "REMARKS"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TEACHERNAME", System.Data.SqlDbType.NVarChar, 30, "TEACHERNAME"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RDate", System.Data.SqlDbType.DateTime, 4, "RDate"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Param1", System.Data.SqlDbType.NVarChar, 35, "STUDENTS.NAME"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FRNAME", System.Data.SqlDbType.NVarChar, 35, "FRNAME"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RollNo", System.Data.SqlDbType.SmallInt, 2, "RollNo"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Param2", System.Data.SqlDbType.NVarChar, 25, "CLASSMAIN.NAME"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SECTIONNAME", System.Data.SqlDbType.NVarChar, 10, "SECTIONNAME"))
        '
        'DaRemarks
        '
        Me.DaRemarks.InsertCommand = Me.SqlInsertCommand1
        Me.DaRemarks.SelectCommand = Me.SqlSelectCommand1
        Me.DaRemarks.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "RemrksVW", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ADMNO", "ADMNO"), New System.Data.Common.DataColumnMapping("REMARKS", "REMARKS"), New System.Data.Common.DataColumnMapping("TEACHERNAME", "TEACHERNAME"), New System.Data.Common.DataColumnMapping("RDate", "RDate"), New System.Data.Common.DataColumnMapping("STUDENTS.NAME", "STUDENTS.NAME"), New System.Data.Common.DataColumnMapping("FRNAME", "FRNAME"), New System.Data.Common.DataColumnMapping("RollNo", "RollNo"), New System.Data.Common.DataColumnMapping("CLASSMAIN.NAME", "CLASSMAIN.NAME"), New System.Data.Common.DataColumnMapping("SECTIONNAME", "SECTIONNAME")})})
        '
        'DsRemarks1
        '
        Me.DsRemarks1.DataSetName = "DsRemarks"
        Me.DsRemarks1.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.DsRemarks1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents lblStdName As System.Web.UI.WebControls.Label
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents txtAdmNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents Cn1 As System.Data.SqlClient.SqlConnection
    Protected WithEvents DaRemarks As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents DsRemarks1 As Whiterose.DsRemarks

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
        ShowStudentName()
        Dim Ssql As String
        Ssql = "SELECT * FROM RemrksVW"
        Ssql &= " where AdmNo=" & txtAdmNo.Text
        Ssql &= " Order by RDate Desc"

        DaRemarks.SelectCommand.CommandText = Ssql
        'lblError.Text = Ssql
        'DaInbox.SelectCommand.Parameters.Item(0).Value = Session.Item("CurrentUsername")
        'DaInbox.SelectCommand.Parameters.Item(1).Value = txtFolderType.Text
        DaRemarks.Fill(DsRemarks1) ', "Inbox")
        DataGrid1.DataBind()
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        If Request.QueryString("AdmNo") <> "" Then
            txtAdmNo.Text = Request.QueryString("AdmNo")
            BindGrid()
        End If
    End Sub

    Private Sub DataGrid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = CDate(e.Item.Cells(0).Text).ToString("MMMM dd, yyyy")
        End If
    End Sub
End Class

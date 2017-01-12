Public Class PathTest
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label

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

            Dim FS As New System.IO.FileStream(Server.MapPath("_dsn") & "\Access_Main.dsn", IO.FileMode.Open)
            Dim TS As New System.IO.StreamReader(FS)

            Label3.Text = TS.ReadToEnd

            FS.Close()

            'lblNews1.Text = "Education Minister Mr. Imran Masood."
            'lblNews1Body.Text = "Awarded QUaid-e-Azam Trophy for Having Largest Nwtwork of Schools in Faisalabad."


        Catch Ex As Exception
            Label1.Text = Ex.Message
        End Try
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        Label1.Text = Server.MapPath("Users")
        Label2.Text = Server.MapPath("_dsn")
        Readnews1()

    End Sub

End Class

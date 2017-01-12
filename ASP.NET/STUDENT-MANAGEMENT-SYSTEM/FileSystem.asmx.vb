Imports System.Web.Services
Imports System.IO

<System.Web.Services.WebService(Namespace := "http://tempuri.org/Whiterose/FileSystem")> _
Public Class FileSystem
    Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region
    <WebMethod()> Public Function DriveNames() As String()
        Try
            Dim S() As String
            S = System.IO.Directory.GetLogicalDrives()
            Return S
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function FolderNames() As String()
        Try
            Dim S() As String
            Dim MYPath As String = Server.MapPath("/Test")
            S = System.IO.Directory.GetDirectories(MYPath)
            Return S
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function FolderName() As String
        Try
            Dim S As String
            Dim MYPath As String = Server.MapPath("FileSystem.asmx")
            S = MYPath
            Return S
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function FileNames(ByVal DPath As String) As String()
        Try
            Dim S() As String
            S = System.IO.Directory.GetFiles(DPath)
            Return S
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function SaveTextFile(ByVal FName As String, ByVal Data As String) As Boolean
        Try
            Dim MYPath As String
            MYPath = Server.MapPath("/FileTest")
            Dim FS As New FileStream(MYPath & "\" & FName & ".txt", FileMode.CreateNew)
            Dim SW As New StreamWriter(FS)
            SW.Write(Data)
            SW.Close()
        Return True
        Catch Ex As Exception
            Dim E As New Exception("Error : " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function SaveBinaryFile(ByVal FName As String, ByVal UserName As String, ByVal FolderName As String, ByVal Data As Byte()) As Boolean
        Try
            Dim MYPath As String
            MYPath = Server.MapPath("/OnlineOffice/Users/" & UserName & "/" & FolderName & "/Inbox")
            Dim FS As New FileStream(MYPath & "\" & FName, FileMode.CreateNew)
            Dim BW As New BinaryWriter(FS)
            BW.Write(Data)
            BW.Close()

            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error : " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function DeleteFile(ByVal FName As String) As Boolean
        Try
            Dim MYPath As String
            MYPath = Server.MapPath("/FileTest")
            File.Delete(MYPath & "\" & FName)
            Return True
        Catch ex As Exception
            Dim E As New Exception("Error : " & ex.Message)
            Throw E

        End Try
    End Function
    ' WEB SERVICE EXAMPLE
    ' The HelloWorld() example service returns the string Hello World.
    ' To build, uncomment the following lines then save and build the project.
    ' To test this web service, ensure that the .asmx file is the start page
    ' and press F5.
    '
    '<WebMethod()> _
    'Public Function HelloWorld() As String
    '   Return "Hello World"
    'End Function

End Class

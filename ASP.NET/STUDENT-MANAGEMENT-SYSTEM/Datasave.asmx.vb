Imports System.Web.Services
Imports System.IO
<System.Web.Services.WebService(Namespace := "http://tempuri.org/Whiterose/Datasave")> _
Public Class Datasave
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
    Public Structure StudentsData
        Dim AdmNo As Double
        Dim StdName As String
        Dim StdFRName As String
        Dim ClassCode As Int16
        Dim SectionCode As Int16
        Dim EmailAddress As String
        Dim Photo As String
        Dim RollNo As Int16
        Dim PassWord As String
    End Structure
    Public Structure RemarksData
        Dim AdmNo As Double
        Dim StdName As String
        Dim FrName As String
        Dim ClassName As String
        Dim SectionName As String
        Dim TeacherName As String
        Dim RDate As Date
        Dim RollNo As Int16
        Dim Remarks As String

    End Structure
    Public Structure ResultsData
        Dim RegNo As Double
        Dim StdName As String
        Dim FrName As String
        Dim ClassName As String
        Dim SectionName As String
        Dim ExmaYear As Int16
        Dim ExamTerm As Int16
        Dim Subject As String
        Dim Obtained As Int16
        Dim OutOf As Int16
        Dim Grade As String
        Dim ExamDate As Date
    End Structure
    <WebMethod()> Public Function SaveNews(ByVal NewsID As Int16, ByVal newsTitle As String, ByVal NewsBody As String) As Boolean
        Dim S As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer

            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Delete from news where NewsID=" & NewsID
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery


            S = "Insert into News(newsID,newsTitle,NewsBody)Values("
            S &= NewsID & ","
            S &= "'" & newsTitle & "',"
            S &= "'" & NewsBody & "'"
            S &= ")"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            S = "Insert into News(newsID,newsTitle,NewsBody)Values("
            S &= NewsID & ","
            S &= "'" & newsTitle & "',"
            S &= "'" & NewsBody & "'"
            S &= ")"
            Dim E As New Exception(S & " -> " & Ex.Message)
            Throw E
        End Try

    End Function
    <WebMethod()> Public Function SaveStudent(ByVal SourceData As StudentsData) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer
            Dim S As String
            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Insert into Students(AdmNo,Name,FRName,ClassCode,SectionCode,Email,Photo,RollNo,Password)Values("
            With SourceData
                S &= .AdmNo & ","
                S &= "'" & .StdName & "',"
                S &= "'" & .StdFRName & "',"
                S &= .ClassCode & ","
                S &= .SectionCode & ","
                S &= "'" & .EmailAddress & "',"
                S &= "'" & .Photo & "',"
                S &= .RollNo & ","
                S &= "'" & .PassWord & "'"

            End With
            S &= ")"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try

    End Function
    <WebMethod()> Public Function SaveRemarks(ByVal SourceData As RemarksData) As Boolean
        Dim S As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer

            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Insert into RemrksVW(AdmNo,StdName,FRName,ClassName,SectionName,TeacherName,RollNo,RDate,Remarks)Values("
            With SourceData
                S &= .AdmNo & ","
                S &= "'" & .StdName & "',"
                S &= "'" & .FrName & "',"
                S &= "'" & .ClassName & "',"
                S &= "'" & .SectionName & "',"
                S &= "'" & .TeacherName & "',"
                S &= .RollNo & ","
                S &= "'" & .RDate & "',"
                S &= "'" & .Remarks & "'"

            End With
            S &= ")"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Dim E As New Exception(S)
            Throw E
        End Try

    End Function
    <WebMethod()> Public Function SaveResults(ByVal SourceData As ResultsData) As Boolean
        Dim S As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer

            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Insert into ExamResultVW(RegNo,StdName,FRName,ClassName,SectionName,ExamDate,Subject,Obtained,outof,Grade,ExamTerm,ExamYear)Values("
            With SourceData
                S &= .RegNo & ","
                S &= "'" & .StdName & "',"
                S &= "'" & .FrName & "',"
                S &= "'" & .ClassName & "',"
                S &= "'" & .SectionName & "',"
                S &= "'" & .ExamDate & "',"
                S &= "'" & .Subject & "',"
                S &= .Obtained & ","
                S &= .OutOf & ","
                S &= "'" & .Grade & "',"
                S &= .ExamTerm & ","
                S &= .ExmaYear

            End With
            S &= ")"
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Dim E As New Exception(S)
            Throw E
        End Try

    End Function
    <WebMethod()> Public Function DeleteStudents() As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer
            Dim S As String
            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Delete From Students"

            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try

    End Function
    <WebMethod()> Public Function DeleteRemarks() As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer
            Dim S As String
            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Delete From RemrksVW"

            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try

    End Function
    <WebMethod()> Public Function DeleteResults() As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer
            Dim S As String
            Dim B As Boolean
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = Bloom1.ConnectionString
                Cn1.Open()
            End If
            S = "Delete From ExamResultVW"

            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery
            If R > 0 Then
                B = True
            Else
                B = False
            End If
            Cn1.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try

    End Function
    <WebMethod()> Public Function GetClassName(ByVal ClassID As Integer) As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
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

            S = "Select * from ClassMain where Code=" & ClassID
            Cmd.CommandText = S
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                S = DR.Item("Name")
            Else
                S = "Wrong Class ID"
            End If
            DR.Close()
            Cn1.Close()
            Return S
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Sub UploadFile(ByVal FileName As String, ByVal B() As Byte)
        Dim FS As FileStream
        Dim DPath As String
        DPath = Bloom1.ServerFolderPath & "\StdImages\" & FileName
        If File.Exists(DPath) Then
            File.Delete(DPath)
        End If
        FS = New FileStream(DPath, FileMode.OpenOrCreate)

        Dim BR As New BinaryWriter(FS)

        'Dim B(BR.BaseStream.Length) As Byte
        'BR.Read(B, 0, BR.BaseStream.Length - 1)
        BR.Write(B)
        BR.Close()


    End Sub
End Class

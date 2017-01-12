Imports System.Web.Services
Imports System.IO

<System.Web.Services.WebService(Namespace := "http://tempuri.org/Whiterose/OLData")> _
Public Class OLData
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
    Public Structure CompanyInfo
        Dim CompanyID As String
        Dim CompanyNumber As String
        Dim SerialNo As Double
        Dim CompanyTitle As String
        Dim CCity As String
        Dim CState As String
        Dim CZip As String
        Dim CAddress As String
        Dim CCountry As String
        Dim Email As String
        Dim AdminPassword As String
        Dim AdminFullName As String
    End Structure
    Public Structure UserInfo
        Dim CompanyID As String
        Dim UserID As String
        Dim UserFullname As String
        Dim Location As String
        Dim Designation As String
        Dim Userpassword As String
    End Structure
    Public Structure FaxMessage
        Dim MsgID As String
        Dim MsgDate As Date
        Dim Sender As String
        Dim CC As String
        Dim BCC As String
        Dim Subject As String
        Dim MessageBody As String
        Dim Attachments As Int16
        Dim Printed As Int16
        Dim UserID As String
        Dim CompanyID As String

    End Structure
    
    <WebMethod()> Public Function AdminLogin(ByVal Password As String, ByVal CompanyID As String) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As Boolean
            Ssql = "Select * from OLCompanies where CompanyID='" & CompanyID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If DR.Item("AdminPassword") = Password Then
                    B = True
                Else
                    B = False
                End If
            Else
                B = False
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function AdminUserName(ByVal CompanyID As String) As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As String
            Ssql = "Select * from OLCompanies where CompanyID='" & CompanyID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("AdminName")) Then B = DR.Item("AdminName")
            Else
                B = ""
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function GetCompanyName(ByVal CompanyID As String) As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As String
            Ssql = "Select * from OLCompanies where CompanyID='" & CompanyID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("CompanyTitle")) Then B = DR.Item("CompanyTitle")
            Else
                B = ""
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function NormalUserName(ByVal CompanyID As String, ByVal UserID As String) As String
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As String
            Ssql = "Select * from OLUsers where CompanyID='" & CompanyID & "' and UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("UserFullName")) Then B = DR.Item("UserFullName")
            Else
                B = ""
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function GetUsersList(ByVal CompanyID As String) As UserInfo()
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            Dim Uinfos() As UserInfo
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As String
            Ssql = "Select * from OLUsers where CompanyID='" & CompanyID & "' Order by UserFullName"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            R = 1
            While DR.Read
                ReDim Preserve Uinfos(R)
                If Not IsDBNull(DR.Item("UserFullName")) Then
                    Uinfos(R).UserFullname = DR.Item("UserFullName")
                End If
                If Not IsDBNull(DR.Item("UserID")) Then
                    Uinfos(R).UserID = DR.Item("UserID")
                End If
                If Not IsDBNull(DR.Item("Designation")) Then
                    Uinfos(R).Designation = DR.Item("Designation")
                End If
                If Not IsDBNull(DR.Item("Location")) Then
                    Uinfos(R).Location = DR.Item("Location")
                End If
                R += 1
            End While
            DR.Close()
            Return Uinfos
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function UserLogin(ByVal UserID As String, ByVal Password As String, ByVal CompanyID As String) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As Boolean
            Ssql = "Select * from OLUsers where CompanyID='" & CompanyID & "' And UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If DR.Item("Password") = Password Then
                    B = True
                Else
                    B = False
                End If
            Else
                B = False
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    Private Function GetSerialNo() As Integer
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim SNo As Integer
            Ssql = "Select Max(SerialNo) as S from OLCompanies"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If IsDBNull(DR.Item("S")) Then
                    SNo = DR.Item("S") + 1
                Else
                    SNo = 1
                End If
            Else
                SNo = 1
            End If
            DR.Close()
            Return SNo
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function CheckCompanyID(ByVal S As String) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As Boolean
            Ssql = "Select * from OLCompanies where CompanyID='" & S & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                B = True
            Else
                B = False
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function CheckUserID(ByVal CompanyID As String, ByVal UserID As String) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim B As Boolean
            Ssql = "Select * from OLUsers where CompanyID='" & CompanyID & "' and UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            Cmd.CommandType = CommandType.Text
            DR = Cmd.ExecuteReader
            If DR.Read Then
                B = True
            Else
                B = False
            End If
            DR.Close()
            Return B
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    <WebMethod()> Public Function SaveUser(ByVal C As UserInfo) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim Sno As Integer



            Ssql = "Insert into OLUsers "
            Ssql &= "(CompanyID,UserID,UserFullName,Password,Location,Designation)"
            Ssql &= "values"
            Ssql &= "("
            Ssql &= "'" & C.CompanyID & "',"
            Ssql &= "'" & C.UserID & "',"
            Ssql &= "'" & C.UserFullname & "',"
            Ssql &= "'" & C.Userpassword & "',"
            Ssql &= "'" & C.Location & "',"
            Ssql &= "'" & C.Designation & "'"
            Ssql &= ")"

            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery

            Cn1.Close()
            CreateUserFolders(C.CompanyID, C.UserID)
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error on Server, Reason : " & Ex.Message)
            Throw E

        End Try
    End Function
    <WebMethod()> Public Function ChangeUser(ByVal C As UserInfo) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim Sno As Integer



            Ssql = "Update OLUsers Set "
            Ssql &= "Password='" & C.Userpassword & "'"
            Ssql &= " where UserID='" & C.UserID & "'"
            Ssql &= " and CompanyID='" & C.CompanyID & "'"


            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery

            Cn1.Close()
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error on Server, Reason : " & Ex.Message)
            Throw E

        End Try
    End Function
    <WebMethod()> Public Function ChangeCompany(ByVal C As CompanyInfo) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim Sno As Integer



            Ssql = "Update OLCompanies Set "
            Ssql &= "AdminPassword='" & C.AdminPassword & "'"
            Ssql &= " where CompanyID='" & C.CompanyID & "'"


            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery

            Cn1.Close()
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error on Server, Reason : " & Ex.Message)
            Throw E

        End Try
    End Function
    Private Function CreateUserFolders(ByVal CompanyID As String, ByVal UserID As String) As Boolean
        Dim MYPath As String
        Dim NewFolder As String
        MYPath = Server.MapPath("/OnlineOffice/" & CompanyID)
        NewFolder = MYPath & "\Users\" & UserID
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\Users\" & UserID & "\Fax"
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\Users\" & UserID & "\Fax\Inbox"
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\Users\" & UserID & "\Fax\Outbox"
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\Users\" & UserID & "\Fax\Sent"
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\Users\" & UserID & "\Fax\Deleted"
        Directory.CreateDirectory(NewFolder)
        Return True
    End Function
    Private Function CreateFolder(ByVal C As String) As Boolean
        Dim MYPath As String
        Dim NewFolder As String
        MYPath = Server.MapPath("/OnlineOffice/")
        NewFolder = MYPath & "\" & C
        Directory.CreateDirectory(NewFolder)
        NewFolder = MYPath & "\" & C & "\Users"
        Directory.CreateDirectory(NewFolder)
        Return True
    End Function
    <WebMethod()> Public Function SaveCompany(ByVal C As CompanyInfo) As Boolean
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim Sno As Integer

            If C.SerialNo > 0 Then
                Sno = C.SerialNo
                Ssql = "Delete from OLCompanies where SerialNo=" & Sno
                Cmd.CommandText = Ssql
                Cmd.CommandType = CommandType.Text
                Cmd.Connection = Cn1
                R = Cmd.ExecuteNonQuery()
            Else
                Sno = GetSerialNo()
                CreateFolder(C.CompanyID)
            End If

            Ssql = "Insert into OLCompanies "
            Ssql &= "(SerialNo,CompanyID,CompanyNumber,CompanyTitle,CCity,CState,CZip,CAddress,CCountry,Email,AdminPassword,AdminName)"
            Ssql &= "values"
            Ssql &= "("
            Ssql &= Sno & ","
            Ssql &= "'" & C.CompanyID & "',"
            Ssql &= "'" & C.CompanyNumber & "',"
            Ssql &= "'" & C.CompanyTitle & "',"
            Ssql &= "'" & C.CCity & "',"
            Ssql &= "'" & C.CState & "',"
            Ssql &= "'" & C.CZip & "',"
            Ssql &= "'" & C.CAddress & "',"
            Ssql &= "'" & C.CCountry & "',"
            Ssql &= "'" & C.Email & "',"
            Ssql &= "'" & C.AdminPassword & "',"
            Ssql &= "'" & C.AdminFullName & "'"
            Ssql &= ")"

            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            R = Cmd.ExecuteNonQuery

            Cn1.Close()
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error on Server, Reason : " & Ex.Message)
            Throw E

        End Try
    End Function
    <WebMethod()> Public Function GetCompanyInfo(ByVal CompanyID As String) As CompanyInfo
        Try
            Dim Cn1 As New SqlClient.SqlConnection
            Dim R As Integer
            Dim J As CompanyInfo
            If Cn1.State = ConnectionState.Open Then
                Cn1.Close()
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            Else
                Cn1.ConnectionString = BloomOL.ConnectionString
                Cn1.Open()
            End If
            Dim Ssql As String
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Ssql = "Select * from OLCOmpanies where CompanyID='" & CompanyID & "'"

            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = Cn1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                If Not IsDBNull(DR.Item("SerialNo")) Then J.SerialNo = DR.Item("SerialNo")
                If Not IsDBNull(DR.Item("CompanyID")) Then J.CompanyID = DR.Item("CompanyID")
                If Not IsDBNull(DR.Item("CompanyNumber")) Then J.CompanyNumber = DR.Item("CompanyNumber")
                If Not IsDBNull(DR.Item("CompanyTitle")) Then J.CompanyTitle = DR.Item("CompanyTitle")
                If Not IsDBNull(DR.Item("CCity")) Then J.CCity = DR.Item("CCity")
                If Not IsDBNull(DR.Item("CState")) Then J.CState = DR.Item("CState")
                If Not IsDBNull(DR.Item("CZip")) Then J.CZip = DR.Item("CZip")
                If Not IsDBNull(DR.Item("CAddress")) Then J.CAddress = DR.Item("CAddress")
                If Not IsDBNull(DR.Item("Email")) Then J.Email = DR.Item("Email")
                If Not IsDBNull(DR.Item("AdminPassword")) Then J.AdminPassword = DR.Item("AdminPassword")
                If Not IsDBNull(DR.Item("AdminName")) Then J.AdminFullName = DR.Item("AdminName")
            End If
            DR.Close()
            Cn1.Close()
            Return J
        Catch Ex As Exception
            Dim E As New Exception("Error on Server, Reason : " & Ex.Message)
            Throw E

        End Try
    End Function
    <WebMethod()> Public Function GetInBoxCount(ByVal CompanyID As String, ByVal UserID As String) As Int16
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim R As Integer

            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If
            Ssql = "Select Count(*) as C from Inbox where Printed=0 and CompanyID='" & CompanyID & "' and UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                R = DR.Item("C")
            Else
                R = 0
            End If
            DR.Close()
            CN1.Close()
            Return R
        Catch Ex As Exception
            Dim E As New Exception("Error in Get Inbox Count " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function GetInBoxMessagePictureNames(ByVal msgID As String) As String()
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim PN() As String
            Dim R As Integer
            Dim I As Int16, P As Int16
            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If
            Ssql = "Select Count(*) as C from InboxFiles where msgID='" & msgID & "'"
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                ReDim PN(DR.Item("C"))
            End If
            DR.Close()

            Ssql = "Select * from InboxFiles where msgID='" & msgID & "'"
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            R = 1
            While DR.Read
                PN(R) = DR.Item("PicturePath")
                R += 1
            End While
            DR.Close()
            CN1.Close()
            Return PN
        Catch Ex As Exception
            Dim E As New Exception("Error in Getting InBox Attachments " & Ex.Message)
            Throw E

        End Try
    End Function
    <WebMethod()> Public Function SetPrintStatus(ByVal msgID As String, ByVal Flag As Boolean) As Boolean
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim R As Integer
            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If
            If Flag = False Then
                Ssql = "Update Inbox Set Printed=1 Where MsgID='" & msgID & "'"
            ElseIf Flag = True Then
                Ssql = "Delete Inbox Where MsgID='" & msgID & "'"
            End If
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            R = Cmd.ExecuteNonQuery
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error in Setting Print Status : " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function GetAllInBoxNewMessages(ByVal CompanyID As String, ByVal UserID As String) As FaxMessage()
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim FM() As FaxMessage
            Dim R As Integer
            Dim I As Int16, P As Int16
            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If
            Ssql = "Select Count(*) as C from Inbox where Printed=0 and CompanyID='" & CompanyID & "' and UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                R = DR.Item("C")
            Else
                R = 0
            End If
            DR.Close()
            If R > 0 Then
                ReDim FM(R)
            Else
                ReDim FM(1)
                FM(1).MsgID = -1
                Return FM
                Exit Function
            End If

            Ssql = "Select * from inBox where Printed=0 and CompanyID='" & CompanyID & "' and UserID='" & UserID & "' Order by msgDate"
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            R = 1
            While DR.Read
                If Not IsDBNull(DR.Item("MsgID")) Then FM(R).MsgID = DR.Item("MsgID")
                If Not IsDBNull(DR.Item("MsgDate")) Then FM(R).MsgDate = DR.Item("MsgDate")
                If Not IsDBNull(DR.Item("Sender")) Then FM(R).Sender = DR.Item("Sender")
                If Not IsDBNull(DR.Item("MessageBody")) Then FM(R).MessageBody = DR.Item("MessageBody")
                If Not IsDBNull(DR.Item("Subject")) Then FM(R).Subject = DR.Item("Subject")
                If Not IsDBNull(DR.Item("Attachments")) Then FM(R).Attachments = DR.Item("Attachments")
                R += 1
            End While
            DR.Close()

            CN1.Close()
            Return FM
        Catch Ex As Exception
            Dim E As New Exception("Error in Get Inbox Messages " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function GetAllInBoxOldMessages(ByVal CompanyID As String, ByVal UserID As String, ByVal Startfrom As Integer) As FaxMessage()
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim DR As SqlClient.SqlDataReader
            Dim FM() As FaxMessage
            Dim R As Integer, MsgCounter As Integer
            Dim I As Int16, P As Int16
            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If
            Ssql = "Select Count(*) as C from Inbox where Printed=1 and CompanyID='" & CompanyID & "' and UserID='" & UserID & "'"
            Cmd.CommandText = Ssql
            Cmd.CommandType = CommandType.Text
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            If DR.Read Then
                R = DR.Item("C")
            Else
                R = 0
            End If
            DR.Close()
            If R > 0 Then
                ReDim FM(R)
            Else
                ReDim FM(1)
                FM(1).MsgID = -1
                Return FM
                Exit Function
            End If

            Ssql = "Select * from inBox where Printed=1 and CompanyID='" & CompanyID & "' and UserID='" & UserID & "' Order by msgDate"
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Ssql
            Cmd.Connection = CN1
            DR = Cmd.ExecuteReader
            R = 1
            MsgCounter = 1
            While DR.Read
                If MsgCounter >= Startfrom Then
                    If Not IsDBNull(DR.Item("MsgID")) Then FM(R).MsgID = DR.Item("MsgID")
                    If Not IsDBNull(DR.Item("MsgDate")) Then FM(R).MsgDate = DR.Item("MsgDate")
                    If Not IsDBNull(DR.Item("Sender")) Then FM(R).Sender = DR.Item("Sender")
                    If Not IsDBNull(DR.Item("MessageBody")) Then FM(R).MessageBody = DR.Item("MessageBody")
                    If Not IsDBNull(DR.Item("Subject")) Then FM(R).Subject = DR.Item("Subject")
                    If Not IsDBNull(DR.Item("Attachments")) Then FM(R).Attachments = DR.Item("Attachments")
                    R += 1
                    If R >= 11 Then
                        Exit While
                    End If
                Else
                    MsgCounter += 1
                End If
                
            End While
            DR.Close()

            CN1.Close()
            Return FM
        Catch Ex As Exception
            Dim E As New Exception("Error in Get Inbox Messages " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function SaveToInBox(ByVal M As FaxMessage) As Boolean
        Try
            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim Path As String
            Dim R As Integer

            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If

            Ssql = "Insert into InBox("
            Ssql &= "MsgID,MsgDate,Sender,CC,BCC,Subject,MessageBody,Attachments,UserID,CompanyID,Printed)Values("
            Ssql &= "'" & M.MsgID & "',"
            Ssql &= "'" & Now & "',"
            Ssql &= "'" & M.Sender & "',"
            Ssql &= "'" & M.CC & "',"
            Ssql &= "'" & M.BCC & "',"
            Ssql &= "'" & M.Subject & "',"
            Ssql &= "'" & M.MessageBody & "',"
            Ssql &= M.Attachments & ","
            Ssql &= "'" & M.UserID & "',"
            Ssql &= "'" & M.CompanyID & "',"
            Ssql &= M.Printed & ")"


            Cmd.CommandText = Ssql
            Cmd.Connection = CN1
            Cmd.CommandType = CommandType.Text
            R = Cmd.ExecuteNonQuery
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error in Sending : " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function GetInBoxAttachment(ByVal CompanyID As String, ByVal UserID As String, ByVal PictureName As String, ByVal FolderType As String) As Byte()
        Try

            Dim MYPath As String
            'Dim FName As String = MSGID & "_" & AttachmentNo & "." & FileExtension
            MYPath = Server.MapPath("/OnlineOffice/" & CompanyID & "/Users/" & UserID & "/" & FolderType & "/Inbox")
            Dim FS As New FileStream(MYPath & "\" & PictureName, FileMode.Open, FileAccess.ReadWrite, System.IO.FileShare.ReadWrite)
            Dim BR As New BinaryReader(FS)
            Dim FInfo As New FileInfo(MYPath & "\" & PictureName)
            Dim D(FInfo.Length) As Byte
            BR.Read(D, 0, D.Length)
            BR.Close()
            Return D
        Catch Ex As Exception
            Dim E As New Exception("Error in Downloading the Attachment " & Ex.Message)
            Throw E
        End Try
    End Function
    <WebMethod()> Public Function SaveInBoxAttachment(ByVal MSGID As String, ByVal AttachmentNo As Int16, ByVal Data() As Byte, ByVal UserID As String, ByVal CompanyID As String, ByVal FolderType As String, ByVal FileExtension As String) As Boolean
        Try

            Dim MYPath As String
            Dim FName As String = MSGID & "_" & AttachmentNo & "." & FileExtension
            MYPath = Server.MapPath("/OnlineOffice/" & CompanyID & "/Users/" & UserID & "/" & FolderType & "/Inbox")
            Dim FS As New FileStream(MYPath & "\" & FName, FileMode.CreateNew, FileAccess.ReadWrite, System.IO.FileShare.ReadWrite)
            Dim BW As New BinaryWriter(FS)
            BW.Write(Data)
            BW.Close()


            Dim Ssql As String
            Dim CN1 As New SqlClient.SqlConnection
            Dim Cmd As New SqlClient.SqlCommand
            Dim Path As String
            Dim R As Integer

            If CN1.State = ConnectionState.Open Then
                CN1.Close()
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            Else
                CN1.ConnectionString = BloomOL.ConnectionString
                CN1.Open()
            End If

            Ssql = "Insert into InBoxFiles("
            Ssql &= "MsgID,PicturePath,UserID,CompanyID)Values("
            Ssql &= "'" & MSGID & "',"
            Ssql &= "'" & FName & "',"
            Ssql &= "'" & UserID & "',"
            Ssql &= "'" & CompanyID & "')"


            Cmd.CommandText = Ssql
            Cmd.Connection = CN1
            Cmd.CommandType = CommandType.Text
            R = Cmd.ExecuteNonQuery
            Return True
        Catch Ex As Exception
            Dim E As New Exception("Error in Sending Attachments : " & Ex.Message)
            Throw E
        End Try
    End Function


End Class

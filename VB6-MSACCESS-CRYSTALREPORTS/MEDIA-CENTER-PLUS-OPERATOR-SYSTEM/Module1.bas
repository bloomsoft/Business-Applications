Attribute VB_Name = "Module1"
Global MyTimes(8) As Date
Global oldVolume As String
Global QuranFolder As String
Global NaatFolder As String
Global SongsFolder As String
Global GPlayTime As Double
Public Sub GetTimes()
Dim DB As Database
Dim TB As Recordset
Dim Ssql As String

Set DB = OpenDatabase(App.Path & "\Main.mdb")
Set TB = DB.OpenRecordset("Select * from Setup")
If Not TB.EOF Then
    MyTimes(0) = CDate(Date & " " & TB.Fields("StartTime").Value)
    MyTimes(1) = CDate(Date & " " & TB.Fields("EndTime").Value)
    MyTimes(2) = CDate(Date & " " & TB.Fields("Fajar").Value)
    MyTimes(3) = CDate(Date & " " & TB.Fields("Duhar").Value)
    MyTimes(4) = CDate(Date & " " & TB.Fields("Asr").Value)
    MyTimes(5) = CDate(Date & " " & TB.Fields("Maghrib").Value)
    MyTimes(6) = CDate(Date & " " & TB.Fields("Isha").Value)
    MyTimes(7) = CDate(Date & " " & TB.Fields("Juma").Value)
    oldVolume = TB.Fields("Volume").Value & ""
    
    QuranFolder = TB.Fields("Quran").Value & ""
    NaatFolder = TB.Fields("Naat").Value & ""
    SongsFolder = TB.Fields("Songs").Value & ""
    
    GPlayTime = Abs(DateDiff("h", MyTimes(0), MyTimes(1)))
End If

TB.Close
DB.Close
End Sub


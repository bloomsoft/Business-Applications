VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1Port 
   Caption         =   "BloomSoft Multiplexor Driver"
   ClientHeight    =   3375
   ClientLeft      =   3840
   ClientTop       =   4890
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6510
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   5880
      Top             =   2760
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   600
      TabIndex        =   47
      Top             =   2160
      Width           =   5055
      Begin VB.CommandButton Command5 
         Caption         =   "Multi"
         Height          =   375
         Left            =   3120
         TabIndex        =   62
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check8 
         Caption         =   "VCR"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Live"
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Sequence"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Zoom"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Still"
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Plus1"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   6
         Left            =   4440
         TabIndex        =   58
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   5
         Left            =   3840
         TabIndex        =   57
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   4
         Left            =   3360
         TabIndex        =   56
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   3
         Left            =   2520
         TabIndex        =   55
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   2
         Left            =   1680
         TabIndex        =   54
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   1
         Left            =   960
         TabIndex        =   53
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   0
         Left            =   360
         TabIndex        =   52
         Top             =   720
         Width           =   135
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Menu"
      Height          =   855
      Left            =   5280
      TabIndex        =   46
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   975
      Left            =   1200
      TabIndex        =   25
      Top             =   840
      Width           =   3975
      Begin VB.OptionButton Option1 
         Caption         =   "10"
         Height          =   375
         Index           =   9
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         Height          =   375
         Index           =   8
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "8"
         Height          =   375
         Index           =   7
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7"
         Height          =   375
         Index           =   6
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6"
         Height          =   375
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         Height          =   375
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         Height          =   375
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2"
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   9
         Left            =   3480
         TabIndex        =   45
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   8
         Left            =   3120
         TabIndex        =   44
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   7
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   6
         Left            =   2400
         TabIndex        =   42
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   5
         Left            =   2040
         TabIndex        =   41
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   4
         Left            =   1680
         TabIndex        =   40
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   3
         Left            =   1320
         TabIndex        =   39
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   2
         Left            =   960
         TabIndex        =   38
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   1
         Left            =   600
         TabIndex        =   37
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   135
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ON"
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2760
      TabIndex        =   23
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5520
      TabIndex        =   22
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "10"
      Top             =   120
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5760
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Stop Switching Cameras"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Text            =   "10"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   5160
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5190
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      MaxLength       =   8
      TabIndex        =   10
      Top             =   5190
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   5220
      TabIndex        =   8
      Top             =   5340
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Write to Data Port"
      Height          =   405
      Left            =   180
      TabIndex        =   7
      Top             =   5700
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   150
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   3495
      ScaleWidth      =   6360
      TabIndex        =   6
      Top             =   6180
      Width           =   6390
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5220
      TabIndex        =   4
      Top             =   6180
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5220
      TabIndex        =   1
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read from Status Port"
      Height          =   375
      Left            =   4740
      TabIndex        =   0
      Top             =   5730
      Width           =   1755
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   120
      TabIndex        =   59
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Goto Camera #"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Last Camera #"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Wait for Seconds Before Switch"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "This project needs your imagination"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1185
      Left            =   2610
      TabIndex        =   14
      Top             =   5190
      Width           =   1845
   End
   Begin VB.Label Label6 
      Caption         =   "="
      Height          =   225
      Left            =   1740
      TabIndex        =   13
      Top             =   5220
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label5 
      Caption         =   "Bin"
      Height          =   255
      Left            =   210
      TabIndex        =   11
      Top             =   5190
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Bin"
      Height          =   255
      Left            =   4740
      TabIndex        =   9
      Top             =   5370
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   5670
      Y2              =   7380
   End
   Begin VB.Label Label4 
      Caption         =   "Dec"
      Height          =   255
      Left            =   4740
      TabIndex        =   5
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Hex"
      Height          =   255
      Left            =   4740
      TabIndex        =   3
      Top             =   6180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Value:"
      Height          =   255
      Left            =   5220
      TabIndex        =   2
      Top             =   5580
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mniShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mniHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mniShutDown 
         Caption         =   "Shut Down Camera Changer"
      End
   End
End
Attribute VB_Name = "Form1Port"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R As Long
Dim Camera As Long
Dim CC As String
Dim G As Long
Dim Intrvl As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Timer1.Enabled = False
Else
    Timer1.Enabled = True
End If
End Sub

Private Sub Check2_Click()
MSComm1.Output = "a"

If Check2.Value = 0 Then Check2.Value = 1
If Check2.Value = 1 Then
    Label12(0).BackColor = vbRed
Else
    Label12(0).BackColor = vbButtonFace
End If
End Sub

Private Sub Check3_Click()
MSComm1.Output = "b"


If Check3.Value = 0 Then Check3.Value = 1
If Check3.Value = 1 Then
    Label12(1).BackColor = vbRed
Else
    Label12(1).BackColor = vbButtonFace
End If

End Sub

Private Sub Check4_Click()
MSComm1.Output = "c"

If Check4.Value = 0 Then Check4.Value = 1
If Check4.Value = 1 Then
    Label12(2).BackColor = vbRed
Else
    Label12(2).BackColor = vbButtonFace
End If

End Sub

Private Sub Check5_Click()
MSComm1.Output = "d"

If Check5.Value = 0 Then Check5.Value = 1
If Check5.Value = 1 Then
    Label12(3).BackColor = vbRed
Else
    Label12(3).BackColor = vbButtonFace
End If

End Sub

Private Sub Check6_Click()

End Sub

Private Sub Check7_Click()
MSComm1.Output = "f"

If Check7.Value = 0 Then Check7.Value = 1
If Check7.Value = 1 Then
    Label12(5).BackColor = vbRed
Else
    Label12(5).BackColor = vbButtonFace
End If

End Sub

Private Sub Check8_Click()
MSComm1.Output = "g"

If Check8.Value = 0 Then Check8.Value = 1
If Check8.Value = 1 Then
    Label12(6).BackColor = vbRed
Else
    Label12(6).BackColor = vbButtonFace
End If

End Sub

Private Sub Command1_Click()
'Value% = DlPortReadPortUchar(&H379)
'Text2.Text = Value%
'Text3.Text = "&H" + Hex(Value%)
'Text4.Text = DecToBin$(Value%)
End Sub

Private Sub Command2_Click()
'DlPortWritePortUchar &H378, 0
End Sub

Private Sub Command3_Click()
MSComm1.Output = "p"
If Command3.Caption = "ON" Then
    Frame1.Enabled = True
    Command3.Caption = "OFF"
    Label13.BackColor = vbRed
ElseIf Command3.Caption = "OFF" Then
    Frame1.Enabled = False
    Option1_Click -1
    Command3.Caption = "ON"
    Label13.BackColor = Command3.BackColor
End If
End Sub

Private Sub Command5_Click()
MSComm1.Output = "e"
Check1.Value = 1
Option1_Click -1
If Check1.Value = 1 Then
    Label12(4).BackColor = vbRed
Else
    Label12(4).BackColor = vbButtonFace
End If

End Sub

Private Sub CheckBoxStatus()
Dim U As String

MSComm1.Output = "s"
DoEvents
If MSComm1.InBufferCount > 0 Then
    U = MSComm1.Input
    If U = "T" Then
        Frame1.Enabled = True
        Command3.Caption = "OFF"
        Label13.BackColor = vbRed
    Else
        Frame1.Enabled = False
        Command3.Caption = "ON"
        Label13.BackColor = vbButtonFace
    End If
End If
End Sub
Private Sub Form_Load()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
MSComm1.PortOpen = True
CheckBoxStatus
Winsock1.LocalPort = 9000
Winsock1.Listen
'CC = "0111"
G = 0
Intrvl = Val(Text7.Text)

   With sysIcon
        .cbSize = LenB(sysIcon)
        .hWnd = Me.hWnd
        .uFlags = NIF_DOALL
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .sTip = Me.Caption
    End With
    Shell_NotifyIcon NIM_ADD, sysIcon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ShellMsg As Long
    
    ShellMsg = X / Screen.TwipsPerPixelX
    Select Case ShellMsg
    Case WM_LBUTTONDBLCLK
        Me.Show
    Case WM_RBUTTONUP
        'Show the menu
        'If gfStarted Then mnuStart.Enabled = False
        PopupMenu mnuFile
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MSComm1.PortOpen = False
End Sub

Private Sub Form_Resize()
'If Me.WindowState = 1 Then
'    If Me.Visible = True Then
'        Me.Visible = False
'        Me.Hide
'    End If
'End If
End Sub

Private Sub mniHide_Click()
Me.Visible = False
Me.Hide

End Sub

Private Sub mniShow_Click()
Me.Visible = True
Me.Show

End Sub

Private Sub mniShutDown_Click()
End
End Sub

Private Sub MSComm1_OnComm()
Text10.Text = MSComm1.Input
End Sub

Private Sub Option1_Click(Index As Integer)

Dim R As Integer
For R = 0 To Label11.Count - 1
    Label11(R).BackColor = vbButtonFace
Next R

If Index = -1 Then
For R = 0 To Label11.Count - 1
    Option1(R).Value = False
Next R
Exit Sub
End If
Label12(4).BackColor = vbButtonFace
MSComm1.Output = CStr(Index + 1)
Label11(Index).BackColor = vbRed
End Sub

Private Sub Text1_Change()
'Text5.Text = BinToDec(Val(Text1.Text))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'DlPortWritePortUchar &H378, 0
R = 0
Do While R < 100
R = R + 1
DoEvents
Loop
'DlPortWritePortUchar &H378, Val(Text5.Text)
End If


End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'j% = BinToDec(Val(Text6.Text))
'MsgBox j%
'DlPortWritePortUlong &H378, Text6.Text
End If
End Sub

Private Sub Text7_Change()
Intrvl = Val(Text7.Text)
End Sub

Private Sub Text9_Change()
If Len(Text9.Text) > 0 Then
G = Text9.Text
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
MSComm1.Output = CStr(G)
Option1(G - 1).Value = True
Option1_Click G - 1
Camera = G
End If
End Sub

Private Sub Timer1_Timer()
'Me.Caption = "Current Interval : " & Intrvl & "Seconds"

If Camera > Intrvl Then
Camera = 0
G = G + 1

MSComm1.Output = CStr(G)
Option1(G - 1).Value = True
Option1_Click G - 1
If G = Val(Text8.Text) Then
    G = 0
End If
'If G = 1 Then
'    Me.Caption = "Camera # 1"
'    DlPortWritePortUlong &H378, "&H0"
'ElseIf G = 2 Then
'    Me.Caption = "Camera # 2"
'    DlPortWritePortUlong &H378, "&H8"
'ElseIf G = 3 Then
'    Me.Caption = "Camera # 3"
'    DlPortWritePortUlong &H378, "&H10"
'ElseIf G = 4 Then
'    Me.Caption = "Camera # 4"
'    DlPortWritePortUlong &H378, "&H18"
'ElseIf G = 5 Then
'    Me.Caption = "Camera # 5"
'    DlPortWritePortUlong &H378, "&H20"
'ElseIf G = 6 Then
'    Me.Caption = "Camera # 6"
'    DlPortWritePortUlong &H378, "&H28"
'ElseIf G = 7 Then
'    Me.Caption = "Camera # 7"
'    DlPortWritePortUlong &H378, "&H30"
'ElseIf G = 8 Then
'    Me.Caption = "Camera # 8"
'    DlPortWritePortUlong &H378, "&H38"
'    G = 0
'End If

Else
Camera = Camera + 1
End If
End Sub

Private Sub Timer2_Timer()
CheckBoxStatus
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim S As String
Dim T As String
Dim O As String
Dim P As Integer
Winsock1.GetData S

If S = "T" Then
    Timer1.Enabled = True
ElseIf InStr(S, "S:") Then
    T = Right(S, Len(S) - 2)
    Intrvl = Val(T)
    Timer1.Enabled = True
    Check1.Value = 0
Else
Timer1.Enabled = False
'DlPortWritePortUlong &H378, S
'Select Case S
'    Case "&H0"
'        Me.Caption = "Camera # 1"
'    Case "&H8"
'        Me.Caption = "Camera # 2"
'    Case "&H10"
'        Me.Caption = "Camera # 3"
'    Case "&H18"
'        Me.Caption = "Camera # 4"
'    Case "&H20"
'        Me.Caption = "Camera # 5"
'    Case "&H28"
'        Me.Caption = "Camera # 6"
'    Case "&H30"
'        Me.Caption = "Camera # 7"
'    Case "&H38"
'        Me.Caption = "Camera # 8"
'End Select
If Len(S) = 1 Then
    O = S
    
    If Trim(O) = "a" Then Check2_Click
    If Trim(O) = "b" Then Check3_Click
    If Trim(O) = "c" Then Check4_Click
    If Trim(O) = "d" Then Check5_Click
    If Trim(O) = "e" Then Command5_Click
    If Trim(O) = "f" Then Check7_Click
    If Trim(O) = "g" Then Check8_Click
    
Else
P = InStr(1, S, ":", vbBinaryCompare)
If P > 0 Then
G = Right(S, Len(S) - P)
End If
If Val(G) > 0 Then
Option1(G - 1).Value = True
'Option1_Click G - 1
End If
Camera = G
End If
End If
End Sub
Private Sub Sleep(Seconds As Integer)
Dim OLDDate As Date
OLDDate = Now
Dim D As Long

Do
D = DateDiff("s", Now, OLDDate)
If Abs(D) >= Seconds Then
    
    Exit Do
End If
DoEvents
Loop

End Sub


VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embedding sound files in a OLE Control"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start the Timer."
      Height          =   735
      Left            =   1560
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0884
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.OLE OLE12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4200
      OleObjectBlob   =   "Form1.frx":093B
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\9.mid"
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4320
      OleObjectBlob   =   "Form1.frx":5953
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\8.mid"
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":A96B
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\5.mid"
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4320
      OleObjectBlob   =   "Form1.frx":F983
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\31.mid"
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":1499B
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\3.mid"
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4320
      OleObjectBlob   =   "Form1.frx":199B3
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\18.mid"
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":1E9CB
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\17.mid"
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4320
      OleObjectBlob   =   "Form1.frx":239E3
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\13.mid"
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":289FB
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\12.mid"
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.OLE OLE3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   4320
      OleObjectBlob   =   "Form1.frx":2DA13
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\11.mid"
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1470
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Media Player has stoped."
      ForeColor       =   &H00C0C0FF&
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404000&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   990
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking if the Media Player is still playing music..."
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.OLE OLE2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":32A2B
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\1.mid"
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":37A43
      ForeColor       =   &H00FFFF80&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4410
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "midfile"
      Height          =   30
      Left            =   1080
      OleObjectBlob   =   "Form1.frx":37ACA
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\OLEMusic\36.mid"
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Menu mnusound 
      Caption         =   "&Turning OLE Controls on/off manually"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''
'You can put anything you want into an OLE
'container and activate it within the .exe,
'provided that the application that created the
'object can be accessed by the users system.
'This is how I did it;
'1) Place an OLE Control on your form.
'2) When the "Insert Object" window appears,
'   click the "Create from File" option button.
'3) Use the "Browse" button to find the music
'   file, then double ckick-it.
'4) Click "OK"
'5) Set the properties of the OLE container to
'   match those that I used in this example.
'Note: By default, the object you insert into the
'OLE container will be embedded in the .exe. If
'you click the "Link" check-box, you will need to
'provide the file along with the .exe
'''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
If StartMe = False Then
StartMe = True
Label1.Visible = False
Label2.Visible = True
Shape1.Visible = True
Label4.Visible = True
Command1.Caption = "Quit the program."
Exit Sub
ElseIf Command1.Caption = "Quit the program." Then
Unload Me
Else
End If

End Sub

Private Sub Form_Activate()
'turn on the Media Player and
'play the contents of OLE1
OLE1.Action = 7
'this hides the Media Player
OLE1.DoVerb -3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'release the Media Player,
'shut off any OLE control that is activated
OLE1.Action = 9
OLE2.Action = 9
OLE3.Action = 9
OLE4.Action = 9
OLE5.Action = 9
OLE6.Action = 9
OLE7.Action = 9
OLE8.Action = 9
OLE9.Action = 9
OLE10.Action = 9
OLE11.Action = 9
OLE12.Action = 9
Unload Me
End Sub

Private Sub mnusound_Click()
'shut off any OLE2 - OLE12 controls that may be activated
OLE2.Action = 9
OLE3.Action = 9
OLE4.Action = 9
OLE5.Action = 9
OLE6.Action = 9
OLE7.Action = 9
OLE8.Action = 9
OLE9.Action = 9
OLE10.Action = 9
OLE11.Action = 9
OLE12.Action = 9
StartMe = False
Command1.Caption = "Start the Timer."
Label1.Visible = False
Shape1.Visible = False
Label2.Visible = False
Shape2.Visible = False
Label3.Visible = False
DoEvents
If OLE1.AppIsRunning = False And quiet = False Then
'activate OLE1 (opening music)
OLE1.Action = 7
'hide the Media Player
OLE1.DoVerb -3
End If
'turn the music on/off
If quiet = False Then
Msg = "The opening music is on, do you want to turn it off?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Title = "OLE1 is activated..."
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
quiet = True
'shut off OLE1
OLE1.Action = 9
Exit Sub
Else
End If
ElseIf quiet = True Then
Msg = "The opening music is off, do you want to turn it back on?"
Style = vbYesNo + vbQuestion + vbDefaultButton1
Title = "OLE1 is not activated..."
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
quiet = False
'turn OLE1 back on
OLE1.Action = 7
'hide the Media Player
OLE1.DoVerb -3
Exit Sub
Else
End If
End If
End Sub

Private Sub Timer1_Timer()
If StartMe = False Then Exit Sub
'turn off the opening music, if it's still playing
OLE1.Action = 9
DoEvents
'''''''''''''''''''''''''''''''''''
'blink the green light
If Shape1.FillColor = &H80FF80 Then
Shape1.FillColor = &H404000
Else
Shape1.FillColor = &H80FF80
End If
'give the user a few seconds to see
'which OLE is activated (Label3)
If Delay = 5 Then
Shape2.Visible = False
Label3.Visible = False
End If
'''''''''''''''''''''''''''''''''''''''''
'check if the Media Player is
'still playing music
If OLE2.AppIsRunning Or _
OLE3.AppIsRunning Or _
OLE4.AppIsRunning Or _
OLE5.AppIsRunning Or _
OLE6.AppIsRunning Or _
OLE7.AppIsRunning Or _
OLE8.AppIsRunning Or _
OLE9.AppIsRunning Or _
OLE10.AppIsRunning Or _
OLE11.AppIsRunning Or _
OLE12.AppIsRunning Then Delay = Delay + 1: Exit Sub
'the music has stopped playing
'so play another music file randomly
Randomize
'OLE2 - OLE12 are the sample files
'so we pick a number FROM 2 - 12
'for simplicity
SoundFile = Int((12 - 2 + 1) * Rnd + 2)
If SoundFile = 2 Then
'activate the contents of the OLE control
OLE2.Action = 7
'hide the Media Player
OLE2.DoVerb -3
ElseIf SoundFile = 3 Then
OLE3.Action = 7
OLE3.DoVerb -3
ElseIf SoundFile = 4 Then
OLE4.Action = 7
OLE4.DoVerb -3
ElseIf SoundFile = 5 Then
OLE5.Action = 7
OLE5.DoVerb -3
ElseIf SoundFile = 6 Then
OLE6.Action = 7
OLE6.DoVerb -3
ElseIf SoundFile = 7 Then
OLE7.Action = 7
OLE7.DoVerb -3
ElseIf SoundFile = 8 Then
OLE8.Action = 7
OLE8.DoVerb -3
ElseIf SoundFile = 9 Then
OLE9.Action = 7
OLE9.DoVerb -3
ElseIf SoundFile = 10 Then
OLE10.Action = 7
OLE10.DoVerb -3
ElseIf SoundFile = 11 Then
OLE11.Action = 7
OLE11.DoVerb -3
ElseIf SoundFile = 12 Then
OLE12.Action = 7
OLE12.DoVerb -3
Else
End If
'show which OLE control is activated
Label3.Caption = "The Media Player has stopped, now playing OLE" & " " & "#" & SoundFile
Shape2.Visible = True
Label3.Visible = True
'reset Delay
Delay = 0
End Sub

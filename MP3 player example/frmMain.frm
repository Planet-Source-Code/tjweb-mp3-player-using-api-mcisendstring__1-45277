VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Z-Ware Productions"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdFF 
      Left            =   2040
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSong 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "C:\Mp3\Song.mp3"
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' To get short filenames out of long filenames
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
' The main heart of mp3 playing :D the MciSendString
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


' Pause Boolean '
Dim Paused As Boolean

Private Sub cmdOpen_Click()
cdFF.Filter = "Mp3 Music (*.mp3)|*.mp3"
cdFF.InitDir = App.Path
cdFF.ShowOpen
If UCase(Right(cdFF.FileName, 4)) = ".MP3" Then
  txtSong = cdFF.FileName
End If
End Sub

Private Sub cmdPause_Click()
If Paused = True Then
  mciSendString "Play Mp3", 0, 0, 0 ' Sets the "Mp3" alias on Play
  Paused = False
Else
  mciSendString "Pause Mp3", 0, 0, 0 ' Sets the "Mp3" alias to Pause
  Paused = True
End If
End Sub

Private Sub cmdPlay_Click()
Dim Temp As String * 255, FileName As String
FileName = txtSong
If UCase(Right(FileName, 4)) = ".MP3" Then ' Checks the file, if its a mp3 :D
  If Dir(FileName) = vbNullString Then Exit Sub ' If File NOT exists, exit the sub
  FileName = GetShortPathName(FileName, Temp, 254) ' Gets the "SHORT" filename of a filename (C:\MyMus~1\SongAbo~1.mp3)
  FileName = Left$(Temp, FileName) ' Safty :D
  mciSendString "Close Mp3", 0, 0, 0 ' Closes the mp3 before opening. in case its playing from before
  mciSendString "Open " & FileName & " Alias Mp3", 0, 0, 0 ' Opens the mp3 in tha aliasname "Mp3"
  mciSendString "Play Mp3", 0, 0, 0 ' Plays the MP3 alias
  Paused = False ' Sets so the pause button works ;)
End If
End Sub

Private Sub cmdStop_Click()
mciSendString "Stop Mp3", 0, 0, 0 ' Stops the mp3.!
End Sub

Private Sub Form_Load()
MsgBox "Hello Dear User. " & vbCrLf & "I really hope this project helps you alot." & vbCrLf & "And i really hope you will rate me at pscode. :P" & vbCrLf & "Enjoy!!!!", vbInformation, "Enjoy!!!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "For more of my examples check:", vbinformaion, "More Examples"
MsgBox "http://pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=7742748056&strAuthorName=Thorleif%20Jacobsen&txtMaxNumberOfEntriesPerPage=25", vbinformaion, "More Examples"
MsgBox "Enjoy!!!", vbinformaion, "More Examples"
End Sub

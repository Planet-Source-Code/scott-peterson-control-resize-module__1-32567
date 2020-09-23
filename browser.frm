VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmviewer 
   Caption         =   "File Viewer"
   ClientHeight    =   8115
   ClientLeft      =   -5805
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HSvol 
      Height          =   135
      LargeChange     =   300
      Left            =   2510
      Max             =   3000
      SmallChange     =   30
      TabIndex        =   7
      Top             =   240
      Width           =   10335
   End
   Begin VB.CheckBox chknext 
      Caption         =   "Play next on media file"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdsetdefault 
      Caption         =   "Save Current Settings"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Width           =   2415
   End
   Begin VB.ListBox lstfileview 
      Height          =   7665
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Hidden          =   -1  'True
      Left            =   0
      OLEDragMode     =   1  'Automatic
      System          =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7695
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   10335
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   7695
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   10305
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub listfilecontents()
lstfileview.Visible = True
lstfileview.Clear
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(File1.Path + "\" + File1.FileName)
Do While a.AtEndOfStream = False
lstfileview.AddItem (a.readline)
Loop
a.Close
End Sub

Private Sub cmdsetdefault_Click()
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(App.Path + "\Default.txt")
a.writeline (Drive1.Drive)
a.writeline (Dir1.Path)
a.writeline (File1.Path)
a.writeline (chknext.Value)
a.writeline (MediaPlayer1.Volume)
a.Close
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_GotFocus()
lstfileview.Visible = False
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive1_GotFocus()
lstfileview.Visible = False
End Sub

Private Sub File1_Click()
lstfileview.Visible = False
MediaPlayer1.Visible = False
Image1.Visible = False
If Right$(File1.FileName, 3) = "jpg" Or Right$(File1.FileName, 3) = "JPG" Or Right$(File1.FileName, 3) = "gif" Or Right$(File1.FileName, 3) = "tif" Or Right$(File1.FileName, 3) = "bmp" Then
    Dim temp As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OpenTextFile(File1.Path + "\" + File1.FileName)
    temp = a.readline
    a.Close
    If Left$(temp, 3) = "GIF" Or Left$(temp, 3) = "ÿØÿ" Or Left$(temp, 2) = "BM" Then
    Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
    Image1.Visible = True
    Else
    Call listfilecontents
    End If
Else
    MediaPlayer1.FileName = File1.Path + "\" + File1.FileName
    MediaPlayer1.Visible = True
    If MediaPlayer1.HasError = True Then
    Call listfilecontents
    End If
End If
Call Form_Resize

End Sub

Private Sub Form_Load()
Call SizeChange.initialize(Me.Name)
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile(App.Path + "\Default.txt")
If Not a.AtEndOfStream Then
    Drive1.Drive = a.readline
    Dir1.Path = a.readline
    File1.Path = a.readline
    chknext.Value = a.readline
    HSvol.Value = 3000 + Val(a.readline)
End If
a.Close
End Sub

Private Sub Form_Resize()
Call SizeChange.SizeChange(Me.Name)
End Sub

Private Sub HSvol_Change()
MediaPlayer1.Volume = HSvol.Value - 3000
End Sub

Private Sub lstfileview_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 And lstfileview.ListIndex >= 0 Then
lstfileview.RemoveItem (lstfileview.ListIndex)
End If
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
If chknext.Value Then
    If File1.ListCount > 0 And File1.ListIndex < (File1.ListCount - 1) Then
        File1.ListIndex = File1.ListIndex + 1
    End If
End If
End Sub

VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form2 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.FileListBox F2 
      Height          =   285
      Left            =   2160
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer T1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   2640
   End
   Begin VB.FileListBox F1 
      Height          =   285
      Left            =   2160
      Pattern         =   "*.gif;*.jpg;*.jpeg;*.tif"
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MediaPlayerCtl.MediaPlayer MD1 
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
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
      PlayCount       =   0
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
      ShowAudioControls=   -1  'True
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
   Begin VB.Image I1 
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
I1.Move 0, 0, Me.Width, Me.Height
Flag = -1
If F1.ListCount > 0 Then
    I1.Picture = LoadPicture(F1.List(0))
Else
    MsgBox "Não existem Imagens para serem visualizadas...Programa vai ser encerrado.", vbInformation, "InforBeta, Lda"
End If
If F2.ListCount > 0 Then
    MD1.FileName = App.Path + "\" + F2.List(0)
    MD1.AutoRewind = True
    MD1.Play
End If
T1.Enabled = True
T1.Interval = 10000
Form1.Show vbModal, Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub T1_Timer()
    Flag = Flag + 1
    Call mudapic(Flag)
End Sub
Public Sub mudapic(varInt As Integer)
    If varInt = -1 Then varInt = 0
    If varInt = F1.ListCount Then varInt = F1.ListCount - 1
    I1.Picture = LoadPicture(F1.List(varInt))
End Sub


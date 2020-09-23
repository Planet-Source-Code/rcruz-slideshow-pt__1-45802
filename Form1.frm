VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjChameleon.chameleonButton c 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   0
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":030A
   End
   Begin prjChameleon.chameleonButton c 
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":0326
   End
   Begin prjChameleon.chameleonButton c 
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":0342
   End
   Begin prjChameleon.chameleonButton c 
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":035E
   End
   Begin prjChameleon.chameleonButton S 
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "S"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":037A
   End
   Begin prjChameleon.chameleonButton p 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "u"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":0396
   End
   Begin prjChameleon.chameleonButton St 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "n"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "Form1.frx":03B2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub c_Click(Index As Integer)
    Select Case Index
    Case 0
        Flag = 0
    Case 1
        Flag = Flag - 1
    Case 2
        Flag = Flag + 1
    Case 3
        Flag = Form2.F1.ListCount - 1
    End Select
    Call Form2.mudapic(Flag)
    Form2.T1.Enabled = False
    Form2.T1.Enabled = True

End Sub

Private Sub Form_Load()
    Me.Left = 100
    Me.Top = 100
    St.Visible = True
    p.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, 161, 2, 0&
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub p_Click()
    p.Visible = False
    St.Visible = True
    Flag = Flag + 1
    Call Form2.mudapic(Flag)
    Form2.T1.Enabled = True
End Sub

Private Sub S_Click(Index As Integer)
    End
End Sub

Private Sub St_Click()
    p.Visible = True
    St.Visible = False
    Form2.T1.Enabled = False
End Sub

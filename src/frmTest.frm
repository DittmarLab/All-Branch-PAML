VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "LRT definition"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3360
      TabIndex        =   8
      Top             =   2400
      Width           =   1452
   End
   Begin VB.TextBox txDf 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   2640
      TabIndex        =   7
      Text            =   "1"
      Top             =   1740
      Width           =   1935
   End
   Begin VB.ComboBox cbH1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   5
      Text            =   "H1"
      Top             =   1260
      Width           =   1935
   End
   Begin VB.ComboBox cbH0 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Text            =   "H0"
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1452
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Degree of freedom (df):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1780
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Definition of likelihood ratio test (LRT):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Alternative hypothesis (H1):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1300
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Null hypothesis (H0):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   820
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Form_Load()
For I = 0 To UBound(Ctls())
    cbH0.AddItem Ctls(I).Name
    cbH1.AddItem Ctls(I).Name
Next
If LRTs(iLRT).Name = "" Then
    cbH0.Text = Ctls(0).Name
    cbH1.Text = Ctls(1).Name
Else
    With LRTs(iLRT)
        cbH0.Text = .H0
        cbH1.Text = .H1
        txDf.Text = .df
    End With
End If
End Sub

Private Sub cmdOK_Click()
For I = 0 To UBound(LRTs())
    If cbH0.Text = LRTs(I).H0 And cbH1.Text = LRTs(I).H1 And txDf.Text = LRTs(I).df Then
        MsgBox "The specified test already exists." & vbCrLf & vbCrLf & "Please specify another one.", vbExclamation
        Exit Sub
    End If
Next
With LRTs(iLRT)
    .H0 = cbH0.Text
    .H1 = cbH1.Text
    .df = CInt(txDf.Text)
    .Name = .H1 & " vs " & .H0 & " (df=" & CStr(.df) & ")"
End With
Unload Me
frmMain.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub


VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "All-Branch PAML - Dittmar Lab"
   ClientHeight    =   7380
   ClientLeft      =   4425
   ClientTop       =   2415
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   6345
   Begin VB.ComboBox cbOutput 
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
      ItemData        =   "frmMain.frx":0000
      Left            =   4440
      List            =   "frmMain.frx":000D
      TabIndex        =   30
      Text            =   "'mlc' only"
      Top             =   6510
      Width           =   1812
   End
   Begin VB.CommandButton cmdAdd2 
      Caption         =   "Add"
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
      Left            =   5040
      TabIndex        =   28
      Top             =   4005
      Width           =   1212
   End
   Begin VB.CommandButton cmdEdit2 
      Caption         =   "Edit"
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
      Left            =   5040
      TabIndex        =   27
      Top             =   4365
      Width           =   1212
   End
   Begin VB.CommandButton cmdDelete2 
      Caption         =   "Delete"
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
      Left            =   5040
      TabIndex        =   26
      Top             =   4725
      Width           =   1212
   End
   Begin VB.ListBox lstLRT 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frmMain.frx":0028
      Left            =   120
      List            =   "frmMain.frx":002F
      TabIndex        =   24
      Top             =   4005
      Width           =   4812
   End
   Begin VB.ComboBox cbSize 
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
      ItemData        =   "frmMain.frx":0044
      Left            =   4440
      List            =   "frmMain.frx":0051
      TabIndex        =   21
      Text            =   "Minimize"
      Top             =   5790
      Width           =   1812
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&About"
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
      Left            =   3720
      TabIndex        =   20
      Top             =   6960
      Width           =   1212
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   6960
      Width           =   1212
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Go!"
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
      Height          =   325
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   1212
   End
   Begin VB.CommandButton cmdWork 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3840
      TabIndex        =   17
      Top             =   6510
      Width           =   312
   End
   Begin VB.TextBox txtWork 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   16
      Top             =   6510
      Width           =   3612
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   2880
      Width           =   1212
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   2520
      Width           =   1212
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   1212
   End
   Begin VB.ListBox lstCtl 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frmMain.frx":006D
      Left            =   120
      List            =   "frmMain.frx":0074
      TabIndex        =   10
      Top             =   2160
      Width           =   4812
   End
   Begin VB.CommandButton cmdCodeml 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3840
      TabIndex        =   9
      Top             =   5790
      Width           =   312
   End
   Begin VB.TextBox txtCodeml 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   8
      Top             =   5790
      Width           =   3615
   End
   Begin VB.CommandButton cmdTree 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   5880
      TabIndex        =   6
      Top             =   1470
      Width           =   312
   End
   Begin VB.TextBox txtTree 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   5
      Top             =   1470
      Width           =   5652
   End
   Begin VB.CommandButton cmdSeq 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   312
   End
   Begin VB.TextBox txtSeq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5652
   End
   Begin VB.Label Label10 
      Caption         =   "Save output:"
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
      Left            =   4440
      TabIndex        =   29
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Likelihood ratio tests:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Program window:"
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
      Left            =   4440
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Working directory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Model sets:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1890
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Codeml program:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Tree file:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Sequence file:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All-Branch PAML: Test for Selection at All Branches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "All-Branch PAML: Test for Selection at All Branches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   100
      TabIndex        =   23
      Top             =   100
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
On Error GoTo ErrDeal:
iCtl = UBound(Ctls()) + 1
ReDim Preserve Ctls(0 To iCtl)
Me.Hide
frmOption.Show
Exit Sub
ErrDeal:
iCtl = 0
ReDim Preserve Ctls(0)
Me.Hide
frmOption.Show
End Sub

Private Sub cmdAdd2_Click()
If ((Not Ctls) = -1) Then
    MsgBox "Please define at least two model sets.", vbExclamation
    Exit Sub
ElseIf UBound(Ctls()) = 0 Then
    MsgBox "Please define at least two model sets.", vbExclamation
    Exit Sub
End If
On Error GoTo ErrDeal2:
iLRT = UBound(LRTs()) + 1
ReDim Preserve LRTs(0 To iLRT)
Me.Hide
frmTest.Show
Exit Sub
ErrDeal2:
iLRT = 0
ReDim Preserve LRTs(0)
Me.Hide
frmTest.Show
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrDeal:
If lstCtl.SelCount = 0 Then
    MsgBox "Please specify a model set to delete.", vbExclamation
Else
    For i = lstCtl.ListCount - 1 To 0 Step -1
        If lstCtl.Selected(i) = True Then
            lstCtl.RemoveItem (i)
            For J = i To UBound(Ctls()) - 1
                Ctls(J) = Ctls(J + 1)
            Next
            ReDim Preserve Ctls(0 To UBound(Ctls()) - 1)
        End If
    Next
End If
Exit Sub
ErrDeal:
Erase Ctls()
lstCtl.Clear
End Sub

Private Sub cmdDelete2_Click()
On Error GoTo ErrDeal2:
If lstLRT.SelCount = 0 Then
    MsgBox "Please specify a test to delete.", vbExclamation
    Exit Sub
End If
Do While i < lstLRT.ListCount
    If lstLRT.Selected(i) = True Then
        lstLRT.RemoveItem (i)
        For J = i To UBound(LRTs()) - 1
            LRTs(J) = LRTs(J + 1)
        Next
        ReDim Preserve LRTs(0 To UBound(LRTs()) - 1)
    Else
        i = i + 1
    End If
Loop
Exit Sub
ErrDeal2:
Erase LRTs()
lstLRT.Clear
End Sub

Private Sub cmdEdit_Click()
If lstCtl.SelCount = 0 Then
    MsgBox "Please specify a model set to edit.", vbExclamation
    Exit Sub
End If
For i = 0 To lstCtl.ListCount
    If lstCtl.Selected(i) = True Then
        iCtl = i
        Exit For
    End If
Next
Me.Hide
frmOption.Show
End Sub

Private Sub cmdEdit2_Click()
If lstLRT.SelCount = 0 Then
    MsgBox "Please specify a test to edit.", vbExclamation
    Exit Sub
End If
For i = 0 To lstLRT.ListCount
    If lstLRT.Selected(i) = True Then
        iLRT = i
        Exit For
    End If
Next
Me.Hide
frmTest.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdHelp_Click()
frmAbout.Show
'Call ShellExecute(Me.hWnd, "Open", App.Path & "\readme.txt", vbNullString, App.Path, 1)
End Sub

Private Sub cmdRun_Click()
If txtSeq.Text = "" Then
    MsgBox "Please specify sequence file.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If txtCodeml.Text = "" Then
    MsgBox "Please specify Codeml program.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If txtWork.Text = "" Then
    MsgBox "Please specify working directory.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If Dir(txtSeq.Text) = "" Then
    MsgBox "The sequence file" & vbCrLf & txtSeq.Text & vbCrLf & "does not exist.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If txtTree.Text <> "" And Dir(txtSeq.Text) = "" Then
    MsgBox "The tree file" & vbCrLf & txtTree.Text & vbCrLf & "does not exist.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If Dir(txtCodeml.Text) = "" Then
    MsgBox "The program file" & vbCrLf & txtCodeml.Text & vbCrLf & "does not exist.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If Dir(txtWork.Text, vbDirectory) = "" Then
    MsgBox "The working directory" & vbCrLf & txtWork.Text & vbCrLf & "does not exist.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
If cbSize.Text = "Normal" Then wSize = SW_NORMAL
If cbSize.Text = "Minimize" Then wSize = SW_MINIMIZE
If cbSize.Text = "Hide" Then wSize = SW_Hide
If cbSize.Text = "Maximize" Then wSize = SW_MAXIMIZE
If cbOutput.Text = "None" Then Outputs = 0
If cbOutput.Text = "'mlc' only" Then Outputs = 1
If cbOutput.Text = "All" Then Outputs = 2

fSequence = txtSeq.Text
fTree = txtTree.Text
fCodeml = txtCodeml.Text
dWork = txtWork.Text
If Right(dWork, 1) <> "\" Then dWork = dWork & "\"
Me.Hide
Load frmProcess
End Sub

Private Sub cmdSeq_Click()
Dim S As String
S = sOpenFile("*.*" + Chr$(0) + "*.*" + Chr$(0))
If S <> "" Then
    txtSeq.Text = S
    S = Left(S, InStrRev(S, "\"))
    If txtWork.Text = "" Then txtWork.Text = S
    If txtCodeml.Text = "" And Dir(S & "codeml.exe") <> "" Then txtCodeml.Text = S & "codeml.exe"
End If
End Sub

Private Sub cmdTree_Click()
Dim S As String
S = sOpenFile("*.*" + Chr$(0) + "*.*" + Chr$(0))
If S <> "" Then
    txtTree.Text = S
    S = Left(S, InStrRev(S, "\"))
    If txtWork.Text = "" Then txtWork.Text = S
    If txtCodeml.Text = "" And Dir(S & "codeml.exe") <> "" Then txtCodeml.Text = S & "codeml.exe"
End If
End Sub

Private Sub cmdCodeml_Click()
Dim S As String
S = sOpenFile("codeml.exe" + Chr$(0) + "codeml.exe" + Chr$(0) + "*.*" + Chr$(0) + "*.*" + Chr$(0))
If S = "" Then Exit Sub
If Right(S, 10) <> "codeml.exe" Then
    If MsgBox("The selected file" & vbCrLf & S & vbCrLf & "looks unlike a correct one." & vbCrLf & vbCrLf & "Do you insist using this file?", vbYesNo + vbQuestion, "Question") = vbYes Then
        txtCodeml.Text = S
    End If
Else
    txtCodeml.Text = S
End If
End Sub

Private Sub cmdWork_Click()
txtWork.Text = sOpenDir
If Right(txtWork.Text, 1) <> "\" Then txtWork.Text = txtWork.Text & "\"
End Sub

Private Sub Form_Activate()
If (Not Ctls) <> -1 Then
    For i = UBound(Ctls()) To LBound(Ctls()) Step -1
        If Ctls(i).Name = "" Then
            If i = LBound(Ctls()) Then
                Erase Ctls()
            Else
                For J = i To UBound(Ctls()) - 1
                    Ctls(J) = Ctls(J + 1)
                Next
                ReDim Preserve Ctls(LBound(Ctls()) To i - 1)
            End If
        End If
    Next
End If
lstCtl.Clear
If (Not Ctls) <> -1 Then
    For i = 0 To UBound(Ctls())
        lstCtl.AddItem Ctls(i).Name
    Next
End If
If (Not LRTs) <> -1 Then
    For i = UBound(LRTs()) To LBound(LRTs()) Step -1
        If LRTs(i).Name = "" Then
            If i = LBound(LRTs()) Then
                Erase LRTs()
            Else
                For J = i To UBound(LRTs()) - 1
                    LRTs(J) = LRTs(J + 1)
                Next
                ReDim Preserve LRTs(LBound(LRTs()) To i - 1)
            End If
        End If
    Next
End If
lstLRT.Clear
If (Not LRTs) <> -1 Then
    For i = 0 To UBound(LRTs())
        lstLRT.AddItem LRTs(i).Name
    Next
End If
End Sub

Private Sub Form_Load()
Dim S As String
S = CurDir
If Right(S, 1) <> "\" Then S = S & "\"
If Dir(S & "codeml.exe") <> "" Then txtCodeml.Text = S & "codeml.exe"
End Sub

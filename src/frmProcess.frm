VERSION 5.00
Begin VB.Form frmProcess 
   Caption         =   "Calculating"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
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
   ScaleHeight     =   7213.583
   ScaleMode       =   0  'User
   ScaleWidth      =   9799.47
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   322
      Left            =   1440
      TabIndex        =   5
      Top             =   6720
      Width           =   1212
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   322
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   1212
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
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
      Height          =   322
      Left            =   4080
      TabIndex        =   2
      Top             =   6720
      Width           =   1212
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   322
      Left            =   2760
      TabIndex        =   1
      Top             =   6720
      Width           =   1212
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6084
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculation results:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2532
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Mlcs() As CodemlRpt         'codeml result (mlc)
Private Chi2s() As LrtRes           'LRT result
Private pauseRun As Boolean         'pause ongoing task
Private stopRun As Boolean          'stop ongoing task

Private Sub Form_Load()
Me.Show
SendMessage txtOut.hWnd, EM_SETTABSTOPS, 1, 48
Call RunPAML
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtOut.Height = Me.ScaleHeight - 1079.769
txtOut.Width = Me.ScaleWidth
cmdCopy.Top = Me.ScaleHeight - 438.562
cmdPause.Top = Me.ScaleHeight - 438.562
cmdStop.Top = Me.ScaleHeight - 438.562
cmdExit.Top = Me.ScaleHeight - 438.562
End Sub

Private Sub Form_GotFocus()
Me.Refresh
txtOut.Refresh
DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtOut.Text
MsgBox "The report has been copied to the clipboard.", vbInformation
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPause_Click()
If pauseRun = True Then
    pauseRun = False
    cmdPause.Caption = "&Pause"
Else
    pauseRun = True
    cmdPause.Caption = "&Resume"
End If
End Sub

Private Sub cmdStop_Click()
stopRun = True
End Sub

Private Sub RunPAML()
On Error Resume Next

'--variable declaration--
Dim i, J As Long                'buffer number variables
Dim S As String                 'bufer string variable
Dim n1, n2, n3, n4, n5 As Long  'buffer number variables
Dim nProt As Long               'number of model sets
Dim nLRT As Long                'number of tests
Dim nBranch As Long             'number of branches
Dim sTree As String             'origin tree without branch label
Dim sSubtree As String          'manipulated tree with branch label
Dim aOmega() As Double          'array of w ratios from w tree
Dim pChi2() As Double           'chi square test posterior possibility
Dim iPosition As Long           'current position in a tree file string
Dim iBranch As Long             'number of branch
pauseRun = False
stopRun = False

'--print task information--
txtOut.Text = "Task starts at " & CStr(Time) & " " & CStr(Date) & vbCrLf
txtOut.Text = txtOut.Text & "Working directory: " & dWork & vbCrLf
Me.Refresh

'--read tree file--
sTree = ""
Dim readTree As Boolean
readTree = False
Open fTree For Input As #1
Do While Not EOF(1)
    Line Input #1, S
    S = Trim(S)
    If Len(S) = 0 Then GoTo Continue_Do
    If Left(S, 1) = "(" Then readTree = True
    If Left(S, 2) = "//" Then Exit Do
    If readTree = True Then sTree = sTree & S
Continue_Do:
Loop
Close #1
If InStr(sTree, ";") > 0 Then sTree = Left(sTree, InStr(sTree, ";"))
Me.Refresh

sTree = Replace(sTree, vbCrLf, ""): sTree = Replace(sTree, vbCr, ""): sTree = Replace(sTree, vbLf, "")
For i = 1 To 20
    sTree = Replace(sTree, "'#" & CStr(i) & "'", ""): sTree = Replace(sTree, "#" & CStr(i), "")
Next
txtOut.Text = txtOut.Text & "Input tree file loaded successfully." & vbCrLf

'--number of branches--
nBranch = 0
For i = 1 To Len(sTree)
    If Mid(sTree, i, 1) = "," Or Mid(sTree, i, 1) = ")" Then
        nBranch = nBranch + 1
    End If
Next
ReDim aOmega(0 To nBranch - 1)
txtOut.Text = txtOut.Text & CStr(nBranch) & " branches to be analyzed." & vbCrLf
Me.Refresh

'--model sets--
txtOut.Text = txtOut.Text & "Model sets:" & vbCrLf
nProt = UBound(Ctls) + 1
For i = 1 To nProt
    txtOut.Text = txtOut.Text & CStr(i) & ":  " & Ctls(i - 1).Name & vbCrLf
Next
nLRT = 0
If (Not LRTs) <> -1 Then
    nLRT = UBound(LRTs) + 1
    txtOut.Text = txtOut.Text & "Likelihood ratio tests:" & vbCrLf
    For i = 1 To nLRT
        txtOut.Text = txtOut.Text & CStr(i) & ":  " & LRTs(i - 1).Name & vbCrLf
    Next
End If
txtOut.Text = txtOut.Text & vbCrLf

'--result array--
ReDim Mlcs(0 To nProt - 1, 0 To nBranch - 1)
If nLRT > 0 Then ReDim Chi2s(0 To nLRT - 1, 0 To nBranch - 1)

'--table header--
txtOut.Text = txtOut.Text & vbTab
For i = 1 To nProt
    txtOut.Text = txtOut.Text & Ctls(i - 1).Name & vbTab
    If Ctls(i - 1).model = 0 Then txtOut.Text = txtOut.Text & vbTab
    If Ctls(i - 1).model = 2 Then txtOut.Text = txtOut.Text & vbTab & vbTab
Next
If nLRT > 0 Then
    For i = 1 To nLRT
        txtOut.Text = txtOut.Text & LRTs(i - 1).Name & vbTab & vbTab
    Next
End If
txtOut.Text = txtOut.Text & vbCrLf

txtOut.Text = txtOut.Text & "Branch"
For i = 1 To nProt
    txtOut.Text = txtOut.Text & vbTab & "lnL"
    If Ctls(i - 1).model = 0 And Ctls(i - 1).NSsites = 0 Then 'model=0: one universal w
        txtOut.Text = txtOut.Text & vbTab & "w"
    ElseIf Ctls(i - 1).model = 2 And Ctls(i - 1).NSsites = 0 Then 'model=2: two w's
        txtOut.Text = txtOut.Text & vbTab & "w0" & vbTab & "w1"
    End If 'model=1: don't display w
Next
If nLRT > 0 Then
    For i = 1 To nLRT
        txtOut.Text = txtOut.Text & vbTab & "chi2" & vbTab & "p"
    Next
End If
txtOut.Text = txtOut.Text & vbCrLf
Me.Refresh
iBranch = 0

'--here we go--

Do
    '--generate sub tree with branch label--
    n1 = InStr(iPosition + 1, sTree, ",")
    n2 = InStr(iPosition + 1, sTree, ")")
    If n1 = 0 And n2 = 0 Then Exit Do
    iBranch = iBranch + 1
    txtOut.Text = txtOut.Text & CStr(iBranch)
    Me.Refresh
    iPosition = 0
    If n1 = 0 Then iPosition = n2
    If n2 = 0 Then iPosition = n1
    If iPosition = 0 Then iPosition = IIf(n1 < n2, n1, n2)
    sSubtree = IIf(iPosition > 1, Left(sTree, iPosition - 1), "") & " #1 " & (Right(sTree, Len(sTree) - iPosition + 1))
    Me.Caption = "Running ... " & CStr(iBranch) & " of " & CStr(nBranch)
    Me.Refresh
    DoEvents
    
    Dim dSubWork As String
    For i = 1 To nProt
        Do While pauseRun = True
            Delay 3
            If stopRun Then Me.Caption = "Stopped": Exit Sub
        Loop
        If stopRun Then Me.Caption = "Stopped": Exit Sub

        '--skip duplicated runs--
        If (Ctls(i - 1).model = "0" Or Ctls(i - 1).model = "1") And iBranch > 1 Then
            Mlcs(i - 1, iBranch - 1) = Mlcs(i - 1, 0)
            txtOut.Text = txtOut.Text & vbTab & "."
            If Ctls(i - 1).model = "0" Then txtOut.Text = txtOut.Text & vbTab & "."
            GoTo nxtFor:
        End If
        
        '--create input files--
        dSubWork = dWork & "m" & CStr(i) & "b" & CStr(iBranch) & "\"
        MkDir dSubWork
        FileCopy fSequence, dSubWork & "inseq"
        Open dSubWork & "insubtree" For Output As #1
        Print #1, sSubtree
        Close #1
        DoEvents
        Open dSubWork & "codeml.ctl" For Output As #1
        With Ctls(i - 1)
            Print #1, "      seqfile = inseq"
            Print #1, "     treefile = insubtree"
            Print #1, "      outfile = mlc"
            Print #1, "        noisy = " & .noisy
            Print #1, "      verbose = " & .verbose
            Print #1, "      runmode = " & .runmode
            Print #1, "      seqtype = " & .seqtype
            Print #1, "    CodonFreq = " & .CodonFreq
            Print #1, "        clock = " & .clock
            Print #1, "       aaDist = " & .aaDist
            Print #1, "   aaRatefile = " & .aaRatefile
            Print #1, "        model = " & .model
            Print #1, "      NSsites = " & .NSsites
            Print #1, "        icode = " & .icode
            Print #1, "    fix_kappa = " & .fix_kappa
            Print #1, "        kappa = " & .Kappa
            Print #1, "    fix_omega = " & .fix_omega
            Print #1, "        omega = " & .omega
            If .fix_alpha <> "" Then
                Print #1, "    fix_alpha = " & .fix_alpha
                Print #1, "        alpha = " & .alpha
            End If
            Print #1, "       Malpha = " & .Malpha
            Print #1, "        ncatG = " & .ncatG
            If .fix_rho <> "" Then
                Print #1, "      fix_rho = " & .fix_rho
                Print #1, "          rho = " & .rho
            End If
            Print #1, "        getSE = " & .getSE
            Print #1, " RateAncestor = " & .RateAncestor
            If .cleandata <> "" Then Print #1, "    cleandata = " & .cleandata
            If .Small_Diff <> "" Then Print #1, "   Small_Diff = " & .Small_Diff
            If .fix_blength <> "" Then Print #1, "  fix_blength = " & .fix_blength
            Print #1, "       method = " & .method
        End With
        Close #1
        DoEvents
        
        '--run codeml.exe--
        ExShell fCodeml, dSubWork, wSize, NORMAL_PRIORITY_CLASS
        DoEvents
        
        '--analyze result--
        With Mlcs(i - 1, iBranch - 1)
        Dim readOmega As Boolean
        readOmega = False
        .Name = Ctls(i - 1).Name
        
        Open dSubWork & "mlc" For Input As #1
        Do While Not EOF(1)
            Line Input #1, S
            S = Replace(S, vbCrLf, ""): S = Replace(S, vbCr, ""): S = Replace(S, vbLf, "")
            '--read lnL--
            If Left(S, 10) = "lnL(ntime:" Then
                n1 = InStr(10, S, "):")
                S = Right(S, Len(S) - n1 - 1)
                n1 = InStrRev(S, " ")
                S = Left(S, n1 - 1)
                .lnL = CDbl(S)
                txtOut.Text = txtOut.Text & vbTab & Format(.lnL, "0.000")
            '--read w1, w0--
            ElseIf Left(S, 6) = " omega" And Right(S, 5) = "fixed" Then
                .Omega_0 = CDbl(Mid(S, 10, Len(S) - 15))
            ElseIf Left(S, 15) = "omega (dN/dS) =" Then
                readOmega = True
                .Omega_0 = CDbl(Right(S, Len(S) - 17))
                If Ctls(i - 1).NSsites = 0 Then txtOut.Text = txtOut.Text & vbTab & Digitrim(.Omega_0)
            ElseIf Left(S, 23) = "w (dN/dS) for branches:" Then
                readOmega = True
                S = Right(S, Len(S) - 25)
                n1 = InStr(1, S, " ")
                If n1 > 0 Then
                    .Omega_1 = CDbl(Left(S, n1 - 1))
                    S = Right(S, Len(S) - n1)
                    n2 = InStr(1, S, " ")
                    If n2 > 0 Then
                        .Omega_0 = CDbl(Left(S, n2 - 1))
                    Else
                        .Omega_0 = CDbl(S)
                        If Ctls(i - 1).NSsites = 0 Then txtOut.Text = txtOut.Text & vbTab & Digitrim(.Omega_0) & vbTab & Digitrim(.Omega_1)
                    End If
                Else
                    .Omega_0 = CDbl(S)
                    If Ctls(i - 1).NSsites = 0 Then txtOut.Text = txtOut.Text & vbTab & Digitrim(.Omega_0)
                End If
            End If
            '--read kappa-- to be added...
        Loop
        Close #1
        If Not readOmega Then
            If Ctls(i - 1).NSsites = 0 Then txtOut.Text = txtOut.Text & vbTab & Digitrim(.Omega_0)
        End If
        End With
        If Outputs = 1 Then
            FileCopy dSubWork & "mlc", Left(dSubWork, Len(dSubWork) - 1) & ".txt"
        End If
        If Outputs < 2 Then
            Kill dSubWork & "*"
            RmDir dSubWork
        End If
        Me.Refresh
nxtFor:
    Next
    If nLRT > 0 Then
        For i = 1 To nLRT
            With Chi2s(i - 1, iBranch - 1)
                .H0 = LRTs(i - 1).H0
                .H1 = LRTs(i - 1).H1
                .df = LRTs(i - 1).df
                For J = 1 To nProt
                    If Mlcs(J - 1, iBranch - 1).Name = LRTs(i - 1).H0 Then .lnL0 = Mlcs(J - 1, iBranch - 1).lnL
                    If Mlcs(J - 1, iBranch - 1).Name = LRTs(i - 1).H1 Then .lnL1 = Mlcs(J - 1, iBranch - 1).lnL
                Next
                .pChi2 = 1 - ChiSquareDistribution(.df, 2 * (.lnL1 - .lnL0))
                txtOut.Text = txtOut.Text & vbTab & Digitrim(2 * (.lnL1 - .lnL0)) & vbTab & Digitrim(.pChi2)
            End With
        Next
        Me.Refresh
    End If
    txtOut.Text = txtOut.Text & vbCrLf
Loop

'--dN/dS tree--
txtOut.Text = txtOut.Text & vbCrLf & "Foreground dN/dS value (w1) trees:" & vbCrLf
For i = 1 To nProt
    If Ctls(i - 1).model = 2 And Ctls(i - 1).NSsites = 0 Then
        txtOut.Text = txtOut.Text & CStr(i) & ": " & Ctls(i - 1).Name & vbCrLf
        Dim w1Tree As String    'foreground dN/dS tree
        iPosition = 0
        For J = 0 To nBranch - 1
            n1 = InStr(iPosition + 1, sTree, ",")
            n2 = InStr(iPosition + 1, sTree, ")")
            n3 = iPosition
            iPosition = 0
            If n1 = 0 Then iPosition = n2
            If n2 = 0 Then iPosition = n1
            If iPosition = 0 Then iPosition = IIf(n1 < n2, n1, n2)
            n1 = InStr(n3 + 1, sTree, ":")
            If n1 = 0 Then
                n2 = iPosition
            Else
                n2 = n1
            End If
            w1Tree = w1Tree & Mid(sTree, n3 + 1, n2 - n3 - 1) & " " & Digitrim(Mlcs(i - 1, J).Omega_1) & " " & Mid(sTree, n2, iPosition - n2 + 1)
        Next
        w1Tree = w1Tree & ";"
        txtOut.Text = txtOut.Text & w1Tree & vbCrLf
        Open dWork & "m" & CStr(i) & "w1.nwk" For Output As #3
        Print #3, w1Tree
        Close #3
    End If
Next

'--p-value tree--
txtOut.Text = txtOut.Text & vbCrLf & "LRT p-value trees:" & vbCrLf
For i = 1 To nLRT
    txtOut.Text = txtOut.Text & CStr(i) & ": " & LRTs(i - 1).Name & vbCrLf
    Dim pTree As String     'LRT p-value tree
    iPosition = 0
    For J = 0 To nBranch - 1
        n1 = InStr(iPosition + 1, sTree, ",")
        n2 = InStr(iPosition + 1, sTree, ")")
        n3 = iPosition
        iPosition = 0
        If n1 = 0 Then iPosition = n2
        If n2 = 0 Then iPosition = n1
        If iPosition = 0 Then iPosition = IIf(n1 < n2, n1, n2)
        n1 = InStr(n3 + 1, sTree, ":")
        If n1 = 0 Then
            n2 = iPosition
        Else
            n2 = n1
        End If
        pTree = pTree & Mid(sTree, n3 + 1, n2 - n3 - 1) & " " & Digitrim(Chi2s(i - 1, J).pChi2) & " " & Mid(sTree, n2, iPosition - n2 + 1)
    Next
    pTree = pTree & ";"
    txtOut.Text = txtOut.Text & pTree & vbCrLf
    Open dWork & "t" & CStr(i) & "p.nwk" For Output As #3
    Print #3, pTree
    Close #3
Next

'--finish up--
txtOut.Text = txtOut.Text & vbCrLf & "Task ends at " & CStr(Time) & " " & CStr(Date) & vbCrLf
Open dWork & "report.txt" For Output As #5
Print #5, txtOut.Text
Close #5
Me.Caption = "Done"
MsgBox "Task completed." & vbCrLf & vbCrLf & "The report has been saved to:" & vbCrLf & dWork & "report.txt", vbInformation
End Sub


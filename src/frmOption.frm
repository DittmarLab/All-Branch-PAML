VERSION 5.00
Begin VB.Form frmOption 
   Caption         =   "Control file"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ckFix_blength 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10560
      TabIndex        =   68
      Top             =   5304
      Value           =   1  'Checked
      Width           =   252
   End
   Begin VB.CheckBox ckSmall_Diff 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8520
      TabIndex        =   67
      Top             =   4608
      Width           =   252
   End
   Begin VB.TextBox txName 
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
      Left            =   240
      TabIndex        =   66
      Top             =   840
      Width           =   5055
   End
   Begin VB.CheckBox ckRho 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10680
      TabIndex        =   64
      Top             =   3156
      Width           =   252
   End
   Begin VB.CheckBox ckAlpha 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10680
      TabIndex        =   63
      Top             =   2796
      Width           =   252
   End
   Begin VB.ComboBox cbMethod 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0000
      Left            =   7200
      List            =   "frmOption.frx":000A
      TabIndex        =   61
      Text            =   "0 - simultaneous"
      ToolTipText     =   "Used with option G, for combined analysis of data from multiple genes or multiple site partitions"
      Top             =   5640
      Width           =   3732
   End
   Begin VB.ComboBox cbFix_blength 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":003A
      Left            =   7200
      List            =   "frmOption.frx":004A
      TabIndex        =   59
      Text            =   "0 - ignore"
      ToolTipText     =   "Type of sequence"
      Top             =   5280
      Width           =   3252
   End
   Begin VB.CheckBox ckCleandata 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8520
      TabIndex        =   58
      Top             =   4956
      Width           =   252
   End
   Begin VB.ComboBox cbCleandata 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":007F
      Left            =   7200
      List            =   "frmOption.frx":0089
      TabIndex        =   56
      Text            =   "0 - No"
      ToolTipText     =   "If sites contain ambiguity characters or gaps"
      Top             =   4920
      Width           =   1212
   End
   Begin VB.TextBox txSmall_Diff 
      Enabled         =   0   'False
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
      Left            =   7200
      TabIndex        =   55
      Text            =   ".5e-6"
      ToolTipText     =   "A small value used in the difference approximation of derivatives"
      Top             =   4560
      Width           =   1212
   End
   Begin VB.TextBox txncatG 
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
      Left            =   9960
      TabIndex        =   53
      Text            =   "1"
      ToolTipText     =   "Number of categories in dG of NSsites models"
      Top             =   4560
      Width           =   972
   End
   Begin VB.ComboBox cbMalpha 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":009E
      Left            =   7200
      List            =   "frmOption.frx":00A8
      TabIndex        =   50
      Text            =   "0 - one alpha"
      ToolTipText     =   "Type of sequence"
      Top             =   4200
      Width           =   3732
   End
   Begin VB.ComboBox cbGetSE 
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
      ItemData        =   "frmOption.frx":00D1
      Left            =   1560
      List            =   "frmOption.frx":00DB
      TabIndex        =   48
      Text            =   "0 - Don't want S.E."
      ToolTipText     =   "Whether estimates of the standard errors of estimated parameters are wanted"
      Top             =   4440
      Width           =   3735
   End
   Begin VB.ComboBox cbMgene 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0111
      Left            =   7200
      List            =   "frmOption.frx":0124
      TabIndex        =   46
      Text            =   "0 - same kappa and pi, different c"
      ToolTipText     =   "Used with option G, for combined analysis of data from multiple genes or multiple site partitions"
      Top             =   3840
      Width           =   3732
   End
   Begin VB.ComboBox cbiCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":01EE
      Left            =   7200
      List            =   "frmOption.frx":01FB
      TabIndex        =   44
      Text            =   "0 - universal code"
      ToolTipText     =   "Genetic code"
      Top             =   3480
      Width           =   3732
   End
   Begin VB.CommandButton cmdaaRatefile 
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
      Height          =   324
      Left            =   10600
      TabIndex        =   43
      Top             =   2400
      Width           =   324
   End
   Begin VB.TextBox txaaRatefile 
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
      Left            =   7200
      TabIndex        =   42
      ToolTipText     =   "Number of separate data sets in the file"
      Top             =   2400
      Width           =   3300
   End
   Begin VB.ComboBox cbaaDist 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":023F
      Left            =   7200
      List            =   "frmOption.frx":0261
      TabIndex        =   39
      Text            =   "0 - equal"
      ToolTipText     =   "Whether equal amino acid distances are assumed (= 0) or Grantham's matrix is used (= 1)"
      Top             =   2040
      Width           =   3732
   End
   Begin VB.CheckBox ckNdata 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8640
      TabIndex        =   38
      Top             =   1716
      Width           =   252
   End
   Begin VB.TextBox txNdata 
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
      Left            =   7200
      TabIndex        =   36
      Text            =   "10"
      ToolTipText     =   "Number of separate data sets in the file"
      Top             =   1680
      Width           =   1332
   End
   Begin VB.ComboBox cbCodonFreq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":02CD
      Left            =   7200
      List            =   "frmOption.frx":02DD
      TabIndex        =   34
      Text            =   "2 - F3X4"
      ToolTipText     =   "The equilibrium codon frequencies in codon substitution model"
      Top             =   1320
      Width           =   3732
   End
   Begin VB.ComboBox cbRunmode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0315
      Left            =   7200
      List            =   "frmOption.frx":0325
      TabIndex        =   32
      Text            =   "0 - user tree"
      ToolTipText     =   "Type of sequence"
      Top             =   960
      Width           =   3732
   End
   Begin VB.ComboBox cbVerbose 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0378
      Left            =   7200
      List            =   "frmOption.frx":0382
      TabIndex        =   30
      Text            =   "0 - concise output"
      ToolTipText     =   "Type of sequence"
      Top             =   600
      Width           =   3732
   End
   Begin VB.ComboBox cbNoisy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":03AF
      Left            =   7200
      List            =   "frmOption.frx":03C2
      TabIndex        =   28
      Text            =   "3"
      ToolTipText     =   "How much output to display on the screen"
      Top             =   240
      Width           =   3732
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export >"
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
      Left            =   1680
      TabIndex        =   27
      Top             =   5040
      Width           =   1452
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import <"
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
      Left            =   240
      TabIndex        =   26
      Top             =   5040
      Width           =   1452
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Advanced >>"
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
      Left            =   3840
      TabIndex        =   25
      Top             =   5040
      Width           =   1452
   End
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
      Left            =   3840
      TabIndex        =   24
      Top             =   5760
      Width           =   1452
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
      Left            =   240
      TabIndex        =   23
      Top             =   5760
      Width           =   1452
   End
   Begin VB.ComboBox cbClock 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":03D5
      Left            =   1560
      List            =   "frmOption.frx":03E5
      TabIndex        =   21
      Text            =   "0 - no clock"
      ToolTipText     =   "Specify Models concerning rate constancy or variation among lineages"
      Top             =   2040
      Width           =   3732
   End
   Begin VB.ComboBox cbSeqtype 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0420
      Left            =   1560
      List            =   "frmOption.frx":042D
      TabIndex        =   19
      Text            =   "1 - codons"
      ToolTipText     =   "Type of sequence"
      Top             =   1680
      Width           =   3732
   End
   Begin VB.ComboBox cbRateAncestor 
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
      ItemData        =   "frmOption.frx":0459
      Left            =   1560
      List            =   "frmOption.frx":0463
      TabIndex        =   18
      Text            =   "0 - Don't reconstruct"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.ComboBox cbFixRho 
      Enabled         =   0   'False
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
      ItemData        =   "frmOption.frx":048F
      Left            =   8520
      List            =   "frmOption.frx":0499
      TabIndex        =   16
      Text            =   "1 - fixed"
      Top             =   3120
      Width           =   2052
   End
   Begin VB.ComboBox cbFixAlpha 
      Enabled         =   0   'False
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
      ItemData        =   "frmOption.frx":04BD
      Left            =   8520
      List            =   "frmOption.frx":04C7
      TabIndex        =   15
      Text            =   "1 - fixed"
      Top             =   2760
      Width           =   2052
   End
   Begin VB.ComboBox cbFixOmega 
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
      ItemData        =   "frmOption.frx":04EB
      Left            =   3000
      List            =   "frmOption.frx":04F5
      TabIndex        =   14
      Text            =   "0 - to be estimated"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.ComboBox cbFixKappa 
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
      ItemData        =   "frmOption.frx":0519
      Left            =   3000
      List            =   "frmOption.frx":0523
      TabIndex        =   13
      Text            =   "0 - to be estimated"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txKappa 
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
      Left            =   1560
      TabIndex        =   12
      Text            =   "2"
      ToolTipText     =   "The transition/transversion ratio of nucleic acid substitution"
      Top             =   3240
      Width           =   1332
   End
   Begin VB.TextBox txOmega 
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
      Left            =   1560
      TabIndex        =   10
      Text            =   "1"
      ToolTipText     =   "dN/dS, the ratio of the non-synonymous substitution rate to the synonymous substitution rate"
      Top             =   3600
      Width           =   1332
   End
   Begin VB.TextBox txAlpha 
      Enabled         =   0   'False
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
      Left            =   7200
      TabIndex        =   8
      Text            =   "1"
      ToolTipText     =   "The shape parameter alpha of the gamma distribution for variable substitution rates across sites"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txRho 
      Enabled         =   0   'False
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
      Left            =   7200
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "The correlation parameter of the auto-discrete-gamma model"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cbNSsites 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":0547
      Left            =   1560
      List            =   "frmOption.frx":0575
      TabIndex        =   4
      Text            =   "0 - one omega"
      Top             =   2760
      Width           =   3732
   End
   Begin VB.ComboBox cbModel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOption.frx":065D
      Left            =   1560
      List            =   "frmOption.frx":066A
      TabIndex        =   2
      Text            =   "0 - one-ratio model"
      Top             =   2400
      Width           =   3732
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   5280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Name this model set:"
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
      Left            =   240
      TabIndex        =   65
      Top             =   480
      Width           =   4092
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "   method = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   62
      Top             =   5670
      Width           =   1215
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "fix_blength = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5640
      TabIndex        =   60
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "cleandata = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   57
      Top             =   4950
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Small_Diff = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5640
      TabIndex        =   54
      Top             =   4590
      Width           =   1455
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "    ncatG = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8760
      TabIndex        =   52
      Top             =   4590
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "    Malpha = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   51
      Top             =   4230
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "getSE ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   49
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "    Mgene = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   47
      Top             =   3870
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "icode  = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5760
      TabIndex        =   45
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "aaRatefile = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5760
      TabIndex        =   41
      Top             =   2430
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "aaDist  = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   6000
      TabIndex        =   40
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "ndata ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   6360
      TabIndex        =   37
      Top             =   1710
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " CodonFreq  = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5760
      TabIndex        =   35
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "runmode = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5760
      TabIndex        =   33
      Top             =   990
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   5520
      X2              =   5520
      Y1              =   120
      Y2              =   6120
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "verbose = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   31
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "    noisy = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   5880
      TabIndex        =   29
      Top             =   270
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "    clock = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   22
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "seqtype = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   20
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "RateAncestor ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -240
      TabIndex        =   17
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "kappa ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   11
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "omega ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   9
      Top             =   3660
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "alpha ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   6360
      TabIndex        =   7
      Top             =   2790
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "rho ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   6360
      TabIndex        =   5
      Top             =   3150
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "NSsites = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   2790
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "model = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Model parameters to use in ""condeml.ctl"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private C As CodemlCtl

Private Function strChop(strIn As String) As String
Dim iPos As Long
Dim sBuffer As String
If Len(strIn) <= 2 Then
    sBuffer = strIn
Else
    iPos = InStr(2, strIn, "-")
    If iPos <> 0 Then
        sBuffer = Left(strIn, iPos - 1)
    Else
        sBuffer = strIn
    End If
End If
If Right(sBuffer, 1) = " " Then sBuffer = Left(sBuffer, Len(sBuffer) - 1)
strChop = sBuffer
End Function

Private Sub cbFixKappa_LostFocus()
If strChop(cbFixKappa.Text) = "0" Then
    txKappa.Enabled = True
Else
    txKappa.Enabled = False
End If
End Sub

Private Sub cbFixOmega_LostFocus()
If strChop(cbFixOmega.Text) = "0" Then
    txOmega.Enabled = True
Else
    txOmega.Enabled = False
End If
End Sub
Private Sub cbFixAlpha_LostFocus()
If strChop(cbFixAlpha.Text) = "0" Then
    txAlpha.Enabled = False
Else
    txAlpha.Enabled = True
End If
End Sub
Private Sub cbFixRho_LostFocus()
If strChop(cbFixRho.Text) = "0" Then
    txRho.Enabled = False
Else
    txRho.Enabled = True
End If
End Sub

Private Sub ckFix_blength_Click()
If ckFix_blength.Value = 1 Then
    cbFix_blength.Enabled = True
Else
    cbFix_blength.Enabled = False
End If
End Sub

Private Sub ckNdata_Click()
If ckNdata.Value = 1 Then
    txNdata.Enabled = True
Else
    txNdata.Enabled = False
End If
End Sub

Private Sub ckSmall_Diff_Click()
If ckSmall_Diff.Value = 1 Then
    txSmall_Diff.Enabled = True
Else
    txSmall_Diff.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdExport_Click()
Dim S As String
S = sSaveFile("*.*" + Chr$(0) + "*.*" + Chr$(0), , , "codeml.ctl")
If S = "" Then Exit Sub
If Dir(S) <> "" Then
    If MsgBox("The export file" & vbCrLf & S & vbCrLf & "already exists." & vbCrLf & vbCrLf & "Overwrite?", vbOKCancel + vbQuestion) = vbCancel Then Exit Sub
End If
Open S For Output As #1
Print #1, "* Title: " & txName.Text
Print #1, "* Date: " & CStr(Time) & CStr(Date)
Print #1, ""
Print #1, "      seqfile = "
Print #1, "     treefile = "
Print #1, "      outfile = "
Print #1, ""
Print #1, "        noisy = " & strChop(cbNoisy.Text)
Print #1, "      verbose = " & strChop(cbVerbose.Text)
Print #1, "      runmode = " & strChop(cbRunmode.Text)
Print #1, ""
Print #1, "      seqtype = " & strChop(cbSeqtype.Text)
Print #1, "    CodonFreq = " & strChop(cbCodonFreq.Text)
Print #1, "        clock = " & strChop(cbClock.Text)
Print #1, "       aaDist = " & strChop(cbaaDist.Text)
Print #1, "   aaRatefile = " & strChop(txaaRatefile.Text)
Print #1, ""
Print #1, "        model = " & strChop(cbModel.Text)
Print #1, "      NSsites = " & strChop(cbNSsites.Text)
Print #1, "        icode = " & strChop(cbiCode.Text)
Print #1, ""
Print #1, "    fix_kappa = " & strChop(cbFixKappa.Text)
Print #1, "        kappa = " & strChop(txKappa.Text)
Print #1, "    fix_omega = " & strChop(cbFixOmega.Text)
Print #1, "        omega = " & strChop(txOmega.Text)
Print #1, ""
If ckAlpha.Value = 1 Then
    Print #1, "    fix_alpha = " & strChop(cbFixKappa.Text)
    Print #1, "        alpha = " & strChop(txKappa.Text)
End If
Print #1, "       Malpha = " & strChop(cbMalpha.Text)
Print #1, "        ncatG = " & strChop(txncatG.Text)
If ckRho.Value = 1 Then
    Print #1, "      fix_rho = " & strChop(cbFixRho.Text)
    Print #1, "          rho = " & strChop(txRho.Text)
End If
Print #1, ""
Print #1, "        getSE = " & strChop(cbGetSE.Text)
Print #1, " RateAncestor = " & strChop(cbRateAncestor.Text)
If ckCleandata.Value = 1 Then Print #1, "    cleandata = " & strChop(cbCleandata.Text)
If ckSmall_Diff.Value = 1 Then Print #1, "   Small_Diff = " & strChop(txSmall_Diff.Text)
If ckFix_blength.Value = 1 Then Print #1, "  fix_blength = " & strChop(cbFix_blength.Text)
Print #1, "       method = " & strChop(cbMethod.Text)
Print #1, ""
Print #1, "* Generated by BatchPAML 0.1"
Close #1
MsgBox "The model set has been sucessfully exported to" & vbCrLf & S & vbCrLf & vbCrLf & "Note: Specify seq/tree/out file before use.", vbInformation
End Sub

Private Sub cbRateAncestor_LostFocus()
If strChop(cbRateAncestor.Text) = "1" Then
cbiCode.Enabled = True
Else
cbiCode.Enabled = False
End If
End Sub

Private Sub ckAlpha_Click()
If ckAlpha.Value = 1 Then
    txAlpha.Enabled = True
    cbFixAlpha.Enabled = True
Else
    txAlpha.Enabled = False
    cbFixAlpha.Enabled = False
End If
End Sub

Private Sub ckCleandata_Click()
If ckCleandata.Value = 1 Then
    cbCleandata.Enabled = True
Else
    cbCleandata.Enabled = False
End If
End Sub

Private Sub ckRho_Click()
If ckRho.Value = 1 Then
    txRho.Enabled = True
    cbFixRho.Enabled = True
Else
    txRho.Enabled = False
    cbFixRho.Enabled = False
End If
End Sub

Private Sub cmdaaRatefile_Click()
txaaRatefile.Text = sOpenFile
End Sub

Private Sub cmdImport_Click()
Dim S As String 'file name
Dim L As String 'line of file
S = sOpenFile("codeml.ctl" + Chr$(0) + "codeml.ctl" + Chr$(0) + "*.*" + Chr$(0) + "*.*" + Chr$(0))
If S = "" Then Exit Sub
If Dir(S) = "" Then
    MsgBox "The import file" & vbCrLf & S & vbCrLf & "doesn't exist.", vbExclamation
    Exit Sub
End If
Open S For Input As #1
Do Until EOF(1)
' ! This program does not recognized Unix coded text file.
    'L = ""
    'Do
    '    L = L & Input(1, #1)
    '    If Right(L, 1) = Chr(10) Or Right(L, 1) = Chr(13) Then Exit Do
    'Loop
    'Do While Left(L, 1) = Chr(10) Or Left(L, 1) = Chr(13)
    '    L = Right(L, Len(L) - 1)
    'Loop
    'If Right(L, 1) = Chr(10) Or Right(L, 1) = Chr(13) Then L = Left(L, Len(L) - 1)
    Line Input #1, L
    If InStr(1, L, "*") <> 0 Then L = Left(L, InStr(1, L, "*") - 1)
    L = Trim(L)
    If LCase(Left(L, 5)) = "noisy" Then cbNoisy.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 7)) = "verbose" Then cbVerbose.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 7)) = "runmode" Then cbRunmode.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 7)) = "seqtype" Then cbSeqtype.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 9)) = "codonfreq" Then cbCodonFreq.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "clock" Then cbClock.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 6)) = "aadist" Then cbaaDist.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 10)) = "aaratefile" Then txaaRatefile.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "model" Then cbModel.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 7)) = "nssites" Then cbNSsites.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "icode" Then cbiCode.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 9)) = "fix_kappa" Then cbFixKappa.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "kappa" Then txKappa.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 9)) = "fix_alpha" Then
        ckAlpha.Value = 1
        cbFixAlpha.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    End If
    If LCase(Left(L, 5)) = "alpha" Then txAlpha.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 9)) = "fix_omega" Then cbFixOmega.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "omega" Then txOmega.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 7)) = "fix_rho" Then
        ckRho.Value = 1
        cbFixRho.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    End If
    If LCase(Left(L, 3)) = "rho" Then txRho.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 6)) = "malpha" Then cbMalpha.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "ncatg" Then txncatG.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 5)) = "getse" Then cbGetSE.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 12)) = "rateancestor" Then cbRateAncestor.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    If LCase(Left(L, 9)) = "cleandata" Then
        ckCleandata.Value = 1
        cbCleandata.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    End If
    If LCase(Left(L, 10)) = "small_diff" Then
        ckSmall_Diff.Value = 1
        txSmall_Diff.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
    End If
    If LCase(Left(L, 11)) = "fix_blength" Then
        cbFix_blength.Text = Trim(Right(L, Len(L) - InStr(1, L, "=")))
        ckFix_blength.Value = 1
    End If
Loop
Close #1
End Sub

Private Sub cmdMore_Click()
If cmdMore.Caption = "Advanced >>" Then
    frmOption.Width = 11328
    cmdMore.Caption = "<< Basic"
Else
    frmOption.Width = 5515
    cmdMore.Caption = "Advanced >>"
End If
End Sub

Private Sub cmdOK_Click()
For i = 0 To UBound(Ctls())
    If txName.Text = Ctls(i).Name And i <> iCtl Then
        MsgBox "The name of the model set " & Chr(34) & txName.Text & Chr(34) & " has been used." & vbCrLf & vbCrLf & "Please specify another one.", vbExclamation
        Exit Sub
    End If
Next
With Ctls(iCtl)
    .Name = strChop(txName.Text)
    .seqtype = strChop(cbSeqtype.Text)
    .clock = strChop(cbClock.Text)
    .model = strChop(cbModel.Text)
    .NSsites = strChop(cbNSsites.Text)
    .fix_kappa = strChop(cbFixKappa.Text)
    .fix_omega = strChop(cbFixOmega.Text)
    .fix_alpha = IIf(ckAlpha.Value = 1, strChop(cbFixAlpha.Text), "")
    .fix_rho = IIf(ckRho.Value = 1, strChop(cbFixRho.Text), "")
    .Kappa = strChop(txKappa.Text)
    .omega = strChop(txOmega.Text)
    .alpha = strChop(txAlpha.Text)
    .rho = strChop(txRho.Text)
    .RateAncestor = strChop(cbRateAncestor.Text)
    .getSE = strChop(cbGetSE.Text)
    .noisy = strChop(cbNoisy.Text)
    .verbose = strChop(cbVerbose.Text)
    .runmode = strChop(cbRunmode.Text)
    .CodonFreq = strChop(cbCodonFreq.Text)
    If ckNdata.Value = 1 Then
        .ndata = strChop(txNdata.Text)
    Else
        .ndata = ""
    End If
    .aaDist = strChop(cbaaDist.Text)
    .aaRatefile = strChop(txaaRatefile.Text)
    .icode = strChop(cbiCode.Text)
    .Mgene = strChop(cbMgene.Text)
    .Malpha = strChop(cbMalpha.Text)
    .ncatG = strChop(txncatG.Text)
    If ckSmall_Diff.Value = 1 Then
        .Small_Diff = strChop(txSmall_Diff.Text)
    Else
        .Small_Diff = ""
    End If
    If ckCleandata.Value = 1 Then
        .cleandata = strChop(cbCleandata.Text)
    Else
        .cleandata = ""
    End If
    If ckFix_blength.Value = 1 Then
        .fix_blength = strChop(cbFix_blength.Text)
    Else
        .fix_blength = ""
    End If
    .method = strChop(cbMethod.Text)
End With
Unload Me
frmMain.Show
End Sub

Private Sub Form_Load()
Me.Width = 5515
Dim B As Boolean    'true - name available
Dim S As String
If Ctls(iCtl).Name = "" Then
    For i = 0 To 99
        S = "H" & i
        B = True
        For J = 0 To UBound(Ctls())
            If S = Ctls(J).Name Then B = False
        Next
        If B = True Then
            txName.Text = S
            Exit For
        End If
    Next
Else
    With Ctls(iCtl)
        txName.Text = .Name
        cbSeqtype.Text = .seqtype
        cbClock.Text = .clock
        cbModel.Text = .model
        cbNSsites.Text = .NSsites
        cbFixKappa.Text = .fix_kappa
        cbFixOmega.Text = .fix_omega
        cbFixAlpha.Text = .fix_alpha
        cbFixRho.Text = .fix_rho
        txKappa.Text = .Kappa
        txOmega.Text = .omega
        txAlpha.Text = .alpha
        txRho.Text = .rho
        cbRateAncestor.Text = .RateAncestor
        cbGetSE.Text = .getSE
        cbNoisy.Text = .noisy
        cbVerbose.Text = .verbose
        cbRunmode.Text = .runmode
        cbCodonFreq.Text = .CodonFreq
        If .ndata = "" Then
            ckNdata.Value = 0
        Else
            ckNdata.Value = 1
            txNdata.Text = .ndata
        End If
        cbaaDist.Text = .aaDist
        txaaRatefile.Text = .aaRatefile
        cbiCode.Text = .icode
        cbMgene.Text = .Mgene
        cbMalpha.Text = .Malpha
        txncatG.Text = .ncatG
        If .Small_Diff = "" Then
            ckSmall_Diff.Value = 0
        Else
            ckSmall_Diff.Value = 1
            txSmall_Diff.Text = .Small_Diff
        End If
        If .cleandata = "" Then
            ckCleandata.Value = 0
        Else
            ckCleandata.Value = 1
            cbCleandata.Text = .cleandata
        End If
        If .fix_blength = "" Then
            ckFix_blength.Value = 0
        Else
            ckFix_blength.Value = 1
            cbFix_blength.Text = .fix_blength
        End If
        cbMethod.Text = .method
    End With
End If
If strChop(cbFixKappa.Text) = "1" Then txKappa.Enabled = False
If strChop(cbFixOmega.Text) = "1" Then txOmega.Enabled = False
If strChop(cbFixAlpha.Text) = "0" Then txAlpha.Enabled = False
If strChop(cbFixRho.Text) = "0" Then txRho.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub


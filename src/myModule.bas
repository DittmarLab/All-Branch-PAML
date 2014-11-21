Attribute VB_Name = "myModule"
Public Type CodemlCtl 'parameter set (codeml.ctl)
    Name As String
    noisy As String
    verbose As String
    runmode As String
    seqtype As String
    CodonFreq As String
    ndata As String
    clock As String
    aaDist As String
    aaRatefile As String
    model As String
    NSsites As String
    icode As String
    Mgene As String
    fix_kappa As String
    Kappa As String
    fix_omega As String
    omega As String
    fix_alpha As String
    alpha As String
    Malpha As String
    ncatG As String
    fix_rho As String
    rho As String
    getSE As String
    RateAncestor As String
    Small_Diff As String
    cleandata As String
    fix_blength As String
    method As String
    df As Long 'degree of freedom
End Type

Public Type LRT 'likelihood ratio test
    Name As String
    H0 As String
    H1 As String
    df As Integer
End Type

Public Type CodemlRpt 'codeml output (mlc)
    Name As String
    File As String
    lnL As Double
    Omega_1 As Double
    Omega_0 As Double
    Kappa As Double
End Type

Public Type LrtRes 'likelihood ratio test result
    Name As String
    H0 As String
    H1 As String
    lnL0 As Double
    lnL1 As Double
    df As Integer
    pChi2 As Double
End Type

Public fSequence, fTree, fCodeml, dWork As String      'source files and target pathes
Public wSize As enSW
Public Outputs As Byte '0 - none, 1 - mlc, 2 - all

Public Ctls() As CodemlCtl
Public iCtl As Long 'currently used parameter set

Public LRTs() As LRT
Public iLRT As Long











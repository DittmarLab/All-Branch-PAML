Attribute VB_Name = "libMath"
'Note:
'The following functions were extracted from the Cephes Math Library
'Release 2.8, copyright by Stephen L. Moshier. See the end of code for
'the original copyright notice and disclaimer.

Public Function ChiSquareDistribution(ByVal v As Double, _
         ByVal X As Double) As Double
    Dim Result As Double

    Result = IncompleteGamma(v / 2#, X / 2#)

    ChiSquareDistribution = Result
End Function

Public Function IncompleteGamma(ByVal a As Double, ByVal X As Double) As Double
    Dim Result As Double
    Dim IGammaEpsilon As Double
    Dim ans As Double
    Dim ax As Double
    Dim C As Double
    Dim r As Double
    Dim Tmp As Double

    IGammaEpsilon = 0.000000000000001
    If X <= 0# Or a <= 0# Then
        Result = 0#
        IncompleteGamma = Result
        Exit Function
    End If
    If X > 1# And X > a Then
        Result = 1# - IncompleteGammaC(a, X)
        IncompleteGamma = Result
        Exit Function
    End If
    ax = a * Log(X) - X - LnGamma(a, Tmp)
    If ax < -709.782712893384 Then
        Result = 0#
        IncompleteGamma = Result
        Exit Function
    End If
    ax = Exp(ax)
    r = a
    C = 1#
    ans = 1#
    Do
        r = r + 1#
        C = C * X / r
        ans = ans + C
    Loop Until C / ans <= IGammaEpsilon
    Result = ans * ax / a

    IncompleteGamma = Result
End Function

Public Function IncompleteGammaC(ByVal a As Double, _
         ByVal X As Double) As Double
    Dim Result As Double
    Dim IGammaEpsilon As Double
    Dim IGammaBigNumber As Double
    Dim IGammaBigNumberInv As Double
    Dim ans As Double
    Dim ax As Double
    Dim C As Double
    Dim yc As Double
    Dim r As Double
    Dim t As Double
    Dim Y As Double
    Dim z As Double
    Dim pk As Double
    Dim pkm1 As Double
    Dim pkm2 As Double
    Dim qk As Double
    Dim qkm1 As Double
    Dim qkm2 As Double
    Dim Tmp As Double

    IGammaEpsilon = 0.000000000000001
    IGammaBigNumber = 4.5035996273705E+15
    IGammaBigNumberInv = 2.22044604925031 * 1E-16
    If X <= 0# Or a <= 0# Then
        Result = 1#
        IncompleteGammaC = Result
        Exit Function
    End If
    If X < 1# Or X < a Then
        Result = 1# - IncompleteGamma(a, X)
        IncompleteGammaC = Result
        Exit Function
    End If
    ax = a * Log(X) - X - LnGamma(a, Tmp)
    If ax < -709.782712893384 Then
        Result = 0#
        IncompleteGammaC = Result
        Exit Function
    End If
    ax = Exp(ax)
    Y = 1# - a
    z = X + Y + 1#
    C = 0#
    pkm2 = 1#
    qkm2 = X
    pkm1 = X + 1#
    qkm1 = z * X
    ans = pkm1 / qkm1
    Do
        C = C + 1#
        Y = Y + 1#
        z = z + 2#
        yc = Y * C
        pk = pkm1 * z - pkm2 * yc
        qk = qkm1 * z - qkm2 * yc
        If qk <> 0# Then
            r = pk / qk
            t = Abs((ans - r) / r)
            ans = r
        Else
            t = 1#
        End If
        pkm2 = pkm1
        pkm1 = pk
        qkm2 = qkm1
        qkm1 = qk
        If Abs(pk) > IGammaBigNumber Then
            pkm2 = pkm2 * IGammaBigNumberInv
            pkm1 = pkm1 * IGammaBigNumberInv
            qkm2 = qkm2 * IGammaBigNumberInv
            qkm1 = qkm1 * IGammaBigNumberInv
        End If
    Loop Until t <= IGammaEpsilon
    Result = ans * ax

    IncompleteGammaC = Result
End Function

Public Function LnGamma(ByVal X As Double, ByRef SgnGam As Double) As Double
    Dim Result As Double
    Dim a As Double
    Dim B As Double
    Dim C As Double
    Dim P As Double
    Dim Q As Double
    Dim u As Double
    Dim w As Double
    Dim z As Double
    Dim I As Long
    Dim LogPi As Double
    Dim LS2PI As Double
    Dim Tmp As Double

    SgnGam = 1#
    LogPi = 1.1447298858494
    LS2PI = 0.918938533204673
    If X < -34# Then
        Q = -X
        w = LnGamma(Q, Tmp)
        P = Int(Q)
        I = Round(P)
        If I Mod 2# = 0# Then
            SgnGam = -1#
        Else
            SgnGam = 1#
        End If
        z = Q - P
        If z > 0.5 Then
            P = P + 1#
            z = P - Q
        End If
        z = Q * Sin(3.14159265358979 * z)
        Result = LogPi - Log(z) - w
        LnGamma = Result
        Exit Function
    End If
    If X < 13# Then
        z = 1#
        P = 0#
        u = X
        Do While u >= 3#
            P = P - 1#
            u = X + P
            z = z * u
        Loop
        Do While u < 2#
            z = z / u
            P = P + 1#
            u = X + P
        Loop
        If z < 0# Then
            SgnGam = -1#
            z = -z
        Else
            SgnGam = 1#
        End If
        If u = 2# Then
            Result = Log(z)
            LnGamma = Result
            Exit Function
        End If
        P = P - 2#
        X = X + P
        B = -1378.25152569121
        B = -38801.6315134638 + X * B
        B = -331612.992738871 + X * B
        B = -1162370.97492762 + X * B
        B = -1721737.0082084 + X * B
        B = -853555.664245765 + X * B
        C = 1#
        C = -351.815701436523 + X * C
        C = -17064.2106651881 + X * C
        C = -220528.590553854 + X * C
        C = -1139334.44367983 + X * C
        C = -2532523.07177583 + X * C
        C = -2018891.41433533 + X * C
        P = X * B / C
        Result = Log(z) + P
        LnGamma = Result
        Exit Function
    End If
    Q = (X - 0.5) * Log(X) - X + LS2PI
    If X > 100000000# Then
        Result = Q
        LnGamma = Result
        Exit Function
    End If
    P = 1# / (X * X)
    If X >= 1000# Then
        Q = Q + ((7.93650793650794 * 0.0001 * P - 2.77777777777778 * 0.001) * P + 8.33333333333333E-02) / X
    Else
        a = 8.11614167470508 * 0.0001
        a = -(5.95061904284301 * 0.0001) + P * a
        a = 7.93650340457717 * 0.0001 + P * a
        a = -(2.777777777301 * 0.001) + P * a
        a = 8.33333333333332 * 0.01 + P * a
        Q = Q + a / X
    End If
    Result = Q

    LnGamma = Result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Cephes Math Library Release 2.8:  June, 2000
'Copyright by Stephen L. Moshier
'
'Contributors:
'    * Sergey Bochkanov (ALGLIB project). Translation from C to
'      pseudocode.
'
'See subroutines comments for additional copyrights.
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are
'met:
'
'- Redistributions of source code must retain the above copyright
'  notice, this list of conditions and the following disclaimer.
'
'- Redistributions in binary form must reproduce the above copyright
'  notice, this list of conditions and the following disclaimer listed
'  in this license in the documentation and/or other materials
'  provided with the distribution.
'
'- Neither the name of the copyright holders nor the names of its
'  contributors may be used to endorse or promote products derived from
'  this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



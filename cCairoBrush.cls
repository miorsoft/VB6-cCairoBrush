VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCairoBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tKnot
    X             As Double
    Y             As Double
    Radius        As Double
    R             As Double
    G             As Double
    B             As Double
    A             As Double
End Type

Private Type tCURVE
    X             As Double
    Y             As Double
    FirstControlPointX As Double  'Control Points
    FirstControlPointY As Double
    SecondControlPointX As Double
    SecondControlPointY As Double
    NX            As Double
    NY            As Double
End Type

Private NP        As Long
Private MaxNP     As Long

Private P()       As tKnot
Private CurveSpine() As tCURVE
Private Curve1()  As tCURVE
Private Curve2()  As tCURVE

Private NX()      As Double
Private NY()      As Double

Private xx()      As Double
Private yy()      As Double
Private out()     As Double
Private tmp()     As Double
Private RHS()     As Double

Private Sub Class_Initialize()
    NP = -1
    MaxNP = -1

End Sub
Friend Sub Clear()
    NP = -1
End Sub

Private Sub pvMakeNormal(ByRef X#, ByRef Y#)
    Dim D#
    D = X * X + Y * Y
    If D Then
        D = 1 / Sqr(D)
        X = X * D: Y = Y * D
    End If
End Sub
Friend Sub AddPoint(ByVal X#, ByVal Y#, ByVal Radius#, ByVal R#, ByVal G#, ByVal B#, Optional A As Double = 1)
    NP = NP + 1
    If NP > MaxNP Then ReDim Preserve P(NP)
    With P(NP)
        .X = X: .Y = Y
        .Radius = Radius
        .R = R: .G = G: .B = B: .A = A
    End With

End Sub

Friend Sub DRAW(CC As cCairoContext, Optional GlobalAlpha As Double = 1, Optional BorderAlpha As Double = 1)
    Dim I&, J&, K&
    Dim DX#, DY#, D#
    Dim NXK#, NYK#

    If NP > MaxNP Then
        MaxNP = NP
        ReDim xx(NP)
        ReDim yy(NP)
        ReDim NX(NP)
        ReDim NY(NP)
        ReDim out(NP)
        ReDim tmp(NP)
        ReDim RHS(NP)
        ReDim CurveSpine(NP)
        ReDim Curve1(NP)
        ReDim Curve2(NP)
    End If


    For I = 1 To NP
        DX = P(I).X - P(I - 1).X
        DY = P(I).Y - P(I - 1).Y
        D = DX * DX + DY * DY
        If D Then
            D = 1# / Sqr(D)
            NX(I) = -DY * D
            NY(I) = DX * D
        End If
    Next

    Curve1(0).X = P(0).X - P(0).Radius * NX(1)
    Curve1(0).Y = P(0).Y - P(0).Radius * NY(1)
    Curve1(0).NX = NX(1)
    Curve1(0).NY = NY(1)
    For I = 1 To NP - 1
        DX = (NX(I + 1) + NX(I)) * 0.5
        DY = (NY(I + 1) + NY(I)) * 0.5
        'Maybe better this?
        '        DX = -(P(I + 1).Y - P(I - 1).Y)
        '        DY = (P(I + 1).X - P(I - 1).X)

        pvMakeNormal DX, DY
        Curve1(I).X = P(I).X - P(I).Radius * DX
        Curve1(I).Y = P(I).Y - P(I).Radius * DY
        Curve1(I).NX = DX
        Curve1(I).NY = DY
    Next
    Curve1(NP).X = P(NP).X - P(NP).Radius * NX(NP)
    Curve1(NP).Y = P(NP).Y - P(NP).Radius * NY(NP)
    Curve1(NP).NX = NX(NP)
    Curve1(NP).NY = NY(NP)
    CalcControlPointsFor Curve1

    Curve2(0).X = P(0).X + P(0).Radius * NX(1)
    Curve2(0).Y = P(0).Y + P(0).Radius * NY(1)
    Curve2(0).NX = NX(1)
    Curve2(0).NY = NY(1)
    For I = 1 To NP - 1
        DX = (NX(I + 1) + NX(I)) * 0.5
        DY = (NY(I + 1) + NY(I)) * 0.5
        'Maybe better this?
        '        DX = -(P(I + 1).Y - P(I - 1).Y)
        '        DY = (P(I + 1).X - P(I - 1).X)

        pvMakeNormal DX, DY
        Curve2(I).X = P(I).X + P(I).Radius * DX
        Curve2(I).Y = P(I).Y + P(I).Radius * DY
        Curve2(I).NX = DX
        Curve2(I).NY = DY
    Next
    Curve2(NP).X = P(NP).X + P(NP).Radius * NX(NP)
    Curve2(NP).Y = P(NP).Y + P(NP).Radius * NY(NP)
    Curve2(NP).NX = NX(NP)
    Curve2(NP).NY = NY(NP)
    CalcControlPointsFor Curve2



    If BorderAlpha <> 1 Then
        For I = 0 To NP
            CurveSpine(I).X = P(I).X
            CurveSpine(I).Y = P(I).Y
        Next
        CurveSpine(0).NX = NX(1)
        CurveSpine(0).NY = NY(1)
        For I = 1 To NP - 1
            DX = (NX(I + 1) + NX(I)) * 0.5
            DY = (NY(I + 1) + NY(I)) * 0.5
            pvMakeNormal DX, DY
            CurveSpine(I).NX = DX
            CurveSpine(I).NY = DY
        Next
        CurveSpine(NP).NX = NX(NP)
        CurveSpine(NP).NY = NY(NP)
        For I = 1 To NP
            NX(I) = CurveSpine(I).NX
            NY(I) = CurveSpine(I).NY
        Next
        CalcControlPointsFor CurveSpine
    End If


    Dim Mesh      As cCairoPattern
    '    Set Mesh = Cairo.CreateMeshPattern

    If BorderAlpha <> 1 Then      ' CON 'Spina Dorsale'
        For I = 0 To NP - 1
            Set Mesh = Cairo.CreateMeshPattern
            K = I + 1
            NXK = NX(K)
            NYK = NY(K)

            Mesh.BeginPatch
            Mesh.MoveTo Curve1(I).X, Curve1(I).Y
            Mesh.CurveTo Curve1(I).FirstControlPointX, Curve1(I).FirstControlPointY, _
                         Curve1(I).SecondControlPointX, Curve1(I).SecondControlPointY, _
                         Curve1(K).X - NYK, Curve1(K).Y + NXK
            Mesh.LineTo CurveSpine(K).X - NYK, CurveSpine(K).Y + NXK
            Mesh.CurveTo CurveSpine(I).SecondControlPointX, CurveSpine(I).SecondControlPointY, _
                         CurveSpine(I).FirstControlPointX, CurveSpine(I).FirstControlPointY, _
                         CurveSpine(I).X, CurveSpine(I).Y
            Mesh.LineTo Curve1(I).X, Curve1(I).Y

            Mesh.SetCornerColorRGBA 0, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha * BorderAlpha
            Mesh.SetCornerColorRGBA 1, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha * BorderAlpha
            Mesh.SetCornerColorRGBA 2, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 3, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha
            Mesh.EndPatch
            CC.Paint 1#, Mesh
        Next
        For I = 0 To NP - 1
            Set Mesh = Cairo.CreateMeshPattern
            K = I + 1
            DX = CurveSpine(I).NX
            DY = CurveSpine(I).NY
            NXK = NX(K)
            NYK = NY(K)
            Mesh.BeginPatch
            Mesh.MoveTo CurveSpine(I).X + DX, CurveSpine(I).Y + DY
            Mesh.CurveTo CurveSpine(I).FirstControlPointX + DX, CurveSpine(I).FirstControlPointY + DX, _
                         CurveSpine(I).SecondControlPointX + DX, CurveSpine(I).SecondControlPointY + DX, _
                         CurveSpine(K).X - NYK + DX, CurveSpine(K).Y + NXK + DY
            Mesh.LineTo Curve2(K).X - NYK, Curve2(K).Y + NXK

            Mesh.CurveTo Curve2(I).SecondControlPointX, Curve2(I).SecondControlPointY, _
                         Curve2(I).FirstControlPointX, Curve2(I).FirstControlPointY, _
                         Curve2(I).X, Curve2(I).Y
            Mesh.LineTo CurveSpine(I).X, CurveSpine(I).Y

            Mesh.SetCornerColorRGBA 0, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 1, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 2, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha * BorderAlpha
            Mesh.SetCornerColorRGBA 3, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha * BorderAlpha
            Mesh.EndPatch
            CC.Paint 1#, Mesh
        Next

    Else

        For I = 0 To NP - 1
            Set Mesh = Cairo.CreateMeshPattern    'Anche Qua per errore alpha
            K = I + 1
            NXK = NX(K)
            NYK = NY(K)
            Mesh.BeginPatch
            Mesh.MoveTo Curve1(I).X, Curve1(I).Y
            Mesh.CurveTo Curve1(I).FirstControlPointX, Curve1(I).FirstControlPointY, _
                         Curve1(I).SecondControlPointX, Curve1(I).SecondControlPointY, _
                         Curve1(K).X - NYK, Curve1(K).Y + NXK
            Mesh.LineTo Curve2(K).X - NYK, Curve2(K).Y + NXK
            Mesh.CurveTo Curve2(I).SecondControlPointX, Curve2(I).SecondControlPointY, _
                         Curve2(I).FirstControlPointX, Curve2(I).FirstControlPointY, _
                         Curve2(I).X, Curve2(I).Y
            Mesh.LineTo Curve1(I).X, Curve1(I).Y

            Mesh.SetCornerColorRGBA 0, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 1, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 2, P(K).R, P(K).G, P(K).B, P(K).A * GlobalAlpha
            Mesh.SetCornerColorRGBA 3, P(I).R, P(I).G, P(I).B, P(I).A * GlobalAlpha
            Mesh.EndPatch
            CC.Paint 1#, Mesh
        Next


    End If



    '''' ------------ DEBUG -----------------
'    CC.SetSourceColor vbBlack
'    CC.SetLineWidth 1
'    For I = 0 To NP
'        CC.SetSourceRGBA P(I).R, P(I).G, P(I).B, 0.75
'        '            CC.Arc CurveSpine(I).FirstControlPointX, CurveSpine(I).FirstControlPointY, 4: CC.Fill
'        '            CC.Arc CurveSpine(I).SecondControlPointX, CurveSpine(I).SecondControlPointY, 4: CC.Fill
'        '            CC.Arc Curve1(I).FirstControlPointX, Curve1(I).FirstControlPointY, 4: CC.Fill
'        '            CC.Arc Curve1(I).SecondControlPointX, Curve1(I).SecondControlPointY, 4: CC.Fill
'        '            CC.Arc Curve2(I).FirstControlPointX, Curve2(I).FirstControlPointY, 4: CC.Fill
'        '            CC.Arc Curve2(I).SecondControlPointX, Curve2(I).SecondControlPointY, 4: CC.Fill
'        CC.Arc P(I).X, P(I).Y, 8: CC.Stroke
'        CC.Arc CurveSpine(I).X, CurveSpine(I).Y, 6: CC.Stroke
'        CC.Arc Curve1(I).X, Curve1(I).Y, 6: CC.Stroke
'        CC.Arc Curve2(I).X, Curve2(I).Y, 6: CC.Stroke
'    Next
    '--------------------------------------------


End Sub



'https://www.codeproject.com/Articles/31859/Draw-a-Smooth-Curve-through-a-Set-of-2D-Points-wit

Private Sub CalcControlPointsFor(Curve() As tCURVE)
    Dim I         As Long

    If (NP = 1) Then


        ' // Special case: Bezier curve should be a straight line.
        ' // 3P1 = 2P0 + P3
        Curve(0).FirstControlPointX = (2 * Curve(0).X + Curve(1).X) / 3
        Curve(0).FirstControlPointY = (2 * Curve(0).Y + Curve(1).Y) / 3
        Curve(0).SecondControlPointX = (2 * Curve(0).FirstControlPointX + Curve(0).X) / 3
        Curve(0).SecondControlPointY = (2 * Curve(0).FirstControlPointX + Curve(0).Y) / 3
    Else

        ' Calculate first Bezier control points
        ' Right hand side vector
        ' Set right hand side X values
        'for (int i = 1; i < NP - 1; ++i)
        For I = 1 To NP - 2
            RHS(I) = 4# * Curve(I).X + 2# * Curve(I + 1).X
        Next
        RHS(0) = Curve(0).X + 2# * Curve(1).X
        RHS(NP - 1) = (8 * Curve(NP - 1).X + Curve(NP).X) * 0.5
        xx() = pvGetFirstControlPoint


        ' Set right hand side Y values
        'for (int i = 1; i < NP - 1; ++i)
        For I = 1 To NP - 2
            RHS(I) = 4 * Curve(I).Y + 2 * Curve(I + 1).Y
        Next
        RHS(0) = Curve(0).Y + 2# * Curve(1).Y
        RHS(NP - 1) = (8# * Curve(NP - 1).Y + Curve(NP).Y) * 0.5
        yy() = pvGetFirstControlPoint


        '        // Fill output arrays.
        For I = 0 To NP - 1
            Curve(I).FirstControlPointX = xx(I)
            Curve(I).FirstControlPointY = yy(I)
            If (I < (NP - 1)) Then
                Curve(I).SecondControlPointX = 2# * Curve(I + 1).X - xx(I + 1)
                Curve(I).SecondControlPointY = 2# * Curve(I + 1).Y - yy(I + 1)
            Else
                Curve(I).SecondControlPointX = (Curve(NP).X + xx(NP - 1)) * 0.5
                Curve(I).SecondControlPointY = (Curve(NP).Y + yy(NP - 1)) * 0.5
            End If
        Next
    End If

End Sub


' <summary>
' Solves a tridiagonal system for one of coordinates (x or y)
' of first Bezier control points.
' <param name="rhs">Right hand side vector.</param>
' <returns>Solution vector.</returns>
Private Function pvGetFirstControlPoint() As Double()
    Dim B         As Double
    Dim I         As Long
    B = 2#
    out(0) = RHS(0) / B
    '// Decomposition and forward substitution.
    For I = 1 To NP
        tmp(I) = 1 / B
        If I < (NP - 1) Then B = 4 - tmp(I) Else B = 3.5 - tmp(I)
        out(I) = (RHS(I) - out(I - 1)) / B
    Next
    For I = 1 To NP - 1
        out(NP - I - 1) = out(NP - I - 1) - tmp(NP - I) * out(NP - I)    ' // Backsubstitution.
    Next
    pvGetFirstControlPoint = out
End Function

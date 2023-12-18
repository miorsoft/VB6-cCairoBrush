VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "CairoBrush"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   5
      Left            =   8160
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   4
      Left            =   8160
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   3
      Left            =   8160
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   2
      Left            =   8160
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   1
      Left            =   8160
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   615
      Index           =   0
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4785
      ScaleWidth      =   5265
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Github Images"
      Height          =   615
      Left            =   8160
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SRF           As cCairoSurface
Dim CC            As cCairoContext
Dim Brush         As cCairoBrush
Dim W&, H&
Attribute H.VB_VarUserMemId = 1073938435
Dim Ima&
Dim I&
Dim X#, Y#
Attribute Y.VB_VarUserMemId = 1073938438
Dim oX#, oY#
Attribute oX.VB_VarUserMemId = 1073938440
Attribute oY.VB_VarUserMemId = 1073938440
Dim Radius        As Double
Attribute Radius.VB_VarUserMemId = 1073938442


Dim R#, G#, B#, A#
Attribute R.VB_VarUserMemId = 1073938443
Attribute G.VB_VarUserMemId = 1073938443
Attribute B.VB_VarUserMemId = 1073938443
Attribute A.VB_VarUserMemId = 1073938443
Dim Poly()        As Single
Attribute Poly.VB_VarUserMemId = 1073938447


Private Sub cmdTEST_Click(Index As Integer)
    Select Case Index
    Case 0
        '    ----------- BRUSH1 --------------
        Brush.Clear
        A = 1
        R = Int(Rnd * 100) * 0.01: G = Int(Rnd * 100) * 0.01: B = Int(Rnd * 100) * 0.01
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 20, R, G, B, A
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            Y = (0.2 + 0.6 * Rnd) * H
            Brush.AddPoint X * 1, Y * 1, 20, R, G, B, A
        Next
        Brush.DRAW CC
        CC.TextOut 5, 5, "cCairoBRUSH (Single Color)" & "  R:" & R & "  G:" & G & "  B:" & B & "  A:" & A
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0

    Case 1

        '    ----------- BRUSH2 --------------
        Brush.Clear
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 20, Rnd, Rnd, Rnd, A
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            Y = (0.2 + 0.6 * Rnd) * H
            Brush.AddPoint X * 1, Y * 1, 20, Rnd, Rnd, Rnd, A
        Next
        Brush.DRAW CC
        CC.TextOut 5, 5, "cCairoBRUSH (Multi Color)"
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0

    Case 2

        '    ----------- BRUSH3 --------------
        Brush.Clear
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 20, Rnd, Rnd, Rnd, Rnd
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            Y = (0.2 + 0.6 * Rnd) * H
            Brush.AddPoint X * 1, Y * 1, 20, Rnd, Rnd, Rnd, Rnd
        Next
        Brush.DRAW CC
        CC.TextOut 5, 5, "cCairoBRUSH (Multi Color and Alpha)"
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0



    Case 3

        '    ----------- BRUSH4 --------------
        Brush.Clear
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 1, Rnd, Rnd, Rnd, A

        I = 0
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            I = I + 1
            Y = (0.2 + 0.6 * Rnd) * H
            Radius = I / 8
            Radius = (1 - Radius) * Radius * 4 * 25
            Brush.AddPoint X * 1, Y * 1, Radius, Rnd, Rnd, Rnd, A
        Next

        Brush.DRAW CC
        CC.TextOut 5, 5, "cCairoBRUSH (Multi Color and Width)"
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0



    Case 4


        '    ----------- BRUSH5 --------------
        Brush.Clear
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 1, Rnd, Rnd, Rnd, Rnd

        I = 0
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            I = I + 1
            Y = (0.2 + 0.6 * Rnd) * H
            Radius = I / 8
            Radius = (1 - Radius) * Radius * 4 * 25
            Brush.AddPoint X * 1, Y * 1, Radius, Rnd, Rnd, Rnd, Rnd
        Next

        Brush.DRAW CC
        CC.TextOut 5, 5, "cCairoBRUSH (Multi Color, Alpha and Width)"
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0



    Case 5



        '    ----------- BRUSH6 --------------
        Brush.Clear
        CC.SetSourceColor vbWhite: CC.Paint
        CC.SetSourceRGBA R, G * 2, B, A
        CC.SetLineWidth 20
        CC.SetLineCap CAIRO_LINE_CAP_ROUND
        CC.MoveTo W * 0.1, H * 0.5
        CC.GetCurrentPoint X, Y
        Brush.AddPoint X * 1, Y * 1, 1, Rnd, Rnd, Rnd, Rnd

        I = 0
        '-------------------------------
        For X = W * 0.2 To W * 0.9 Step W * 0.1
            I = I + 1
            Y = (0.2 + 0.6 * Rnd) * H
            Radius = I / 8
            Radius = (1 - Radius) * Radius * 4 * 25
            Brush.AddPoint X * 1, Y * 1, Radius, Rnd, Rnd, Rnd, 0.1 + 0.9 * Rnd
        Next
        Brush.DRAW CC, 1, 0.5     '<---- Border Alpha (is multiplied by Base Alpha)
        CC.TextOut 5, 5, "cCairoBRUSH (Multi Color, Alpha and Width + Border Alpha)"
        PIC.Picture = SRF.Picture
        SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
        Ima = Ima + 1
        If Ima > 8 Then Ima = 0

    End Select

End Sub

Private Sub Form_Load()
    W = 520
    H = 284

    PIC.Width = W
    PIC.Height = H

    Set SRF = Cairo.CreateSurface(W, H)
    Set CC = SRF.CreateContext

    Set Brush = New cCairoBrush


    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.SetLineJoin CAIRO_LINE_JOIN_ROUND

    For I = 0 To cmdTEST.Count - 1

        cmdTEST.Item(I).Caption = "TEST " & I + 1
    Next

End Sub


Private Sub Command1_Click()
    Ima = 0

    '-----------------------------------------------------------------------
    R = 0.8
    G = 0.1
    B = 0
    A = 0.5


    '----------- LINE By LINE --------------
    CC.SetSourceColor vbWhite: CC.Paint
    CC.SetSourceRGBA R, G, B, A
    CC.SetLineWidth 20
    oX = W * 0.1: oY = H * 0.5
    '-------------------------------
    For X = W * 0.2 To W * 0.9 Step W * 0.1
        Y = (0.2 + 0.6 * Rnd) * H
        CC.MoveTo oX, oY
        CC.LineTo X, Y
        oX = X
        oY = Y
        CC.Stroke
    Next
    CC.TextOut 5, 5, "LINE by LINE (multiple strokes)" & "  R:" & R & "  G:" & G & "  B:" & B & "  A:" & A
    PIC.Picture = SRF.Picture
    SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
    Ima = Ima + 1


    '----------- LINETO --------------
    CC.SetSourceColor vbWhite: CC.Paint
    CC.SetSourceRGBA R, G, B, A
    CC.SetLineWidth 20
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.MoveTo W * 0.1, H * 0.5
    '-------------------------------
    For X = W * 0.2 To W * 0.9 Step W * 0.1
        Y = (0.2 + 0.6 * Rnd) * H
        CC.LineTo X, Y
    Next
    CC.Stroke
    CC.TextOut 5, 5, "LINETO (Single stroke)" & "  R:" & R & "  G:" & G & "  B:" & B & "  A:" & A
    PIC.Picture = SRF.Picture
    SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
    Ima = Ima + 1

    '    ----------- Polygon --------------
    ReDim Poly(17)
    CC.SetSourceColor vbWhite: CC.Paint
    CC.SetSourceRGBA R, G * 2, B, A
    CC.SetLineWidth 20
    CC.SetLineCap CAIRO_LINE_CAP_ROUND
    CC.MoveTo W * 0.1, H * 0.5
    CC.GetCurrentPoint X, Y
    Poly(0) = X
    Poly(1) = Y
    I = 0
    '-------------------------------
    For X = W * 0.2 To W * 0.9 Step W * 0.1
        I = I + 2
        Y = (0.2 + 0.6 * Rnd) * H
        Poly(I) = X
        Poly(I + 1) = Y
    Next
    CC.PolygonSingle Poly(), , splNormal, True, True
    CC.Stroke
    CC.TextOut 5, 5, "PolygonSingle (Single stroke)" & "  R:" & R & "  G:" & G & "  B:" & B & "  A:" & A
    PIC.Picture = SRF.Picture
    SRF.WriteContentToJpgFile App.Path & "\Images\" & Format(Ima, "000") & ".jpg"
    Ima = Ima + 1


    '******************************************************************
    'cCAIROBRUSH
    '******************************************************************



    cmdTEST_Click (0)
    cmdTEST_Click (1)
    cmdTEST_Click (2)
    cmdTEST_Click (3)
    cmdTEST_Click (4)
    cmdTEST_Click (5)






End Sub


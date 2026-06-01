Attribute VB_Name = "CONTOUR_PIECE"
Option Explicit

' ===================================================
' CONTOUR-PIECE.bas  |  Macro : CP
' Contour de piece avec mi-epaisseur cloisons auto
' Calque cloisons par defaut : "Cloisons"
'
' INSTALLATION GeoGex :
'   1. Taper VBAIDE dans GeoGex
'   2. Fichier > Nouveau projet
'   3. Insertion > Module
'   4. Coller tout ce code dans le module
'   5. Fichier > Enregistrer sous -> CONTOUR-PIECE.dvb
'   6. Fermer VBAIDE
'   7. Pour charger : VBALOAD -> choisir CONTOUR-PIECE.dvb
'   8. Pour lancer : VBARUN -> CP
' ===================================================

'--- Constantes de detection --------------------------
Private Const TOL_COIN As Double = 0.01   ' dist <= = coincident (sur la ligne)
Private Const TOL_PAR  As Double = 0.985  ' cos angle min parallelisme (~10 deg)
Private Const MAX_EP   As Double = 2#     ' epaisseur max cherchee (unites dessin)

'--- Stockage global des segments cloisons ------------
Private gN   As Long
Private gAx() As Double, gAy() As Double
Private gBx() As Double, gBy() As Double

'======================================================
'  GEOMETRIE 2D
'======================================================

Private Function Cross2(ax As Double, ay As Double, _
                        bx As Double, by As Double) As Double
    Cross2 = ax * by - ay * bx
End Function

Private Function Dot2(ax As Double, ay As Double, _
                      bx As Double, by As Double) As Double
    Dot2 = ax * bx + ay * by
End Function

Private Sub NormVec(ByRef vx As Double, ByRef vy As Double)
    Dim k As Double
    k = Sqr(vx * vx + vy * vy)
    If k > 0.000001 Then vx = vx / k: vy = vy / k
End Sub

Private Function IsPar(d1x As Double, d1y As Double, _
                       d2x As Double, d2y As Double) As Boolean
    Dim nx1 As Double, ny1 As Double
    Dim nx2 As Double, ny2 As Double
    nx1 = d1x: ny1 = d1y: Call NormVec(nx1, ny1)
    nx2 = d2x: ny2 = d2y: Call NormVec(nx2, ny2)
    IsPar = (Abs(Dot2(nx1, ny1, nx2, ny2)) > TOL_PAR)
End Function

' Distance perpendiculaire point P -> droite infinie (A,B)
Private Function DPL(px As Double, py As Double, _
                     ax As Double, ay As Double, _
                     bx As Double, by As Double) As Double
    Dim dx As Double, dy As Double, L As Double
    dx = bx - ax: dy = by - ay
    L = Sqr(dx * dx + dy * dy)
    If L < 0.000001 Then DPL = 999999: Exit Function
    DPL = Abs((dx * (py - ay) - dy * (px - ax)) / L)
End Function

' Parametre de projection de P sur segment [A,B]
Private Function ProjT(px As Double, py As Double, _
                       ax As Double, ay As Double, _
                       bx As Double, by As Double) As Double
    Dim dx As Double, dy As Double, L2 As Double
    dx = bx - ax: dy = by - ay
    L2 = dx * dx + dy * dy
    If L2 < 0.000001 Then ProjT = 0.5: Exit Function
    ProjT = ((px - ax) * dx + (py - ay) * dy) / L2
End Function

' Intersection deux droites P1+t*D1 et P2+s*D2
Private Sub Isect2(p1x As Double, p1y As Double, _
                   d1x As Double, d1y As Double, _
                   p2x As Double, p2y As Double, _
                   d2x As Double, d2y As Double, _
                   ByRef rx As Double, ByRef ry As Double)
    Dim dpx As Double, dpy As Double
    Dim den As Double, t As Double
    dpx = p2x - p1x: dpy = p2y - p1y
    den = Cross2(d1x, d1y, d2x, d2y)
    If Abs(den) < 0.000001 Then
        rx = (p1x + p2x) / 2: ry = (p1y + p2y) / 2
    Else
        t = Cross2(dpx, dpy, d2x, d2y) / den
        rx = p1x + t * d1x: ry = p1y + t * d1y
    End If
End Sub

' Aire signee du polygone (> 0 = CCW)
Private Function SignedArea(xs() As Double, ys() As Double, _
                             n As Long) As Double
    Dim s As Double, i As Long, j As Long
    For i = 0 To n - 1
        j = (i + 1) Mod n
        s = s + xs(i) * ys(j) - xs(j) * ys(i)
    Next i
    SignedArea = s / 2
End Function

'======================================================
'  ENTITES DESSIN
'======================================================

Private Sub EnsureLayer(LayName As String)
    On Error Resume Next
    Dim oL As AcadLayer
    Set oL = ThisDrawing.Layers(LayName)
    If oL Is Nothing Then
        Set oL = ThisDrawing.Layers.Add(LayName)
    End If
    On Error GoTo 0
End Sub

Private Sub AddSeg(ax As Double, ay As Double, _
                   bx As Double, by As Double)
    ReDim Preserve gAx(gN), gAy(gN), gBx(gN), gBy(gN)
    gAx(gN) = ax: gAy(gN) = ay
    gBx(gN) = bx: gBy(gN) = by
    gN = gN + 1
End Sub

' Charger tous les segments LINE/LWPOLYLINE du calque cloisons
Private Function LoadCloisons(LayName As String) As Boolean
    gN = 0

    Dim ss As AcadSelectionSet
    On Error Resume Next
    ThisDrawing.SelectionSets("_CPSS").Delete
    Set ss = ThisDrawing.SelectionSets.Add("_CPSS")
    On Error GoTo 0
    If ss Is Nothing Then
        MsgBox "Impossible de creer un SelectionSet.", vbCritical, "CP"
        LoadCloisons = False: Exit Function
    End If

    Dim FT(1) As Integer, FV(1) As Variant
    FT(0) = 8:  FV(0) = LayName
    FT(1) = 0:  FV(1) = "LINE,LWPOLYLINE"
    ss.SelectAll FT, FV

    Dim ent As AcadEntity
    For Each ent In ss
        Select Case TypeName(ent)
        Case "AcadLine"
            Dim lin As AcadLine: Set lin = ent
            AddSeg lin.StartPoint(0), lin.StartPoint(1), _
                   lin.EndPoint(0), lin.EndPoint(1)
        Case "AcadLWPolyline"
            Dim pl As AcadLWPolyline: Set pl = ent
            Dim C() As Double: C = pl.Coordinates
            Dim nP As Long: nP = (UBound(C) + 1) \ 2
            Dim k As Long
            For k = 0 To nP - 2
                AddSeg C(k * 2), C(k * 2 + 1), _
                       C((k + 1) * 2), C((k + 1) * 2 + 1)
            Next k
            If pl.Closed Then
                AddSeg C((nP - 1) * 2), C((nP - 1) * 2 + 1), C(0), C(1)
            End If
        End Select
    Next ent
    ss.Delete

    If gN = 0 Then
        MsgBox "Aucun segment sur le calque [" & LayName & "]." & Chr(10) & _
               "Verifiez le nom du calque.", vbExclamation, "CP"
        LoadCloisons = False
    Else
        MsgBox gN & " segment(s) charge(s) sur [" & LayName & "].", _
               vbInformation, "CP"
        LoadCloisons = True
    End If
End Function

' Calcul de l'offset pour un segment contour [v1,v2]
Private Function CalcOffset(v1x As Double, v1y As Double, _
                             v2x As Double, v2y As Double) As Double
    Dim mx As Double, my As Double
    Dim sdx As Double, sdy As Double
    Dim coin As Boolean, minD As Double
    Dim i As Long, dd As Double, pt As Double

    mx = (v1x + v2x) / 2: my = (v1y + v2y) / 2
    sdx = v2x - v1x: sdy = v2y - v1y
    coin = False: minD = 999999

    For i = 0 To gN - 1
        If IsPar(sdx, sdy, gBx(i) - gAx(i), gBy(i) - gAy(i)) Then
            dd = DPL(mx, my, gAx(i), gAy(i), gBx(i), gBy(i))
            If dd < MAX_EP Then
                pt = ProjT(mx, my, gAx(i), gAy(i), gBx(i), gBy(i))
                If pt > -0.3 And pt < 1.3 Then
                    If dd < TOL_COIN Then
                        coin = True
                    ElseIf dd < minD Then
                        minD = dd
                    End If
                End If
            End If
        End If
    Next i

    If coin And minD < 999999 Then
        CalcOffset = minD / 2
    Else
        CalcOffset = 0
    End If
End Function

' Formate un Double avec point decimal (independant de la locale)
Private Function FmtD(v As Double) As String
    FmtD = Replace(CStr(v), ",", ".")
End Function

'======================================================
'  COMMANDE PRINCIPALE CP
'======================================================

Public Sub CP()
    '--- Parametres ---
    Dim layC As String
    layC = InputBox("Calque des cloisons :", "Contour Piece", "Cloisons")
    If layC = "" Then Exit Sub

    Dim layO As String
    layO = InputBox("Calque de sortie :", "Contour Piece", "CONTOURS")
    If layO = "" Then Exit Sub

    '--- Charger les segments cloisons (1 fois) ---
    If Not LoadCloisons(layC) Then Exit Sub
    Call EnsureLayer(layO)

    Dim nDone As Long: nDone = 0

    '--- Boucle de clic ---
    Dim continuer As Boolean: continuer = True
    Do While continuer

        Dim ptVar As Variant
        On Error Resume Next
        Err.Clear
        ptVar = ThisDrawing.Utility.GetPoint(, _
                "Cliquer dans une piece (Echap pour terminer) : ")
        If Err.Number <> 0 Then
            continuer = False
            On Error GoTo 0
        Else
            On Error GoTo 0

            '--- Lancer BOUNDARY et recuperer la polyligne ---
            Dim nBefore As Long
            nBefore = ThisDrawing.ModelSpace.Count

            Dim ptStr As String
            ptStr = FmtD(ptVar(0)) & "," & FmtD(ptVar(1))
            ThisDrawing.SendCommand "_-BOUNDARY " & ptStr & " " & Chr(13)

            If ThisDrawing.ModelSpace.Count <= nBefore Then
                MsgBox "Zone non fermee — recliquez dans la piece.", _
                       vbExclamation, "CP"
                GoTo Suite
            End If

            Dim lastEnt As AcadEntity
            Set lastEnt = ThisDrawing.ModelSpace( _
                          ThisDrawing.ModelSpace.Count - 1)

            If TypeName(lastEnt) <> "AcadLWPolyline" Then
                MsgBox "BOUNDARY a cree une region (pas une polyligne)." & Chr(10) & _
                       "Recommencez.", vbExclamation, "CP"
                lastEnt.Delete
                GoTo Suite
            End If

            Dim bPoly As AcadLWPolyline: Set bPoly = lastEnt
            Dim Co() As Double: Co = bPoly.Coordinates
            Dim n As Long: n = (UBound(Co) + 1) \ 2
            bPoly.Delete

            If n < 3 Then
                MsgBox "Contour invalide.", vbExclamation, "CP"
                GoTo Suite
            End If

            '--- Extraire vertices XY ---
            Dim xs() As Double, ys() As Double
            ReDim xs(n - 1), ys(n - 1)
            Dim i As Long
            For i = 0 To n - 1
                xs(i) = Co(i * 2): ys(i) = Co(i * 2 + 1)
            Next i

            '--- Orienter CCW ---
            If SignedArea(xs, ys, n) < 0 Then
                Dim tmpX() As Double, tmpY() As Double
                ReDim tmpX(n - 1), tmpY(n - 1)
                For i = 0 To n - 1
                    tmpX(i) = xs(n - 1 - i)
                    tmpY(i) = ys(n - 1 - i)
                Next i
                xs = tmpX: ys = tmpY
            End If

            '--- Offset par segment ---
            Dim offs() As Double
            ReDim offs(n - 1)
            Dim i2 As Long
            For i = 0 To n - 1
                i2 = (i + 1) Mod n
                offs(i) = CalcOffset(xs(i), ys(i), xs(i2), ys(i2))
            Next i

            '--- Droites decalees vers l'exterieur (dans la cloison) ---
            Dim oPx() As Double, oPy() As Double
            Dim oDx() As Double, oDy() As Double
            ReDim oPx(n - 1), oPy(n - 1)
            ReDim oDx(n - 1), oDy(n - 1)
            For i = 0 To n - 1
                i2 = (i + 1) Mod n
                Dim dx As Double, dy As Double, kk As Double
                dx = xs(i2) - xs(i): dy = ys(i2) - ys(i)
                kk = Sqr(dx * dx + dy * dy)
                If kk > 0.000001 Then dx = dx / kk: dy = dy / kk
                ' Normale droite = vers exterieur pour polygone CCW
                Dim nx As Double, ny As Double
                nx = dy: ny = -dx
                oPx(i) = xs(i) + offs(i) * nx
                oPy(i) = ys(i) + offs(i) * ny
                oDx(i) = xs(i2) - xs(i)
                oDy(i) = ys(i2) - ys(i)
            Next i

            '--- Nouveaux sommets = intersection droites adjacentes ---
            Dim nxs() As Double, nys() As Double
            ReDim nxs(n - 1), nys(n - 1)
            Dim ip As Long, rx As Double, ry As Double
            For i = 0 To n - 1
                ip = (i - 1 + n) Mod n
                Call Isect2(oPx(ip), oPy(ip), oDx(ip), oDy(ip), _
                            oPx(i), oPy(i), oDx(i), oDy(i), rx, ry)
                nxs(i) = rx: nys(i) = ry
            Next i

            '--- Creer la polyligne finale ---
            Dim pts() As Double
            ReDim pts(n * 2 - 1)
            For i = 0 To n - 1
                pts(i * 2) = nxs(i): pts(i * 2 + 1) = nys(i)
            Next i
            Dim newPoly As AcadLWPolyline
            Set newPoly = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
            newPoly.Closed = True
            newPoly.Layer = layO
            ThisDrawing.Regen acAllViewports
            nDone = nDone + 1

        End If
Suite:
        Err.Clear
    Loop

    If nDone > 0 Then
        MsgBox nDone & " contour(s) trace(s) sur [" & layO & "].", _
               vbInformation, "CP"
    End If
End Sub

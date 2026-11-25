Attribute VB_Name = "Belt_Pulley_Builder1"
'==============================================================
' SolidWorks 2022 macro: timing belt pulley
'   HTD 3M, HTD 8M (ISO 13050 profiles H3M, H8M), 2GT, 3GT
' FILE 1 of 2 - MAIN CODE. Paste into a standard module (Module1).
' Inputs come from userform "frmHTD3M" (see FILE 2).
'
' GROOVE MODEL (all profiles)
'   hg  - groove depth (outside diameter to groove bottom)
'   X   - half-length of the flat at the groove bottom (0 = none)
'   R1  - groove bottom radius
'   phi - half groove angle (flank angle to groove axis)
'   rt  - radius at the groove crest
'   Joints are tangent: [flat] - R1 arc - straight flank - rt arc - OD land.
'   Outside diameter: do = d - 2a + N',  d = z*P/pi.
'
' DATA SOURCES
'   HTD 3M, HTD 8M: ISO 13050:2022, Table 14 (groove profile dimensions by
'     number of grooves), Table 9 (a = 0.381 / 0.686 mm). For H8M the
'     outside-diameter correction N' (Table 16 was not available) is taken
'     from the standard pulley sizes of Table 15 and interpolated.
'   2GT, 3GT (Gates GT2 family): the groove profile is proprietary and not
'     published by Gates or in a standard. Used here: a = 0.254 / 0.381 mm
'     (published pitch line differentials), groove depth 0.764 / 1.14 mm and
'     groove width at the OD 1.494 / 2.27 mm (widely used CAD values), with a
'     circular bottom, short flanks and a small crest radius chosen to
'     reproduce these two dimensions. THIS IS AN APPROXIMATION - check it
'     against your belt / vendor drawing and edit GrooveParams if needed.
'==============================================================
Option Explicit

' Belt profiles
Public Const PROF_HTD3M As Long = 1
Public Const PROF_HTD8M As Long = 2
Public Const PROF_2GT As Long = 3
Public Const PROF_3GT As Long = 4

Private Const K_EM As Long = 0, K_GM As Long = 1, K_TM As Long = 2, K_BM As Long = 3
Private Const K_BP As Long = 4, K_TP As Long = 5, K_GP As Long = 6, K_EP As Long = 7
Private Const K_CM As Long = 8, K_CP As Long = 9, K_FM As Long = 10, K_FP As Long = 11

Public swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks
    frmHTD3M.Show
End Sub

Public Function Pi() As Double
    Pi = 4# * Atn(1#)
End Function

' Pitch, pitch line differential, tooth count limits and name of a profile
Public Sub ProfileInfo(ByVal prof As Long, ByRef pitch As Double, ByRef pld As Double, _
                       ByRef minZ As Long, ByRef maxZ As Long, ByRef pname As String)
    Select Case prof
        Case PROF_HTD8M
            pitch = 8#: pld = 0.686: minZ = 22: maxZ = 200: pname = "HTD 8M (ISO 13050 H8M)"
        Case PROF_2GT
            pitch = 2#: pld = 0.254: minZ = 12: maxZ = 200: pname = "2GT (GT2 2 mm, approximate groove)"
        Case PROF_3GT
            pitch = 3#: pld = 0.381: minZ = 12: maxZ = 200: pname = "3GT (GT2 3 mm, approximate groove)"
        Case Else
            pitch = 3#: pld = 0.381: minZ = 12: maxZ = 200: pname = "HTD 3M (ISO 13050 H3M)"
    End Select
End Sub

' Groove parameters by profile and number of grooves
Private Sub GrooveParams(ByVal prof As Long, ByVal n As Long, ByRef hg As Double, ByRef x As Double, _
                         ByRef R1 As Double, ByRef phiDeg As Double, ByRef rt As Double)
    Select Case prof
        Case PROF_HTD8M
            ' ISO 13050:2022 Table 14, H8M
            Select Case n
                Case Is <= 27           ' 22 to 27 grooves
                    hg = 3.295: x = 0#: R1 = 2.675: phiDeg = 11.3: rt = 0.874
                Case Is <= 89           ' 28 to 89 grooves
                    hg = 3.604: x = 0#: R1 = 2.629: phiDeg = 7#:   rt = 1.024
                Case Else               ' 90 to 200 grooves
                    hg = 3.63:  x = 0#: R1 = 2.639: phiDeg = 6.6:  rt = 1.008
            End Select
        Case PROF_2GT
            ' approximation, see header
            hg = 0.764: x = 0#: R1 = 0.62: phiDeg = 12#: rt = 0.1
        Case PROF_3GT
            ' approximation, see header
            hg = 1.14: x = 0#: R1 = 0.95: phiDeg = 12#: rt = 0.15
        Case Else
            ' ISO 13050:2022 Table 14, H3M
            Select Case n
                Case Is <= 13           ' 10 to 13 grooves
                    hg = 1.19:  x = 0.029: R1 = 0.991: phiDeg = 15: rt = 0.181
                Case Is <= 25           ' 14 to 25 grooves
                    hg = 1.179: x = 0.112: R1 = 0.889: phiDeg = 9:  rt = 0.229
                Case Is <= 80           ' 26 to 80 grooves
                    hg = 1.219: x = 0.028: R1 = 0.927: phiDeg = 8:  rt = 0.191
                Case Else               ' 81 to 200 grooves
                    hg = 1.234: x = 0.074: R1 = 0.925: phiDeg = 4:  rt = 0.301
            End Select
    End Select
End Sub

' Outside diameter correction N' (mm). H8M: from ISO 13050 Table 15 standard pulleys
' (28 grooves +0.15, 30 +0.11, 32 +0.04, 34 and more 0), linear in between. Others: 0.
Private Function NPrime(ByVal prof As Long, ByVal n As Long) As Double
    If prof <> PROF_HTD8M Then Exit Function
    Select Case n
        Case Is <= 26
            NPrime = 0#
        Case Is <= 28
            NPrime = 0.15 * (n - 26) / 2#
        Case Is <= 30
            NPrime = 0.15 + (0.112 - 0.15) * (n - 28) / 2#
        Case Is <= 32
            NPrime = 0.112 + (0.042 - 0.112) * (n - 30) / 2#
        Case Is <= 34
            NPrime = 0.042 * (34 - n) / 2#
        Case Else
            NPrime = 0#
    End Select
End Function

' Pitch diameter and outside radius (mm)
Private Sub PulleySize(ByVal prof As Long, ByVal n As Long, ByRef PD As Double, ByRef Ro As Double)
    Dim pitch As Double, pld As Double, mnz As Long, mxz As Long, nm As String
    ProfileInfo prof, pitch, pld, mnz, mxz, nm
    PD = n * pitch / Pi()
    Ro = PD / 2# - pld + NPrime(prof, n) / 2#
End Sub

' Short text for the userform label
Public Function PulleyInfoText(ByVal prof As Long, ByVal n As Long) As String
    Dim pitch As Double, pld As Double, mnz As Long, mxz As Long, nm As String
    Dim hg As Double, x As Double, R1 As Double, ph As Double, rt As Double
    Dim PD As Double, Ro As Double
    ProfileInfo prof, pitch, pld, mnz, mxz, nm
    PulleyInfoText = nm & vbCrLf & "Pitch " & Format(pitch, "0.0") & " mm, a = " & Format(pld, "0.000") & _
                     " mm, teeth " & mnz & " ... " & mxz
    If n >= mnz And n <= mxz Then
        GrooveParams prof, n, hg, x, R1, ph, rt
        PulleySize prof, n, PD, Ro
        PulleyInfoText = PulleyInfoText & vbCrLf & "Pitch dia " & Format(PD, "0.000") & ", outside dia " & _
                         Format(2# * Ro, "0.000") & ", root dia " & Format(2# * (Ro - hg), "0.000") & _
                         " mm; max bore " & Format(2# * (Ro - hg) - 1#, "0.00") & " mm"
    End If
End Function

' Returns "" if inputs are acceptable, otherwise an error text
Public Function ValidatePulley(ByVal prof As Long, ByVal n As Long, ByVal w As Double, ByVal bore As Double) As String
    Dim pitch As Double, pld As Double, mnz As Long, mxz As Long, nm As String
    Dim hg As Double, x As Double, R1 As Double, ph As Double, rt As Double
    Dim PD As Double, Ro As Double, Rroot As Double
    ProfileInfo prof, pitch, pld, mnz, mxz, nm
    If n < mnz Or n > mxz Then
        ValidatePulley = nm & ": number of teeth must be between " & mnz & " and " & mxz & ".": Exit Function
    End If
    If w <= 0 Then ValidatePulley = "Width must be greater than zero.": Exit Function
    GrooveParams prof, n, hg, x, R1, ph, rt
    PulleySize prof, n, PD, Ro
    Rroot = Ro - hg
    If bore <= 0 Or bore > 2# * Rroot - 1# Then
        ValidatePulley = "Bore diameter must be > 0 and <= " & Format(2# * Rroot - 1#, "0.00") & " mm for " & n & " teeth."
    End If
End Function

Private Sub Rot(ByVal x As Double, ByVal y As Double, ByVal a As Double, _
                ByRef rx As Double, ByRef ry As Double)
    rx = x * Cos(a) - y * Sin(a)
    ry = x * Sin(a) + y * Cos(a)
End Sub

' Builds the pulley. Returns True on success; msg holds info or error text.
Public Function BuildPulley(ByVal prof As Long, ByVal n As Long, ByVal w As Double, _
                            ByVal bore As Double, ByRef msg As String) As Boolean
    On Error GoTo EH
    Dim swModel As SldWorks.ModelDoc2
    Dim skMgr As SldWorks.SketchManager
    Dim sketchOpen As Boolean
    Dim errNum As Long, errDesc As String

    msg = ValidatePulley(prof, n, w, bore)
    If msg <> "" Then Exit Function
    If swApp Is Nothing Then Set swApp = Application.SldWorks

    Dim pitch As Double, pld As Double, mnz As Long, mxz As Long, pname As String
    ProfileInfo prof, pitch, pld, mnz, mxz, pname

    '---------------- Groove geometry in the local groove frame (mm) ----------------
    ' x = radial (outward), y = tangential, groove axis = +x
    Dim hg As Double, XF As Double, R1 As Double, phiDeg As Double, rt As Double
    GrooveParams prof, n, hg, XF, R1, phiDeg, rt

    Dim PD As Double, Ro As Double, theta As Double, phi As Double
    PulleySize prof, n, PD, Ro
    theta = 2# * Pi() / n
    phi = phiDeg * Pi() / 180#

    Dim hasFlat As Boolean
    hasFlat = (XF > 0.0005)

    Dim xb As Double, xc As Double
    xb = Ro - hg                           ' groove bottom
    xc = xb + R1                           ' x of the R1 arc centers

    ' tangent point flank / R1 arc (+y side); arc center is (xc, +XF)
    Dim Tx As Double, Ty As Double
    Tx = xc - R1 * Sin(phi): Ty = XF + R1 * Cos(phi)

    Dim nx As Double, ny As Double, dx As Double, dy As Double
    nx = -Sin(phi): ny = Cos(phi)          ' flank normal (toward material)
    dx = Cos(phi): dy = Sin(phi)           ' flank direction (outward)

    ' crest arc rt: tangent to flank line and to the OD circle
    Dim P0x As Double, P0y As Double, Rp As Double, b As Double, disc As Double, s As Double
    P0x = Tx + rt * nx: P0y = Ty + rt * ny
    Rp = Ro - rt
    b = P0x * dx + P0y * dy
    disc = b * b - (P0x * P0x + P0y * P0y) + Rp * Rp
    If disc < 0 Then msg = "Profile cannot be solved (crest radius).": Exit Function
    s = -b + Sqr(disc)
    If s <= 0 Then msg = "Profile cannot be solved (flank length <= 0).": Exit Function

    Dim Fx As Double, Fy As Double, Gx As Double, Gy As Double, Ex As Double, Ey As Double, fl As Double
    Fx = P0x + s * dx: Fy = P0y + s * dy
    Gx = Fx - rt * nx: Gy = Fy - rt * ny
    fl = Sqr(Fx * Fx + Fy * Fy)
    Ex = Fx * Ro / fl: Ey = Fy * Ro / fl

    Dim landMM As Double
    landMM = (theta / 2# - Atn(Fy / Fx)) * 2# * Ro
    If landMM < 0.05 Then
        msg = "Grooves overlap for " & n & " teeth - land on OD is too small.": Exit Function
    End If

    '---------------- Global points for every tooth (meters) ----------------
    Dim pt() As Double
    ReDim pt(0 To n - 1, 0 To 11, 0 To 1)
    Dim i As Long, a As Double, rx As Double, ry As Double
    For i = 0 To n - 1
        a = i * theta
        Rot Ex, -Ey, a, rx, ry:  pt(i, K_EM, 0) = rx / 1000#: pt(i, K_EM, 1) = ry / 1000#
        Rot Gx, -Gy, a, rx, ry:  pt(i, K_GM, 0) = rx / 1000#: pt(i, K_GM, 1) = ry / 1000#
        Rot Tx, -Ty, a, rx, ry:  pt(i, K_TM, 0) = rx / 1000#: pt(i, K_TM, 1) = ry / 1000#
        Rot xb, -XF, a, rx, ry:  pt(i, K_BM, 0) = rx / 1000#: pt(i, K_BM, 1) = ry / 1000#
        Rot xb, XF, a, rx, ry:   pt(i, K_BP, 0) = rx / 1000#: pt(i, K_BP, 1) = ry / 1000#
        Rot Tx, Ty, a, rx, ry:   pt(i, K_TP, 0) = rx / 1000#: pt(i, K_TP, 1) = ry / 1000#
        Rot Gx, Gy, a, rx, ry:   pt(i, K_GP, 0) = rx / 1000#: pt(i, K_GP, 1) = ry / 1000#
        Rot Ex, Ey, a, rx, ry:   pt(i, K_EP, 0) = rx / 1000#: pt(i, K_EP, 1) = ry / 1000#
        Rot xc, -XF, a, rx, ry:  pt(i, K_CM, 0) = rx / 1000#: pt(i, K_CM, 1) = ry / 1000#
        Rot xc, XF, a, rx, ry:   pt(i, K_CP, 0) = rx / 1000#: pt(i, K_CP, 1) = ry / 1000#
        Rot Fx, -Fy, a, rx, ry:  pt(i, K_FM, 0) = rx / 1000#: pt(i, K_FM, 1) = ry / 1000#
        Rot Fx, Fy, a, rx, ry:   pt(i, K_FP, 0) = rx / 1000#: pt(i, K_FP, 1) = ry / 1000#
    Next i

    '---------------- New part ----------------
    Dim tmpl As String
    tmpl = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
    If Len(tmpl) = 0 Then msg = "No default part template set (Tools > Options > Default Templates).": Exit Function
    Set swModel = swApp.NewDocument(tmpl, 0, 0, 0)
    If swModel Is Nothing Then msg = "Could not create a new part.": Exit Function

    ' first reference plane in the tree = Front plane (works in any UI language)
    Dim f As SldWorks.Feature, plane As SldWorks.Feature
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then Set plane = f: Exit Do
        Set f = f.GetNextFeature
    Loop
    If plane Is Nothing Then msg = "Front plane not found.": Exit Function

    '---------------- Sketch ----------------
    Set skMgr = swModel.SketchManager
    swModel.ClearSelection2 True
    plane.Select2 False, 0
    skMgr.InsertSketch True
    sketchOpen = True
    skMgr.AddToDB = True

    Dim j As Long
    For i = 0 To n - 1
        j = (i + 1) Mod n
        ' crest arc "-" side: OD point -> flank end (CCW about the arc center)
        skMgr.CreateArc pt(i, K_FM, 0), pt(i, K_FM, 1), 0, _
                        pt(i, K_EM, 0), pt(i, K_EM, 1), 0, _
                        pt(i, K_GM, 0), pt(i, K_GM, 1), 0, 1
        ' flank "-" side
        skMgr.CreateLine pt(i, K_GM, 0), pt(i, K_GM, 1), 0, pt(i, K_TM, 0), pt(i, K_TM, 1), 0
        If hasFlat Then
            ' R1 arc "-" side, flat, R1 arc "+" side
            skMgr.CreateArc pt(i, K_CM, 0), pt(i, K_CM, 1), 0, _
                            pt(i, K_BM, 0), pt(i, K_BM, 1), 0, _
                            pt(i, K_TM, 0), pt(i, K_TM, 1), 0, 1
            skMgr.CreateLine pt(i, K_BM, 0), pt(i, K_BM, 1), 0, pt(i, K_BP, 0), pt(i, K_BP, 1), 0
            skMgr.CreateArc pt(i, K_CP, 0), pt(i, K_CP, 1), 0, _
                            pt(i, K_TP, 0), pt(i, K_TP, 1), 0, _
                            pt(i, K_BP, 0), pt(i, K_BP, 1), 0, 1
        Else
            ' one groove bottom arc, CCW from the "+" tangent point to the "-" one through the bottom
            skMgr.CreateArc pt(i, K_CP, 0), pt(i, K_CP, 1), 0, _
                            pt(i, K_TP, 0), pt(i, K_TP, 1), 0, _
                            pt(i, K_TM, 0), pt(i, K_TM, 1), 0, 1
        End If
        ' flank "+" side
        skMgr.CreateLine pt(i, K_TP, 0), pt(i, K_TP, 1), 0, pt(i, K_GP, 0), pt(i, K_GP, 1), 0
        ' crest arc "+" side: flank end -> OD point
        skMgr.CreateArc pt(i, K_FP, 0), pt(i, K_FP, 1), 0, _
                        pt(i, K_GP, 0), pt(i, K_GP, 1), 0, _
                        pt(i, K_EP, 0), pt(i, K_EP, 1), 0, 1
        ' land on the outside diameter to the next groove
        skMgr.CreateArc 0, 0, 0, _
                        pt(i, K_EP, 0), pt(i, K_EP, 1), 0, _
                        pt(j, K_EM, 0), pt(j, K_EM, 1), 0, 1
    Next i

    ' bore
    skMgr.CreateCircleByRadius 0, 0, 0, bore / 2000#

    ' pitch circle as construction geometry (reference for the drawing)
    Dim seg As SldWorks.SketchSegment
    Set seg = skMgr.CreateCircleByRadius(0, 0, 0, PD / 2000#)
    If Not seg Is Nothing Then seg.ConstructionGeometry = True

    skMgr.AddToDB = False
    skMgr.InsertSketch True
    sketchOpen = False

    '---------------- Boss-extrude ----------------
    Dim skFeat As SldWorks.Feature, boss As SldWorks.Feature
    Set skFeat = swModel.FeatureByPositionReverse(0)
    skFeat.Name = "Pulley_Tooth_Profile"
    swModel.ClearSelection2 True
    skFeat.Select2 False, 0

    Set boss = swModel.FeatureManager.FeatureExtrusion2( _
        True, False, False, 0, 0, w / 1000#, 0, _
        False, False, False, False, 0, 0, _
        False, False, False, False, True, True, True, 0, 0, False)
    If boss Is Nothing Then msg = "Boss-extrude failed.": Exit Function
    boss.Name = "Pulley_Body"

    swModel.ShowNamedView2 "", 7           ' isometric
    swModel.ViewZoomtofit2

    msg = pname & " pulley created." & vbCrLf & _
          "Teeth: " & n & vbCrLf & _
          "Pitch diameter: " & Format(PD, "0.000") & " mm" & vbCrLf & _
          "Outside diameter: " & Format(2# * Ro, "0.000") & " mm" & vbCrLf & _
          "Groove depth: " & Format(hg, "0.000") & " mm, R1 = " & Format(R1, "0.000") & _
          " mm, crest R = " & Format(rt, "0.000") & " mm, groove angle = " & Format(2# * phiDeg, "0.0") & " deg" & vbCrLf & _
          "Width: " & w & " mm" & vbCrLf & _
          "Bore: " & bore & " mm"
    If prof = PROF_2GT Or prof = PROF_3GT Then
        msg = msg & vbCrLf & "Note: GT groove is an approximation - verify against your belt/vendor drawing."
    End If
    BuildPulley = True
    Exit Function

EH:
    errNum = Err.Number: errDesc = Err.Description      ' read BEFORE On Error (which clears Err)
    On Error Resume Next
    If Not skMgr Is Nothing Then
        skMgr.AddToDB = False
        If sketchOpen Then skMgr.InsertSketch True
    End If
    msg = "Error " & errNum & ": " & errDesc
    BuildPulley = False
End Function


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGear 
   Caption         =   "Involute gear"
   ClientHeight    =   8340.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   OleObjectBlob   =   "frmGear.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' FILE 2 of 2 - USERFORM CODE. Paste into the code window of
' userform "frmGear" (right-click the form > View Code).
'
' Create the form (Insert > UserForm), set (Name) = frmGear,
' Caption = "Involute gear", then place the controls below and set
' their (Name) exactly as listed:
'
'   TextBox       txtModule      normal module, mm              (label "Module")
'   TextBox       txtTeeth       number of teeth
'   TextBox       txtWidth       gear thickness, mm
'   OptionButton  optExternal    Caption "External"   GroupName = grpMesh
'   OptionButton  optInternal    Caption "Internal"   GroupName = grpMesh
'   TextBox       txtBore        shaft bore, mm (external only)
'   TextBox       txtOuterD      outside diameter, mm (internal only)
'   TextBox       txtAngle       pressure angle, deg
'   CheckBox      chkHelical     Caption "Helical gear"
'   TextBox       txtHelixAngle  helix angle, deg (helical only)
'   OptionButton  optRight       Caption "Right hand"  GroupName = grpHand
'   OptionButton  optLeft        Caption "Left hand"   GroupName = grpHand
'   -- additional parameters --
'   TextBox       txtShift       profile shift coefficient x (external only)
'   TextBox       txtAddendum    addendum coefficient ha*
'   ComboBox      cmbProfile     ISO 53 rack profile A / B / C / D
'   TextBox       txtClearance   clearance coefficient c*   (set by cmbProfile)
'   TextBox       txtFilletR     root fillet radius coefficient (set by cmbProfile)
'   -- info and buttons --
'   Label         lblHint        (make it 4-5 lines high; shows diameters,
'                                minimum outer diameter / maximum bore, fillet
'                                radius, minimum profile shift, warnings)
'   CommandButton btnCreate      Caption "Create gear"
'   CommandButton btnCancel      Caption "Cancel"
'==============================================================
Option Explicit

Private mLoading As Boolean

Private Sub Frame2_Click()

End Sub

Private Sub UserForm_Initialize()
    mLoading = True
    txtModule.Text = "2"
    txtTeeth.Text = "30"
    txtWidth.Text = "20"
    optExternal.Value = True
    txtBore.Text = "10"
    txtOuterD.Text = ""
    txtAngle.Text = "20"
    chkHelical.Value = False
    txtHelixAngle.Text = "15"
    optRight.Value = True
    txtShift.Text = "0"
    txtAddendum.Text = "1"
    cmbProfile.Style = 2                       ' fmStyleDropDownList
    cmbProfile.List = Array("A", "B", "C", "D")
    cmbProfile.Value = "A"
    ApplyProfile
    mLoading = False
    UpdateUI
    UpdateHint
End Sub

' ISO 53:1998 basic rack profiles: clearance coefficient c*, fillet coefficient rhoF*
Private Sub ApplyProfile()
    Select Case cmbProfile.Value
        Case "A": txtClearance.Text = "0.25": txtFilletR.Text = "0.38"
        Case "B": txtClearance.Text = "0.25": txtFilletR.Text = "0.3"
        Case "C": txtClearance.Text = "0.25": txtFilletR.Text = "0.25"
        Case "D": txtClearance.Text = "0.4":  txtFilletR.Text = "0.39"
    End Select
End Sub

Private Function ToNum(ByVal s As String) As Double
    ToNum = Val(Replace(Trim$(s), ",", "."))     ' accepts "2,5" and "2.5"
End Function

Private Sub ReadParams(p As GearParams)
    p.ModuleMn = ToNum(txtModule.Text)
    p.Teeth = CLng(ToNum(txtTeeth.Text))
    p.FaceWidth = ToNum(txtWidth.Text)
    p.IsInternal = optInternal.Value
    p.BoreDia = ToNum(txtBore.Text)
    p.OuterDia = ToNum(txtOuterD.Text)
    p.PressAngleDeg = ToNum(txtAngle.Text)
    p.Helical = chkHelical.Value
    p.HelixAngleDeg = ToNum(txtHelixAngle.Text)
    p.RightHand = optRight.Value
    p.ProfileShift = ToNum(txtShift.Text)
    p.AddendumCoef = ToNum(txtAddendum.Text)
    p.ClearanceCoef = ToNum(txtClearance.Text)
    p.FilletCoef = ToNum(txtFilletR.Text)
End Sub

Private Sub UpdateUI()
    Dim isInt As Boolean
    isInt = optInternal.Value
    txtBore.Enabled = Not isInt
    txtOuterD.Enabled = isInt
    txtShift.Enabled = Not isInt            ' internal gears are built with x = 0
    txtHelixAngle.Enabled = chkHelical.Value
    optRight.Enabled = chkHelical.Value
    optLeft.Enabled = chkHelical.Value
End Sub

Private Sub UpdateHint()
    If mLoading Then Exit Sub
    Dim p As GearParams, g As GearGeom
    ReadParams p
    lblHint.Caption = GearHint(p)
    ' pre-fill the outer diameter with the recommended value if it is empty
    If p.IsInternal And Len(Trim$(txtOuterD.Text)) = 0 Then
        If CalcGear(p, g) = "" Then txtOuterD.Text = Format(RecOuterDia(g, p), "0.00")
    End If
End Sub

Private Sub cmbProfile_Change()
    If mLoading Then Exit Sub
    ApplyProfile
    UpdateHint
End Sub

Private Sub optExternal_Click()
    UpdateUI
    UpdateHint
End Sub

Private Sub optInternal_Click()
    UpdateUI
    UpdateHint
End Sub

Private Sub chkHelical_Click()
    UpdateUI
    UpdateHint
End Sub

Private Sub txtModule_Change()
    UpdateHint
End Sub

Private Sub txtTeeth_Change()
    UpdateHint
End Sub

Private Sub txtAngle_Change()
    UpdateHint
End Sub

Private Sub txtHelixAngle_Change()
    UpdateHint
End Sub

Private Sub txtShift_Change()
    UpdateHint
End Sub

Private Sub txtAddendum_Change()
    UpdateHint
End Sub

Private Sub txtClearance_Change()
    UpdateHint
End Sub

Private Sub txtFilletR_Change()
    UpdateHint
End Sub

Private Sub btnCreate_Click()
    Dim p As GearParams, msg As String
    ReadParams p

    msg = ValidateGear(p)
    If msg <> "" Then
        MsgBox msg, vbExclamation, "Involute gear"
        Exit Sub
    End If

    Me.Hide
    If BuildGear(p, msg) Then
        MsgBox msg, vbInformation, "Involute gear"
    Else
        MsgBox msg, vbCritical, "Involute gear"
    End If
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHTD3M 
   Caption         =   "Pulley Builder"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "frmHTD3M.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHTD3M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' FILE 2 of 2 - USERFORM CODE. Paste into the code window of
' userform "frmHTD3M" (right-click the form > View Code).
' Replace the old content of the form code with this file.
'
' Controls on the form (set the Name property exactly):
'   OptionButton  optHTD3M   Caption "HTD 3M"   GroupName = grpProfile
'   OptionButton  optHTD8M   Caption "HTD 8M"   GroupName = grpProfile
'   OptionButton  opt2GT     Caption "2GT"      GroupName = grpProfile
'   OptionButton  opt3GT     Caption "3GT"      GroupName = grpProfile
'   TextBox       txtTeeth   number of teeth
'   TextBox       txtWidth   pulley width, mm
'   TextBox       txtBore    bore diameter, mm
'   Label         lblInfo    (about 4 lines high: profile info, diameters, max bore)
'   CommandButton btnCreate  Caption "Create pulley"
'   CommandButton btnCancel  Caption "Cancel"
'==============================================================
Option Explicit

Private mLoading As Boolean

Private Sub UserForm_Initialize()
    mLoading = True
    optHTD3M.Value = True
    txtTeeth.Text = "30"
    txtWidth.Text = "10"
    txtBore.Text = "8"
    mLoading = False
    UpdateInfo
End Sub

Private Function ToNum(ByVal s As String) As Double
    ToNum = Val(Replace(Trim$(s), ",", "."))   ' accepts "8,5" and "8.5"
End Function

Private Function CurrentProfile() As Long
    If optHTD8M.Value Then
        CurrentProfile = PROF_HTD8M
    ElseIf opt2GT.Value Then
        CurrentProfile = PROF_2GT
    ElseIf opt3GT.Value Then
        CurrentProfile = PROF_3GT
    Else
        CurrentProfile = PROF_HTD3M
    End If
End Function

Private Sub UpdateInfo()
    If mLoading Then Exit Sub
    Dim n As Long
    n = CLng(ToNum(txtTeeth.Text))
    lblInfo.Caption = PulleyInfoText(CurrentProfile(), n)
End Sub

' When the profile changes, bring the number of teeth into the allowed range
Private Sub FixTeeth()
    Dim pitch As Double, pld As Double, mnz As Long, mxz As Long, nm As String
    Dim n As Long
    ProfileInfo CurrentProfile(), pitch, pld, mnz, mxz, nm
    n = CLng(ToNum(txtTeeth.Text))
    If n < mnz Or n > mxz Then
        mLoading = True
        txtTeeth.Text = CStr(IIf(n < mnz, mnz, mxz))
        mLoading = False
    End If
End Sub

Private Sub optHTD3M_Click()
    FixTeeth
    UpdateInfo
End Sub

Private Sub optHTD8M_Click()
    FixTeeth
    UpdateInfo
End Sub

Private Sub opt2GT_Click()
    FixTeeth
    UpdateInfo
End Sub

Private Sub opt3GT_Click()
    FixTeeth
    UpdateInfo
End Sub

Private Sub txtTeeth_Change()
    UpdateInfo
End Sub

Private Sub btnCreate_Click()
    Dim n As Double, w As Double, bore As Double
    Dim msg As String, prof As Long

    prof = CurrentProfile()
    n = ToNum(txtTeeth.Text)
    w = ToNum(txtWidth.Text)
    bore = ToNum(txtBore.Text)

    If n <> Int(n) Then
        MsgBox "Number of teeth must be an integer.", vbExclamation, "Timing pulley"
        txtTeeth.SetFocus
        Exit Sub
    End If

    msg = ValidatePulley(prof, CLng(n), w, bore)
    If msg <> "" Then
        MsgBox msg, vbExclamation, "Timing pulley"
        Exit Sub
    End If

    Me.Hide
    If BuildPulley(prof, CLng(n), w, bore, msg) Then
        MsgBox msg, vbInformation, "Timing pulley"
    Else
        MsgBox msg, vbCritical, "Timing pulley"
    End If
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

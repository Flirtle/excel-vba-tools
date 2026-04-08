'testing
Option Explicit
Public myRibbon                         As IRibbonUI
Public MyTag                            As String
Public MyBtnTag                         As String

Dim CP1Protected As Boolean
Dim CP2Protected As Boolean
Dim CP3Protected As Boolean

Sub RibbonOnload(ribbon As IRibbonUI)
    'Create a ribbon instance for use in this project
    Set myRibbon = ribbon
    myRibbon.ActivateTab "CP1_ITP_controls"
    CP1Protected = True
    CP2Protected = True
    CP3Protected = True
End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    If MyTag = "show" Then
        visible = True
    Else
        If control.Tag Like MyTag Then
            visible = True
        Else
            visible = False
        End If
    End If
End Sub
Sub tabActivate(ByVal control As IRibbonControl)
    myRibbon.ActivateTab (control.ID)
End Sub

'DropDown onAction
Sub myMacroDD(ByVal control As IRibbonControl, _
    selectedID As String, _
    selectedIndex As Integer)
    On Error Resume Next
    Dim strList                         As String
    Dim strMacro                        As String
    
    Select Case control.ID
        Case "DD1"
            strList = "CP1addDelITC2"
        Case "DD2"
            strList = "CP1setSched2"
        Case "DD3"
            strList = "CP1Import2"
        Case "DD4"
            strList = "CP2addDelITC2"
        Case "DD5"
            strList = "CP2setSched2"
        Case "DD6"
            strList = "CP2Import2"
        Case "DD7"
            strList = "CP3addDelITC2"
        Case "DD8"
            strList = "CP3setSched2"
        Case "DD9"
            strList = "CP3Import2"
    End Select
    strMacro = ThisWorkbook.Names(strList) _
               .RefersToRange.Rows(CLng(selectedIndex)).Value
    Application.Run (strMacro)
    'Restore control to original state
    myRibbon.InvalidateControl control.ID
End Sub

'Callback for DropDown getItemCount
Sub GetItemCount(ByVal control As IRibbonControl, ByRef count)
    Dim strList                         As String
    Select Case control.ID
 
        Case "DD1"
            strList = "CP1addDelITC"
        Case "DD2"
            strList = "CP1setSched"
        Case "DD3"
            strList = "CP1Import"
        Case "DD4"
            strList = "CP2addDelITC"
        Case "DD5"
            strList = "CP2setSched"
        Case "DD6"
            strList = "CP2Import"
        Case "DD7"
            strList = "CP3addDelITC"
        Case "DD8"
            strList = "CP3setSched"
        Case "DD9"
            strList = "CP3Import"
    End Select
    count = ThisWorkbook.Names(strList) _
            .RefersToRange.Rows.count
End Sub

'Callback for DropDown getItemLabel
Sub GetItemLabel(ByVal control As IRibbonControl, _
    Index As Integer, ByRef label)
    Dim rngML                           As Range
    Dim strList                         As String
    
    Select Case control.ID
        Case "DD1"
            strList = "CP1addDelITC"
        Case "DD2"
            strList = "CP1setSched"
        Case "DD3"
            strList = "CP1Import"
        Case "DD4"
            strList = "CP2addDelITC"
        Case "DD5"
            strList = "CP2setSched"
        Case "DD6"
            strList = "CP2Import"
        Case "DD7"
            strList = "CP3addDelITC"
        Case "DD8"
            strList = "CP3setSched"
        Case "DD9"
            strList = "CP3Import"
    End Select

    Set rngML = ThisWorkbook.Names(strList) _
        .RefersToRange
    label = rngML.Cells(Index + 1)

End Sub

'Callback for DropDown getSelectedItemIndex
Sub GetSelItemIndex(ByVal control As IRibbonControl, ByRef Index)
    'Ensure first item in dropdown is displayed.
    Select Case control.ID
        Case Is = "DD1"
            Index = 0
        Case Is = "DD2"
            Index = 0
        Case Is = "DD3"
            Index = 0
        Case Is = "DD4"
            Index = 0
        Case Is = "DD5"
            Index = 0
        Case Is = "DD6"
            Index = 0
        Case Is = "DD7"
            Index = 0
        Case Is = "DD8"
            Index = 0
        Case Is = "DD9"
            Index = 0
        Case Else
    End Select
End Sub

Sub RefreshRibbon(Tag As String, Optional TabID As String, Optional BTNID As String)
    If myRibbon Is Nothing Then
        'MsgBox ("Error, Save/Restart your workbook")
    Else
        myRibbon.Invalidate
        If TabID <> "" Then myRibbon.ActivateTab TabID

    End If

End Sub

'Callback for SaveRegister onAction
Sub SaveRegister(control As IRibbonControl)
    ThisWorkbook.Save
End Sub

'Callback for BTNSaveClose onAction
Sub SaveClose(control As IRibbonControl)
    Dim answer                          As Integer
    Dim wbCount                         As Integer
    Dim wb                              As Workbook

    answer = MsgBox("Are you sure you want to save and close?", vbQuestion + vbYesNo + vbDefaultButton2, "Save and close")
    If answer = vbYes Then
        ThisWorkbook.Save
        Application.DisplayAlerts = False

        wbCount = 0
        For Each wb In Workbooks
            If wb.Name <> ThisWorkbook.Name Then
                wbCount = wbCount + 1
            End If
        Next wb

        If wbCount = 1 Then
            Application.Quit
        Else

            ThisWorkbook.Close SaveChanges:=True
            Application.DisplayAlerts = True
        End If

    End If

End Sub

'Callback for BTNCreateCP1Docs onAction
Sub UpdateCP01(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "UpdateCP1"
    Application.Run (strMacro)
End Sub
'Callback for BTNCreateCP2Docs onAction
Sub UpdateCP02(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "UpdateCP2"
    Application.Run (strMacro)
End Sub
'Callback for BTNCreateCP3Docs onAction
Sub UpdateCP03(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "UpdateCP3"
    Application.Run (strMacro)
End Sub
'Callback for BTNResetCP1 onAction
Sub resetCP01(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "resetCP1"
    Application.Run (strMacro)
End Sub
'Callback for BTNResetCP1 onAction
Sub resetCP02(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "resetCP2"
    Application.Run (strMacro)
End Sub        'Callback for BTNResetCP1 onAction
Sub resetCP03(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "resetCP3"
    Application.Run (strMacro)
End Sub
'Callback for BTNUnprotectCP1 onAction
Sub unprotectCP01(control As IRibbonControl)
    CP1Protected = False
    Application.Run "UnprotectCP1"
'    myRibbon.Invalidate
End Sub
'Callback for BTNPprotectCP1 onAction
Sub protectCP01(control As IRibbonControl)
    CP1Protected = True
    Application.Run "protectCP1"
    myRibbon.Invalidate
End Sub

'Callback for BTNUnprotectCP2 onAction
Sub unprotectCP02(control As IRibbonControl)
    CP2Protected = False
    Application.Run "UnprotectCP2"
    myRibbon.Invalidate
End Sub

'Callback for BTNPprotectCP2 onAction
Sub protectCP02(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "protectCP2"
    Application.Run (strMacro)
End Sub        'Callback for BTNUnprotectCP2 onAction
Sub unprotectCP03(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "unprotectCP3"
    Application.Run (strMacro)
End Sub
'Callback for BTNPprotectCP3 onAction
Sub protectCP03(control As IRibbonControl)
    Dim strMacro                        As String
    strMacro = "protectCP3"
    Application.Run (strMacro)
End Sub

' CP1 visibility
Sub GetProtectCP1Visible(control As IRibbonControl, ByRef visible)
    visible = Not CP1Protected
End Sub

Sub GetUnprotectCP1Visible(control As IRibbonControl, ByRef visible)
    visible = CP1Protected
End Sub

' CP2 visibility
Sub GetProtectCP2Visible(control As IRibbonControl, ByRef visible)
    visible = Not CP2Protected
End Sub

Sub GetUnprotectCP2Visible(control As IRibbonControl, ByRef visible)
    visible = CP2Protected
End Sub

' CP3 visibility
Sub GetProtectCP3Visible(control As IRibbonControl, ByRef visible)
    visible = Not CP3Protected
End Sub

Sub GetUnprotectCP3Visible(control As IRibbonControl, ByRef visible)
    visible = CP3Protected
End Sub


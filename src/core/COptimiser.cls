'''=============================================================================
''' Module Name:
''' ---------------------------------------------
''' https://github.com/Flirtle/excel-vba-tools/
''' ---------------------------------------------
''' The Unlicence
'''=============================================================================

'*******************************************************************************
'' This is run before doing a long VBA calculation, to ensure it runs as quickly as possible.
''  Public Methods:
''  - OptimiseOn 'run before any long operations. Make sure to run an error handler to optimiseOff, so original document status is returned
''  - OptimiseOff 'should be run again after those same operations
'*******************************************************************************

'*******************************************************************************
'' Examples
''
''OptimiseOn
''On Error GoTo CleanUp
''
''' long-running code here

''CleanUp:
''OptimiseOff
''If Err.Number <> 0 Then Err.Raise Err.Number
''
'*******************************************************************************

Option Explicit

Private oldCalculation As XlCalculation
Private oldStatusBar As Variant

Private Sub Class_Initialize()
	With Application
		oldCalculation = .Calculation
		oldStatusBar = .StatusBar

		.Calculation = xlCalculationManual
		.ScreenUpdating = False
		.EnableEvents = False
    .DisplayAlerts = False
		.StatusBar = "VBA Macro Running..."
    End With
End Sub

Private Sub Class_Terminate()
    With Application
        .Calculation = oldCalculation
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False

        If oldCalculation = xlCalculationAutomatic Then
            .Calculate
        End If
        .ScreenUpdating = True
    End With
End Sub

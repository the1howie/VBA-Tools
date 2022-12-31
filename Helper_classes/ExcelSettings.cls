VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Aurel Nicolae 2021
' Based on Paul Kelly's example (https://excelmacromastery.com/) with my own additions.
' Purpose: Turn Off functionality that slows down Excel.

Option Explicit

Private calculation_ As XlCalculation
Private displayStatus_ As Boolean
Private enableEvents_ As Boolean
Private screenUpdating_ As Boolean
Private displayAlerts_ As Boolean

Private Sub Class_Initialize()

    Call Backup
    
End Sub

Private Sub Class_Terminate()

    Call Restore
    
End Sub

Public Sub Backup()
' Purpose: Back current settings
    
    calculation_ = Application.calculation
    displayStatus_ = Application.DisplayStatusBar
    enableEvents_ = Application.enableEvents
    screenUpdating_ = Application.screenUpdating
    displayAlerts_ = Application.DisplayAlerts
    
End Sub

Public Sub Restore()
' Purpose: Restore saved settings

    Application.calculation = calculation_
    Application.DisplayStatusBar = displayStatus_
    Application.enableEvents = enableEvents_
    Application.screenUpdating = screenUpdating_
    Application.DisplayAlerts = displayAlerts_
    
End Sub

Public Sub Disable()
' Purpose: Turn OFF automatic calculations, screen updating, etc.

    Application.calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.enableEvents = False
    Application.screenUpdating = False
    Application.DisplayAlerts = False

End Sub

Public Sub Enable()
' Purpose: Turn ON automatic calculations, screen updating, etc.

    Application.calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

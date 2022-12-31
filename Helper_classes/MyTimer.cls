VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStartTime As Date
Private mEndTime As Date

Public Sub Class_Initialize()
    mStartTime = VBA.Now
    mEndTime = vbNull
End Sub

Public Property Let StartTime(Value As Date)
    mStartTime = Value
End Property

Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Let EndTime(Value As Date)
    mEndTime = Value
End Property

Public Property Get EndTime() As Date
    EndTime = mEndTime
End Property

Private Function fCalcElapsedTime() As Date
    If mEndTime = vbNull Then
        fCalcElapsedTime = VBA.Now - mStartTime
    Else
        fCalcElapsedTime = mEndTime - mStartTime
    End If
End Function

Public Function ElapsedTime() As Date
    ElapsedTime = fCalcElapsedTime
End Function

Public Function ElapsedTimeInSeconds() As String
    Dim eTime As Date
    eTime = fCalcElapsedTime
    
    Dim strTime As String
    strTime = Int(CSng(eTime * 24 * 3600)) & " Seconds"
    
    ElapsedTimeInSeconds = strTime
End Function

Public Function ElapsedTimeInMilliseconds() As String
    Dim eTime As Date
    eTime = fCalcElapsedTime
    
    Dim strTime As String
    strTime = Int(CSng(eTime * 24 * 3600 * 1000)) & " Milliseconds"
    
    ElapsedTimeInMilliseconds = strTime
End Function

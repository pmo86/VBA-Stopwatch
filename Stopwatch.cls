VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As LongPtr
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As LongPtr
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long
#End If

Private m_elapsed As Variant 'Decimal
Private m_ticksPerSecond As Currency
Private m_startTick As Currency
Private m_endTick As Currency
Private m_isRunning As Boolean

Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

Public Property Get Elapsed() As Variant
        Elapsed = m_elapsed
End Property

Public Property Get ElapsedMilliseconds() As Variant
        ElapsedMilliseconds = m_elapsed * 1000
End Property

Private Sub Class_Initialize()
    If QueryPerformanceFrequency(m_ticksPerSecond) = 0 Then
        Err.Raise Number:=vbObjectError + 234, Description:="Could not determine frequency!"
    End If
End Sub

Public Sub Start()
    If IsRunning = False Then
        m_isRunning = True
        QueryPerformanceCounter m_startTick
    Else
        Err.Raise Number:=vbObjectError + 236, Description:="Cannot execute Start(): Stopwatch is currently running!"
    End If
End Sub

Public Sub Finish()
    If IsRunning Then
        m_isRunning = False
        QueryPerformanceCounter m_endTick
        m_elapsed = CDec((m_endTick - m_startTick) / m_ticksPerSecond)
    Else
        Err.Raise Number:=vbObjectError + 235, Description:="Cannot execute Finish(): Stopwatch is not running!"
    End If
End Sub

Public Sub Restart()
    ResetTimer
    StartTimer
End Sub

Public Sub Reset()
    m_isRunning = False
    m_elapsed = 0
End Sub

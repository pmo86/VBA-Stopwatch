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

Private Enum Interval
    Second = 1
    Millisecond = 1000
    Microsecond = 1000000
    Nanosecond = 1000000000
End Enum

Private m_Elapsed As Variant 'Decimal
Private m_ticksPerSecond As Currency
Private m_startTick As Currency
Private m_endTick As Currency
Private m_IsRunning As Boolean

Public Property Get IsRunning() As Boolean
    IsRunning = m_IsRunning
End Property

Public Property Get Elapsed() As Variant
    If m_IsRunning Then
        Elapsed = 0
    Else
        Elapsed = m_Elapsed
    End If
End Property

Public Property Get ElapsedMilliseconds() As Variant
    If m_IsRunning Then
        ElapsedMilliseconds = 0
    Else
        ElapsedMilliseconds = m_Elapsed * Interval.Millisecond
    End If
End Property

Private Property Get ElapsedMicroseconds() As Variant 'not used
    ElapsedMicroseconds = m_Elapsed * Interval.Microsecond
End Property

Private Property Get ElapsedNanoseconds() As Variant 'not used
    ElapsedNanoseconds = m_Elapsed * Interval.Nanosecond
End Property

Private Sub Class_Initialize()
    If QueryPerformanceFrequency(m_ticksPerSecond) = 0 Then
        Err.Raise Number:=vbObjectError + 234, Description:="Could not determine frequency!"
    End If
End Sub

Public Sub Start()
    If m_IsRunning = False Then
        m_IsRunning = True
        QueryPerformanceCounter m_startTick
    Else
        Err.Raise Number:=vbObjectError + 236, Description:="Cannot execute Start(): Stopwatch is currently running!"
    End If
End Sub

Public Sub Finish()
    If m_IsRunning Then
        m_IsRunning = False
        QueryPerformanceCounter m_endTick
        m_Elapsed = CDec((m_endTick - m_startTick) / m_ticksPerSecond)
    Else
        Err.Raise Number:=vbObjectError + 235, Description:="Cannot execute Finish(): Stopwatch is not running!"
    End If
End Sub

Public Sub Restart()
    ResetTimer
    StartTimer
End Sub

Public Sub Reset()
    m_IsRunning = False
    m_Elapsed = 0
End Sub

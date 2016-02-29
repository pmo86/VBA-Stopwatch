Attribute VB_Name = "Test"
Sub Test1()
    Dim t As New Stopwatch
    
    t.Start
    
    For i = 1 To 99999999
        ' do something
    Next
    
    t.Finish
    
    Debug.Print "Seconds: " & t.Elapsed
    Debug.Print "Milliseconds: " & t.ElapsedMilliseconds
End Sub

Sub Test2()
    Dim sw As Stopwatch
    Dim d As Object, i As Long, k As Long
    Dim c As Collection
    
    Set sw = New Stopwatch
    k = 1000000
    
    'start dictionary test
    sw.Start
    Set d = CreateObject("Scripting.Dictionary")
    
    For i = 1 To k
        d.Add CStr(i), "a"
    Next
    sw.Finish
    Debug.Print "Dictionary: " & Round(sw.Elapsed, 2) & " seconds"
    
    sw.Reset 'reset stopwatch
    Set d = Nothing
    
    'start collection test
    Set c = New Collection
    
    sw.Start
    For i = 1 To k
        c.Add "a", CStr(i)
    Next
    sw.Finish
    
    Debug.Print "Collection: " & Round(sw.Elapsed, 2) & " seconds"
End Sub

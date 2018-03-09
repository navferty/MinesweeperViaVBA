Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public MinesDict As New Dictionary
Public AllField As Range

Public Sub Minesweeper()
'C3:Q17
Dim r As Range

Set AllField = ActiveSheet.Range("C3:Q17")

Application.EnableEvents = False
With ActiveSheet.Range("C3:Q17")
    .Interior.Color = vbWhite
    .Value = vbNullString
End With
Application.EnableEvents = True

MinesDict.RemoveAll

Do While MinesDict.Count < 25
    Randomize
    Set r = ActiveSheet.Cells(3 + Rnd() * 14, 3 + Rnd() * 14)
    If Not MinesDict.Exists(r.Address) Then
        MinesDict.Add r.Address, r
    End If
Loop
End Sub

Public Sub ThatsAll()
    Dim c As Range
    Dim v As Variant
    For Each c In AllField.Cells
        If c.Interior.Color = vbRed And Not MinesDict.Exists(c.Address) Then
            Beep 600, 700
            MsgBox "Nope =(" & vbCrLf & "You've flagged emty cell"
            Exit Sub
        End If
    Next
    
    For Each v In MinesDict.Keys
        Set c = MinesDict.Item(v)
        If c.Interior.Color <> vbRed Then
            Beep 600, 700
            MsgBox "Nope =(" & vbCrLf & "You haven't flagged all mines"
            Exit Sub
        End If
    Next
    
    AxelFoley
    
    MsgBox "You're cool as cucumber! =))"
End Sub

Public Sub StepOnCell(ByRef c As Range)
    Dim v As Variant
    Dim n As Long
    
    If Not CheckOnField(c) Then Exit Sub
    
    If MinesDict.Exists(c.Address) Then
        Beep 300, 1000
        MsgBox "Boom!"
        For Each v In MinesDict.Keys
            MinesDict.Item(v).Interior.Color = vbRed
        Next
        MinesDict.RemoveAll
    Else
        n = 0
        If MinesDict.Exists(c.Offset(1, 0).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(1, 1).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(0, 1).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(-1, 1).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(-1, 0).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(-1, -1).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(0, -1).Address) Then n = n + 1
        If MinesDict.Exists(c.Offset(1, -1).Address) Then n = n + 1
        c.Value = n
        c.Interior.Color = vbGreen
        
        If n = 0 Then
            StepOnCell c.Offset(1, 0)
            StepOnCell c.Offset(1, 1)
            StepOnCell c.Offset(0, 1)
            StepOnCell c.Offset(-1, 1)
            StepOnCell c.Offset(-1, 0)
            StepOnCell c.Offset(-1, -1)
            StepOnCell c.Offset(0, -1)
            StepOnCell c.Offset(1, -1)
        End If
    End If
End Sub

Public Function CheckOnField(ByRef r As Range) As Boolean
CheckOnField = False
If AllField Is Nothing Then Exit Function
If r.Cells(1, 1).Interior.Color = vbGreen Then Exit Function
If r.Cells(1, 1).Interior.Color = vbRed Then Exit Function
If Intersect(r, AllField) Is Nothing Then Exit Function
If Union(r, AllField).Address <> AllField.Address Then Exit Function
CheckOnField = True
End Function

Private Sub AxelFoley()
    Beep 659, 460
    Beep 784, 340
    Beep 659, 230
    Beep 659, 110
    Beep 880, 230
    Beep 659, 230
    Beep 587, 230
    Beep 659, 460
    Beep 988, 340
    Beep 659, 230
    Beep 659, 110
    Beep 1047, 230
    Beep 988, 230
    Beep 784, 230
    Beep 659, 230
    Beep 988, 230
    Beep 1318, 230
    Beep 659, 110
    Beep 587, 230
    Beep 587, 110
    Beep 494, 230
    Beep 740, 230
    Beep 659, 460
End Sub

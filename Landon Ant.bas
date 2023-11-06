Attribute VB_Name = "Ä£¿é1"
Option Explicit
Sub move()
    'Dim dinput As Object
    'Selection.Interior.ColorIndex = 2
    'Set dinput = Application.InputBox("Select your aray:", Type:=8)
    'MsgBox Range(Selection.Address).Column
    'Dim i As Integer
    Dim row_start As Integer, row_end As Integer, col_start As Integer, col_end As Integer
    row_start = 2
    row_end = 500
    col_start = 2
    col_end = 500
    
    Dim step As Integer
    step = 0
    
    While (Range(Selection.Address).Column < col_end) And _
    (Range(Selection.Address).Column > col_start) And _
    (Range(Selection.Address).Row < row_end) And _
    (Range(Selection.Address).Row > row_start)
        If Selection.Interior.ColorIndex = 2 Then
            Selection.Interior.ColorIndex = 1
            If Selection.Value = "R" Then
                Selection.Offset(1, 0).Select
                Selection.Value = "D"
            ElseIf Selection.Value = "D" Then
                Selection.Offset(0, -1).Select
                Selection.Value = "L"
            ElseIf Selection.Value = "L" Then
                Selection.Offset(-1, 0).Select
                Selection.Value = "U"
            ElseIf Selection.Value = "U" Then
                Selection.Offset(0, 1).Select
                Selection.Value = "R"
            End If
        ElseIf Selection.Interior.ColorIndex = 1 Then
            Selection.Interior.ColorIndex = 2
            If Selection.Value = "R" Then
                Selection.Offset(-1, 0).Select
                Selection.Value = "U"
            ElseIf Selection.Value = "U" Then
                Selection.Offset(0, -1).Select
                Selection.Value = "L"
            ElseIf Selection.Value = "L" Then
                Selection.Offset(1, 0).Select
                Selection.Value = "D"
            ElseIf Selection.Value = "D" Then
                Selection.Offset(0, 1).Select
                Selection.Value = "R"
            End If
        End If
        step = step + 1
        Application.Wait Now() + VBA.TimeValue("00:00:01")
    Wend
    MsgBox step
End Sub

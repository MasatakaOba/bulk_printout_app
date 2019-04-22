Attribute VB_Name = "print"
Sub かんたん一括印刷くん起動()
    
    Dim ws As Worksheet, flag As Boolean
    
    UserForm1.Show vbModeless
    
    For Each ws In Worksheets
        
        If ws.Name = "印刷リスト" Then flag = True
    
    Next ws
    
    If flag = True Then
    
        Sheets("印刷リスト").Visible = True
        Sheets("印刷リスト").Activate
        Columns("A:A").Select
        Range("A1").Select
        MsgBox "「印刷リスト」シートに" & vbCrLf & "必要なリストを貼り付けてください。"
        Exit Sub
    
    Else
    
        ActiveWorkbook.Sheets.Add
        
        On Error Resume Next
        
        ActiveSheet.Name = "印刷リスト"
        Columns("A:A").Select
        Range("A1").Interior.Color = 65535
        Range("B1") = "←ここから値で貼付け！"
        Range("A1").Select
        
        MsgBox "「印刷リスト」シートに" & vbCrLf & "必要なリストを貼り付けてください。"
    
    End If
    
End Sub

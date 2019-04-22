VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "かんたん一括印刷くん"
   ClientHeight    =   4396
   ClientLeft      =   49
   ClientTop       =   392
   ClientWidth     =   4711
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim rw As Long
    Dim msg0 As Integer
        
    '■ポップアップ画面非表示・イベント発生不可・画面描写中止
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
        
    rw = Cells(Rows.Count, 1).End(xlUp).row 'A列の最も下のセルの行の場所を取得
    On Error GoTo myerror 'A列に何も入力されていなかった場合、myerrorへ
    
    ListBox1.List = Range(Cells(1, 1), Cells(rw, 1)).Value
    MsgBox "登録件数は" & ListBox1.ListCount & "件です。" & vbCrLf & vbCrLf _
    & "印刷するシートの代入先のセルを選択し、" & vbCrLf & "「印刷スタート！」のボタンを押してください。"
    
    Exit Sub
        
myerror:     MsgBox "A1セルから貼り付けてください"

    '■ポップアップ画面非表示・イベント発生不可・画面描写中止
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Private Sub CommandButton2_Click()

    Dim S As String
    Dim msg0 As Integer
    Dim msg1 As Integer
    Dim msg2 As Integer
    Dim i As Long
    S = ActiveCell.Address

    If ListBox1.ListCount = 0 Then
    
        MsgBox "リストを読み込んでください。"
        
    Else
    
        msg0 = MsgBox("代入先は" & S & "でよろしいですか？", vbYesNo + vbQuestion, "確認")
        
        If msg0 = vbYes Then
        
            For i = 0 To 4
            
                ActiveCell.Value = ListBox1.List(i)
                Range(S) = Replace(Range(S), "", "")
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False  'シートを印刷(印刷ボタンを押す)Copies:=1(1枚ずつ印刷)
                On Error GoTo myerror1
                
            Next
myerror1:
                msg1 = MsgBox("正しく印刷されていますか？", vbYesNo + vbQuestion, "確認")
                
            If msg1 = vbYes Then
            
                For i = 5 To ListBox1.ListCount - 1
                
                    ActiveCell.Value = ListBox1.List(i)
                    Range(S) = Replace(Range(S), "", "")
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                    IgnorePrintAreas:=False  'シートを印刷(印刷ボタンを押す)Copies:=1(1枚ずつ印刷)
    
                Next
                
                msg2 = MsgBox("印刷が完了しました！" & vbCrLf & "「印刷リスト」シートを削除しますか？", vbYesNo + vbQuestion, "確認")
                
                If msg2 = vbYes Then
                
                    Application.DisplayAlerts = False
                    Sheets("印刷リスト").Delete
                    On Error GoTo myerror2
                    Unload UserForm2
                    Application.DisplayAlerts = True
                    
                Else
                
                    End If
                
            Else
                
                MsgBox "もう一度設定しなおしてください。"
                
            End If
                
        Else
        
            MsgBox "もう一度セルを選択して下さい。"
        
        End If
    
    End If
    
    Exit Sub
    
myerror2:
    MsgBox "「印刷リスト」シートはこのファイルにありません。必要に応じて削除ください。"
    
End Sub

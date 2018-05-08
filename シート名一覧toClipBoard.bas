Attribute VB_Name = "シート名一覧toClipBoard"
'注意
' このモジュールは、DataObjectを使用してクリップボードに文字列を送ります。
' DataObjectを使用するには「Microsoft Forms 2.0 Object Library」への参照が必要です。
' Visual Basic Editorのメニューから［ツール］→［参照設定］コマンドを選択し
' ［参照設定］ダイアログボックスで「Microsoft Forms 2.0 Object Library」にチェックを入れて、
' ［OK］ボタンをクリックし、参照設定を行います。
'
' 「参照可能なライブラリ ファイル」のリストにない場合は、
' ［参照設定］ダイアログボックスで［参照］ボタンをクリックして
' 「C:\WINNT(または Windows)\system32\FM20.DLL」を選択します。r

'
'
'開いているブックのシート一覧をクリップボードに貼り付けます
'クリップボードへの貼り付けはsetClipBoadのコメントを参照
Sub シート名一覧toClipBoard()
    'シート名の文字列を保持します
    Dim workSheetNames As String
      
    For Each targetWorkSheet In Sheets
        workSheetNames = workSheetNames & targetWorkSheet.Name & vbCrLf
    
    Next
    
    'クリップボードに設定します
    setClipBoad (workSheetNames)

End Sub

'
'　文字列をクリップボードに貼り付けます
'＜説明＞
' ［ツール］→［参照設定］で「Microsoft Forms 2.0 Object Library」に
' チェックして使用する。
'［参照可能なライブラリ］のリストにない場合は［参照設定］
'ダイアログボックスで［参照］ボタンをクリックして
'「C:\Windows\system32\FM20.DLL」を選択する
Function setClipBoad(strValue As String)

    Dim CB As New DataObject
    With CB
        .SetText strValue
        .PutInClipboard
    End With

End Function

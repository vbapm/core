Attribute VB_Name = "JapaneseExample"
' 日本語のサンプルモジュール
' このモジュールはエンコーディング検出のテスト用です
' Windows-932 (Windows-31J) でエンコードされています

' よく使われる日本語の単語やフレーズ:
' こんにちは世界 おはようございます
' ありがとうございます すみません
' 今日はいい天気ですね

Function 計算する(値 As Double) As Double
    ' 入力された値を処理します
    Dim 結果 As Double
    結果 = 値 * 1.1
    計算する = 結果
End Function

Sub メッセージを表示()
    Dim 挨拶 As String
    Dim 名前 As String
    名前 = Environ("USERNAME")
    挨拶 = "こんにちは、" & 名前 & "さん！"
    MsgBox 挨拶, vbInformation, "ご挨拶"
End Sub

Function 漢字チェック(文字列 As String) As Boolean
    ' 文字列に漢字が含まれているかチェック
    Dim i As Integer
    For i = 1 To Len(文字列)
        If AscW(Mid(文字列, i, 1)) >= &H4E00 Then
            漢字チェック = True
            Exit Function
        End If
    Next i
    漢字チェック = False
End Function

Sub データ処理()
    ' データの処理を行います
    Application.StatusBar = "処理中..."
    ' ここにデータ処理のコードを記述
    Application.StatusBar = False
End Sub

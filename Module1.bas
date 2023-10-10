Attribute VB_Name = "Module1"
Option Explicit
Sub シート名書出開始()
    Application.ScreenUpdating = False
    Dim 文 As String
    Dim 起点フォルダ As Folder
    Dim 始時 As Date, 終時 As Date
    Dim 終行 As Long
    Dim FSO As New FileSystemObject
    文 = "起点フォルダ下にある全てのフォルダ内のExcelファイルのシート名を書出します。" & vbCrLf & vbCrLf & "処理を開始してよろしいですか？"
    文 = 文 & vbCrLf & "※起点フォルダ設定次第では処理に時間が掛かります！"
    If MsgBox(文, vbYesNo) = vbYes Then
        Set 起点フォルダ = FSO.GetFolder(Sheets("シート名検査").Cells(2, 1))
        始時 = Timer
        実行中.Show vbModeless
        実行中.Repaint
        Call シート名書出(起点フォルダ)
        With Sheets("シート名検査")
            終行 = .Cells(Rows.Count, 1).End(xlUp).Row
            .Range("A5:L" & Rows.Count).Borders.LineStyle = False
            .Range("A5:L" & 終行).Borders.LineStyle = True
        End With
        終時 = Timer
        Unload 実行中
        MsgBox "書出完了" & vbCrLf & "処理時間：" & 終時 - 始時
    End If
    Application.ScreenUpdating = True
End Sub
Sub シート名書出(フォルダ As Folder)
    Dim ファイル As File
    Dim 書込(1 To 1, 1 To 12)
    Dim 終行 As Long, 回 As Long, 列 As Long, シート数 As Long
    Dim シート As Worksheet
    With Sheets("シート名検査")
        For Each ファイル In フォルダ.Files
            If ファイル.Name <> ThisWorkbook.Name And InStrRev(ファイル.Name, ".xls") > 0 And InStrRev(ファイル.Name, "~$") = 0 Then
                Workbooks.Open ファイル.Path, UpdateLinks:=0
                書込(1, 1) = フォルダ.Path
                書込(1, 2) = ファイル.Name
                シート数 = Workbooks(ファイル.Name).Sheets.Count
                If シート数 <= 10 Then
                    列 = 3
                    For Each シート In Workbooks(ファイル.Name).Sheets
                        書込(1, 列) = シート.Name
                        列 = 列 + 1
                    Next
                    Else: 書込(1, 3) = "※シート数超過"
                End If
                Workbooks(ファイル.Name).Close SaveChanges:=False
                終行 = .Cells(Rows.Count, 1).End(xlUp).Row
                Range(.Cells(終行 + 1, 1), .Cells(終行 + 1, 12)) = 書込
                Erase 書込
            End If
        Next
    End With
    For Each フォルダ In フォルダ.SubFolders
        Call シート名書出(フォルダ)
    Next
End Sub
Sub シート名書出クリア()
    With Sheets("シート名検査")
        Range(.Cells(5, 1), .Cells(Rows.Count, 12)).ClearContents
        Range(.Cells(5, 1), .Cells(Rows.Count, 12)).Borders.LineStyle = False
    End With
End Sub
Sub シート連続書出()
    Application.ScreenUpdating = False
    Dim 終行 As Long, 終列 As Long, 行 As Long, 列 As Long, 添行 As Long, 基終行 As Long
    Dim 文 As String
    Dim 始時 As Date, 終時 As Date
    Dim 基配列, 配列
    With Sheets("展開検査設定")
        終行 = .Cells(Rows.Count, 4).End(xlUp).Row
        If 終行 = 1 Then MsgBox "「展開検査設定」シートにデータがありません": Exit Sub
        文 = "ファイル" & 終行 - 1 & "件分について連続で「展開書出」シートに書出します" & vbCrLf & vbCrLf & "処理を開始してよろしいですか？"
        文 = 文 & vbCrLf & "※展開シートの1列目で最下行を算定します"
        If MsgBox(文, vbYesNo) = vbNo Then Exit Sub
        始時 = Timer
        実行中.Show vbModeless
        実行中.Repaint
        ReDim リスト(2 To 終行, 1 To 3)
        For 行 = 2 To 終行
            リスト(行, 1) = .Cells(行, 2) & "\" & .Cells(行, 3)
            リスト(行, 2) = .Cells(行, 3)
            リスト(行, 3) = .Cells(行, 4)
        Next
    End With
    For 添行 = 2 To 終行
        Workbooks.Open リスト(添行, 1)
        With Workbooks(リスト(添行, 2))
            With Sheets(リスト(添行, 3))
                終行 = .Cells(Rows.Count, 1).End(xlUp).Row
                終列 = .Cells(1, Columns.Count).End(xlToLeft).Column
                If 終行 = 1 And 終列 = 1 Then
                    ReDim 基配列(1 To 1, 1 To 1)
                    Else: 基配列 = Range(.Cells(1, 1), .Cells(終行, 終列))
                End If
            End With
            ReDim 配列(1 To 終行, 1 To 終列 + 1)
            For 行 = 1 To 終行
                配列(行, 1) = .Name
                For 列 = 1 To 終列
                    配列(行, 列 + 1) = 基配列(行, 列)
                Next
            Next
        End With
        Workbooks(リスト(添行, 2)).Close
        With ThisWorkbook.Sheets("展開書出")
            基終行 = .Cells(Rows.Count, 1).End(xlUp).Row
            Range(.Cells(基終行 + 1, 1), .Cells(基終行 + 終行, 終列 + 1)) = 配列
        End With
    Next
    終時 = Timer
    Unload 実行中
    MsgBox "書出完了" & vbCrLf & "処理時間：" & 終時 - 始時
    Application.ScreenUpdating = True
End Sub
Sub 検査設定リストクリア()
    With Sheets("展開検査設定")
        Range(.Cells(2, 2), .Cells(Rows.Count, 4)).ClearContents
    End With
End Sub
Sub 展開書出クリア()
    With Sheets("展開書出")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
    End With
End Sub

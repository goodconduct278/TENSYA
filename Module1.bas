Attribute VB_Name = "Module1"
'==============================================================================
' 転記ツール メインモジュール
' Module_Transfer
'
' 機能：
'   - コピー元Excel → 転記先Excel への自動転記
'   - 設定表（設定シート）から動作パラメータを読み込み
'   - 転記先の行数が不足する場合は3行ブロック単位で自動追加
'   - SUM式を完全再生成
'   - バックアップ作成（必須）
'
' フェーズ1（最小版）+ 行追加機能
'   ※明細ブロック判定は単純な「数量・名称・単位がある行」のみ
'   ※複数行明細の検出は将来フェーズ
'==============================================================================

'------ ファイルパス（メインシートのボタンから設定）------
Public g_srcFilePath As String   ' コピー元ファイルパス
Public g_dstFilePath As String   ' 転記先ファイルパス

'------ 設定値（LoadSettingsで設定シートから読み込み）------
Private g_srcSheetName  As String
Private g_srcStartRow   As Long
Private g_srcColName    As Long  ' 列番号（数値）に変換して保持
Private g_srcColSpec    As Long
Private g_srcColQty     As Long
Private g_srcColUnit    As Long
Private g_srcColPrice   As Long

Private g_dstSheetName  As String
Private g_dstStartRow   As Long
Private g_dstColName    As Long
Private g_dstColSpec    As Long
Private g_dstColQty     As Long
Private g_dstColUnit    As Long
Private g_dstColPrice   As Long
Private g_dstColAmount  As Long
Private g_sumKeyword    As String
Private g_doBackup      As Boolean

'------ ログ蓄積用 ------
Private g_logBuffer As String


'==============================================================================
' 診断：コピー元データの読み取り結果をダンプ
' 設定の列指定が正しいか確認するため、開始行から最終行までをログシートに出力する
'==============================================================================
Public Sub Debug_コピー元ダンプ()

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim r As Long, lastRow As Long
    Dim logWs As Worksheet
    Dim outRow As Long
    Dim hasError As Boolean
    Dim errMsg As String
    hasError = False

    If g_srcFilePath = "" Then
        MsgBox "先にコピー元ファイルを選択してください", vbExclamation
        Exit Sub
    End If

    If Not LoadSettings() Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    Set srcWb = Workbooks.Open(g_srcFilePath, ReadOnly:=True)

    If Not SheetExists(srcWb, g_srcSheetName) Then
        MsgBox "コピー元シートが見つかりません：[" & g_srcSheetName & "]" & vbCrLf & vbCrLf & _
               "このファイルにあるシート：" & vbCrLf & ListSheetNames(srcWb), vbExclamation
        GoTo CleanUp
    End If

    Set srcWs = srcWb.Sheets(g_srcSheetName)

    ' 最終行検出（空シート対策）
    Dim findResult As Range
    Set findResult = srcWs.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious)
    If findResult Is Nothing Then
        MsgBox "コピー元シート [" & g_srcSheetName & "] にデータがありません。" & vbCrLf & _
               "シート名の設定を確認してください。", vbExclamation
        GoTo CleanUp
    End If
    lastRow = findResult.Row

    ' ログシート確保
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("ログ")
    On Error GoTo ErrHandler
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logWs.Name = "ログ"
    End If

    outRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 2
    If outRow < 3 Then outRow = 3

    ' ヘッダー
    logWs.Cells(outRow, 1).Value = "===== コピー元ダンプ " & _
        Format(Now, "yyyy/mm/dd hh:nn:ss") & " ====="
    outRow = outRow + 1
    logWs.Cells(outRow, 1).Value = "シート：[" & g_srcSheetName & "]" & _
        " 開始行：" & g_srcStartRow & " 最終行：" & lastRow
    outRow = outRow + 1
    logWs.Cells(outRow, 1).Value = "行"
    logWs.Cells(outRow, 2).Value = "名称(" & NumToColLetter(g_srcColName) & ")"
    logWs.Cells(outRow, 3).Value = "仕様(" & NumToColLetter(g_srcColSpec) & ")"
    logWs.Cells(outRow, 4).Value = "数量(" & NumToColLetter(g_srcColQty) & ")"
    logWs.Cells(outRow, 5).Value = "単位(" & NumToColLetter(g_srcColUnit) & ")"
    logWs.Cells(outRow, 6).Value = "単価(" & NumToColLetter(g_srcColPrice) & ")"
    logWs.Cells(outRow, 7).Value = "判定"
    outRow = outRow + 1

    Dim nameVal As String, specVal As String, unitVal As String
    Dim qtyVal As Variant, priceVal As Variant
    Dim verdict As String
    Dim isAnchor As Boolean

    For r = g_srcStartRow To lastRow
        nameVal = Trim(CStr(srcWs.Cells(r, g_srcColName).Value))
        specVal = Trim(CStr(srcWs.Cells(r, g_srcColSpec).Value))
        qtyVal = srcWs.Cells(r, g_srcColQty).Value
        unitVal = Trim(CStr(srcWs.Cells(r, g_srcColUnit).Value))
        priceVal = srcWs.Cells(r, g_srcColPrice).Value

        ' アンカー判定（多行ブロック対応版と同じロジック）
        isAnchor = (IsNumeric(qtyVal) And qtyVal <> "" And unitVal <> "")

        verdict = ""
        If isAnchor Then
            verdict = "★アンカー（ブロック2行目）"
        Else
            ' アンカーでない場合、上下のアンカーに連結される可能性を判定
            Dim aboveAnchor As Boolean, belowAnchor As Boolean
            aboveAnchor = False: belowAnchor = False

            ' 直下がアンカー → 上の行として連結
            If r < lastRow Then
                Dim nextQty As Variant, nextUnit As String
                nextQty = srcWs.Cells(r + 1, g_srcColQty).Value
                nextUnit = Trim(CStr(srcWs.Cells(r + 1, g_srcColUnit).Value))
                If IsNumeric(nextQty) And nextQty <> "" And nextUnit <> "" Then
                    belowAnchor = True
                End If
            End If
            ' 直上がアンカー → 下の行として連結
            If r > g_srcStartRow Then
                Dim prevQty As Variant, prevUnit As String
                prevQty = srcWs.Cells(r - 1, g_srcColQty).Value
                prevUnit = Trim(CStr(srcWs.Cells(r - 1, g_srcColUnit).Value))
                If IsNumeric(prevQty) And prevQty <> "" And prevUnit <> "" Then
                    aboveAnchor = True
                End If
            End If

            If belowAnchor And (nameVal <> "" Or specVal <> "") Then
                verdict = "→ 次行ブロックの1行目（連結）"
            ElseIf aboveAnchor And nameVal = "" And specVal <> "" Then
                verdict = "→ 前行ブロックの3行目（連結）"
            ElseIf aboveAnchor And nameVal <> "" Then
                verdict = "[次ブロックの開始扱い／前ブロックには非連結]"
            Else
                verdict = "[孤立／無視]"
            End If
        End If

        logWs.Cells(outRow, 1).Value = r
        logWs.Cells(outRow, 2).Value = nameVal
        logWs.Cells(outRow, 3).Value = specVal
        logWs.Cells(outRow, 4).Value = qtyVal
        logWs.Cells(outRow, 5).Value = unitVal
        logWs.Cells(outRow, 6).Value = priceVal
        logWs.Cells(outRow, 7).Value = verdict
        outRow = outRow + 1
    Next r

    ' 列幅調整
    logWs.Columns("A:G").AutoFit

    GoTo CleanUp

ErrHandler:
    hasError = True
    errMsg = "Err#" & Err.Number & " : " & Err.Description

CleanUp:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If hasError Then
        MsgBox "ダンプ中にエラーが発生しました。" & vbCrLf & vbCrLf & errMsg, _
               vbCritical, "ダンプ失敗"
    Else
        MsgBox "「ログ」シートにコピー元の中身をダンプしました。" & vbCrLf & _
               "各行の判定理由を確認してください。", vbInformation, "ダンプ完了"
    End If
End Sub


'==============================================================================
' 初回セットアップ（1回だけ実行する）
' メイン・設定・ログの3シートを作成し、設定表のデフォルト値を投入する
'==============================================================================
Public Sub Setup_転記ツール初期化()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--- メインシート ---
    If Not SheetExistsInBook(wb, "メイン") Then
        wb.Sheets.Add(Before:=wb.Sheets(1)).Name = "メイン"
    End If
    With wb.Sheets("メイン")
        .Cells.Clear
        .Range("A1").Value = "転記ツール"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A2").Value = "コピー元ファイル："
        .Range("A3").Value = "転記先ファイル："
        .Range("A2:A3").Font.Bold = True
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 80
    End With

    '--- 設定シート（デフォルト値を投入）---
    If Not SheetExistsInBook(wb, "設定") Then
        wb.Sheets.Add(After:=wb.Sheets("メイン")).Name = "設定"
    End If
    With wb.Sheets("設定")
        .Cells.Clear
        .Range("A1").Value = "設定項目"
        .Range("B1").Value = "値"
        .Range("A1:B1").Font.Bold = True

        Dim settings As Variant
        settings = Array( _
            Array("コピー元シート名", "内訳(3)"), _
            Array("コピー元開始行", 6), _
            Array("コピー元名称列", "A"), _
            Array("コピー元仕様列", "B"), _
            Array("コピー元数量列", "C"), _
            Array("コピー元単位列", "D"), _
            Array("コピー元単価列", "E"), _
            Array("転記先シート名", "防水工事"), _
            Array("転記先開始行", 8), _
            Array("転記先名称列", "D"), _
            Array("転記先仕様列", "E"), _
            Array("転記先数量列", "C"), _
            Array("転記先単位列", "G"), _
            Array("転記先単価列", "E"), _
            Array("転記先金額列", "F"), _
            Array("合計行キーワード", "合計"), _
            Array("バックアップ作成", "TRUE") _
        )

        Dim i As Long
        For i = 0 To UBound(settings)
            .Cells(i + 2, 1).Value = settings(i)(0)
            .Cells(i + 2, 2).Value = settings(i)(1)
        Next i

        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 20
    End With

    '--- ログシート ---
    If Not SheetExistsInBook(wb, "ログ") Then
        wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).Name = "ログ"
    End If
    With wb.Sheets("ログ")
        .Cells.Clear
        .Range("A1").Value = "ログ"
        .Range("A1").Font.Bold = True
        .Columns("A").ColumnWidth = 100
    End With

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "初期化完了。" & vbCrLf & vbCrLf & _
           "「設定」シートで値を確認・調整してから、" & vbCrLf & _
           "「メイン」シートにボタンを配置してください。", _
           vbInformation, "セットアップ完了"
End Sub


'==============================================================================
' このブック内のシート存在確認（補助）
'==============================================================================
Private Function SheetExistsInBook(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExistsInBook = (Not ws Is Nothing)
    On Error GoTo 0
End Function


'==============================================================================
' メイン処理（ボタンに登録）
'==============================================================================
Public Sub Main_転記実行()

    Dim srcWb As Workbook, dstWb As Workbook
    Dim srcWs As Worksheet, dstWs As Worksheet
    Dim srcData As Variant     ' 読み取った明細データ（2次元配列）
    Dim dataCount As Long      ' 明細件数
    Dim addedRows As Long      ' 追加した行数

    ' 画面更新停止
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    g_logBuffer = ""
    Call WriteLog("===== 転記処理開始 =====")
    Call WriteLog("実行日時：" & Format(Now, "yyyy/mm/dd hh:nn:ss"))

    '--- 1. ファイルパスの確認 ---
    If g_srcFilePath = "" Then
        Call ShowError("コピー元ファイルを選択してください")
        GoTo CleanUp
    End If
    If g_dstFilePath = "" Then
        Call ShowError("転記先ファイルを選択してください")
        GoTo CleanUp
    End If
    Call WriteLog("コピー元：" & g_srcFilePath)
    Call WriteLog("転記先　：" & g_dstFilePath)

    '--- 2. 設定読み込み ---
    If Not LoadSettings() Then
        GoTo CleanUp
    End If

    '--- 3. バックアップ作成（必須） ---
    If g_doBackup Then
        If Not BackupDestinationFile() Then
            Call ShowError("バックアップを作成できませんでした。転記を中止します。")
            GoTo CleanUp
        End If
    End If

    '--- 4. ファイルを開く ---
    Set srcWb = Workbooks.Open(g_srcFilePath, ReadOnly:=True)
    Set dstWb = Workbooks.Open(g_dstFilePath)

    '--- 5. シート存在確認 ---
    If Not SheetExists(srcWb, g_srcSheetName) Then
        Call ShowError("コピー元シートが見つかりません：" & g_srcSheetName & vbCrLf & vbCrLf & _
                       "このファイルにあるシート：" & vbCrLf & ListSheetNames(srcWb))
        GoTo CleanUp
    End If
    If Not SheetExists(dstWb, g_dstSheetName) Then
        Call ShowError("転記先シートが見つかりません：" & g_dstSheetName & vbCrLf & vbCrLf & _
                       "このファイルにあるシート：" & vbCrLf & ListSheetNames(dstWb))
        GoTo CleanUp
    End If

    Set srcWs = srcWb.Sheets(g_srcSheetName)
    Set dstWs = dstWb.Sheets(g_dstSheetName)

    '--- 6. コピー元データ読み取り ---
    dataCount = ReadSourceData(srcWs, srcData)
    Call WriteLog("読み取り件数：" & dataCount & "件")

    If dataCount = 0 Then
        Call ShowError("転記できる明細がありませんでした")
        GoTo CleanUp
    End If

    '--- 7. 既存データクリア ---
    Call ClearDestinationData(dstWs)
    Call WriteLog("既存データクリア完了")

    '--- 8. 行数確認・追加 ---
    addedRows = EnsureRows(dstWs, dataCount)
    If addedRows > 0 Then
        Call WriteLog("行追加：" & addedRows & "行（" & (addedRows / 3) & "ブロック）")
    Else
        Call WriteLog("行追加：なし")
    End If

    '--- 9. 転記実行 ---
    Call WriteToDestination(dstWs, srcData, dataCount)
    Call WriteLog("転記件数：" & dataCount & "件")

    '--- 10. SUM式再生成 ---
    Dim newFormula As String
    newFormula = RebuildSumFormula(dstWs)
    Call WriteLog("更新後SUM式：" & newFormula)

    '--- 11. 保存 ---
    dstWb.Save
    Call WriteLog("転記先ファイル保存完了")

    Call WriteLog("===== 転記処理正常終了 =====")
    Call FlushLog

    MsgBox dataCount & "件の転記が完了しました。", vbInformation, "転記完了"

CleanUp:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    ' dstWbは保存済みのため閉じる（必要に応じてコメントアウト）
    If Not dstWb Is Nothing Then dstWb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Call ShowError("予期しないエラーが発生しました：" & Err.Description)
    Call WriteLog("エラー：" & Err.Description)
    Call FlushLog
    Resume CleanUp

End Sub


'==============================================================================
' コピー元ファイル選択（ボタンに登録）
'==============================================================================
Public Sub SelectSourceFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "コピー元（内訳明細）ファイルを選択"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            g_srcFilePath = .SelectedItems(1)
            ' メインシートが存在する場合のみ表示用にパスを書き込む
            On Error Resume Next
            ThisWorkbook.Sheets("メイン").Range("B2").Value = g_srcFilePath
            On Error GoTo 0
            MsgBox "コピー元を選択しました：" & vbCrLf & g_srcFilePath, _
                   vbInformation, "選択完了"
        End If
    End With
End Sub


'==============================================================================
' 転記先ファイル選択（ボタンに登録）
'==============================================================================
Public Sub SelectDestinationFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "転記先（見積書）ファイルを選択"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            g_dstFilePath = .SelectedItems(1)
            ' メインシートが存在する場合のみ表示用にパスを書き込む
            On Error Resume Next
            ThisWorkbook.Sheets("メイン").Range("B3").Value = g_dstFilePath
            On Error GoTo 0
            MsgBox "転記先を選択しました：" & vbCrLf & g_dstFilePath, _
                   vbInformation, "選択完了"
        End If
    End With
End Sub


'==============================================================================
' 設定シートから設定値を読み込む
' 戻り値：成功=True、失敗=False
'==============================================================================
Private Function LoadSettings() As Boolean
    Dim ws As Worksheet
    Dim missing As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("設定")

    ' 各項目をA列で検索してB列の値を取得
    g_srcSheetName = GetSettingValue(ws, "コピー元シート名")
    g_srcStartRow = CLng(val(GetSettingValue(ws, "コピー元開始行")))
    g_srcColName = ColLetterToNum(GetSettingValue(ws, "コピー元名称列"))
    g_srcColSpec = ColLetterToNum(GetSettingValue(ws, "コピー元仕様列"))
    g_srcColQty = ColLetterToNum(GetSettingValue(ws, "コピー元数量列"))
    g_srcColUnit = ColLetterToNum(GetSettingValue(ws, "コピー元単位列"))
    g_srcColPrice = ColLetterToNum(GetSettingValue(ws, "コピー元単価列"))

    g_dstSheetName = GetSettingValue(ws, "転記先シート名")
    g_dstStartRow = CLng(val(GetSettingValue(ws, "転記先開始行")))
    g_dstColName = ColLetterToNum(GetSettingValue(ws, "転記先名称列"))
    g_dstColSpec = ColLetterToNum(GetSettingValue(ws, "転記先仕様列"))
    g_dstColQty = ColLetterToNum(GetSettingValue(ws, "転記先数量列"))
    g_dstColUnit = ColLetterToNum(GetSettingValue(ws, "転記先単位列"))
    g_dstColPrice = ColLetterToNum(GetSettingValue(ws, "転記先単価列"))
    g_dstColAmount = ColLetterToNum(GetSettingValue(ws, "転記先金額列"))
    g_sumKeyword = GetSettingValue(ws, "合計行キーワード")
    g_doBackup = (UCase(GetSettingValue(ws, "バックアップ作成")) = "TRUE")

    ' バリデーション
    missing = ""
    If g_srcSheetName = "" Then missing = missing & "コピー元シート名 "
    If g_srcStartRow <= 0 Then missing = missing & "コピー元開始行 "
    If g_dstSheetName = "" Then missing = missing & "転記先シート名 "
    If g_dstStartRow <= 0 Then missing = missing & "転記先開始行 "
    If g_sumKeyword = "" Then missing = missing & "合計行キーワード "

    ' 列番号バリデーション（ColLetterToNum が 0 を返した場合は不正値）
    If g_srcColName <= 0 Then missing = missing & "コピー元名称列(不正) "
    If g_srcColSpec <= 0 Then missing = missing & "コピー元仕様列(不正) "
    If g_srcColQty <= 0 Then missing = missing & "コピー元数量列(不正) "
    If g_srcColUnit <= 0 Then missing = missing & "コピー元単位列(不正) "
    If g_srcColPrice <= 0 Then missing = missing & "コピー元単価列(不正) "
    If g_dstColName <= 0 Then missing = missing & "転記先名称列(不正) "
    If g_dstColSpec <= 0 Then missing = missing & "転記先仕様列(不正) "
    If g_dstColQty <= 0 Then missing = missing & "転記先数量列(不正) "
    If g_dstColUnit <= 0 Then missing = missing & "転記先単位列(不正) "
    If g_dstColPrice <= 0 Then missing = missing & "転記先単価列(不正) "
    If g_dstColAmount <= 0 Then missing = missing & "転記先金額列(不正) "

    If missing <> "" Then
        Call ShowError("設定表を確認してください：" & missing)
        LoadSettings = False
        Exit Function
    End If

    LoadSettings = True
    Exit Function

ErrHandler:
    Call ShowError("設定読み込み中にエラー：" & Err.Description)
    LoadSettings = False
End Function


'==============================================================================
' 設定シートから1項目の値を取得
'==============================================================================
Private Function GetSettingValue(ws As Worksheet, key As String) As String
    Dim foundCell As Range
    Set foundCell = ws.Columns("A").Find(What:=key, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        GetSettingValue = ""
    Else
        GetSettingValue = CStr(ws.Cells(foundCell.Row, 2).Value)
    End If
End Function


'==============================================================================
' 列文字（"A", "B", "AA"等）を列番号（1, 2, 27等）に変換
'==============================================================================
Private Function ColLetterToNum(colLetter As String) As Long
    On Error Resume Next
    If colLetter = "" Then
        ColLetterToNum = 0
        Exit Function
    End If
    ColLetterToNum = Range(colLetter & "1").Column
    On Error GoTo 0
End Function


'==============================================================================
' ワークブックのシート名を改行区切りで列挙（補助）
'==============================================================================
Private Function ListSheetNames(wb As Workbook) As String
    Dim s As String
    Dim ws As Worksheet
    s = ""
    For Each ws In wb.Sheets
        s = s & "・[" & ws.Name & "]" & vbCrLf
    Next ws
    ListSheetNames = s
End Function


'==============================================================================
' シートが存在するか確認
'==============================================================================
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function


'==============================================================================
' バックアップ作成
' 戻り値：成功=True、失敗=False
'==============================================================================
Private Function BackupDestinationFile() As Boolean
    Dim originalPath As String
    Dim folderPath As String
    Dim fileName As String
    Dim ext As String
    Dim dotPos As Long
    Dim backupPath As String

    On Error GoTo ErrHandler

    originalPath = g_dstFilePath
    folderPath = Left(originalPath, InStrRev(originalPath, "\"))
    fileName = Mid(originalPath, InStrRev(originalPath, "\") + 1)
    dotPos = InStrRev(fileName, ".")
    ext = Mid(fileName, dotPos)
    fileName = Left(fileName, dotPos - 1)

    backupPath = folderPath & fileName & "_バックアップ_" & _
                 Format(Now, "yyyymmdd_hhnn") & ext

    ' ファイルが他プロセス（Excel等）で開かれているかチェック
    If IsFileOpen(originalPath) Then
        Call WriteLog("バックアップ失敗：転記先ファイルが開かれています")
        MsgBox "転記先ファイルが開かれています。" & vbCrLf & _
               "Excelで開いている場合は閉じてから再実行してください。" & vbCrLf & vbCrLf & _
               "ファイル：" & originalPath, _
               vbExclamation, "バックアップ失敗"
        BackupDestinationFile = False
        Exit Function
    End If

    FileCopy originalPath, backupPath
    Call WriteLog("バックアップ作成：" & backupPath)
    BackupDestinationFile = True
    Exit Function

ErrHandler:
    Dim errMsg As String
    errMsg = "Err#" & Err.Number & " : " & Err.Description
    Call WriteLog("バックアップ失敗：" & errMsg)
    MsgBox "バックアップ作成に失敗しました。" & vbCrLf & vbCrLf & _
           errMsg & vbCrLf & vbCrLf & _
           "コピー元：" & originalPath & vbCrLf & _
           "コピー先：" & backupPath, _
           vbExclamation, "バックアップ失敗"
    BackupDestinationFile = False
End Function


'==============================================================================
' ファイルが他プロセスで開かれているか確認（補助）
'==============================================================================
Private Function IsFileOpen(filePath As String) As Boolean
    Dim ff As Integer
    Dim errNum As Long

    On Error Resume Next
    ff = FreeFile
    Open filePath For Binary Access Read Lock Read Write As #ff
    errNum = Err.Number
    Close #ff
    On Error GoTo 0

    ' エラー70（書き込みできません）または55（ファイルは既に開かれています）
    If errNum = 70 Or errNum = 55 Then
        IsFileOpen = True
    Else
        IsFileOpen = False
    End If
End Function


'==============================================================================
' コピー元データを読み取って配列に格納（多行ブロック対応版）
' 戻り値：明細件数
'
' 配列構造（dataArr）：
'   dataArr(i, 0) = 名称1（ブロック上の行）
'   dataArr(i, 1) = 名称2（アンカー行）
'   dataArr(i, 2) = 仕様1（ブロック上の行）
'   dataArr(i, 3) = 仕様2（アンカー行）
'   dataArr(i, 4) = 仕様3（ブロック下の行）
'   dataArr(i, 5) = 数量
'   dataArr(i, 6) = 単位
'   dataArr(i, 7) = 単価
'
' ブロック判定ルール：
'   - アンカー行：数量が数値 AND 単位に値あり
'   - 上の行：数量なし AND (名称または仕様に値) → ブロックに含める
'   - 下の行：数量なし AND 名称が空 AND 仕様に値 → ブロックに含める
'     （下の行に名称があれば次ブロックの始まりとみなして除外）
'==============================================================================
Private Function ReadSourceData(ws As Worksheet, ByRef dataArr As Variant) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cnt As Long
    Dim tempArr() As Variant
    Dim findResult As Range

    Set findResult = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious)
    If findResult Is Nothing Then
        ReadSourceData = 0
        Exit Function
    End If
    lastRow = findResult.Row

    ' 0:name1, 1:name2, 2:spec1, 3:spec2, 4:spec3, 5:qty, 6:unit, 7:price
    ReDim tempArr(1 To lastRow, 0 To 7)

    Dim qtyVal As Variant, unitVal As String
    Dim name1 As String, name2 As String
    Dim spec1 As String, spec2 As String, spec3 As String
    Dim priceVal As Variant
    Dim aboveQty As Variant, belowQty As Variant
    Dim aboveName As String, aboveSpec As String
    Dim belowName As String, belowSpec As String

    cnt = 0
    For r = g_srcStartRow To lastRow
        qtyVal = ws.Cells(r, g_srcColQty).Value
        unitVal = Trim(CStr(ws.Cells(r, g_srcColUnit).Value))

        ' アンカー判定：数量が数値 AND 単位に値あり
        If IsNumeric(qtyVal) And qtyVal <> "" And unitVal <> "" Then

            ' アンカー行の値
            name2 = Trim(CStr(ws.Cells(r, g_srcColName).Value))
            spec2 = Trim(CStr(ws.Cells(r, g_srcColSpec).Value))
            priceVal = ws.Cells(r, g_srcColPrice).Value

            ' 上の行をチェック
            name1 = "": spec1 = ""
            If r > g_srcStartRow Then
                aboveQty = ws.Cells(r - 1, g_srcColQty).Value
                If Not (IsNumeric(aboveQty) And aboveQty <> "") Then
                    aboveName = Trim(CStr(ws.Cells(r - 1, g_srcColName).Value))
                    aboveSpec = Trim(CStr(ws.Cells(r - 1, g_srcColSpec).Value))
                    If aboveName <> "" Or aboveSpec <> "" Then
                        name1 = aboveName
                        spec1 = aboveSpec
                    End If
                End If
            End If

            ' 下の行をチェック（名称が空で仕様だけある場合のみ採用）
            ' ただし、その次の行(r+2)がアンカーなら、下の行は次ブロックの上の行として
            ' 優先的に扱うため、こちらでは採用しない（重複防止）
            spec3 = ""
            If r < lastRow Then
                belowQty = ws.Cells(r + 1, g_srcColQty).Value
                If Not (IsNumeric(belowQty) And belowQty <> "") Then
                    belowName = Trim(CStr(ws.Cells(r + 1, g_srcColName).Value))
                    belowSpec = Trim(CStr(ws.Cells(r + 1, g_srcColSpec).Value))
                    If belowName = "" And belowSpec <> "" Then
                        ' 次の行(r+2)が次のアンカーかチェック
                        Dim nextQty As Variant
                        Dim nextUnit As String
                        Dim isNextAnchor As Boolean
                        isNextAnchor = False
                        If r + 2 <= lastRow Then
                            nextQty = ws.Cells(r + 2, g_srcColQty).Value
                            nextUnit = Trim(CStr(ws.Cells(r + 2, g_srcColUnit).Value))
                            If IsNumeric(nextQty) And nextQty <> "" _
                               And nextUnit <> "" Then
                                isNextAnchor = True
                            End If
                        End If
                        If Not isNextAnchor Then
                            spec3 = belowSpec
                        End If
                    End If
                End If
            End If

            ' 仕様4行目以上の検出（警告）
            ' spec3 を採用した上で、さらに r+2 に仕様データがあれば取りこぼし
            If spec3 <> "" And r + 2 <= lastRow Then
                Dim fourthQty As Variant
                Dim fourthName As String, fourthSpec As String
                fourthQty = ws.Cells(r + 2, g_srcColQty).Value
                fourthName = Trim(CStr(ws.Cells(r + 2, g_srcColName).Value))
                fourthSpec = Trim(CStr(ws.Cells(r + 2, g_srcColSpec).Value))
                If Not (IsNumeric(fourthQty) And fourthQty <> "") _
                   And fourthName = "" And fourthSpec <> "" Then
                    Call WriteLog("警告：行" & r & "の明細は仕様が4行以上あります。3行目以降は省略されました。")
                End If
            End If

            cnt = cnt + 1
            tempArr(cnt, 0) = name1
            tempArr(cnt, 1) = name2
            tempArr(cnt, 2) = spec1
            tempArr(cnt, 3) = spec2
            tempArr(cnt, 4) = spec3
            tempArr(cnt, 5) = qtyVal
            tempArr(cnt, 6) = unitVal
            tempArr(cnt, 7) = priceVal
        End If
    Next r

    If cnt = 0 Then
        ReadSourceData = 0
        Exit Function
    End If

    ' サイズ調整
    ReDim dataArr(1 To cnt, 0 To 7)
    Dim i As Long, j As Long
    For i = 1 To cnt
        For j = 0 To 7
            dataArr(i, j) = tempArr(i, j)
        Next j
    Next i

    ReadSourceData = cnt
End Function


'==============================================================================
' 転記先の既存データをクリア（値のみ、書式・数式は維持）
'==============================================================================
Private Sub ClearDestinationData(ws As Worksheet)
    Dim totalRow As Long
    Dim endRow As Long
    Dim r As Long

    totalRow = FindTotalRow(ws)
    If totalRow > 0 Then
        endRow = totalRow - 1
    Else
        ' 合計行が無い場合：データ最終行を使用
        Dim findResult As Range
        Set findResult = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                                        SearchDirection:=xlPrevious)
        If findResult Is Nothing Then Exit Sub
        endRow = findResult.Row
    End If

    ' 名称・仕様・数量・単位・単価の列を値のみクリア
    ' （金額列はIF式が入っているのでクリアしない）
    For r = g_dstStartRow To endRow
        Call SafeSetValue(ws.Cells(r, g_dstColName), "")
        Call SafeSetValue(ws.Cells(r, g_dstColSpec), "")
        Call SafeSetValue(ws.Cells(r, g_dstColQty), "")
        Call SafeSetValue(ws.Cells(r, g_dstColUnit), "")
        Call SafeSetValue(ws.Cells(r, g_dstColPrice), "")
    Next r
End Sub


'==============================================================================
' 合計行を検索
' 戻り値：合計行の行番号（見つからない場合は0）
'==============================================================================
Private Function FindTotalRow(ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Columns("A").Find(What:=g_sumKeyword, LookAt:=xlPart)
    If foundCell Is Nothing Then
        FindTotalRow = 0
    Else
        FindTotalRow = foundCell.Row
    End If
End Function


'==============================================================================
' 行数を確認し、不足している場合は3行ブロック単位で行を追加
' 戻り値：追加した行数（合計行が無い場合は0を返し処理スキップ）
'==============================================================================
Private Function EnsureRows(ws As Worksheet, neededBlocks As Long) As Long
    Dim totalRow As Long
    Dim currentBlocks As Long
    Dim shortBlocks As Long
    Dim addRows As Long
    Dim insertRow As Long
    Dim copySrc As Range

    totalRow = FindTotalRow(ws)
    If totalRow = 0 Then
        ' 合計行が無いシートは行追加・SUM更新ともスキップ
        Call WriteLog("合計行なし：行追加処理をスキップします")
        EnsureRows = 0
        Exit Function
    End If

    ' 現在のブロック数（合計行直前まで）
    currentBlocks = (totalRow - g_dstStartRow) \ 3
    shortBlocks = neededBlocks - currentBlocks

    If shortBlocks <= 0 Then
        EnsureRows = 0
        Exit Function
    End If

    addRows = shortBlocks * 3
    insertRow = totalRow  ' 合計行の直前に挿入

    ' 直上の3行ブロック（最後のブロック）をコピー元として使う
    Set copySrc = ws.Range(ws.Cells(totalRow - 3, 1), _
                            ws.Cells(totalRow - 1, ws.Columns.Count))

    Dim i As Long
    For i = 1 To shortBlocks
        ' 合計行の直前に3行ブロックを挿入
        copySrc.Copy
        ws.Rows(insertRow & ":" & (insertRow + 2)).Insert Shift:=xlDown
        ' 挿入後、元のcopySrcは下に押し下がっているので、新しい合計行を再取得
        totalRow = FindTotalRow(ws)
        ' 次の挿入位置を更新
        insertRow = totalRow
        ' コピー元範囲も再設定（最後のブロック）
        Set copySrc = ws.Range(ws.Cells(totalRow - 3, 1), _
                                ws.Cells(totalRow - 1, ws.Columns.Count))
    Next i

    Application.CutCopyMode = False

    ' 追加したブロックの値だけクリア（数式は残す）
    Dim r As Long
    For r = totalRow - addRows To totalRow - 1
        Call SafeSetValue(ws.Cells(r, g_dstColName), "")
        Call SafeSetValue(ws.Cells(r, g_dstColSpec), "")
        Call SafeSetValue(ws.Cells(r, g_dstColQty), "")
        Call SafeSetValue(ws.Cells(r, g_dstColUnit), "")
        Call SafeSetValue(ws.Cells(r, g_dstColPrice), "")
    Next r

    EnsureRows = addRows
End Function


'==============================================================================
' 転記先へデータを書き込む（多行ブロック対応版）
'
' 1明細 = 3行ブロック
'   行N   ：名称(空)、仕様1、数量、単位、単価
'   行N+1 ：名称1、仕様2
'   行N+2 ：名称2、仕様3
'
' 名称は1行下にずらして配置（理想レイアウトに合わせる）
' 数量・単位・単価はブロック1行目に書き込む（金額式が1行目を参照するため）
' 金額列（F列）には書き込まない（IF式が自動計算）
'
' 配列構造：
'   0:name1, 1:name2, 2:spec1, 3:spec2, 4:spec3, 5:qty, 6:unit, 7:price
'==============================================================================
Private Sub WriteToDestination(ws As Worksheet, dataArr As Variant, dataCount As Long)
    Dim i As Long
    Dim writeRow As Long

    writeRow = g_dstStartRow

    For i = 1 To dataCount
        ' 1行目：名称は空、仕様1、数量、単位、単価
        Call SafeSetValue(ws.Cells(writeRow, g_dstColSpec), dataArr(i, 2))   ' spec1
        Call SafeSetValue(ws.Cells(writeRow, g_dstColQty), dataArr(i, 5))    ' qty
        Call SafeSetValue(ws.Cells(writeRow, g_dstColUnit), dataArr(i, 6))   ' unit
        Call SafeSetValue(ws.Cells(writeRow, g_dstColPrice), dataArr(i, 7))  ' price

        ' 2行目：名称1、仕様2
        Call SafeSetValue(ws.Cells(writeRow + 1, g_dstColName), dataArr(i, 0))  ' name1
        Call SafeSetValue(ws.Cells(writeRow + 1, g_dstColSpec), dataArr(i, 3))  ' spec2

        ' 3行目：名称2、仕様3
        Call SafeSetValue(ws.Cells(writeRow + 2, g_dstColName), dataArr(i, 1))  ' name2
        Call SafeSetValue(ws.Cells(writeRow + 2, g_dstColSpec), dataArr(i, 4))  ' spec3

        ' 次のブロックへ
        writeRow = writeRow + 3
    Next i
End Sub


'==============================================================================
' セルに値を安全に書き込む（結合セル対応）
' 結合範囲の場合は左上のセルに書き込む
'==============================================================================
Private Sub SafeSetValue(c As Range, val As Variant)
    On Error Resume Next
    If c.MergeCells Then
        c.MergeArea.Cells(1, 1).Value = val
    Else
        c.Value = val
    End If
    On Error GoTo 0
End Sub


'==============================================================================
' SUM式を完全再生成
'
' 開始行から3行おきに金額列のセルアドレスを収集し、
' SUM式を組み立てて合計行の金額列に書き込む。
'
' 戻り値：生成したSUM式の文字列
'==============================================================================
Private Function RebuildSumFormula(ws As Worksheet) As String
    Dim totalRow As Long
    Dim r As Long
    Dim cellList As String
    Dim colLetter As String
    Dim formula As String

    totalRow = FindTotalRow(ws)
    If totalRow = 0 Then
        Call WriteLog("合計行なし：SUM式更新をスキップします")
        RebuildSumFormula = "(SUM式更新スキップ)"
        Exit Function
    End If

    ' 金額列の列文字を取得
    colLetter = NumToColLetter(g_dstColAmount)

    ' 開始行から3行おきにセルアドレスを収集
    cellList = ""
    For r = g_dstStartRow To totalRow - 1 Step 3
        If cellList <> "" Then cellList = cellList & ","
        cellList = cellList & colLetter & r
    Next r

    formula = "=SUM(" & cellList & ")"
    ws.Cells(totalRow, g_dstColAmount).formula = formula

    RebuildSumFormula = formula
End Function


'==============================================================================
' 列番号を列文字に変換（例：1→"A", 27→"AA"）
'==============================================================================
Private Function NumToColLetter(colNum As Long) As String
    Dim s As String
    Dim n As Long
    n = colNum
    s = ""
    Do While n > 0
        s = Chr(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    NumToColLetter = s
End Function


'==============================================================================
' ログを蓄積（メモリ上）
'==============================================================================
Private Sub WriteLog(msg As String)
    Dim ts As String
    ts = Format(Now, "yyyy/mm/dd hh:nn:ss")
    g_logBuffer = g_logBuffer & ts & " " & msg & vbCrLf
End Sub


'==============================================================================
' ログを「ログ」シートに書き出す
'==============================================================================
Private Sub FlushLog()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ログ")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "ログ"
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 1 Then lastRow = 1

    ws.Cells(lastRow, 1).Value = g_logBuffer
End Sub


'==============================================================================
' エラーメッセージ表示
'==============================================================================
Private Sub ShowError(msg As String)
    MsgBox msg, vbExclamation, "転記ツール"
    Call WriteLog("エラー：" & msg)
End Sub



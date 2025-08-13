Attribute VB_Name = "modInvoiceBuild"
Option Explicit

' Win32 API（必要なら残す）
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

'=== Globals ===
Public gExportFolderPath As String
'Public processingForm As Object   '未使用ならこのまま

'=== Config (共通設定) ===
Public Const TAX_RATE As Double = 0.1
Public Const DST_FIRST As Long = 15
Public Const DST_LAST  As Long = 272

'=== Sheet names ===
Public Const SH_SALES    As String = "売上台帳"
Public Const SH_TEMPLATE As String = "請求書フォーマット"
Public Const SH_LIST     As String = "請求書対象リスト"

'=== Column mapping ===
Public Const COL_SRC_DATE     As Long = 2   'B
Public Const COL_SRC_CLIENTID As Long = 8   'H
Public Const COL_PASTE_DATE   As Long = 2   'B
Public Const COL_PASTE_AMT    As Long = 6   'F
Public Const COL_PASTE_TAX    As Long = 7   'G

'=== H1'：2行目から連番を記載するバージョン ===
Sub H1_GetClientList()
    Dim wsList As Worksheet
    Dim lastRow As Long
    Dim i As Long, n As Long

    Set wsList = ThisWorkbook.Worksheets(SH_LIST)
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    wsList.Range("I2:I" & lastRow).ClearContents  'H2行目から初期化に変更
    n = 1
    For i = 2 To lastRow                          'H2行目から開始に変更
        If wsList.Cells(i, "A").Value <> "" Then
            wsList.Cells(i, "I").Value = Format(n, "00")
            n = n + 1
        End If
    Next i
End Sub
'=== H2（極小）：請求書対象リストA列(ClientID)で売上台帳H列をフィルタするだけ（Ver67-Min） ===

Public Sub H2_FilterByClientID_H(ByVal clientID As Variant)
    Dim wsSrc As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim fieldIdx As Long

    ' (1) シート取得
    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    ' (2) 売上台帳の最終行・最終列
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub

    ' (3) フィルタ範囲（A1～最終セル）
    Set rng = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))

    ' (4) 既存フィルタ解除
    On Error Resume Next
    If wsSrc.FilterMode Then wsSrc.ShowAllData
    wsSrc.AutoFilterMode = False
    On Error GoTo 0

    ' (5) H列（=8列目）で絞り込み
    fieldIdx = 8 - rng.Column + 1   'rng開始列がA想定
    rng.AutoFilter Field:=fieldIdx, Criteria1:=clientID
End Sub
'=== (前処理) 売上台帳のA列を mm/dd にする（ご要望どおりA列。B列も念のため） ===
Public Sub PrepSalesDateFormat()
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    On Error Resume Next
    wsSrc.Columns("A").NumberFormat = "mm/dd"   '★ご指定：A列を最初に書式変更
    wsSrc.Columns("B").NumberFormat = "mm/dd"   '★レイアウト差異対策（B列が日付の場合にも対応）
    On Error GoTo 0
End Sub

'=== (後処理) 税・集計・印刷範囲・罫線・日付を仕上げる ===
' 仕様：H2から独立
'  1) G15:G[明細末] に 行別消費税 = F×0.10（数式→値）
'  2) 「小計」「消費税合計」のラベル行を F/G の同じ行に配置、
'     その下の行に金額（F=小計金額、G=消費税合計金額）
'  3) C11 に 合計 = 小計金額 + 消費税合計金額
'  4) 印刷範囲：明細が 46 行目以下なら B1:G48、超える場合は B1:G[消費税合計金額の行]
'  5) 罫線は印刷範囲の下端まで
'  6) B列（日付）は明細～合計行まで mm/dd
Public Sub ApplyTaxTotalsAndPrintAdv(ByVal ws As Worksheet, ByVal DST_FIRST As Long, ByVal DST_LAST As Long)
    Dim pasteLast As Long                 ' 明細の最終行（B列基準）
    Dim sumLabelRow As Long               ' ラベル行（小計/消費税合計）
    Dim sumValueRow As Long               ' 金額行（小計金額/消費税合計金額）
    Dim printEndRow As Long               ' 印刷範囲の下端行

    ' (A) 明細の最終行（B列基準）
    pasteLast = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If pasteLast < DST_FIRST Then pasteLast = DST_FIRST - 1    ' 明細なし

    ' (B) ラベル・金額の行番号を決定
    If pasteLast >= DST_FIRST Then
        sumLabelRow = pasteLast + 1
    Else
        sumLabelRow = DST_FIRST
    End If
    sumValueRow = sumLabelRow + 1

    ' (C) B列：日付を mm/dd に
    If sumValueRow >= DST_FIRST Then
        ws.Range("B" & DST_FIRST & ":B" & sumValueRow).NumberFormat = "mm/dd"
    End If

    ' (D) G列：行別の消費税（F×0.10）を計算 → 値貼り
    If pasteLast >= DST_FIRST Then
        With ws.Range("G" & DST_FIRST & ":G" & pasteLast)
            .FormulaR1C1 = "=RC[-1]*" & CStr(TAX_RATE)
            .Value = .Value
        End With
    End If

    ' (E) 下段のラベルと金額
    ws.Range("F" & sumLabelRow).Value = "小計"
    ws.Range("G" & sumLabelRow).Value = "消費税合計"
    If pasteLast >= DST_FIRST Then
        ws.Range("F" & sumValueRow).Formula = "=SUM(F" & DST_FIRST & ":F" & pasteLast & ")"
        ws.Range("G" & sumValueRow).Formula = "=SUM(G" & DST_FIRST & ":G" & pasteLast & ")"
    Else
        ws.Range("F" & sumValueRow).Value = 0
        ws.Range("G" & sumValueRow).Value = 0
    End If

    ' (F) 請求書合計（ヘッダ C11）= 小計金額 + 消費税合計金額
    ws.Range("C11").Formula = "=F" & sumValueRow & "+G" & sumValueRow

    ' (G) 印刷範囲
    If pasteLast <= 46 Then
        printEndRow = 48                            ' 固定：B1:G48
    Else
        printEndRow = sumValueRow                   ' 消費税合計金額の行まで
    End If
    ws.PageSetup.PrintArea = "B1:G" & printEndRow

    ' (H) 罫線（B14～G：印刷範囲下端まで）
    With ws.Range("B14:G" & printEndRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub
Public Sub H3_CopyToInvoiceSheet(ByVal clientRow As Range)  'Ver64-CP2 + 追記(前/後処理)
    Const DST_FIRST As Long = 15, DST_LAST As Long = 272

    Dim wsSrc As Worksheet, wsDst As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, visCnt As Long, n As Long
    Dim renban As String

    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    Set wsDst = ThisWorkbook.Worksheets(SH_TEMPLATE)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ★(0)【前処理】売上台帳：A列の日付書式を mm/dd に（ご要望の通り）
    Call PrepSalesDateFormat

    ' ① H1 に対象行(A:I)を貼付け（テンプレ上）
    wsDst.Range("H1:P1").ClearContents
    clientRow.Resize(1, 9).Copy Destination:=wsDst.Range("H1")

    ' ② 明細（A:E 可視セル）→ B15（テンプレ上）
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    visCnt = WorksheetFunction.Subtotal(103, wsSrc.Range("A2:A" & lastRow))
    wsDst.Range("B" & DST_FIRST & ":F" & DST_LAST).ClearContents
    If visCnt > 0 Then
        n = WorksheetFunction.Min(visCnt, DST_LAST - DST_FIRST + 1)
        wsSrc.Range("A2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        wsDst.Range("B" & DST_FIRST).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        If DST_FIRST + visCnt - 1 > DST_LAST Then
            wsDst.Range("B" & (DST_FIRST + n) & ":F" & DST_LAST).ClearContents
        End If
    End If

    ' ③ テンプレを複製 → 連番で命名（I列想定）
    renban = Trim(clientRow.Cells(1, "I").Value)  ' 例：01, 02...
    If renban = "" Then renban = "請求書_" & Format(Now, "yyyymmdd_HHmmss")
    wsDst.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error Resume Next
    wsNew.Name = renban
    On Error GoTo 0

    ' ★(4)【後処理】新シートで ①日付(mm/dd) ②G列=行別消費税(F×0.1)
    '                 ③小計/消費税合計(ラベル行＋金額行)
    '                 ④印刷範囲(条件付き) ⑤罫線 を一括仕上げ
    Call ApplyTaxTotalsAndPrintAdv(wsNew, DST_FIRST, DST_LAST)

    ' ★(5)【任意】テンプレ汚れ防止：テンプレの明細をクリア（次回のコピー元をクリーンに）
    '    ※テンプレートに明細を残したくない場合だけ有効化してください
    'wsDst.Range("B" & DST_FIRST & ":G" & DST_LAST).ClearContents

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub H4_ExportAllPDFs() '(Ver62)
    Dim wsList As Worksheet
    Dim lastRow As Long, i As Long
    Dim clientName As String, renban As String
    Dim outputFolder As String
    Dim pdfName As String

    ' (1)▼シート「請求書対象リスト」の参照
    Set wsList = ThisWorkbook.Worksheets(SH_LIST)
    ' (2)▼保存先フォルダを作成（同じフォルダ内に「請求書PDF_yyyymmdd_HHmm」）
    Dim basePath As String
    basePath = ThisWorkbook.Path
    outputFolder = basePath & "\請求書PDF_" & Format(Now, "yyyymmdd_HHmm")
    
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' (3)▼最終行を取得
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    ' (4)▼ループしてPDF出力
    For i = 2 To lastRow
        renban = wsList.Cells(i, "I").Value
        clientName = wsList.Cells(i, "C").Value

        ' (4')▼連番・社名どちらか空ならスキップ
        If Trim(renban) = "" Or Trim(clientName) = "" Then GoTo SkipRow

        ' (4'')▼対象シートの存在確認（シート名＝連番）
        If Not SheetExists(renban) Then GoTo SkipRow

        ' (4''')▼PDFファイル名作成
        pdfName = outputFolder & "\" & renban & "_" & clientName & ".pdf"

        ' (4'''')▼PDF出力（1シートのみ）
        ThisWorkbook.Sheets(renban).ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pdfName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

SkipRow:
    Next i

    MsgBox "PDF出力が完了しました：" & vbCrLf & outputFolder, vbInformation
End Sub

'（2025/07/04 11:45）
Sub H5_MakeAllInvoices()
    Dim wsList As Worksheet
    Dim lastRow As Long, i As Long
    Dim clientName As String, renban As String
    Dim mark As String

    '▼対象：請求書対象リストシート
    Set wsList = ThisWorkbook.Worksheets(SH_LIST)

    '▼最終行（A列）
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow 'H2行目から処理
        If wsList.Cells(i, "A").Value <> "" Then

            renban = wsList.Cells(i, "I").Value
            clientName = wsList.Cells(i, "C").Value

            '★新規(1)：連番またはClientNameが空ならスキップ
            If Trim(renban) = "" Or Trim(clientName) = "" Then GoTo SkipRow

            '▼交互エフェクトで処理中を表示
            mark = IIf(i Mod 2 = 0, "?", "?")
            Application.StatusBar = mark & " 処理中：" & renban & " - " & clientName

            '▼処理を実行
            ' ▼処理を実行（ClientIDで抽出→テンプレ複製）
            Dim clientID As Variant
            clientID = wsList.Cells(i, "A").Value   '請求書対象リスト A列=ClientID
            Call H2_FilterByClientID_H(clientID)    '← フィルタ & 印刷/罫線/日付を事前適用
            Call H3_CopyToInvoiceSheet(wsList.Rows(i)) '← H1貼付→テンプレ複製→新シートにB:F貼付＆税/集計
            
           ' Call H4_ExportInvoicePDF(renban, clientName)

SkipRow:
        End If
    Next i

    Application.StatusBar = False
End Sub


Sub DeleteBlankRowsInSelection() 'スペース行削除
    Dim rng As Range
    Dim cell As Range
    Dim rowCheck As Range
    Dim i As Long

    On Error Resume Next
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "選択範囲が正しくありません。", vbExclamation
        Exit Sub
    End If

    ' 下から上にループ（上からだと行がずれて削除ミスの原因になります）
    For i = rng.Rows.Count To 1 Step -1
        Set rowCheck = rng.Rows(i)
        If Application.WorksheetFunction.CountA(rowCheck) = 0 Then
            rowCheck.EntireRow.Delete
        End If
    Next i
End Sub
'=== 入口：請求書作成→PDF出力まで一気通貫（既定でPDFも出す） ===
Public Sub Call_AllInvoiceMacros(Optional ByVal ExportPDF As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "請求書作成を開始します..."

    ' 1) 連番付与（請求書対象リスト）
    H1_GetClientList

    ' 2) 請求書作成（内部で H2→H3。H3で仕上げ&印刷範囲も設定済）
    Application.StatusBar = "請求書シートを作成中..."
    H5_MakeAllInvoices

    ' 3) 念のため：請求書シートのB列を mm/dd（既にH3で設定してても安全に再適用）
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left$(ws.Name, 4) = "請求書" Then
            On Error Resume Next
            ws.Columns("B").NumberFormat = "mm/dd"
            On Error GoTo 0
        End If
    Next

    ' 4) PDF出力（任意・既定=ON）
    If ExportPDF Then
        Application.StatusBar = "PDF出力中..."
        H4_ExportAllPDFs
    End If

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Call_AllInvoiceMacros でエラー：" & Err.Description, vbCritical
End Sub

'=== テスト用マクロ：Accessから呼び出して動作確認 ===
Public Sub TestMacro()
    MsgBox " おつかれさま！マクロは正常に動いています！", vbInformation, "テスト完了"
End Sub

' 進捗フォームを表示
Public Sub ShowUserForm()
    UserForm1.Show vbModeless
End Sub
Public Sub BlinkProgressLabel()
    With UserForm1
        If .Visible Then
            If .lblProgress.ForeColor = vbBlack Then
                .lblProgress.ForeColor = vbRed
            Else
                .lblProgress.ForeColor = vbBlack
            End If
            Application.OnTime Now + TimeValue("00:00:01"), "BlinkProgressLabel"
        End If
    End With
End Sub

Sub ListSheetNamesToSheet1()
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    Dim i As Integer
    
    ' Sheet1が存在するか確認
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo 0
    
    ' Sheet1が見つからない場合はエラーメッセージ
    If targetSheet Is Nothing Then
        MsgBox "シート名「Sheet1」が見つかりません。先に作成してください。", vbExclamation
        Exit Sub
    End If
    
    ' Sheet1の内容をクリア
    targetSheet.Cells.ClearContents
    
    ' シート名をA列に記載
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        targetSheet.Cells(i, 1).Value = ws.Name
        i = i + 1
    Next ws

    MsgBox "シート名の一覧を「Sheet1」に記載しました！", vbInformation
End Sub
Public Sub ShowBlinkingForm()
    With UserForm1
        .lblProgress.ForeColor = vbRed
        .Show vbModeless
        Application.OnTime Now + TimeValue("00:00:01"), "BlinkProgressLabel"
    End With
End Sub
Public Sub CloseProgressForm()
    On Error Resume Next
    Application.OnTime EarliestTime:=Now, Procedure:="BlinkProgressLabel", Schedule:=False
    Unload UserForm1
End Sub
' (Ver1) 指定されたシート名が存在するかどうか判定する関数
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet          ' 各ワークシートを格納するための変数を宣言
    SheetExists = False         ' 初期値として、シートは存在しないと仮定

    ' ThisWorkbook（このマクロが含まれるブック）内のすべてのシートを順に確認
    For Each ws In ThisWorkbook.Sheets
        ' シート名が一致するかどうかを判定
        If ws.Name = sheetName Then
            SheetExists = True  ' 一致するシートが見つかった場合、Trueを設定
            Exit Function       ' 処理を終了（以降のループは不要）
        End If
    Next ws                     ' 次のシートへ

End Function

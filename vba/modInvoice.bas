Attribute VB_Name = "modInvoiceBuild"
Option Explicit

' Win32 API�i�K�v�Ȃ�c���j
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

'=== Globals ===
Public gExportFolderPath As String
'Public processingForm As Object   '���g�p�Ȃ炱�̂܂�

'=== Config (���ʐݒ�) ===
Public Const TAX_RATE As Double = 0.1
Public Const DST_FIRST As Long = 15
Public Const DST_LAST  As Long = 272

'=== Sheet names ===
Public Const SH_SALES    As String = "����䒠"
Public Const SH_TEMPLATE As String = "�������t�H�[�}�b�g"
Public Const SH_LIST     As String = "�������Ώۃ��X�g"

'=== Column mapping ===
Public Const COL_SRC_DATE     As Long = 2   'B
Public Const COL_SRC_CLIENTID As Long = 8   'H
Public Const COL_PASTE_DATE   As Long = 2   'B
Public Const COL_PASTE_AMT    As Long = 6   'F
Public Const COL_PASTE_TAX    As Long = 7   'G

'=== H1'�F2�s�ڂ���A�Ԃ��L�ڂ���o�[�W���� ===
Sub H1_GetClientList()
    Dim wsList As Worksheet
    Dim lastRow As Long
    Dim i As Long, n As Long

    Set wsList = ThisWorkbook.Worksheets(SH_LIST)
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    wsList.Range("I2:I" & lastRow).ClearContents  'H2�s�ڂ��珉�����ɕύX
    n = 1
    For i = 2 To lastRow                          'H2�s�ڂ���J�n�ɕύX
        If wsList.Cells(i, "A").Value <> "" Then
            wsList.Cells(i, "I").Value = Format(n, "00")
            n = n + 1
        End If
    Next i
End Sub
'=== H2�i�ɏ��j�F�������Ώۃ��X�gA��(ClientID)�Ŕ���䒠H����t�B���^���邾���iVer67-Min�j ===

Public Sub H2_FilterByClientID_H(ByVal clientID As Variant)
    Dim wsSrc As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim fieldIdx As Long

    ' (1) �V�[�g�擾
    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    ' (2) ����䒠�̍ŏI�s�E�ŏI��
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub

    ' (3) �t�B���^�͈́iA1�`�ŏI�Z���j
    Set rng = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))

    ' (4) �����t�B���^����
    On Error Resume Next
    If wsSrc.FilterMode Then wsSrc.ShowAllData
    wsSrc.AutoFilterMode = False
    On Error GoTo 0

    ' (5) H��i=8��ځj�ōi�荞��
    fieldIdx = 8 - rng.Column + 1   'rng�J�n��A�z��
    rng.AutoFilter Field:=fieldIdx, Criteria1:=clientID
End Sub
'=== (�O����) ����䒠��A��� mm/dd �ɂ���i���v�]�ǂ���A��BB����O�̂��߁j ===
Public Sub PrepSalesDateFormat()
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    On Error Resume Next
    wsSrc.Columns("A").NumberFormat = "mm/dd"   '�����w��FA����ŏ��ɏ����ύX
    wsSrc.Columns("B").NumberFormat = "mm/dd"   '�����C�A�E�g���ّ΍�iB�񂪓��t�̏ꍇ�ɂ��Ή��j
    On Error GoTo 0
End Sub

'=== (�㏈��) �ŁE�W�v�E����͈́E�r���E���t���d�グ�� ===
' �d�l�FH2����Ɨ�
'  1) G15:G[���ז�] �� �s�ʏ���� = F�~0.10�i�������l�j
'  2) �u���v�v�u����ō��v�v�̃��x���s�� F/G �̓����s�ɔz�u�A
'     ���̉��̍s�ɋ��z�iF=���v���z�AG=����ō��v���z�j
'  3) C11 �� ���v = ���v���z + ����ō��v���z
'  4) ����͈́F���ׂ� 46 �s�ڈȉ��Ȃ� B1:G48�A������ꍇ�� B1:G[����ō��v���z�̍s]
'  5) �r���͈���͈͂̉��[�܂�
'  6) B��i���t�j�͖��ׁ`���v�s�܂� mm/dd
Public Sub ApplyTaxTotalsAndPrintAdv(ByVal ws As Worksheet, ByVal DST_FIRST As Long, ByVal DST_LAST As Long)
    Dim pasteLast As Long                 ' ���ׂ̍ŏI�s�iB���j
    Dim sumLabelRow As Long               ' ���x���s�i���v/����ō��v�j
    Dim sumValueRow As Long               ' ���z�s�i���v���z/����ō��v���z�j
    Dim printEndRow As Long               ' ����͈͂̉��[�s

    ' (A) ���ׂ̍ŏI�s�iB���j
    pasteLast = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If pasteLast < DST_FIRST Then pasteLast = DST_FIRST - 1    ' ���ׂȂ�

    ' (B) ���x���E���z�̍s�ԍ�������
    If pasteLast >= DST_FIRST Then
        sumLabelRow = pasteLast + 1
    Else
        sumLabelRow = DST_FIRST
    End If
    sumValueRow = sumLabelRow + 1

    ' (C) B��F���t�� mm/dd ��
    If sumValueRow >= DST_FIRST Then
        ws.Range("B" & DST_FIRST & ":B" & sumValueRow).NumberFormat = "mm/dd"
    End If

    ' (D) G��F�s�ʂ̏���ŁiF�~0.10�j���v�Z �� �l�\��
    If pasteLast >= DST_FIRST Then
        With ws.Range("G" & DST_FIRST & ":G" & pasteLast)
            .FormulaR1C1 = "=RC[-1]*" & CStr(TAX_RATE)
            .Value = .Value
        End With
    End If

    ' (E) ���i�̃��x���Ƌ��z
    ws.Range("F" & sumLabelRow).Value = "���v"
    ws.Range("G" & sumLabelRow).Value = "����ō��v"
    If pasteLast >= DST_FIRST Then
        ws.Range("F" & sumValueRow).Formula = "=SUM(F" & DST_FIRST & ":F" & pasteLast & ")"
        ws.Range("G" & sumValueRow).Formula = "=SUM(G" & DST_FIRST & ":G" & pasteLast & ")"
    Else
        ws.Range("F" & sumValueRow).Value = 0
        ws.Range("G" & sumValueRow).Value = 0
    End If

    ' (F) ���������v�i�w�b�_ C11�j= ���v���z + ����ō��v���z
    ws.Range("C11").Formula = "=F" & sumValueRow & "+G" & sumValueRow

    ' (G) ����͈�
    If pasteLast <= 46 Then
        printEndRow = 48                            ' �Œ�FB1:G48
    Else
        printEndRow = sumValueRow                   ' ����ō��v���z�̍s�܂�
    End If
    ws.PageSetup.PrintArea = "B1:G" & printEndRow

    ' (H) �r���iB14�`G�F����͈͉��[�܂Łj
    With ws.Range("B14:G" & printEndRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub
Public Sub H3_CopyToInvoiceSheet(ByVal clientRow As Range)  'Ver64-CP2 + �ǋL(�O/�㏈��)
    Const DST_FIRST As Long = 15, DST_LAST As Long = 272

    Dim wsSrc As Worksheet, wsDst As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, visCnt As Long, n As Long
    Dim renban As String

    Set wsSrc = ThisWorkbook.Worksheets(SH_SALES)
    Set wsDst = ThisWorkbook.Worksheets(SH_TEMPLATE)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ��(0)�y�O�����z����䒠�FA��̓��t������ mm/dd �Ɂi���v�]�̒ʂ�j
    Call PrepSalesDateFormat

    ' �@ H1 �ɑΏۍs(A:I)��\�t���i�e���v����j
    wsDst.Range("H1:P1").ClearContents
    clientRow.Resize(1, 9).Copy Destination:=wsDst.Range("H1")

    ' �A ���ׁiA:E ���Z���j�� B15�i�e���v����j
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

    ' �B �e���v���𕡐� �� �A�ԂŖ����iI��z��j
    renban = Trim(clientRow.Cells(1, "I").Value)  ' ��F01, 02...
    If renban = "" Then renban = "������_" & Format(Now, "yyyymmdd_HHmmss")
    wsDst.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error Resume Next
    wsNew.Name = renban
    On Error GoTo 0

    ' ��(4)�y�㏈���z�V�V�[�g�� �@���t(mm/dd) �AG��=�s�ʏ����(F�~0.1)
    '                 �B���v/����ō��v(���x���s�{���z�s)
    '                 �C����͈�(�����t��) �D�r�� ���ꊇ�d�グ
    Call ApplyTaxTotalsAndPrintAdv(wsNew, DST_FIRST, DST_LAST)

    ' ��(5)�y�C�Ӂz�e���v������h�~�F�e���v���̖��ׂ��N���A�i����̃R�s�[�����N���[���Ɂj
    '    ���e���v���[�g�ɖ��ׂ��c�������Ȃ��ꍇ�����L�������Ă�������
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

    ' (1)���V�[�g�u�������Ώۃ��X�g�v�̎Q��
    Set wsList = ThisWorkbook.Worksheets(SH_LIST)
    ' (2)���ۑ���t�H���_���쐬�i�����t�H���_���Ɂu������PDF_yyyymmdd_HHmm�v�j
    Dim basePath As String
    basePath = ThisWorkbook.Path
    outputFolder = basePath & "\������PDF_" & Format(Now, "yyyymmdd_HHmm")
    
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' (3)���ŏI�s���擾
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    ' (4)�����[�v����PDF�o��
    For i = 2 To lastRow
        renban = wsList.Cells(i, "I").Value
        clientName = wsList.Cells(i, "C").Value

        ' (4')���A�ԁE�Ж��ǂ��炩��Ȃ�X�L�b�v
        If Trim(renban) = "" Or Trim(clientName) = "" Then GoTo SkipRow

        ' (4'')���ΏۃV�[�g�̑��݊m�F�i�V�[�g�����A�ԁj
        If Not SheetExists(renban) Then GoTo SkipRow

        ' (4''')��PDF�t�@�C�����쐬
        pdfName = outputFolder & "\" & renban & "_" & clientName & ".pdf"

        ' (4'''')��PDF�o�́i1�V�[�g�̂݁j
        ThisWorkbook.Sheets(renban).ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pdfName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

SkipRow:
    Next i

    MsgBox "PDF�o�͂��������܂����F" & vbCrLf & outputFolder, vbInformation
End Sub

'�i2025/07/04 11:45�j
Sub H5_MakeAllInvoices()
    Dim wsList As Worksheet
    Dim lastRow As Long, i As Long
    Dim clientName As String, renban As String
    Dim mark As String

    '���ΏہF�������Ώۃ��X�g�V�[�g
    Set wsList = ThisWorkbook.Worksheets(SH_LIST)

    '���ŏI�s�iA��j
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow 'H2�s�ڂ��珈��
        If wsList.Cells(i, "A").Value <> "" Then

            renban = wsList.Cells(i, "I").Value
            clientName = wsList.Cells(i, "C").Value

            '���V�K(1)�F�A�Ԃ܂���ClientName����Ȃ�X�L�b�v
            If Trim(renban) = "" Or Trim(clientName) = "" Then GoTo SkipRow

            '�����݃G�t�F�N�g�ŏ�������\��
            mark = IIf(i Mod 2 = 0, "?", "?")
            Application.StatusBar = mark & " �������F" & renban & " - " & clientName

            '�����������s
            ' �����������s�iClientID�Œ��o���e���v�������j
            Dim clientID As Variant
            clientID = wsList.Cells(i, "A").Value   '�������Ώۃ��X�g A��=ClientID
            Call H2_FilterByClientID_H(clientID)    '�� �t�B���^ & ���/�r��/���t�����O�K�p
            Call H3_CopyToInvoiceSheet(wsList.Rows(i)) '�� H1�\�t���e���v���������V�V�[�g��B:F�\�t����/�W�v
            
           ' Call H4_ExportInvoicePDF(renban, clientName)

SkipRow:
        End If
    Next i

    Application.StatusBar = False
End Sub


Sub DeleteBlankRowsInSelection() '�X�y�[�X�s�폜
    Dim rng As Range
    Dim cell As Range
    Dim rowCheck As Range
    Dim i As Long

    On Error Resume Next
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "�I��͈͂�����������܂���B", vbExclamation
        Exit Sub
    End If

    ' �������Ƀ��[�v�i�ォ�炾�ƍs������č폜�~�X�̌����ɂȂ�܂��j
    For i = rng.Rows.Count To 1 Step -1
        Set rowCheck = rng.Rows(i)
        If Application.WorksheetFunction.CountA(rowCheck) = 0 Then
            rowCheck.EntireRow.Delete
        End If
    Next i
End Sub
'=== �����F�������쐬��PDF�o�͂܂ň�C�ʊсi�����PDF���o���j ===
Public Sub Call_AllInvoiceMacros(Optional ByVal ExportPDF As Boolean = True)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "�������쐬���J�n���܂�..."

    ' 1) �A�ԕt�^�i�������Ώۃ��X�g�j
    H1_GetClientList

    ' 2) �������쐬�i������ H2��H3�BH3�Ŏd�グ&����͈͂��ݒ�ρj
    Application.StatusBar = "�������V�[�g���쐬��..."
    H5_MakeAllInvoices

    ' 3) �O�̂��߁F�������V�[�g��B��� mm/dd�i����H3�Őݒ肵�ĂĂ����S�ɍēK�p�j
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left$(ws.Name, 4) = "������" Then
            On Error Resume Next
            ws.Columns("B").NumberFormat = "mm/dd"
            On Error GoTo 0
        End If
    Next

    ' 4) PDF�o�́i�C�ӁE����=ON�j
    If ExportPDF Then
        Application.StatusBar = "PDF�o�͒�..."
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
    MsgBox "Call_AllInvoiceMacros �ŃG���[�F" & Err.Description, vbCritical
End Sub

'=== �e�X�g�p�}�N���FAccess����Ăяo���ē���m�F ===
Public Sub TestMacro()
    MsgBox " �����ꂳ�܁I�}�N���͐���ɓ����Ă��܂��I", vbInformation, "�e�X�g����"
End Sub

' �i���t�H�[����\��
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
    
    ' Sheet1�����݂��邩�m�F
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo 0
    
    ' Sheet1��������Ȃ��ꍇ�̓G���[���b�Z�[�W
    If targetSheet Is Nothing Then
        MsgBox "�V�[�g���uSheet1�v��������܂���B��ɍ쐬���Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' Sheet1�̓��e���N���A
    targetSheet.Cells.ClearContents
    
    ' �V�[�g����A��ɋL��
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        targetSheet.Cells(i, 1).Value = ws.Name
        i = i + 1
    Next ws

    MsgBox "�V�[�g���̈ꗗ���uSheet1�v�ɋL�ڂ��܂����I", vbInformation
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
' (Ver1) �w�肳�ꂽ�V�[�g�������݂��邩�ǂ������肷��֐�
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet          ' �e���[�N�V�[�g���i�[���邽�߂̕ϐ���錾
    SheetExists = False         ' �����l�Ƃ��āA�V�[�g�͑��݂��Ȃ��Ɖ���

    ' ThisWorkbook�i���̃}�N�����܂܂��u�b�N�j���̂��ׂẴV�[�g�����Ɋm�F
    For Each ws In ThisWorkbook.Sheets
        ' �V�[�g������v���邩�ǂ����𔻒�
        If ws.Name = sheetName Then
            SheetExists = True  ' ��v����V�[�g�����������ꍇ�ATrue��ݒ�
            Exit Function       ' �������I���i�ȍ~�̃��[�v�͕s�v�j
        End If
    Next ws                     ' ���̃V�[�g��

End Function

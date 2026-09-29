Attribute VB_Name = "FormatRenewalSheets"
' ============================================================================
'  갱신보험료 납입 엑셀 - 인쇄용 서식 정리 매크로 (Python 없이 엑셀에서 바로 실행)
'
'  사용법
'   1) 엑셀 열기 → Alt + F11 (VBA 편집기)
'   2) 메뉴 [파일] > [파일 가져오기] 로 이 .bas 파일 가져오기
'      (또는 [삽입] > [모듈] 후 아래 내용 전체 복사/붙여넣기)
'   3) Alt + F8 →  FormatActiveWorkbook   : 지금 열려 있는 파일만 정리
'                  FormatFolder           : 폴더 안 .xlsx 전체 일괄 정리
'
'  * 값·수식은 건드리지 않고 서식만 바꿉니다.
'  * 일괄 실행 전에 폴더를 통째로 복사해 두는 것을 권장합니다 (되돌리기 불가).
' ============================================================================
Option Explicit

Private Const HEADER_BG As Long = 16182489      ' RGB(217, 232, 247) 밝은 하늘색
Private Const HEADER_FG As Long = 6373407       ' RGB(31, 56, 100)  진한 남색
Private Const LINE_COLOR As Long = 14732223     ' RGB(191, 206, 224) 연회색 테두리
Private Const BODY_FG As Long = 2500134         ' RGB(38, 38, 38)
Private Const AMOUNT_THRESHOLD As Double = 1000 ' 이 값 이상 숫자는 금액으로 보고 우측 정렬
Private Const MAX_COL_WIDTH As Double = 42


Public Sub FormatActiveWorkbook()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then FormatOneSheet ws
    Next ws
    Application.ScreenUpdating = True
    MsgBox "서식 정리 완료: " & ActiveWorkbook.Name, vbInformation
End Sub


Public Sub FormatFolder()
    Dim folderPath As String, fileName As String
    Dim wb As Workbook, ws As Worksheet, cnt As Long

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "갱신보험료 엑셀이 들어있는 폴더를 선택하세요"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        If Left$(fileName, 2) <> "~$" Then
            Set wb = Workbooks.Open(folderPath & fileName)
            For Each ws In wb.Worksheets
                If ws.Visible = xlSheetVisible Then FormatOneSheet ws
            Next ws
            wb.Close SaveChanges:=True
            cnt = cnt + 1
        End If
        fileName = Dir
    Loop

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox cnt & "개 파일 서식 정리 완료.", vbInformation
End Sub


Private Sub FormatOneSheet(ws As Worksheet)
    Dim rng As Range, c As Range, col As Range
    Dim lastRow As Long, lastCol As Long

    If ws.UsedRange Is Nothing Then Exit Sub
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' --- 전체 기본 서식 (가운데 정렬) -----------------------------------
    With rng
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Color = BODY_FG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = LINE_COLOR
    End With

    ' --- 1행 제목 --------------------------------------------------------
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = HEADER_BG
        .Font.Bold = True
        .Font.Color = HEADER_FG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30

    ' --- 금액 셀만 오른쪽 정렬 ------------------------------------------
    If lastRow >= 2 Then
        For Each c In ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
            If IsAmountCell(c) Then
                c.HorizontalAlignment = xlRight
                c.IndentLevel = 1
                c.WrapText = False
            End If
        Next c
    End If

    ' --- 열 너비 / 행 높이 ----------------------------------------------
    rng.Columns.AutoFit
    For Each col In rng.Columns
        If col.ColumnWidth > MAX_COL_WIDTH Then col.ColumnWidth = MAX_COL_WIDTH
        If col.ColumnWidth < 8 Then col.ColumnWidth = 8
    Next col

    ' --- 보기 / 인쇄 ------------------------------------------------------
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False

    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With

    ws.Range("A1").Select
End Sub


Private Function IsAmountCell(c As Range) As Boolean
    Dim v As Variant, s As String
    v = c.Value2
    If IsEmpty(v) Then Exit Function
    If IsDate(c.Value) Then Exit Function
    If InStr(c.NumberFormat, "%") > 0 Then Exit Function

    If IsNumeric(v) And Not VarType(v) = vbString Then
        If InStr(c.NumberFormat, "#,##") > 0 Then IsAmountCell = True: Exit Function
        IsAmountCell = (Abs(CDbl(v)) >= AMOUNT_THRESHOLD)
        Exit Function
    End If

    ' 문자열로 들어간 금액 ("1,234,000", "USD 5,000" 등)
    s = Trim$(CStr(v))
    If InStr(s, ",") > 0 Then
        s = Replace(Replace(Replace(Replace(s, ",", ""), "USD", ""), "원", ""), "$", "")
        IsAmountCell = IsNumeric(Trim$(s))
    End If
End Function

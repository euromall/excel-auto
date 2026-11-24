Attribute VB_Name = "ShipScan"
'==============================================================
' 출고 체크 매크로 v1 (2026-11-24)
' - 대상: 적하 시트(PTS_YYMM). 실행할 때 열려 있는 시트(ActiveSheet)를 사용
' - 송장번호(HAWB) 열 전체를 자동 검색 (범위 선택 불필요)
' - 스캔하면 출고일/스캔시각을 열에 기록 + 스캔로그 시트에 한 줄 추가
' - 기존 ScanCheckDuplicates 매크로는 건드리지 않음
'==============================================================
Option Explicit

' ---- 설정 (열 위치가 바뀌면 여기만 수정) ----
Private Const COL_HAWB As String = "B"        ' 송장번호(HAWB)
Private Const COL_DATE As String = "A"        ' 적하 작성일
Private Const COL_SENDER As String = "C"      ' 송하인
Private Const COL_ORDER As String = "F"       ' 관리번호
Private Const COL_RECV As String = "I"        ' 수령인
Private Const COL_ITEM As String = "AC"       ' 상품명
Private Const COL_QTY As String = "AF"        ' 수량
Private Const COL_SHIP As String = "CA"       ' [신규] 출고일
Private Const COL_SCANAT As String = "CB"     ' [신규] 스캔시각
Private Const COL_NOTE As String = "CC"       ' [신규] 출고차수/MAWB 메모
Private Const HEADER_ROW As Long = 5          ' 열 제목 행
Private Const FIRST_ROW As Long = 7           ' 데이터 시작 행
Private Const LOG_SHEET As String = "스캔로그"
Private Const VIEW_SHEET As String = "출고현황"

' 셀 값을 비교용 송장번호 문자열로 변환 (숫자/문자 혼재 대응)
Private Function HawbKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        HawbKey = Format$(v, "0")
    ElseIf VarType(v) = vbString Then
        HawbKey = Replace(Trim$(v), " ", "")
    End If
End Function

Private Function LastDataRow(ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, COL_HAWB).End(xlUp).Row
    If LastDataRow < FIRST_ROW Then LastDataRow = FIRST_ROW
End Function

Private Sub EnsureHeaders(ws As Worksheet)
    If Len(ws.Range(COL_SHIP & HEADER_ROW).Value) = 0 Then ws.Range(COL_SHIP & HEADER_ROW).Value = "출고일"
    If Len(ws.Range(COL_SCANAT & HEADER_ROW).Value) = 0 Then ws.Range(COL_SCANAT & HEADER_ROW).Value = "스캔시각"
    If Len(ws.Range(COL_NOTE & HEADER_ROW).Value) = 0 Then ws.Range(COL_NOTE & HEADER_ROW).Value = "출고메모"
End Sub

Private Function GetLog(wb As Workbook) As Worksheet
    Dim lg As Worksheet
    On Error Resume Next
    Set lg = wb.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If lg Is Nothing Then
        Set lg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        lg.Name = LOG_SHEET
        lg.Range("A1:H1").Value = Array("스캔시각", "시트", "송장번호", "결과", "출고일", "행", "출고메모", "사용자")
        lg.Rows(1).Font.Bold = True
    End If
    Set GetLog = lg
End Function

Private Sub WriteLog(lg As Worksheet, ws As Worksheet, ByVal hawb As String, ByVal result As String, _
                     ByVal shipDate As Variant, ByVal rowList As String, ByVal note As String)
    Dim r As Long
    r = lg.Cells(lg.Rows.Count, "A").End(xlUp).Row + 1
    lg.Cells(r, 1).Value = Now
    lg.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    lg.Cells(r, 2).Value = ws.Name
    lg.Cells(r, 3).NumberFormat = "@"
    lg.Cells(r, 3).Value = hawb
    lg.Cells(r, 4).Value = result
    If IsDate(shipDate) Then
        lg.Cells(r, 5).Value = CDate(shipDate)
        lg.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
    End If
    lg.Cells(r, 6).NumberFormat = "@"
    lg.Cells(r, 6).Value = rowList
    lg.Cells(r, 7).Value = note
    lg.Cells(r, 8).Value = Application.UserName
End Sub

' HAWB -> 행 번호 목록(Collection) 사전
Private Function BuildIndex(ws As Worksheet) As Object
    Dim idx As Object, data As Variant, i As Long, k As String, lastRow As Long
    Set idx = CreateObject("Scripting.Dictionary")
    lastRow = LastDataRow(ws)
    If lastRow < FIRST_ROW Then Set BuildIndex = idx: Exit Function
    data = ws.Range(COL_HAWB & FIRST_ROW & ":" & COL_HAWB & lastRow).Value2
    If Not IsArray(data) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = data
        data = one
    End If
    For i = 1 To UBound(data, 1)
        k = HawbKey(data(i, 1))
        If Len(k) >= 6 Then
            If Not idx.Exists(k) Then idx.Add k, New Collection
            idx.Item(k).Add FIRST_ROW + i - 1
        End If
    Next i
    Set BuildIndex = idx
End Function

Private Function AskDate(ByVal prompt As String) As Variant
    Dim s As Variant
    s = Application.InputBox(prompt & vbCrLf & "(형식: 2026-11-24)", "출고일", Format$(Date, "yyyy-mm-dd"), Type:=2)
    If VarType(s) = vbBoolean Then AskDate = Empty: Exit Function
    If Not IsDate(s) Then
        MsgBox "날짜 형식이 아닙니다: " & s, vbExclamation
        AskDate = Empty
        Exit Function
    End If
    AskDate = CDate(s)
End Function

'==============================================================
' 1) 출고 스캔: 송장 바코드를 스캔하면 출고일 기록
'==============================================================
Public Sub ShipScan_Start()
    Dim ws As Worksheet, lg As Worksheet, idx As Object
    Dim shipDate As Variant, note As Variant, raw As Variant, key As String
    Dim rowsCol As Collection, rw As Variant, rowList As String, firstRow As Long
    Dim already As Variant, msg As String, nOk As Long, nDup As Long, nMiss As Long
    Dim pLeft As Double, pTop As Double

    Set ws = ActiveSheet
    If MsgBox("이 시트에서 출고 스캔을 시작합니다: " & ws.Name, vbOKCancel Or vbQuestion, "출고 스캔") = vbCancel Then Exit Sub
    shipDate = AskDate("출고일을 입력하세요.")
    If IsEmpty(shipDate) Then Exit Sub
    note = Application.InputBox("출고메모(출고차수, MAWB 등). 없으면 비워 두고 확인.", "출고메모", "", Type:=2)
    If VarType(note) = vbBoolean Then note = ""

    EnsureHeaders ws
    Set lg = GetLog(ws.Parent)
    Set idx = BuildIndex(ws)
    If idx.Count = 0 Then MsgBox COL_HAWB & "열에 송장번호가 없습니다.", vbExclamation: Exit Sub

    pLeft = Application.Left + Application.Width * 0.55
    pTop = Application.Top + Application.Height * 0.5
    msg = "출고일 " & Format$(shipDate, "yyyy-mm-dd") & " | 송장 바코드를 스캔하세요."
    Do
        raw = Application.InputBox(msg & vbCrLf & "종료: Esc", "출고 스캔 (" & ws.Name & ")", Left:=pLeft, Top:=pTop, Type:=2)
        If VarType(raw) = vbBoolean Then Exit Do
        key = HawbKey(CStr(raw))
        If Len(key) = 0 Then GoTo NextScan

        If Not idx.Exists(key) Then
            nMiss = nMiss + 1
            Beep
            WriteLog lg, ws, key, "미일치", shipDate, "", CStr(note)
            msg = "[미일치] " & key & " - 이 시트에 없는 송장번호입니다." & vbCrLf & "다음 송장을 스캔하세요."
            GoTo NextScan
        End If

        Set rowsCol = idx.Item(key)
        firstRow = rowsCol.Item(1)
        rowList = ""
        For Each rw In rowsCol
            rowList = rowList & IIf(Len(rowList) > 0, ",", "") & CStr(rw)
        Next rw
        already = ws.Range(COL_SHIP & firstRow).Value

        If Not IsEmpty(already) And Len(CStr(already)) > 0 Then
            nDup = nDup + 1
            Beep
            WriteLog lg, ws, key, "중복", shipDate, rowList, "기존 출고일 " & CStr(already)
            Application.Goto ws.Range(COL_HAWB & firstRow), True
            If MsgBox("이미 출고 처리된 송장입니다." & vbCrLf & vbCrLf & _
                      "송장번호: " & key & vbCrLf & "기존 출고일: " & Format$(already, "yyyy-mm-dd") & vbCrLf & _
                      "행: " & rowList & vbCrLf & vbCrLf & "확인: 계속 스캔    취소: 종료", _
                      vbExclamation Or vbOKCancel, "중복 스캔") = vbCancel Then Exit Do
            msg = "[중복] " & key & " (기존 " & Format$(already, "yyyy-mm-dd") & ")" & vbCrLf & "다음 송장을 스캔하세요."
            GoTo NextScan
        End If

        For Each rw In rowsCol
            With ws.Range(COL_SHIP & rw)
                .Value = shipDate
                .NumberFormat = "yyyy-mm-dd"
            End With
            With ws.Range(COL_SCANAT & rw)
                .Value = Now
                .NumberFormat = "yyyy-mm-dd hh:mm"
            End With
            If Len(CStr(note)) > 0 Then ws.Range(COL_NOTE & rw).Value = note
            ws.Range(COL_HAWB & rw).Interior.Color = RGB(198, 239, 206)
            ws.Range(COL_SHIP & rw).Interior.Color = RGB(198, 239, 206)
        Next rw
        nOk = nOk + 1
        WriteLog lg, ws, key, "출고", shipDate, rowList, CStr(note)
        Application.Goto ws.Range(COL_HAWB & firstRow), True
        msg = "[출고] " & key & " | 행 " & rowList & " | " & CStr(ws.Range(COL_ORDER & firstRow).Value) & vbCrLf & _
              "이번 작업: 출고 " & nOk & " / 중복 " & nDup & " / 미일치 " & nMiss & vbCrLf & "다음 송장을 스캔하세요."
NextScan:
    Loop
    MsgBox "스캔 종료" & vbCrLf & "출고 " & nOk & "건 / 중복 " & nDup & "건 / 미일치 " & nMiss & "건" & vbCrLf & _
           "기록: '" & LOG_SHEET & "' 시트", vbInformation, "출고 스캔"
End Sub

'==============================================================
' 2) 출고 취소: 선택한 행의 출고일/스캔시각/메모 지우기
'==============================================================
Public Sub ShipScan_Cancel()
    Dim ws As Worksheet, lg As Worksheet, area As Range, rw As Range, n As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    If MsgBox("선택한 행의 출고 기록을 지웁니다. 계속할까요?", vbOKCancel Or vbQuestion, "출고 취소") = vbCancel Then Exit Sub
    Set lg = GetLog(ws.Parent)
    Dim sel As Range
    Set sel = Application.Intersect(Selection, ws.UsedRange)
    If sel Is Nothing Then Exit Sub
    For Each area In sel.Areas
        For Each rw In area.Rows
            If rw.Row >= FIRST_ROW Then
                If Len(CStr(ws.Range(COL_SHIP & rw.Row).Value)) > 0 Then
                    WriteLog lg, ws, HawbKey(ws.Range(COL_HAWB & rw.Row).Value), "취소", ws.Range(COL_SHIP & rw.Row).Value, CStr(rw.Row), ""
                    ws.Range(COL_SHIP & rw.Row & ":" & COL_NOTE & rw.Row).ClearContents
                    If ws.Range(COL_SHIP & rw.Row).Interior.Color = RGB(198, 239, 206) Then ws.Range(COL_SHIP & rw.Row).Interior.Pattern = xlNone
                    If ws.Range(COL_HAWB & rw.Row).Interior.Color = RGB(198, 239, 206) Then ws.Range(COL_HAWB & rw.Row).Interior.Pattern = xlNone
                    n = n + 1
                End If
            End If
        Next rw
    Next area
    MsgBox n & "행의 출고 기록을 지웠습니다.", vbInformation
End Sub

'==============================================================
' 3) 출고일 일괄 입력: 밀린 분 처리용 (선택한 행 중 출고일이 빈 행만)
'==============================================================
Public Sub ShipScan_BulkDate()
    Dim ws As Worksheet, lg As Worksheet, area As Range, rw As Range
    Dim shipDate As Variant, note As Variant, n As Long, k As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    shipDate = AskDate("선택한 행에 넣을 출고일을 입력하세요." & vbCrLf & "(출고일이 이미 있는 행은 건너뜀)")
    If IsEmpty(shipDate) Then Exit Sub
    note = Application.InputBox("출고메모(예: 일괄입력, MAWB). 없으면 비워 두고 확인.", "출고메모", "일괄입력", Type:=2)
    If VarType(note) = vbBoolean Then note = "일괄입력"
    EnsureHeaders ws
    Set lg = GetLog(ws.Parent)
    Dim sel As Range
    Set sel = Application.Intersect(Selection, ws.UsedRange)
    If sel Is Nothing Then Exit Sub
    For Each area In sel.Areas
        For Each rw In area.Rows
            If rw.Row >= FIRST_ROW Then
                k = HawbKey(ws.Range(COL_HAWB & rw.Row).Value)
                If Len(k) >= 6 And Len(CStr(ws.Range(COL_SHIP & rw.Row).Value)) = 0 Then
                    ws.Range(COL_SHIP & rw.Row).Value = shipDate
                    ws.Range(COL_SHIP & rw.Row).NumberFormat = "yyyy-mm-dd"
                    ws.Range(COL_NOTE & rw.Row).Value = note
                    WriteLog lg, ws, k, "일괄입력", shipDate, CStr(rw.Row), CStr(note)
                    n = n + 1
                End If
            End If
        Next rw
    Next area
    MsgBox n & "행에 출고일 " & Format$(shipDate, "yyyy-mm-dd") & "을(를) 입력했습니다.", vbInformation
End Sub

'==============================================================
' 4) 출고현황 시트 만들기: 출고일 순 정렬, 미출고는 맨 아래
'    (원본 적하 시트의 행 순서는 바꾸지 않음)
'==============================================================
Public Sub ShipScan_BuildView()
    Dim src As Worksheet, vw As Worksheet, lastRow As Long, i As Long, r As Long, n As Long
    Dim cols As Variant, heads As Variant, c As Long, k As String
    Dim hawbArr As Variant, out() As Variant, colData() As Variant
    Dim nShip As Long, nWait As Long

    Set src = ActiveSheet
    If src.Name = VIEW_SHEET Or src.Name = LOG_SHEET Then
        MsgBox "적하 시트(PTS_...)를 연 상태에서 실행하세요.", vbExclamation: Exit Sub
    End If
    cols = Array(COL_SHIP, COL_DATE, COL_HAWB, COL_ORDER, COL_SENDER, COL_RECV, COL_ITEM, COL_QTY, COL_NOTE, COL_SCANAT)
    heads = Array("출고일", "작성일", "송장번호", "관리번호", "송하인", "수령인", "상품명", "수량", "출고메모", "스캔시각")
    lastRow = LastDataRow(src)
    If lastRow <= FIRST_ROW Then MsgBox "송장번호가 있는 행이 없습니다.", vbExclamation: Exit Sub
    hawbArr = src.Range(COL_HAWB & FIRST_ROW & ":" & COL_HAWB & lastRow).Value2

    ReDim colData(0 To UBound(cols))
    For c = 0 To UBound(cols)
        colData(c) = src.Range(cols(c) & FIRST_ROW & ":" & cols(c) & lastRow).Value2
    Next c

    ' 송장번호가 있는 행 수 세기
    For i = 1 To UBound(hawbArr, 1)
        If Len(HawbKey(hawbArr(i, 1))) >= 6 Then n = n + 1
    Next i
    If n = 0 Then MsgBox "송장번호가 있는 행이 없습니다.", vbExclamation: Exit Sub

    ReDim out(1 To n, 1 To UBound(cols) + 2)
    r = 0
    For i = 1 To UBound(hawbArr, 1)
        k = HawbKey(hawbArr(i, 1))
        If Len(k) >= 6 Then
            r = r + 1
            For c = 0 To UBound(cols)
                out(r, c + 1) = colData(c)(i, 1)
            Next c
            out(r, 3) = k
            out(r, UBound(cols) + 2) = FIRST_ROW + i - 1   ' 원본 행
            If IsEmpty(out(r, 1)) Or Len(CStr(out(r, 1))) = 0 Then nWait = nWait + 1 Else nShip = nShip + 1
        End If
    Next i

    Application.ScreenUpdating = False
    On Error Resume Next
    Set vw = src.Parent.Worksheets(VIEW_SHEET)
    On Error GoTo 0
    If vw Is Nothing Then
        Set vw = src.Parent.Worksheets.Add(After:=src.Parent.Worksheets(src.Parent.Worksheets.Count))
        vw.Name = VIEW_SHEET
    End If
    If vw.AutoFilterMode Then vw.AutoFilterMode = False
    vw.Cells.Clear
    vw.Range("A1").Value = "원본: " & src.Name & " | 만든 시각: " & Format$(Now, "yyyy-mm-dd hh:mm") & _
                           " | 출고 " & nShip & "행 / 미출고 " & nWait & "행"
    For c = 0 To UBound(heads)
        vw.Cells(2, c + 1).Value = heads(c)
    Next c
    vw.Cells(2, UBound(heads) + 2).Value = "원본행"
    vw.Rows(2).Font.Bold = True
    vw.Range("C3").Resize(n, 1).NumberFormat = "@"
    vw.Range("A3").Resize(n, UBound(cols) + 2).Value = out
    vw.Range("A3").Resize(n, 2).NumberFormat = "yyyy-mm-dd"
    vw.Range("J3").Resize(n, 1).NumberFormat = "yyyy-mm-dd hh:mm"

    ' 정렬: 출고일 오름차순(빈칸은 Excel이 항상 맨 아래로 보냄) → 작성일 → 원본행
    With vw.Sort
        .SortFields.Clear
        .SortFields.Add Key:=vw.Range("A3").Resize(n, 1), Order:=xlAscending
        .SortFields.Add Key:=vw.Range("B3").Resize(n, 1), Order:=xlAscending
        .SortFields.Add Key:=vw.Cells(3, UBound(cols) + 2).Resize(n, 1), Order:=xlAscending
        .SetRange vw.Range("A2").Resize(n + 1, UBound(cols) + 2)
        .Header = xlYes
        .Apply
    End With

    ' 미출고 행 표시
    For i = 3 To n + 2
        If Len(CStr(vw.Cells(i, 1).Value)) = 0 Then
            vw.Cells(i, 1).Value = "미출고"
            vw.Range(vw.Cells(i, 1), vw.Cells(i, UBound(cols) + 2)).Interior.Color = RGB(255, 235, 156)
        End If
    Next i
    vw.Range("A2").Resize(1, UBound(cols) + 2).AutoFilter
    vw.Columns("A:K").AutoFit
    vw.Columns("G").ColumnWidth = 60
    Application.ScreenUpdating = True
    vw.Activate
    vw.Range("A3").Select
    MsgBox "출고현황을 만들었습니다." & vbCrLf & "출고 " & nShip & "행 / 미출고 " & nWait & "행", vbInformation
End Sub

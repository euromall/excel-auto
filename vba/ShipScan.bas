Attribute VB_Name = "ShipScan"
'==============================================================
' 출고 체크 매크로 v3 (2026-11-28)
' - 대상: 실행할 때 열려 있는 적하 시트(PTS_YYMM)
' - 먼저 출고할 행들을 선택(여러 행, Ctrl로 떨어진 구간도 가능)한 뒤 실행
' - 선택한 행 안에서만 송장번호(B열)를 찾음. 선택 밖은 처리하지 않음
' - 스캔되면: A~C열(송장과 양옆)만 연두색 + AI열에 출고일. 나머지 열의 기존 색은 그대로
' - AI열에 이미 값이 있는 행은 덮어쓰지 않음 (출고일이면 중복, 다른 값이면 건너뜀)
' - 별도 기록 시트는 만들지 않음. 기존 ScanCheckDuplicates 매크로는 그대로
'==============================================================
Option Explicit

' ---- 설정 (열 위치가 바뀌면 여기만 수정) ----
Private Const COL_HAWB As String = "B"        ' 송장번호(HAWB)
Private Const COL_SHIP As String = "AI"       ' 출고일을 넣는 열
Private Const MARK_FROM As String = "A"       ' 색칠 시작 열
Private Const MARK_TO As String = "C"         ' 색칠 끝 열
Private Const FIRST_ROW As Long = 7           ' 데이터 시작 행
Private Const SHIP_COLOR As Long = 13434828   ' 연두색 RGB(204,255,204)
Private Const DATE_FORMAT As String = "yyyy-mm-dd"

' 셀 값을 비교용 송장번호 문자열로 변환 (숫자/문자 혼재 대응)
Private Function HawbKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbLong, vbInteger, vbCurrency, vbDecimal
            HawbKey = Format$(v, "0")
        Case vbString
            HawbKey = Replace(Trim$(v), " ", "")
    End Select
End Function

' 선택 영역의 행 번호(데이터 행만, 중복 없이) -> Dictionary(행번호 -> True)
Private Function SelectedRows(ws As Worksheet) As Object
    Dim d As Object, area As Range, r As Long, lastRow As Long
    Set d = CreateObject("Scripting.Dictionary")
    If TypeName(Selection) <> "Range" Then Set SelectedRows = d: Exit Function
    If Not Selection.Worksheet Is ws Then Set SelectedRows = d: Exit Function
    lastRow = ws.Cells(ws.Rows.Count, COL_HAWB).End(xlUp).Row
    For Each area In Selection.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            If r > lastRow Then Exit For            ' 행 전체 선택 시 빈 아래쪽은 건너뜀
            If r >= FIRST_ROW Then d(r) = True
        Next r
    Next area
    Set SelectedRows = d
End Function

' 선택한 행 안에서만: 송장번호 -> 행 번호 목록(Collection)
Private Function BuildIndex(ws As Worksheet, selRows As Object) As Object
    Dim idx As Object, r As Variant, k As String
    Set idx = CreateObject("Scripting.Dictionary")
    For Each r In selRows.Keys
        k = HawbKey(ws.Range(COL_HAWB & r).Value2)
        If Len(k) >= 6 Then
            If Not idx.Exists(k) Then idx.Add k, New Collection
            idx.Item(k).Add CLng(r)
        End If
    Next r
    Set BuildIndex = idx
End Function

Private Function NeedRows(ws As Worksheet) As Object
    Dim selRows As Object
    Set selRows = SelectedRows(ws)
    If selRows.Count < 2 Then
        MsgBox "출고할 행을 먼저 선택하세요." & vbCrLf & _
               "(오늘 블록처럼 여러 행을 선택. 떨어진 구간은 Ctrl을 누른 채 선택)", vbExclamation, "출고 체크"
        Set NeedRows = Nothing
    Else
        Set NeedRows = selRows
    End If
End Function

Private Function AskDate(ByVal prompt As String) As Variant
    Dim s As Variant
    s = Application.InputBox(prompt & vbCrLf & "(형식: 2026-11-27)", "출고일", Format$(Date, DATE_FORMAT), Type:=2)
    If VarType(s) = vbBoolean Then AskDate = Empty: Exit Function
    If Not IsDate(s) Then
        MsgBox "날짜 형식이 아닙니다: " & s, vbExclamation
        AskDate = Empty
        Exit Function
    End If
    AskDate = CDate(s)
End Function

' AI열 상태: 0 = 비어 있음, 1 = 출고일 있음, 2 = 다른 값 있음
Private Function ShipState(ws As Worksheet, ByVal r As Long) As Integer
    Dim v As Variant
    v = ws.Range(COL_SHIP & r).Value
    If IsEmpty(v) Then
        ShipState = 0
    ElseIf VarType(v) = vbString And Len(Trim$(CStr(v))) = 0 Then
        ShipState = 0
    ElseIf VarType(v) = vbDate Then
        ShipState = 1
    Else
        ShipState = 2
    End If
End Function

' 색칠할 칸 (예: A7:C7)
Private Function MarkRange(ws As Worksheet, ByVal r As Long) As Range
    Set MarkRange = ws.Range(MARK_FROM & r & ":" & MARK_TO & r)
End Function

Private Sub MarkShipped(ws As Worksheet, ByVal r As Long, ByVal shipDate As Date)
    MarkRange(ws, r).Interior.Color = SHIP_COLOR
    With ws.Range(COL_SHIP & r)
        .NumberFormat = DATE_FORMAT
        .Value = shipDate
    End With
End Sub

'==============================================================
' 1) 출고 스캔: 선택한 행 안에서 송장 바코드를 찾아 출고 표시
'==============================================================
Public Sub ShipScan_Start()
    Dim ws As Worksheet, selRows As Object, idx As Object
    Dim shipDate As Variant, raw As Variant, key As String
    Dim rw As Variant, rowList As String, state As Integer
    Dim nDone As Long, nShip As Long, nDup As Long, nOther As Long, nMiss As Long
    Dim msg As String, pLeft As Double, pTop As Double

    Set ws = ActiveSheet
    Set selRows = NeedRows(ws)
    If selRows Is Nothing Then Exit Sub
    Set idx = BuildIndex(ws, selRows)
    If idx.Count = 0 Then MsgBox "선택한 행의 " & COL_HAWB & "열에 송장번호가 없습니다.", vbExclamation: Exit Sub
    If MsgBox(ws.Name & " 시트, 선택한 " & selRows.Count & "행(송장 " & idx.Count & "개)에서 출고 스캔을 시작합니다.", _
              vbOKCancel Or vbQuestion, "출고 스캔") = vbCancel Then Exit Sub
    shipDate = AskDate("출고일을 입력하세요.")
    If IsEmpty(shipDate) Then Exit Sub

    pLeft = Application.Left + Application.Width * 0.55
    pTop = Application.Top + Application.Height * 0.5
    msg = "출고일 " & Format$(shipDate, DATE_FORMAT) & " | 송장 바코드를 스캔하세요."
    Do
        raw = Application.InputBox(msg & vbCrLf & "진행: " & nDone & " / " & idx.Count & "   종료: Esc", _
                                   "출고 스캔 (" & ws.Name & ")", Left:=pLeft, Top:=pTop, Type:=2)
        If VarType(raw) = vbBoolean Then Exit Do
        key = HawbKey(CStr(raw))
        If Len(key) = 0 Then GoTo NextScan

        If Not idx.Exists(key) Then
            nMiss = nMiss + 1
            Beep
            msg = "[선택 범위에 없음] " & key & " - 처리하지 않았습니다."
            GoTo NextScan
        End If

        rowList = ""
        nShip = 0
        For Each rw In idx.Item(key)
            rowList = rowList & IIf(Len(rowList) > 0, ",", "") & rw
            state = ShipState(ws, CLng(rw))
            If state = 0 Then
                MarkShipped ws, CLng(rw), CDate(shipDate)
                nShip = nShip + 1
            ElseIf state = 1 Then
                nDup = nDup + 1
            Else
                nOther = nOther + 1
            End If
        Next rw

        If nShip > 0 Then
            nDone = nDone + 1
            msg = "[출고] " & key & " → " & rowList & "행"
            Application.Goto ws.Range(COL_HAWB & idx.Item(key)(1)), False
        ElseIf ShipState(ws, CLng(idx.Item(key)(1))) = 1 Then
            Beep
            msg = "[중복] " & key & " - 이미 출고 " & Format$(ws.Range(COL_SHIP & idx.Item(key)(1)).Value, DATE_FORMAT)
        Else
            Beep
            msg = "[건너뜀] " & key & " - " & COL_SHIP & "열에 다른 값이 있어 덮어쓰지 않았습니다 (" & rowList & "행)"
        End If
NextScan:
    Loop

    MsgBox "출고 처리: 송장 " & nDone & "개" & vbCrLf & _
           "중복(이미 출고): " & nDup & "행" & vbCrLf & _
           COL_SHIP & "열 다른 값으로 건너뜀: " & nOther & "행" & vbCrLf & _
           "선택 범위에 없음: " & nMiss & "건", vbInformation, "출고 스캔 종료"
End Sub

'==============================================================
' 2) 밀린 분: 선택한 행 전체를 같은 출고일로 처리
'==============================================================
Public Sub ShipScan_BulkDate()
    Dim ws As Worksheet, selRows As Object, shipDate As Variant, r As Variant
    Dim nShip As Long, nDup As Long, nOther As Long, nNoHawb As Long

    Set ws = ActiveSheet
    Set selRows = NeedRows(ws)
    If selRows Is Nothing Then Exit Sub
    shipDate = AskDate("선택한 " & selRows.Count & "행에 넣을 출고일을 입력하세요.")
    If IsEmpty(shipDate) Then Exit Sub
    For Each r In selRows.Keys
        If Len(HawbKey(ws.Range(COL_HAWB & r).Value2)) = 0 Then
            nNoHawb = nNoHawb + 1
        Else
            Select Case ShipState(ws, CLng(r))
                Case 0: MarkShipped ws, CLng(r), CDate(shipDate): nShip = nShip + 1
                Case 1: nDup = nDup + 1
                Case Else: nOther = nOther + 1
            End Select
        End If
    Next r
    MsgBox "출고 처리: " & nShip & "행" & vbCrLf & _
           "이미 출고(그대로 둠): " & nDup & "행" & vbCrLf & _
           COL_SHIP & "열 다른 값으로 건너뜀: " & nOther & "행" & vbCrLf & _
           "송장번호 없는 행: " & nNoHawb & "행", vbInformation, "일괄 출고"
End Sub

'==============================================================
' 3) 취소: 선택한 행의 출고일과 A~C열 색상을 지움 (출고일이 있는 행만)
'==============================================================
Public Sub ShipScan_Cancel()
    Dim ws As Worksheet, selRows As Object, r As Variant, n As Long

    Set ws = ActiveSheet
    Set selRows = SelectedRows(ws)
    If selRows.Count = 0 Then MsgBox "취소할 행을 선택하세요.", vbExclamation: Exit Sub
    If MsgBox("선택한 " & selRows.Count & "행 중 출고일이 있는 행의 출고일과 " & MARK_FROM & "~" & MARK_TO & "열 색상을 지웁니다.", _
              vbOKCancel Or vbQuestion, "출고 취소") = vbCancel Then Exit Sub
    For Each r In selRows.Keys
        If ShipState(ws, CLng(r)) = 1 Then
            ws.Range(COL_SHIP & r).ClearContents
            MarkRange(ws, CLng(r)).Interior.ColorIndex = xlNone
            n = n + 1
        End If
    Next r
    MsgBox "취소: " & n & "행", vbInformation, "출고 취소"
End Sub

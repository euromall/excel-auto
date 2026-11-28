Attribute VB_Name = "ShipScan"
'==============================================================
' 출고 체크 매크로 v4 (2026-11-28)
' - 대상: 실행할 때 열려 있는 적하 시트(PTS_YYMM)
' - 먼저 출고할 행들을 선택(여러 행, Ctrl로 떨어진 구간도 가능)한 뒤 실행
' - 선택한 행 안에서만 송장번호(B열)를 찾음. 선택 밖은 처리하지 않음
' - 스캔되면: A~C열 연두색 + AI열에 출고일. 나머지 열의 기존 색은 그대로
' - 같은 송장이 선택 안에서 떨어진 행에 있으면: A~C열 주황색(확인 필요) + 출고일
' - AI열에 이미 값이 있는 행은 덮어쓰지 않음 (출고일이면 중복, 다른 값이면 건너뜀)
' - 스캔 중 "-" 입력: 방금 스캔한 1건 되돌리기
' - Esc로 끝낼 때: [유지] 결과 그대로 / [취소] 이번 스캔 전 상태로 (원래 색 복원)
' - 파일 저장은 하지 않음 (사용자가 직접 저장)
' - 별도 기록 시트는 만들지 않음. 기존 ScanCheckDuplicates 매크로는 그대로
'==============================================================
Option Explicit

' ---- 설정 (열 위치가 바뀌면 여기만 수정) ----
Private Const COL_HAWB As String = "B"        ' 송장번호(HAWB)
Private Const COL_SHIP As String = "AI"       ' 출고일을 넣는 열
Private Const MARK_FROM As String = "A"       ' 색칠 시작 열
Private Const MARK_TO As String = "C"         ' 색칠 끝 열
Private Const FIRST_ROW As Long = 7           ' 데이터 시작 행
Private Const SHIP_COLOR As Long = 13434828   ' 연두색 RGB(204,255,204) = 출고
Private Const SEP_COLOR As Long = 10079487    ' 주황색 RGB(255,204,153) = 떨어진 행, 확인 필요
Private Const DATE_FORMAT As String = "yyyy-mm-dd"
Private Const UNDO_KEY As String = "-"        ' 스캔 중 이것을 입력하면 방금 것 되돌리기

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
    s = Application.InputBox(prompt & vbCrLf & "(형식: 2026-11-28)", "출고일", Format$(Date, DATE_FORMAT), Type:=2)
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

' 행 번호 목록을 작은 순서로 정렬한 배열(0부터)
Private Function SortedRows(col As Collection) As Variant
    Dim a() As Long, i As Long, j As Long, t As Long
    ReDim a(0 To col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    For i = 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= 0
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
    SortedRows = a
End Function

' 같은 송장의 행들이 서로 붙어 있지 않으면 True
Private Function IsSeparated(hits As Variant) As Boolean
    IsSeparated = (UBound(hits) > 0) And (hits(UBound(hits)) - hits(0) <> UBound(hits))
End Function

Private Function RowText(hits As Variant) As String
    Dim i As Long
    For i = 0 To UBound(hits)
        RowText = RowText & IIf(i > 0, ",", "") & hits(i)
    Next i
End Function

' 색칠할 칸 (예: A7:C7)
Private Function MarkRange(ws As Worksheet, ByVal r As Long) As Range
    Set MarkRange = ws.Range(MARK_FROM & r & ":" & MARK_TO & r)
End Function

Private Sub MarkShipped(ws As Worksheet, ByVal r As Long, ByVal shipDate As Date, ByVal fillColor As Long)
    MarkRange(ws, r).Interior.Color = fillColor
    With ws.Range(COL_SHIP & r)
        .NumberFormat = DATE_FORMAT
        .Value = shipDate
    End With
End Sub

' 바꾸기 전 상태 기억: Array(행, A~C 각 칸 색(-1=색 없음), AI 값, AI 표시형식)
Private Function SaveState(ws As Worksheet, ByVal r As Long) As Variant
    Dim cols() As Variant, i As Long, c As Range
    ReDim cols(0 To MarkRange(ws, r).Cells.Count - 1)
    i = 0
    For Each c In MarkRange(ws, r).Cells
        If c.Interior.ColorIndex = xlNone Then
            cols(i) = -1
        Else
            cols(i) = c.Interior.Color
        End If
        i = i + 1
    Next c
    SaveState = Array(r, cols, ws.Range(COL_SHIP & r).Value, ws.Range(COL_SHIP & r).NumberFormat)
End Function

Private Sub RestoreState(ws As Worksheet, st As Variant)
    Dim r As Long, cols As Variant, i As Long, c As Range
    r = st(0)
    cols = st(1)
    i = 0
    For Each c In MarkRange(ws, r).Cells
        If cols(i) = -1 Then
            c.Interior.ColorIndex = xlNone
        Else
            c.Interior.Color = cols(i)
        End If
        i = i + 1
    Next c
    With ws.Range(COL_SHIP & r)
        .NumberFormat = st(3)
        If IsEmpty(st(2)) Then
            .ClearContents
        Else
            .Value = st(2)
        End If
    End With
End Sub

' 스캔 1건 되돌리기. grp(1)=송장번호, grp(2)=떨어짐 여부, grp(3~)=행별 이전 상태
Private Sub UndoGroup(ws As Worksheet, grp As Collection)
    Dim i As Long
    For i = grp.Count To 3 Step -1
        RestoreState ws, grp(i)
    Next i
End Sub

Private Sub UndoAll(ws As Worksheet, history As Collection)
    Dim i As Long
    For i = history.Count To 1 Step -1
        UndoGroup ws, history(i)
    Next i
End Sub

'==============================================================
' 1) 출고 스캔: 선택한 행 안에서 송장 바코드를 찾아 출고 표시
'==============================================================
Public Sub ShipScan_Start()
    Dim ws As Worksheet, selRows As Object, idx As Object, k As Variant
    Dim shipDate As Variant, raw As Variant, key As String
    Dim hits As Variant, i As Long, state As Integer, isSep As Boolean
    Dim history As Collection, grp As Collection
    Dim nDup As Long, nOther As Long, nMiss As Long, nSepKeys As Long
    Dim dupNow As Long, otherNow As Long, nSepDone As Long, sepList As String
    Dim msg As String, pLeft As Double, pTop As Double, ans As VbMsgBoxResult

    Set ws = ActiveSheet
    Set selRows = NeedRows(ws)
    If selRows Is Nothing Then Exit Sub
    Set idx = BuildIndex(ws, selRows)
    If idx.Count = 0 Then MsgBox "선택한 행의 " & COL_HAWB & "열에 송장번호가 없습니다.", vbExclamation: Exit Sub
    For Each k In idx.Keys
        If IsSeparated(SortedRows(idx.Item(k))) Then nSepKeys = nSepKeys + 1
    Next k

    msg = ws.Name & " 시트, 선택한 " & selRows.Count & "행(송장 " & idx.Count & "개)에서 출고 스캔을 시작합니다."
    If nSepKeys > 0 Then
        msg = msg & vbCrLf & vbCrLf & "주의: 같은 송장이 떨어진 행에 있는 경우 " & nSepKeys & "개" & vbCrLf & _
              "(스캔하면 주황색으로 표시됩니다. 붙여 주거나 확인하세요)"
    End If
    If MsgBox(msg, vbOKCancel Or vbQuestion, "출고 스캔") = vbCancel Then Exit Sub
    shipDate = AskDate("출고일을 입력하세요.")
    If IsEmpty(shipDate) Then Exit Sub

    Set history = New Collection
    pLeft = Application.Left + Application.Width * 0.55
    pTop = Application.Top + Application.Height * 0.5
    msg = "출고일 " & Format$(shipDate, DATE_FORMAT) & " | 송장 바코드를 스캔하세요."
    Do
        raw = Application.InputBox(msg & vbCrLf & "진행: " & history.Count & " / " & idx.Count & _
                                   "   방금 것 되돌리기: " & UNDO_KEY & "   끝내기: Esc", _
                                   "출고 스캔 (" & ws.Name & ")", Left:=pLeft, Top:=pTop, Type:=2)

        ' ---- Esc: 끝내기 (유지 / 취소 / 계속) ----
        If VarType(raw) = vbBoolean Then
            If history.Count = 0 Then Exit Do
            ans = MsgBox("이번 스캔 " & history.Count & "건을 어떻게 할까요?" & vbCrLf & vbCrLf & _
                         "[예]  유지 - 결과를 그대로 둡니다 (저장은 직접 하세요)" & vbCrLf & _
                         "[아니요]  취소 - 이번 스캔을 모두 원래 상태로 돌립니다" & vbCrLf & _
                         "[취소]  스캔을 계속합니다", vbYesNoCancel Or vbQuestion, "출고 스캔 끝내기")
            If ans = vbYes Then Exit Do
            If ans = vbNo Then
                If MsgBox("이번 스캔 " & history.Count & "건을 모두 취소하고 스캔 전 상태로 돌립니다." & vbCrLf & _
                          "(A~C열 원래 색과 AI열 값이 복원됩니다)" & vbCrLf & vbCrLf & "계속할까요?", _
                          vbYesNo Or vbExclamation Or vbDefaultButton2, "출고 스캔 취소") = vbYes Then
                    UndoAll ws, history
                    MsgBox "이번 스캔 " & history.Count & "건을 모두 취소했습니다. 스캔 전 상태입니다.", _
                           vbInformation, "출고 스캔 취소"
                    Exit Sub
                End If
            End If
            msg = "스캔을 계속합니다. 송장 바코드를 스캔하세요."
            GoTo NextScan
        End If

        ' ---- "-": 방금 스캔한 1건 되돌리기 ----
        If Trim$(CStr(raw)) = UNDO_KEY Then
            If history.Count = 0 Then
                msg = "[되돌리기] 되돌릴 스캔이 없습니다."
            Else
                Set grp = history(history.Count)
                UndoGroup ws, grp
                history.Remove history.Count
                msg = "[되돌림] " & grp(1) & " - 스캔 전 상태로 돌렸습니다."
            End If
            GoTo NextScan
        End If

        key = HawbKey(CStr(raw))
        If Len(key) = 0 Then GoTo NextScan

        If Not idx.Exists(key) Then
            nMiss = nMiss + 1
            Beep
            msg = "[선택 범위에 없음] " & key & " - 처리하지 않았습니다."
            GoTo NextScan
        End If

        hits = SortedRows(idx.Item(key))
        isSep = IsSeparated(hits)
        Set grp = New Collection
        grp.Add key
        grp.Add isSep
        dupNow = 0
        otherNow = 0
        For i = 0 To UBound(hits)
            state = ShipState(ws, CLng(hits(i)))
            If state = 0 Then
                grp.Add SaveState(ws, CLng(hits(i)))
                If isSep Then
                    MarkShipped ws, CLng(hits(i)), CDate(shipDate), SEP_COLOR
                Else
                    MarkShipped ws, CLng(hits(i)), CDate(shipDate), SHIP_COLOR
                End If
            ElseIf state = 1 Then
                dupNow = dupNow + 1
            Else
                otherNow = otherNow + 1
            End If
        Next i
        nDup = nDup + dupNow
        nOther = nOther + otherNow

        If grp.Count > 2 Then
            history.Add grp
            If isSep Then
                Beep
                msg = "[분리됨 - 확인 필요] " & key & " → " & RowText(hits) & "행 (주황색, 출고일 입력함)"
            Else
                msg = "[출고] " & key & " → " & RowText(hits) & "행"
            End If
            Application.Goto ws.Range(COL_HAWB & hits(0)), False
        ElseIf dupNow > 0 Then
            Beep
            msg = "[중복] " & key & " - 이미 출고 " & Format$(ws.Range(COL_SHIP & hits(0)).Value, DATE_FORMAT)
        Else
            Beep
            msg = "[건너뜀] " & key & " - " & COL_SHIP & "열에 다른 값이 있어 덮어쓰지 않았습니다 (" & RowText(hits) & "행)"
        End If
NextScan:
    Loop

    For i = 1 To history.Count
        If history(i)(2) Then
            nSepDone = nSepDone + 1
            If nSepDone <= 15 Then sepList = sepList & vbCrLf & "  " & history(i)(1)
        End If
    Next i
    If nSepDone > 15 Then sepList = sepList & vbCrLf & "  ... 외 " & (nSepDone - 15) & "개"

    MsgBox "출고 처리: 송장 " & history.Count & "개" & vbCrLf & _
           "  그중 떨어진 행(주황, 확인 필요): " & nSepDone & "개" & sepList & vbCrLf & _
           "중복(이미 출고): " & nDup & "행" & vbCrLf & _
           COL_SHIP & "열 다른 값으로 건너뜀: " & nOther & "행" & vbCrLf & _
           "선택 범위에 없음: " & nMiss & "건" & vbCrLf & vbCrLf & _
           "파일은 저장하지 않았습니다. 확인 후 직접 저장하세요.", vbInformation, "출고 스캔 종료"
End Sub

'==============================================================
' 2) 밀린 분: 선택한 행 전체를 같은 출고일로 처리
'==============================================================
Public Sub ShipScan_BulkDate()
    Dim ws As Worksheet, selRows As Object, idx As Object, k As Variant
    Dim shipDate As Variant, hits As Variant, i As Long, isSep As Boolean
    Dim nShip As Long, nSep As Long, nDup As Long, nOther As Long, nIndexed As Long

    Set ws = ActiveSheet
    Set selRows = NeedRows(ws)
    If selRows Is Nothing Then Exit Sub
    Set idx = BuildIndex(ws, selRows)
    shipDate = AskDate("선택한 " & selRows.Count & "행에 넣을 출고일을 입력하세요.")
    If IsEmpty(shipDate) Then Exit Sub
    For Each k In idx.Keys
        hits = SortedRows(idx.Item(k))
        isSep = IsSeparated(hits)
        For i = 0 To UBound(hits)
            nIndexed = nIndexed + 1
            Select Case ShipState(ws, CLng(hits(i)))
                Case 0
                    If isSep Then
                        MarkShipped ws, CLng(hits(i)), CDate(shipDate), SEP_COLOR
                        nSep = nSep + 1
                    Else
                        MarkShipped ws, CLng(hits(i)), CDate(shipDate), SHIP_COLOR
                    End If
                    nShip = nShip + 1
                Case 1
                    nDup = nDup + 1
                Case Else
                    nOther = nOther + 1
            End Select
        Next i
    Next k
    MsgBox "출고 처리: " & nShip & "행" & vbCrLf & _
           "  그중 떨어진 행(주황, 확인 필요): " & nSep & "행" & vbCrLf & _
           "이미 출고(그대로 둠): " & nDup & "행" & vbCrLf & _
           COL_SHIP & "열 다른 값으로 건너뜀: " & nOther & "행" & vbCrLf & _
           "송장번호 없는 행: " & (selRows.Count - nIndexed) & "행" & vbCrLf & vbCrLf & _
           "파일은 저장하지 않았습니다.", vbInformation, "일괄 출고"
End Sub

'==============================================================
' 3) 출고 취소(나중에): 선택한 행의 출고일과 A~C열 색상을 지움 (출고일이 있는 행만)
'    스캔 중 취소와 달리 원래 색은 기억하지 않으므로 A~C열은 "색 없음"이 됨
'==============================================================
Public Sub ShipScan_Cancel()
    Dim ws As Worksheet, selRows As Object, r As Variant, n As Long

    Set ws = ActiveSheet
    Set selRows = SelectedRows(ws)
    If selRows.Count = 0 Then MsgBox "취소할 행을 선택하세요.", vbExclamation: Exit Sub
    If MsgBox("선택한 " & selRows.Count & "행 중 출고일이 있는 행의 출고일과 " & MARK_FROM & "~" & MARK_TO & "열 색상을 지웁니다." & vbCrLf & _
              "(" & MARK_FROM & "~" & MARK_TO & "열은 색 없음이 됩니다)", _
              vbOKCancel Or vbExclamation, "출고 취소") = vbCancel Then Exit Sub
    For Each r In selRows.Keys
        If ShipState(ws, CLng(r)) = 1 Then
            ws.Range(COL_SHIP & r).ClearContents
            MarkRange(ws, CLng(r)).Interior.ColorIndex = xlNone
            n = n + 1
        End If
    Next r
    MsgBox "취소: " & n & "행", vbInformation, "출고 취소"
End Sub

'==============================================================
' 4) 색 정렬: 선택한 한 구간의 행들을 B열 색 기준으로 정렬
'    순서: 색 없음(미출고, 다른 색 포함) -> 주황(떨어진 행) -> 연두(출고)
'    같은 색 안에서는 원래 순서 유지. 행 전체(모든 열)가 함께 움직임
'==============================================================
Public Sub ShipScan_SortByColor()
    Dim ws As Worksheet, area As Range, rTop As Long, rBot As Long, lastRow As Long
    Dim lastCol As Long, tmpCol As Long, r As Long, rank As Long, c As Range
    Dim n0 As Long, n1 As Long, n2 As Long

    Set ws = ActiveSheet
    If TypeName(Selection) <> "Range" Then MsgBox "정렬할 행을 선택하세요.", vbExclamation: Exit Sub
    If Not Selection.Worksheet Is ws Then Exit Sub
    If Selection.Areas.Count <> 1 Then
        MsgBox "색 정렬은 붙어 있는 한 구간만 선택했을 때 할 수 있습니다." & vbCrLf & _
               "(Ctrl로 떨어진 구간을 여러 개 선택하면 정렬할 수 없습니다)", vbExclamation, "색 정렬"
        Exit Sub
    End If
    If ws.FilterMode Then
        MsgBox "필터가 걸려 있습니다. 필터를 모두 해제한 뒤 실행하세요.", vbExclamation, "색 정렬"
        Exit Sub
    End If

    Set area = Selection.Areas(1)
    lastRow = ws.Cells(ws.Rows.Count, COL_HAWB).End(xlUp).Row
    rTop = area.Row
    If rTop < FIRST_ROW Then rTop = FIRST_ROW
    rBot = area.Row + area.Rows.Count - 1
    If rBot > lastRow Then rBot = lastRow
    If rBot - rTop < 1 Then MsgBox "정렬할 행을 2행 이상 선택하세요.", vbExclamation, "색 정렬": Exit Sub

    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    tmpCol = lastCol + 1
    If tmpCol > ws.Columns.Count Then MsgBox "정렬에 쓸 빈 열이 없습니다.", vbExclamation, "색 정렬": Exit Sub

    If MsgBox(rTop & "~" & rBot & "행(" & (rBot - rTop + 1) & "행)을 " & COL_HAWB & "열 색 기준으로 정렬합니다." & vbCrLf & _
              "순서: 색 없음 → 주황(떨어진 행) → 연두(출고)" & vbCrLf & vbCrLf & _
              "행 전체가 함께 움직이며, 정렬은 되돌릴 수 없습니다." & vbCrLf & _
              "먼저 저장해 두는 것을 권합니다. 계속할까요?", _
              vbYesNo Or vbExclamation Or vbDefaultButton2, "색 정렬") = vbNo Then Exit Sub

    ' 임시 열에 순서 번호 = 색 순위 * 1000000 + 원래 순서
    For r = rTop To rBot
        Set c = ws.Range(COL_HAWB & r)
        If c.Interior.ColorIndex = xlNone Then
            rank = 0
        ElseIf c.Interior.Color = SEP_COLOR Then
            rank = 1
        ElseIf c.Interior.Color = SHIP_COLOR Then
            rank = 2
        Else
            rank = 0
        End If
        Select Case rank
            Case 0: n0 = n0 + 1
            Case 1: n1 = n1 + 1
            Case 2: n2 = n2 + 1
        End Select
        ws.Cells(r, tmpCol).Value = rank * 1000000# + (r - rTop)
    Next r

    On Error GoTo Fail
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(rTop, tmpCol), ws.Cells(rBot, tmpCol)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(rTop, 1), ws.Cells(rBot, tmpCol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
        .SortFields.Clear
    End With
    On Error GoTo 0
    ws.Range(ws.Cells(rTop, tmpCol), ws.Cells(rBot, tmpCol)).ClearContents

    MsgBox "정렬 완료 (" & rTop & "~" & rBot & "행)" & vbCrLf & _
           "색 없음: " & n0 & "행" & vbCrLf & _
           "주황(떨어진 행): " & n1 & "행" & vbCrLf & _
           "연두(출고): " & n2 & "행" & vbCrLf & vbCrLf & _
           "파일은 저장하지 않았습니다.", vbInformation, "색 정렬"
    Exit Sub

Fail:
    ws.Range(ws.Cells(rTop, tmpCol), ws.Cells(rBot, tmpCol)).ClearContents
    MsgBox "정렬하지 못했습니다: " & Err.Description & vbCrLf & _
           "(병합된 셀이 있으면 정렬할 수 없습니다)", vbExclamation, "색 정렬"
End Sub

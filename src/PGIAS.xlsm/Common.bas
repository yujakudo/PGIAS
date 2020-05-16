Attribute VB_Name = "Common"
Option Explicit

'   シート移動履歴
Dim shHistory(10) As Object
Dim shHistoryPos(2) As Integer


'   List objectの取得
Public Function getListObject(ByVal table As Variant) As ListObject
    Select Case TypeName(table)
        Case "ListObject"
            Set getListObject = table
        Case "Worksheet"
            Set getListObject = table.ListObjects(1)
        Case "Range"
            Set getListObject = table.ListObject
        Case Else
            Set getListObject = Nothing
    End Select
End Function

'   テーブル上の、該セルと同じ行にあるセルを得る
Public Function getColumn(ByVal cname As String, ByVal cel As Range) As Range
    Dim row As Long
    On Error GoTo Err
    With cel.ListObject
        row = cel.row - .DataBodyRange.row + 1
        Set getColumn = .ListColumns(cname).DataBodyRange.cells(row, 1)
    End With
    Exit Function
Err:
    If Not cel.ListObject Is Nothing Then _
        MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(cel.ListObject.name, cname))
End Function

'   テーブル上の、列タイトルより列番号を得る
Public Function getColumnIndex(ByVal cname As String, ByVal table As Variant) As Long
    On Error GoTo Err
    With getListObject(table)
        getColumnIndex = WorksheetFunction.Match(cname, .HeaderRowRange, 0)
    End With
    Exit Function
Err:
    MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(getListObject(table).name, cname))
End Function

'   テーブル上の、列タイトルから列番号を得るハッシュまたは配列を取得する
Public Function getColumnIndexes(ByVal table As Variant, _
                        Optional ByVal attrs As Variant = Nothing) As Variant
    Dim col, num, i As Long
    Dim vals() As Long
    If Not IsArray(attrs) Then
        Set getColumnIndexes = CreateObject("Scripting.Dictionary")
        With getListObject(table).HeaderRowRange
            For col = 1 To .count
                getColumnIndexes.item(.cells(1, col).text) = col
            Next
        End With
    Else
        num = UBound(attrs)
        ReDim vals(num)
        On Error GoTo Err
        With getListObject(table)
            For i = 0 To num
                vals(i) = WorksheetFunction.Match( _
                        attrs(i), .HeaderRowRange, 0)
            Next
        End With
        On Error GoTo 0
        getColumnIndexes = vals
    End If
    Exit Function
Err:
    MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(getListObject(table).name, attrs(i)))
End Function

'   セルと同行の値を取得する
Public Function getRowValues(ByVal cel As Range, _
                        Optional ByVal attrs As Variant = Nothing, _
                        Optional ByVal colLimit As Variant = Nothing) As Variant
    Dim val() As Variant
    Dim dataNum, i, row, col As Long
    
    With cel.ListObject
        row = cel.row - .DataBodyRange.row + 1
        If Not IsArray(attrs) Then
            Set getRowValues = CreateObject("Scripting.Dictionary")
            If Not IsArray(colLimit) Then colLimit = _
                    Array(1, .DataBodyRange.columns.count)
            If Not IsNumeric(colLimit(0)) Then
                colLimit(0) = getColumnIndex(colLimit(0), cel)
            End If
            If Not IsNumeric(colLimit(1)) Then
                colLimit(1) = getColumnIndex(colLimit(1), cel)
            End If
            For col = colLimit(0) To colLimit(1)
                getRowValues.item(.HeaderRowRange.cells(1, col).text) = _
                    .DataBodyRange.cells(row, col).value
            Next
        Else
            dataNum = UBound(attrs)
            ReDim val(dataNum)
            On Error GoTo Err
            For i = 0 To dataNum
                val(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).value
            Next
            On Error GoTo 0
            getRowValues = val
        End If
    End With
    Exit Function
Err:
    MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(cel.ListObject.name, attrs(i)))
End Function

'   テーブル上の列を検索し、行番号を取得する
Public Function searchRow(ByVal key As Variant, ByVal column As Variant, _
                            ByVal table As Variant, _
                    Optional ByVal ignoreError As Boolean = False) As Long
    If key = "" Then
        Exit Function
    End If
    On Error GoTo Err
    With getListObject(table)
        searchRow = WorksheetFunction.Match(key, .ListColumns(column).DataBodyRange, 0)
    End With
    Exit Function
Err:
    If Not ignoreError Then _
        MsgBox msgstr(msgKeyDoesNotExist, Array(getListObject(table).name, column, key))
End Function

'   テーブルを検索し、指定の列の値を取得する
Public Function seachAndGetValues(ByVal key As Variant, ByVal column As Variant, _
                    ByVal table As Variant, _
                    Optional ByVal attrs As Variant = Nothing) As Variant
    Dim val() As Variant
    Dim dataNum, i, row, col As Long
    If key = "" Then Exit Function
    With getListObject(table)
        row = WorksheetFunction.Match(key, .ListColumns(column).DataBodyRange, 0)
        If Not IsArray(attrs) Then
            seachAndGetValues = CreateObject("Scripting.Dictionary")
            For col = 1 To .DataBodyRange.columns.count
                seachAndGetValues.item(.HeaderRowRange.cells(1, col)) = _
                    .DataBodyRange.cells(row, col).value
            Next
        Else
            dataNum = UBound(attrs)
            ReDim val(dataNum)
            On Error GoTo Err
            For i = 0 To dataNum
                val(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).value
            Next
            On Error GoTo 0
            seachAndGetValues = val
        End If
    End With
    Exit Function
Err:
    MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(getListObject(table).name, attrs(i)))
End Function


'   指定列範囲の値を得る
Public Function getSettings(ByVal rng As Variant, _
                Optional orientation As XlOrientation = xlHorizontal, _
                Optional idx As Integer = 0) As Object
    Dim row, col As Long
    If Not IsObject(rng) Then Set rng = Range(rng)
    Set getSettings = CreateObject("Scripting.Dictionary")
    With rng
        If orientation = xlHorizontal Then
            For col = 1 To .columns.count
                Call setSettingValue(getSettings, _
                        .cells(1, col).text, .cells(2 + idx, col))
            Next
            Exit Function
        Else
            For row = 1 To .rows.count
                Call setSettingValue(getSettings, _
                        .cells(row, 1).text, .cells(row, 2 + idx).value)
            Next
        End If
    End With
End Function

Private Sub setSettingValue(ByRef settings As Object, ByVal key As String, _
                    ByVal value As Variant)
    If right(key, 2) = "_b" Then
        If IsNumeric(value) Then
            If value Then value = True Else value = False
        Else
            value = LCase(value)
            If value = "on" Or value = "true" Then value = True Else value = False
        End If
    End If
    settings(key) = value
End Sub

Public Sub setSettings(ByVal rng As Variant, _
                    ByRef settings As Object, _
            Optional orientation As XlOrientation = xlHorizontal, _
            Optional idx As Integer = 0)
    Dim row, col As Long
    Dim key As String
    If settings Is Nothing Then Exit Sub
    If Not IsObject(rng) Then Set rng = Range(rng)
    With rng
        If orientation = xlHorizontal Then
            For col = 1 To .columns.count
                key = .cells(1, col).text
                If settings.exists(key) Then
                    .cells(2 + idx, col).value = settings(key)
                End If
            Next
            Exit Sub
        Else
            For row = 1 To .rows.count
                key = .cells(row, 1).text
                If settings.exists(key) Then
                    .cells(row, 2 + idx).value = settings(key)
                End If
            Next
        End If
    End With
End Sub

'   マクロ高速化
Function doMacro(Optional ByVal comment As String = "", _
                    Optional ByVal disableEvent As Boolean = True, _
                    Optional ByVal totalReset As Boolean = False)
    Static PrevCalculation As Long
    Static nest As Long
    Static msg(100) As String
    
    If comment <> "" Then
        If nest = 0 Then
            nest = 1
            Application.ScreenUpdating = False
            PrevCalculation = Application.Calculation
            Application.Calculation = xlCalculationManual
            If disableEvent Then Application.EnableEvents = False
        Else
            nest = nest + 1
        End If
        msg(nest) = comment
        Application.StatusBar = comment
    Else
        nest = nest - 1
        If nest < 1 Or totalReset Then
            nest = 0
            Application.ScreenUpdating = True
            Application.Calculation = PrevCalculation
            Application.StatusBar = ""
            Application.EnableEvents = True
        Else
            Application.StatusBar = msg(nest)
        End If
    End If
End Function

'   強制リセット
Public Sub a_resetDoMacro()
    Call doMacro(, , True)
End Sub

'   イベントの有効/無効
Sub enableEvent(Optional ByVal enable As Boolean = True)
    Static ee As Boolean
    If Not enable Then
        ee = Application.EnableEvents
        Application.EnableEvents = False
    Else
        Application.EnableEvents = True
    End If
End Sub

'プログレスバーもどき
Sub dspProgress( _
        Optional ByVal msg As String = "", _
        Optional ByVal den As Long = -1)
    Static cnt As Long
    Static prog As Long
    Static deno As Long
    Static message As String
    Dim ratio As Long
    
    If den > 0 Then
        deno = den: prog = 0: cnt = 0
        If msg = "" Then msg = Application.StatusBar
        message = msg
    ElseIf den = 0 Then
        deno = 0: prog = 0: cnt = 0
        message = ""
        Application.StatusBar = ""
        Exit Sub
    Else
        If msg <> "" Then
            message = msg
        ElseIf deno > 0 Then
            cnt = cnt + 1
            ratio = cnt * 20 / deno
            If ratio > 20 Then ratio = 20
            If prog = ratio Then Exit Sub
            prog = ratio
        End If
    End If
    If deno > 0 Then
        Application.StatusBar = _
            String(prog, "●") & String(20 - prog, "○") & " " & message
    End If
End Sub

'   CSVでセーブ
Sub saveCsv(ByVal fh As Variant, ByVal rng As Range, _
            Optional ByVal head As Range = Nothing)
    Dim i As Long
    Dim fileName As String

    If head Is Nothing Then
        Set head = rng.Areas(1).rows(1).Offset(-1, 0)
        i = 2
        While i <= rng.Areas.count
            Set head = Union(head, rng.Areas(i).rows(1).Offset(-1, 0))
            i = i + 1
        Wend
    End If
    If Not IsNumeric(fh) Then
        If fh = "" Then Exit Sub
        fileName = fh
        fh = FreeFile
        On Error GoTo errProc
        Open fileName For Output As #fh
        On Error GoTo 0
    End If
    Print #fh, joinRow(head)
    For i = 1 To rng.rows.count
        Print #fh, joinRow(rng, i)
    Next
    If fileName <> "" Then
        Close #fh
    End If
    Exit Sub
errProc:
    MsgBox "Error occured when open " & fileName
    Close #fh
End Sub

'   ヘッダの読み込みまで
Private Function readHeader(ByRef head As Range, ByRef cell As Range, _
            ByRef fh As Variant, ByRef fileName As String, _
            ByRef colHead As Variant) As Boolean
    Dim line As String
    Dim found As Range
    Dim i As Long
    Dim lhead As Variant
    
    readHeader = False
    If head Is Nothing Then
        Set head = cell.Offset(-1, 0)
        Set head = Range(head, head.End(xlToRight))
    End If
    If Not IsNumeric(fh) Then
        If fh = "" Then Exit Function
        fileName = fh
        fh = FreeFile
        On Error GoTo errProc
        Open fileName For Input As #fh
        On Error GoTo 0
    End If
    While line = ""
        Line Input #fh, line
    Wend
    lhead = splitLine(line)
    colHead = Array()
    ReDim colHead(UBound(lhead))
    For i = 0 To UBound(lhead)
        Set found = head.Find(What:=lhead(i), LookAt:=xlWhole)
        If found Is Nothing Then
            colHead(i) = 0
        Else
            colHead(i) = found.column - head.column + 1
        End If
    Next
    readHeader = True
    Exit Function
errProc:
    MsgBox "Error occured when open " & fileName
    Close #fh
End Function

'   CSVのロード
Public Sub loadCsv(ByVal fh As Variant, ByVal cell As Range, _
            Optional ByVal head As Range = Nothing)
    Dim fileName As String
    Dim i As Long
    Dim colHead, words As Variant
    Dim line As String
    
    If Not readHeader(head, cell, fh, fileName, colHead) Then
        Exit Sub
    End If
    Do Until EOF(fh)
        Line Input #fh, line
        If line = "" Then Exit Do
        words = splitLine(line)
        For i = 0 To UBound(words)
            If colHead(i) Then
                cell.Offset(0, colHead(i) - 1).value = words(i)
            End If
        Next
        Set cell = cell.Offset(1, 0)
    Loop
    If fileName <> "" Then
        Close #fh
    End If
End Sub

'   行内のデータをコンマ区切りの文字列にする
Private Function joinRow(ByVal rng As Range, _
            Optional ByVal row As Long = 1) As String
    Dim ai, col As Long
    joinRow = ""
    For ai = 1 To rng.Areas.count
        With rng.Areas(ai)
            For col = 1 To .columns.count
                If joinRow <> "" Then joinRow = joinRow & ","
                joinRow = joinRow & """" & .cells(row, col).text & """"
            Next
        End With
    Next
End Function

'   コンマ区切りを分割（ダブルクォーテーションを考慮）
Private Function splitLine(ByVal line As String) As Variant
    Dim arr As Variant
    Dim idx, i As Long
    Dim word, words() As String
    idx = 0
    arr = Split(line, ",")
    ReDim words(UBound(arr))
    For i = 0 To UBound(arr)
        If word = "" Then
            word = arr(i)
        Else
            word = word & "," & arr(i)
        End If
        If left(word, 1) = Chr(34) And right(word, 1) = Chr(34) Then
            word = Mid(word, 2, Len(word) - 2)
        End If
        If left(word, 1) <> Chr(34) Then
                words(idx) = word
                idx = idx + 1
                word = ""
        End If
    Next
    ReDim Preserve words(idx - 1)
    splitLine = words
End Function

'   マージインポート
'   読み込むデータと書き込み先はkeyの列で同順にソートされていること
Public Function margeCsv(ByVal fh As Variant, ByVal cell As Range, _
            ByVal key As String, _
            Optional ByVal head As Range = Nothing, _
            Optional ByVal isTest As Boolean = False, _
            Optional ByVal callback As Variant = Nothing) As String
    Dim fileName As String
    Dim i, keycol, keyidx As Long
    Dim colHead, words, oldVal, newVal As Variant
    Dim line, log, diff, slog As String
    Dim keyRng, bottom, cel As Range
    
    If Not readHeader(head, cell, fh, fileName, colHead) Then
        Exit Function
    End If
    '   キー列タイトルの列インデックス（0から）とデータインデックス
    For i = 0 To UBound(colHead)
        If head.cells(1, colHead(i)).value = key Then
            keycol = colHead(i) - 1
            keyidx = i
            Exit For
        End If
    Next
    Set bottom = cell.Offset(0, keycol).End(xlDown)
    Set keyRng = Range(cell.Offset(0, keycol), bottom)
    
    Do Until EOF(fh)
        Line Input #fh, line
        If line = "" Then Exit Do
        words = splitLine(line)
        Set cel = keyRng.Find(What:=words(keyidx), After:=bottom, LookAt:=xlWhole)
        If cel Is Nothing Then
            '   新規キー値。行を挿入して書き込み
            Set bottom = bottom.Offset(1, 0)
            Set keyRng = Range(keyRng.cells(1, 1), bottom)
            Set cel = bottom.Offset(0, -keycol)
            If Not isTest Then
                For i = 0 To UBound(words)
                    If colHead(i) Then
                        cel.Offset(0, colHead(i) - 1).value = words(i)
                    End If
                Next
            End If
            log = log & "Added " & words(keyidx) & vbCrLf
        Else
            '   既存キー値。各データを比較し、異なる場合はコールバック
            '   コールバック先で上書きするかも
            Set cel = cel.Offset(0, -keycol)
            diff = ""
            For i = 0 To UBound(words)
                If colHead(i) Then
                    oldVal = cel.Offset(0, colHead(i) - 1).value
                    If IsNumeric(words(i)) Then newVal = val(words(i)) Else newVal = words(i)
                    If oldVal <> newVal Then
                        slog = ""
                        If IsArray(callback) Then
                            slog = CallByName(callback(0), callback(1), VbMethod, _
                                Array(cel.Offset(0, colHead(i) - 1), newVal, callback(2)))
                        End If
                        If slog = "" Then
                            slog = oldVal & "->" & newVal
                        End If
                        If slog <> "-" Then
                            If diff <> "" Then diff = diff & ","
                            diff = diff & head.cells(1, colHead(i)).text _
                                    & "(" & slog & ")"
                        End If
                    End If
                End If
            Next
            If diff <> "" Then
                log = log & "Difference at """ & words(keyidx) _
                        & """ " & diff & vbCrLf
            End If
        End If
    Loop
    If fileName <> "" Then
        Close #fh
    End If
    margeCsv = log
End Function


'   ダイアログを開いてファイル名を聞き、ファイルを開く
Public Function openFileWithDialog(filter As String, _
                Optional isSave As Boolean = False, _
                Optional ByVal baseName As String = "") As Integer
    Dim fh As Integer
    Dim fileName As String
    
    openFileWithDialog = -1
    fileName = fileDialog(filter, isSave, baseName)
    If fileName = "" Then Exit Function
    fh = FreeFile
    On Error GoTo errProc
    If isSave Then
        Open fileName For Output As #fh
    Else
        Open fileName For Input As #fh
    End If
    On Error GoTo 0
    openFileWithDialog = fh
    Exit Function
errProc:
    MsgBox "Error occured when open " & fileName
    Close #fh
End Function

'   ファイルを開くのダイアログ
Function fileDialog(ByVal filter As String, _
                Optional ByVal isSave As Boolean = False, _
                Optional ByVal baseName As String = "") As String
    Dim cur, fn As String
    cur = CurDir
    ChDrive left(ThisWorkbook.path, 2)
    ChDir ThisWorkbook.path & "\"
    If isSave Then
        fn = Date
        fn = Replace(fn, "/", "")
        fn = Replace(fn, ":", "")
        fn = Replace(fn, " ", "_")
        If baseName <> "" Then fn = baseName & "_" & fn
        fileDialog = Application.GetSaveAsFilename(fn, filter)
    Else
        fileDialog = Application.GetOpenFilename(filter)
    End If
    ChDrive left(cur, 2)
    ChDir cur
    If fileDialog = "False" Then fileDialog = ""
End Function
                
'   sprintf
Public Function msgstr(ByVal s As String, ByVal var As Variant) As String
    Dim i As Long
    If Not IsArray(var) Then var = Array(var)
    For i = 0 To UBound(var)
        s = Replace(s, "{" & Trim(i) & "}", var(i))
    Next
    msgstr = s
End Function

'   名前の定義からシートの取得
Public Function getSheetsByName(ByVal name As String)
    Dim sh, sheets() As Worksheet
    Dim num As Integer
    ReDim sheets(Worksheets.count)
    For Each sh In Worksheets
        If checkNameInSheet(sh, name) Then
            Set sheets(num) = sh
            num = num + 1
        End If
    Next
    ReDim Preserve sheets(num - 1)
    getSheetsByName = sheets
End Function

Public Function checkNameInSheet(ByVal sh As Worksheet, _
                        ByVal name As String) As Boolean
    Dim rng As Range
    On Error GoTo Err
    Set rng = sh.Range(name)
    checkNameInSheet = True
    Exit Function
Err:
    checkNameInSheet = False
End Function

'   時間文字列の取得
Public Function getTimeStr(ByVal stime As Long, _
                Optional ByVal delimiter As String = ":") As String
    Dim sec, min, hour As Integer
    
    sec = stime Mod 60
    stime = (stime - sec) / 60
    min = stime Mod 60
    hour = (stime - min) / 60
    If delimiter = ":" Then
        getTimeStr = right("0" & Trim(hour), 2) & ":" _
                    & right("0" & Trim(min), 2) & ":" _
                    & right("0" & Trim(sec), 2)
    Else
        If hour > 0 Then getTimeStr = Trim(hour) & "ﾟ"
        If getTimeStr <> "" Or min > 0 Then getTimeStr = getTimeStr & Trim(min) & "'"
        getTimeStr = getTimeStr & Trim(sec) & """"
    End If
    Trim (min) & "'" & right("0" & Trim(sec), 2) & """"
End Function

'   シートが変わったので履歴に記憶
Public Sub onSheetChange(ByVal sh As Object)
    Dim npos As Integer
    npos = nextHistoryPos(shHistoryPos(2))
    '   現在位置が途中で、かつシートが同じ
    If shHistoryPos(2) <> shHistoryPos(0) And shHistory(npos) Is sh Then
        shHistoryPos(2) = npos
        Exit Sub
    End If
    Set shHistory(npos) = sh
    shHistoryPos(0) = npos
    shHistoryPos(2) = npos
    If shHistory(shHistoryPos(1)) Is Nothing Then
        shHistoryPos(1) = npos
    ElseIf npos = shHistoryPos(1) Then
        shHistoryPos(1) = nextHistoryPos(npos)
    End If
End Sub

'   シートをまたいで移動
Public Sub jumpTo(ByVal sbj As Variant, Optional ByVal log As Boolean = True)
    Dim sh As Worksheet
    Dim cel As Range
    Dim state As Boolean
    If TypeName(sbj) = "Range" Then
        Set sh = sbj.Parent
        Set cel = sbj
    Else
        Set sh = sbj
        Set cel = Nothing
    End If
    If Not ActiveSheet Is sh Then
        state = Application.EnableEvents
        Application.EnableEvents = False
        sh.Activate
        Application.EnableEvents = state
        If log Then Call onSheetChange(sh)
    End If
    If Not cel Is Nothing Then Application.Goto cel
End Sub

Public Sub historyReset()
    shHistoryPos(0) = 0
    shHistoryPos(1) = 0
    shHistoryPos(2) = 0
End Sub

'   履歴を前に移動
Public Sub historyForward()
Attribute historyForward.VB_ProcData.VB_Invoke_Func = "m\n14"
    Call moveHistory(1)
End Sub

'   履歴を後に移動
Public Sub historyBackward()
Attribute historyBackward.VB_ProcData.VB_Invoke_Func = "n\n14"
    Call moveHistory(-1)
End Sub

'   移動
Private Sub moveHistory(ByVal dir As Integer)
    Dim sh As Object
    Dim state As Boolean
    If (dir > 0 And shHistoryPos(2) = shHistoryPos(0)) _
            Or (dir < 0 And shHistoryPos(2) = shHistoryPos(1)) Then Exit Sub
    shHistoryPos(2) = nextHistoryPos(shHistoryPos(2), dir)
    Set sh = shHistory(shHistoryPos(2))
    If sh Is Nothing Then Set sh = trancateHistory(dir)
    If sh Is Nothing Then Exit Sub
    state = Application.EnableEvents
    Application.EnableEvents = False
    sh.Activate
    ActiveCell.Activate
    Application.EnableEvents = state
End Sub

Private Function trancateHistory(ByVal dir As Integer) As Worksheet
    Dim pos, lidx, npos As Integer
    If dir > 0 Then lidx = 0 Else lidx = 1
    pos = shHistoryPos(2)
    '   現在位置が終端であれば、現在位置と終端を一つ戻す
    If pos = shHistory(lidx) Then
        While shHistory(shHistoryPos(2)) Is Nothing
            If pos = shHistory(1 - lidx) Then Exit Function
            shHistoryPos(2) = nextHistoryPos(pos, -dir)
            shHistoryPos(lidx) = shHistoryPos(2)
        Wend
        trancateHistory = shHistory(shHistoryPos(2))
        Exit Function
    End If
    While pos <> shHistoryPos(lidx)
        npos = nextHistoryPos(pos, dir)
        Set shHistory(pos) = shHistory(npos)
        pos = npos
    Wend
    shHistoryPos(lidx) = nextHistoryPos(pos, -dir)
    trancateHistory = shHistory(shHistoryPos(2))
End Function

'   ポインタの移動
Private Function nextHistoryPos(ByVal pos As Integer, _
            Optional ByVal dir As Integer = 1) As Integer
    If dir > 0 Then
        If pos = UBound(shHistory) Then nextHistoryPos = 0 Else nextHistoryPos = pos + 1
    Else
        If pos = 0 Then nextHistoryPos = UBound(shHistory) Else nextHistoryPos = pos - 1
    End If
End Function


'   コンマ区切りの文字列の連結
Public Function joinStrList(ByVal sl As Variant, _
                    Optional ByVal dir As Integer = 1) As String
    Dim lim(1), i As Integer
    If dir > 0 Then
        dir = 1: lim(0) = 0: lim(1) = UBound(sl)
    Else
        dir = -1: lim(1) = 0: lim(0) = UBound(sl)
    End If
    joinStrList = ""
    For i = lim(0) To lim(1) Step dir
        If IsArray(sl(i)) Then
            sl(i) = joinStrList(sl(i), dir)
        End If
        If sl(i) <> "" Then
            If joinStrList <> "" Then joinStrList = joinStrList & ","
            joinStrList = joinStrList & sl(i)
        End If
    Next
End Function

'   フィルター解除
Public Sub resetTableFilter(ByVal table As Variant)
    Dim LC As ListColumn
    With getListObject(table)
        .Sort.SortFields.Clear
        On Error GoTo EachColumn
        .Parent.ShowAllData
        Exit Sub
EachColumn:
        For Each LC In .ListColumns
            .Range.AutoFilter LC.DataBodyRange.column
        Next
    End With
End Sub

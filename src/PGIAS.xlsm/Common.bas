Attribute VB_Name = "Common"
Option Explicit

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
                getColumnIndexes.item(.cells(1, col).Text) = col
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
                getRowValues.item(.HeaderRowRange.cells(1, col).Text) = _
                    .DataBodyRange.cells(row, col).Value
            Next
        Else
            dataNum = UBound(attrs)
            ReDim val(dataNum)
            On Error GoTo Err
            For i = 0 To dataNum
                val(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).Value
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
                    ByVal table As Variant) As Long
    If key = "" Then
        Exit Function
    End If
    On Error GoTo Err
    With getListObject(table)
        searchRow = WorksheetFunction.Match(key, .ListColumns(column).DataBodyRange, 0)
    End With
    Exit Function
Err:
    MsgBox msgstr(msgColumnDoesNotExistOnTable, Array(getListObject(table).name, column))
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
                    .DataBodyRange.cells(row, col).Value
            Next
        Else
            dataNum = UBound(attrs)
            ReDim val(dataNum)
            On Error GoTo Err
            For i = 0 To dataNum
                val(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).Value
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
Function getSettings(ByVal rng As Variant, _
                Optional orientation As XlOrientation = xlHorizontal) As Object
    Dim row, col As Long
    If Not IsObject(rng) Then Set rng = Range(rng)
    Set getSettings = CreateObject("Scripting.Dictionary")
    With rng
        If orientation = xlHorizontal Then
            For col = 1 To .columns.count
                getSettings.item(.cells(1, col).Text) = .cells(2, col).Value
            Next
        Else
            For row = 1 To .rows.count
                getSettings.item(.cells(row, 1).Text) = .cells(row, 2).Value
            Next
        End If
    End With
End Function

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

Public Sub a_resetDoMacro()
    Call doMacro(, , True)
End Sub


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
                cell.Offset(0, colHead(i) - 1).Value = words(i)
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
                joinRow = joinRow & """" & .cells(row, col).Text & """"
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
    Dim tcel As Range
    
    If Not readHeader(head, cell, fh, fileName, colHead) Then
        Exit Function
    End If
    '   キー列タイトルの列インデックス（0から）とデータインデックス
    keycol = getColumnIndex(key, cell) - head.column
    For i = 0 To UBound(colHead)
        If head.cells(1, i + 1).Value = key Then
            keyidx = i
            Exit For
        End If
    Next
    
    Do Until EOF(fh)
        Line Input #fh, line
        If line = "" Then Exit Do
        words = splitLine(line)
        oldVal = cell.Offset(0, keycol).Value
        '   キーが異なる場合、キー列を先読みして一致をさがす。
        If oldVal <> words(keyidx) Then
            Set tcel = cell.Offset(1, keycol)
            Do While tcel.Value <> ""
                If tcel.Value = words(keyidx) Then
                    Set cell = tcel.Offset(0, -keycol)
                    oldVal = cell.Offset(0, keycol).Value
                    Exit Do
                End If
                Set tcel = tcel.Offset(1, 0)
            Loop
        End If
        If oldVal <> words(keyidx) Then
            '   新規キー値。行を挿入して書き込み
            If Not isTest Then
                Range(cell, cell.Offset(0, head.columns.count - 1)).Insert (xlShiftDown)
                For i = 0 To UBound(words)
                    If colHead(i) And Not isTest Then
                        cell.Offset(0, colHead(i) - 1).Value = words(i)
                    End If
                Next
            End If
            log = log & "Added " & words(keyidx) & vbCrLf
        Else
            '   既存キー値。各データを比較し、異なる場合はコールバック
            '   コールバック先で上書きするかも
            diff = ""
            For i = 0 To UBound(words)
                If colHead(i) Then
                    oldVal = cell.Offset(0, colHead(i) - 1).Value
                    If IsNumeric(words(i)) Then newVal = val(words(i)) Else newVal = words(i)
                    If oldVal <> newVal Then
                        slog = ""
                        If IsArray(callback) Then
                            slog = CallByName(callback(0), callback(1), VbMethod, _
                                Array(cell.Offset(0, colHead(i) - 1), newVal, callback(2)))
                        End If
                        If slog = "" Then
                            slog = oldVal & "->" & newVal
                        End If
                        If diff <> "" Then diff = diff & ","
                        diff = diff & head.cells(1, colHead(i)).Text _
                                & "(" & slog & ")"
                    End If
                End If
            Next
            If diff <> "" Then
                log = log & "Difference at """ & words(keyidx) _
                        & """ " & diff & vbCrLf
            End If
        End If
        Set cell = cell.Offset(1, 0)
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
    ChDrive left(ThisWorkbook.Path, 2)
    ChDir ThisWorkbook.Path & "\"
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

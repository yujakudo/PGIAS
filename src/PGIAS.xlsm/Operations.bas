Attribute VB_Name = "Operations"

Option Explicit

'   Win32
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long

'   入力規則のリストをセルに設定する
Sub setInputList(Optional ByVal target As Range = Nothing, _
                        Optional ByVal lst As String = "", _
                        Optional ByVal showList = False)
    Static lastTarget As Range
    '   先の規則をクリア
    If Not lastTarget Is Nothing Then
        On Error GoTo ValidationClearSkip
        lastTarget.Validation.Delete
ValidationClearSkip:
        On Error GoTo 0
        Set lastTarget = Nothing
    End If
    If target Is Nothing Then Exit Sub
    '   規則の設定
    With target
        .Validation.Delete
        If lst <> "" Then
            .Validation.Add Type:=xlValidateList, Formula1:=lst
            If .text <> "" And Not InStr(lst, .text) > 0 Then .value = ""
            Set lastTarget = target
            If showList Then SendKeys "%{Down}"
        End If
    End With
End Sub

'   種族名の予測変換候補の入力規則設定
Public Function speciesExpectation(ByVal target As Range) As Boolean
    Dim txt, part(1) As String
    Dim pos As Integer
    Dim rng As Range
    If IsNull(target.text) Then Exit Function
    txt = target.text
    txt = Replace(txt, "（", "(")
    pos = InStr(txt, "(")
    If pos > 0 Then
        part(0) = left(txt, pos)
        part(1) = Replace(Mid(txt, pos + 1), "）", ")")
        txt = StrConv(part(0), vbKatakana) & part(1)
    Else
        txt = StrConv(txt, vbKatakana)
    End If
    Set rng = shSpecies.ListObjects(1).ListColumns(SPEC_Name).Range
    speciesExpectation = rangeExpectation(target, rng, txt)
End Function

'   個体の予測変換候補の入力規則設定
Public Function individualExpectation(ByVal target As Range) As Boolean
    Dim txt As String
    Dim rng As Range
    Set rng = shIndividual.ListObjects(1).ListColumns(IND_Nickname).Range
    individualExpectation = rangeExpectation(target, rng)
    If Not individualExpectation Then
        txt = StrConv(target.text, vbKatakana)
        individualExpectation = rangeExpectation(target, rng, txt)
    End If
End Function

'   range指定での入力規則設定
Private Function rangeExpectation(ByVal target As Range, _
                ByVal rng As Range, _
        Optional ByVal txt As String = "") As Boolean
    Dim first, cell As Range
    Dim cand(), stmp As String
    Dim clim, cnum, i As Integer
    
    rangeExpectation = False
    If txt = "" Then txt = target.text
    Set first = rng.Find(txt, LookAt:=xlPart)
    If first Is Nothing Then Exit Function
    Set cell = first
    cnum = 0
    clim = 10
    ReDim cand(clim)
    Do
        If InStr(cell.text, txt) = 1 Then
            If cnum > clim Then
                clim = clim * 2
                ReDim Preserve cand(clim)
            End If
            cand(cnum) = cell.text
            i = cnum - 1
            cnum = cnum + 1
            Do While i >= 0
                If StrComp(cand(i), cand(i + 1), vbTextCompare) <= 0 Then Exit Do
                stmp = cand(i): cand(i) = cand(i + 1): cand(i + 1) = stmp
                i = i - 1
            Loop
        End If
        Set cell = rng.FindNext(cell)
    Loop While cell <> first
    If cnum = 0 Then Exit Function
    '   最初の候補に設定
    Call enableEvent(False)
    target.value = cand(0)
    '   複数候補があるならリストに設定
    
    If cnum > 1 Then
        ReDim Preserve cand(cnum - 1)
        Call setInputList(target, Join(cand, ","))
        On Error GoTo Continue
        target.Select
        SendKeys "%{Down}"
Continue:
        On Error GoTo 0
    End If
    Call enableEvent(True)
    rangeExpectation = True
End Function

'   わざの選択。入力規則の設定
Sub AtkSelected(ByVal target As Range, _
                Optional ByVal atkClass As Integer = -1, _
                Optional ByVal species As String = "")
    Dim lo As ListObject
    Dim cname As String
    Dim lst As Variant
    
    If GetAsyncKeyState(vbKeyEscape) Then
        Call setInputList(target)
        Exit Sub
    End If
    Set lo = target.Parent.ListObjects(1)
    cname = lo.HeaderRowRange.cells(target.column).text
    If atkClass < 0 Then atkClass = getAtkClassIndex(cname)
    If species = "" Then species = getColumn(C_SpeciesName, target).text
    If Not speciesExists(species) Then Exit Sub
    lst = getAtkNames(species, True, True)
    '   何らかの入力があり、リストにないなら設定しないで終了
    If target.text <> "" And InStr(lst(atkClass), target.text) < 1 Then Exit Sub
    ' ゲージ2と目標技には、未選択に戻す空白を追加
    If cname = IND_SpecialAtk2 Or cname = IND_TargetNormalAtk _
        Or cname = IND_TargetSpecialAtk Then lst(atkClass) = "　," & lst(atkClass)
    Call setInputList(target, lst(atkClass))
End Sub

'   わざを見るボタンの処理
Sub ClickShowAttack()
    Dim colTitle As String
    Dim sh As Worksheet
    
    With ActiveCell
        If .ListObject Is Nothing Then Exit Sub
        If .row < .ListObject.DataBodyRange.row Then Exit Sub
        Call doMacro(msgSelectingAttack)
        If selectSpeciesForAtkTable() Then
            colTitle = .ListObject.HeaderRowRange.cells(1, .column).text
            Set sh = shNormalAttack
            If getAtkClassIndex(colTitle) = C_IdSpecialAtk Then
                Set sh = shSpecialAttack
            End If
            Call jumpTo(sh)
        End If
        Call doMacro
    End With
End Sub

'   種族が選択されわざ表を加工する
Function selectSpeciesForAtkTable() As Boolean
    Dim species As String
    selectSpeciesForAtkTable = False
    species = getSpeciesFromCell()
    If species = "" Then Exit Function
    Call setSpeciesOnAtk(species)
    Call ShowCorrColumns(True)
    Call filterAtkBySpecies(species)
    selectSpeciesForAtkTable = True
End Function

'   セルより種族名を得る
Function getSpeciesFromCell(Optional ByVal cel As Range = Nothing) As String
    Dim row, col, lrow As Long
    getSpeciesFromCell = ""
    '   行、種族名の取得と、検証
    If cel Is Nothing Then Set cel = ActiveCell
    If IsEmpty(cel.ListObject) Then Exit Function
    getSpeciesFromCell = getColumn(C_SpeciesName, cel).text
End Function

'   種族の選択をリセット
Sub deselectSpecies()
    Call setSpeciesOnAtk("")
    Call filterAtkReset
    Call ShowCorrColumns(False)
End Sub

'   わざの表を種族でフィルタ
Function filterAtkBySpecies(ByVal species As String) As Boolean
    Dim atkCol, tbls As Variant
    Dim i As Long
    
    filterAtkBySpecies = False
    tbls = Array(TBL_NormalAtk, TBL_SpecialAtk)
    atkCol = getAtkNames(species, False, True)
    If Not IsArray(atkCol) Then Exit Function
    For i = 0 To 1
        Range(tbls(i)).AutoFilter Field:=1, _
            Criteria1:=atkCol(i), Operator:=xlFilterValues
    Next
    filterAtkBySpecies = True
End Function

'   種族でのフィルタのリセット
Sub filterAtkReset()
    Range(TBL_NormalAtk).AutoFilter Field:=1
    Range(TBL_SpecialAtk).AutoFilter Field:=1
End Sub

'   補正の列の表示/非表示
Public Sub ShowCorrColumns(ByVal isShow As Boolean)
    Dim rng As Range
    Dim sh(2) As Worksheet
    Dim col, i, cols As Long
    Set sh(0) = shNormalAttack
    Set sh(1) = shSpecialAttack
    cols = 5
    For i = 0 To 1
        col = getColumnIndex(ATK_typeMatch, sh(i))
        With sh(i).ListObjects(1).DataBodyRange
            Set rng = Range(.cells(1, col), .cells(1, col + cols - 1))
            If isShow Then
                rng.EntireColumn.Hidden = False
            Else
                rng.EntireColumn.Hidden = True
            End If
        End With
    Next
End Sub

'   わざのシートに種族名を書き込む
Private Sub setSpeciesOnAtk(ByVal species As String)
    Dim dspCol, stype As Variant
    Dim i, j, gcol As Long
    Dim cur As Worksheet
    
    Set cur = ActiveSheet
    dspCol = Array(R_NormalAtkSpeciesSelect, R_SpecialAtkSpeciesSelect)
    stype = Array("", "")
    If species <> "" Then
        stype = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    End If
    For i = 0 To 1
        With Range(dspCol(i))
            .value = species
            .Offset(0, 1).value = stype(0)
            .Offset(0, 1).Font.Color = getTypeColor(stype(0))
            .Offset(0, 2).value = stype(1)
            .Offset(0, 2).Font.Color = getTypeColor(stype(1))
            '   わざのシートをアクティブにして、カーソルを見えるように移動
            With .Parent
                .Activate
                gcol = 2
'                If species <> "" Then
'                    gcol = WorksheetFunction.Match(ATK_typeMatch, _
'                            .ListObjects(1).HeaderRowRange, 0)
'                    Application.Calculate '要検討
'                End If
                Application.Goto .cells(1, gcol), True
            End With
        End With
    Next
    cur.Activate
End Sub

'   タイプ、わざの色付け
Sub setTypeColorsOnTableColumns(ByVal table As Variant, _
            ByVal columns As Variant, _
            Optional ByVal atkClass As Variant = "", _
            Optional ByVal isCsv As Boolean = False)
    Dim lo As ListObject
    Dim i, row, pos, comma As Long
    Dim col As Variant
    
    Set lo = getListObject(table)
    If IsNumeric(atkClass) Then atkClass = atkClassArray()(atkClass)
    For i = 0 To UBound(columns)
        col = columns(i)
        If Not IsNumeric(col) Then
            col = getColumnIndex(col, lo)
        End If
        With lo.DataBodyRange
            For row = 1 To .rows.count
                If row = 254 Then
                    row = row
                End If
                Call setTypeColorsOnCell(.cells(row, col), atkClass, isCsv)
            Next
        End With
    Next
End Sub

Public Sub setTypeColorsOnCell(ByVal cell As Range, _
            Optional ByVal atkClass As String = "", _
            Optional ByVal isCsv As Boolean = False)
    Dim stp, val As String
    Dim cc, comma, pos As Long
    
    If cell.text = "" Then Exit Sub
    If isCsv Then
        cell.Font.Color = rgbBlack
        val = cell.text
        comma = 0
        On Error GoTo Err
        Do While comma <= Len(val)
            pos = comma + 1
            comma = InStr(pos, val, ",")
            If comma < 1 Then comma = Len(val) + 1
            stp = Trim(Mid(val, pos, comma - pos))
            If atkClass <> "" Then stp = getAtkAttr(atkClass, stp, ATK_Type)
            cc = getTypeColor(stp)
            cell.Characters(start:=pos, Length:=comma - pos).Font.Color = cc
        Loop
        On Error GoTo 0
    Else
        stp = cell.text
        If atkClass <> "" Then stp = getAtkAttr(atkClass, stp, ATK_Type)
        cc = getTypeColor(stp)
        If cc Then
            cell.Font.Color = cc
        Else
            cell.Font.Color = rgbBlack
        End If
    End If
    Exit Sub
Err:
End Sub

'   わざの変更。色を付ける
Sub AtkChange(ByVal target As Range, _
        Optional ByVal isInput As Boolean = True, _
        Optional ByVal atkClass As Variant = -1, _
        Optional ByVal species As String = "")
    Dim typeColor As Long
    Dim atkType, atk As String
    Dim cel As Range
    
    '   クラス（通常/ゲージ）の取得
    If atkClass < 0 Then atkClass = getAttackClassByHeader(target)
    typeColor = 0
    atk = target.text
    If atk <> "" Then
        On Error GoTo unknownAttack
        atkType = getAtkAttr(atkClass, atk, C_Type)
        On Error GoTo 0
        typeColor = getTypeColor(atkType)
    End If
    If typeColor Then
        target.Font.Color = typeColor
        If isInput Then
            On Error GoTo addAttack
            If target.Validation.Type = xlValidateList Then Exit Sub
addAttack:
            On Error GoTo 0
            If species = "" Then species = getSpeciesFromCell(target)
            Call shSpecies.addAttackToSpecies(atkClass, atk, species)
        End If
    Else
        target.Font.ColorIndex = 1
    End If
    Exit Sub
unknownAttack:
    MsgBox msgstr(msgUnknownAttackName, Array(atkClass, atk))
    target.value = ""
End Sub

'   ヘッダタイトルより、技クラス（通常/ゲージ）の取得
Private Function getAttackClassByHeader(ByVal target As Range) As Integer
    Dim cel As Range
    Dim atkClass As String
    If target.ListObject Is Nothing Then
        '   Targetがテーブルの外なので、範囲内まで遡る
        Set cel = target.Offset(-1, 0)
        While cel.ListObject Is Nothing And cel.row > 1
            Set cel = cel.Offset(-1, 0)
        Wend
        If cel.row = 1 Then Exit Function
        atkClass = cel.ListObject.HeaderRowRange(1, target.column).text
    Else
        atkClass = target.ListObject.HeaderRowRange(1, target.column).text
    End If
    If InStr(atkClass, C_SpecialAttack) > 0 Then
        getAttackClassByHeader = C_IdSpecialAtk
    Else
        getAttackClassByHeader = C_IdNormalAtk
    End If
End Function

'   天候の変更。色を変える
Public Sub weatherChange(ByVal target As Range)
    Dim idx As Integer
    On Error GoTo Err
    idx = WorksheetFunction.Match(target.text, _
            Range(R_WeatherTable).columns(1), 0)
    On Error GoTo 0
    target.Font.Color = Range(R_WeatherTable).cells(idx, 1).Font.Color
    Exit Sub
Err:
    target.Font.ColorIndex = 1
End Sub

'   テーブルのヘッダタイトルのサフィックスの表示切替
Public Sub switchHeaderSuffixes(Optional ByVal show As Boolean = False)
    Dim shi As Long
    For shi = 1 To Worksheets.count
        If Worksheets(shi).ListObjects.count > 0 Then
            Call switchHeaderSuffixesATable( _
                    Worksheets(shi).ListObjects(1), show)
        End If
    Next
End Sub

Public Sub switchHeaderSuffixesATable(ByVal lo As ListObject, _
                Optional ByVal show As Boolean = False)
    Dim col, cc, defc, pos As Long
    With lo.HeaderRowRange
        defc = .cells(1, 1).Font.Color
        For col = 1 To .columns.count
            With .cells(1, col)
                cc = defc
                If Not show Then cc = .Interior.Color
                pos = InStr(.text, "_")
                If pos > 0 Then
                    .Characters(start:=pos).Font.Color = cc
                End If
            End With
        Next
    End With
End Sub

'   ソート
Public Sub sortTable(ByVal table As Variant, ByVal cols As Variant, _
            Optional ByVal order As XlSortOrder = xlAscending)
    Dim i As Long
    Dim lo As ListObject
    
    If Not IsArray(cols) Then cols = Array(cols)
    Set lo = getListObject(table)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    With lo.Sort
        With .SortFields
            .Clear
            For i = 0 To UBound(cols)
                .Add key:=lo.ListColumns(cols(i)).DataBodyRange, _
                    SortOn:=xlSortOnValues, _
                    order:=order, _
                    DataOption:=xlSortNormal
            Next
        End With
        .header = xlYes
        .MatchCase = False
        .orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'   時間と日付の書き込み
Public Sub setTimeAndDate(ByVal rng As Range, ByVal start As Double)
    Dim stime As Double
    stime = Timer - start
    rng.Areas(1).value = getTimeStr(stime, "'")
    rng.Areas(2).value = Now
End Sub

'   個体値の16進数を10進数に変換する
Public Sub decimalizeIndivValue(ByVal target As Range)
    Dim c As Integer
    c = Asc(UCase(left(target.text, 1))) - 55
    If 9 < c And c < 16 Then
        target.value = c
    ElseIf target.value = 0 And target.text <> "0" Then
        target.ClearContents
    End If
End Sub

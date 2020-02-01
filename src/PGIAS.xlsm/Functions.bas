Attribute VB_Name = "Functions"
Option Explicit

'   通常わざ、ゲージわざの文字の配列
Public Function atkClassArray() As Variant
    atkClassArray = Array(SPEC_NormalAttack, SPEC_SpecialAttack)
End Function

'   文字列よりわざクラスのインデックスを得る
Public Function getAtkClassIndex(ByVal cname As String)
    getAtkClassIndex = C_IdNormalAtk
    If InStr(cname, C_SpecialAttack) > 0 Then
        getAtkClassIndex = C_IdSpecialAtk
    End If
End Function

Public Function getAtkClassName(ByVal idx As Integer)
    getAtkClassName = Array(C_NormalAttack, C_SpecialAttack)(idx)
End Function

'   番号よりタイプ名の取得
Public Function getTypeName(ByVal no As Integer) As String
    getTypeName = ""
    With Range(R_Type)
        If 0 < no And no <= .rows.count Then
            getTypeName = .cells(no, 1).Text
        End If
    End With
End Function

'   番号よりタイプの略号の取得
Public Function getTypeSynonym(ByVal no As Integer) As String
    getTypeSynonym = ""
    With Range(R_TypeSynonym)
        If 0 < no And no <= .rows.count Then
            getTypeSynonym = .cells(no, 1).Text
        End If
    End With
End Function

'   番号よりタイプの色の取得（名前も可）
Public Function getTypeColor(ByVal no As Variant) As Long
    getTypeColor = 0
    If Not IsNumeric(no) Then no = getTypeIndex(no)
    With Range(R_Type)
        If 0 < no And no <= .rows.count Then
            getTypeColor = .cells(no, 1).Interior.Color
        End If
    End With
End Function

'   タイプ名より番号の取得
Public Function getTypeIndex(ByVal str As String) As Integer
    On Error GoTo Err
    getTypeIndex = WorksheetFunction.Match(str, Range(R_Type), 0)
Err:
End Function

Public Function getWeatherIndex(ByVal str As String) As Integer
    On Error GoTo Err
    getWeatherIndex = WorksheetFunction.Match(str, Range(R_WeatherTable).columns(1), 0)
Err:
End Function

'   タイプの全数
Public Function typesNum() As Long
    typesNum = Range(R_Type).cells.count
End Function


'   タイプの配列の取得
Public Function getTypeArray() As Variant
    Dim arr() As String
    Dim types As Range
    Dim row As Long
    Set types = Range(R_Type)
    ReDim arr(types.rows.count)
    For row = 1 To types.rows.count
        arr(row) = types.cells(row, 1).Text
    Next
    getTypeArray = arr
End Function

'   タイプより種族の候補範囲の取得
Public Function getSpecCandidate(ByVal type1 As String, ByVal type2 As String) As String
    Dim row, scol, tcol As Long
    Dim tstr As String
    Dim cell As Excel.Range
    getSpecCandidate = ""
    tstr = joinTypeName(type1, type2)
    If tstr = "" Then Exit Function
    With shClassifiedByType.ListObjects(1)
        getSpecCandidate = .DataBodyRange.cells( _
            WorksheetFunction.Match(tstr, .ListColumns(CBT_Type).DataBodyRange, 0), _
            WorksheetFunction.Match(CBT_Species, .HeaderRowRange, 0))
    End With
End Function

'   セルのある行にて、それ以降の文字列をカンマ区切りで結合して返す
'   lcolが0のときは空白でストップ
Public Function getListStr(ByVal Target As Excel.Range, ByVal lcol As Long) As String
    Dim row, col As Long
    Dim str As String
    With Target.Parent
        row = Target.row
        col = Target.column
        Do While lcol = 0 Or col <= lcol
            str = .cells(row, col).Text
            If str <> "" Then
                getListStr = getListStr & "," & str
            ElseIf lcol = 0 Then
                Exit Do
            End If
            col = col + 1
        Loop
        getListStr = Mid(getListStr, 2)
    End With
End Function


'   タイプ名をカンマ区切りで結合する
Public Function joinTypeName(ByVal type1 As String, _
                    Optional ByVal type2 As String = "") As String
    Dim t1, t2, tmpn As Long
    Dim tmps As String
    joinTypeName = ""
    t1 = getTypeIndex(type1)
    t2 = getTypeIndex(type2)
    If (t2 <> 0 And t2 < t1) Or t1 = 0 Then
        tmps = type1: type1 = type2: type2 = tmps
        tmpn = t1: t1 = t2: t2 = tmpn
    End If
    If t1 = 0 Then Exit Function
    joinTypeName = type1
    If t2 <> 0 And t2 <> t1 Then joinTypeName = type1 & "," & type2
End Function

'   種族名から連結したタイプ文字列を得る
Public Function getJoinedTypeName(ByVal species) As String
    Dim ts As Variant
    ts = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    getJoinedTypeName = joinTypeName(ts(0), ts(1))
End Function

'   タイプ2の入力候補の文字列を得る
Public Function getType2Candidate(ByVal type1 As String) As String
    Dim type2, tstr As String
    Dim i, lcol, lrow, idx As Long
    Dim cell, tnames As Excel.Range
    getType2Candidate = "　"
    If getTypeIndex(type1) = 0 Then Exit Function
    With shClassifiedByType
        lrow = .cells(1, 1).End(xlDown).row
        Set tnames = .Range(.cells(1, 1), .cells(lrow, 1))
        For i = 1 To typesNum()
            type2 = getTypeName(i)
            If type1 <> type2 Then
                tstr = joinTypeName(type1, type2)
                Set cell = tnames.Find(tstr, LookAt:=xlWhole)
                If Not cell Is Nothing Then getType2Candidate = getType2Candidate & "," & type2
            End If
        Next
    End With
End Function

'   種族名は有効か？
Public Function speciesExists(ByVal species As String) As Boolean
    speciesExists = False
    On Error GoTo Err
    Call WorksheetFunction.Match(species, _
        shSpecies.ListObjects(1).ListColumns(SPEC_Name).DataBodyRange, 0)
    speciesExists = True
Err:
End Function

'  種族の属性の取得
Public Function getSpcAttr(ByVal species As String, ByVal attr As String) As Variant
    With shSpecies.ListObjects(1)
        getSpcAttr = .DataBodyRange.cells( _
                WorksheetFunction.Match(species, .ListColumns(SPEC_Name).DataBodyRange, 0), _
                WorksheetFunction.Match(attr, .HeaderRowRange, 0))
    End With
End Function

'  種族の属性を複数取得する（すべてをハッシュか、一部を配列にて）
Public Function getSpcAttrs(ByVal species As String, _
                Optional ByVal attrs As Variant = Nothing) As Variant
    Dim row, col, i, num As Long
    Dim vals() As Variant
    With shSpecies.ListObjects(1)
        On Error GoTo Err
        row = WorksheetFunction.Match(species, .ListColumns(SPEC_Name).DataBodyRange, 0)
        If Not IsArray(attrs) Then
            Set getSpcAttrs = CreateObject("Scripting.Dictionary")
            For col = 1 To .HeaderRowRange.count
                getSpcAttrs.item(.HeaderRowRange.cells(1, col).Text) = _
                    .DataBodyRange.cells(row, col).Value
            Next
        Else
            num = UBound(attrs)
            ReDim vals(num)
            For i = 0 To num
                vals(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).Value
            Next
            getSpcAttrs = vals
        End If
    End With
    Exit Function
Err:
    MsgBox msgstr(msgKeyDoesNotExist, Array(shSpecies.name, SPEC_Name, species))
End Function

Public Function getAtkTable(ByVal atkClass As Variant) As ListObject
    Dim sh As Worksheet
    If IsNumeric(atkClass) Then
        Set sh = Array(shNormalAttack, shSpecialAttack)(atkClass)
    Else
        Set sh = Worksheets(atkClass)
    End If
    Set getAtkTable = sh.ListObjects(1)
End Function

'  技の属性の取得
Public Function getAtkAttr(ByVal atkClass As Variant, ByVal atkName As String, ByVal attr As String) As Variant
    With getAtkTable(atkClass)
            getAtkAttr = .DataBodyRange.cells( _
                WorksheetFunction.Match(atkName, .ListColumns(ATK_Name).DataBodyRange, 0), _
                WorksheetFunction.Match(attr, .HeaderRowRange, 0))
    End With
End Function

'  技の属性の複数取得（すべてをハッシュか、一部を配列にて）
Public Function getAtkAttrs(ByVal atkClass As Variant, _
                ByVal atkName As String, _
                Optional ByVal attrs As Variant = Nothing) As Variant
    Dim row, col, i, num As Long
    Dim vals() As Variant
    With getAtkTable(atkClass)
        row = WorksheetFunction.Match(atkName, .ListColumns(ATK_Name).DataBodyRange, 0)
        If Not IsArray(attrs) Then
            Set getAtkAttrs = CreateObject("Scripting.Dictionary")
            For col = 1 To .HeaderRowRange.count
                getAtkAttrs.item(.HeaderRowRange.cells(1, col).Text) = _
                    .DataBodyRange.cells(row, col).Value
            Next
        Else
            num = UBound(attrs)
            ReDim vals(num)
            For i = 0 To num
                vals(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).Value
            Next
            getAtkAttrs = vals
        End If
    End With
End Function

'   天候ブーストの取得
Public Function weatherBoost(ByVal tp As String) As String
    Dim idx As Long
    idx = getTypeIndex(tp)
    If idx Then weatherBoost = Range(R_WeatherBoost).cells(idx, 1).Text
End Function

'   天候ブーストの取得（２つ、種族用）
Public Function weatherBoost2(ByVal tp1 As String, ByVal tp2 As String) As String
    Dim b(2), stmp As String
    b(1) = weatherBoost(tp1): b(2) = weatherBoost(tp2)
    If b(2) <> "" And b(2) < b(1) Then stmp = b(1): b(1) = b(2): b(2) = stmp
    If b(1) = 0 Then Exit Function
    weatherBoost2 = b(1)
    If b(2) <> "" And b(2) <> b(1) Then weatherBoost2 = b(1) & "," & b(2)
End Function

'   ブーストするか？
Public Function doesBoost(ByVal weather As String, ByVal boost As String) As Boolean
    weather = "," & weather & ","
    boost = "," & boost & ","
    doesBoost = False
    If InStr(boost, weather) > 0 Then doesBoost = True
End Function

'   PLの推定、種族名とHPより(推定値の一番小さいもの)
Public Function getPL( _
            ByVal species As String, ByVal CP As Long, ByVal HP As Long, _
            ByVal iatk As Long, ByVal idef As Long, ByVal ihp As Long) As Double
    Dim thp, cpg, min, max, ref, PL, cpm, lPL, uPL As Double
    Dim row As Long
    Dim attrs As Variant
    If species = "" Or CP = 0 Or HP = 0 Then Exit Function
    attrs = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
    thp = ihp + attrs(2)
    cpg = (iatk + attrs(0)) * Sqr(idef + attrs(1)) * Sqr(thp) / 10
    min = HP / thp: ref = Sqr(CP / cpg)
    If min < ref Then min = ref
    max = (HP + 1) / thp: ref = Sqr((CP + 1) / cpg)
    If max > ref Then max = ref
    With shCpm.ListObjects(1)
        On Error GoTo Err
        row = WorksheetFunction.Match(min, .ListColumns("CPM").DataBodyRange, -1)
        getPL = .ListColumns("PL").DataBodyRange.cells(row, 1)
        row = WorksheetFunction.Match(max, .ListColumns("CPM").DataBodyRange, -1)
        '上限のチェック
        uPL = .ListColumns("PL").DataBodyRange.cells(row, 1)
        If getPL > uPL Then getPL = 0
    End With
    Exit Function
'    With Range("CPM表")
'        For row = 1 To .Rows.Count
'            cpm = .cells(row, 2).Value
'            If min <= cpm And cpm < max Then
'                getPL = .cells(row, 1).Value
'                Exit Function
'            End If
'        Next
'    End With
Err:
    getPL = 0
End Function

'空白チェック
Public Function check(p1 As String, _
                    Optional p2 As String = "1", _
                    Optional p3 As String = "1", _
                    Optional p4 As String = "1", _
                    Optional p5 As String = "1", _
                    Optional p6 As String = "1", _
                    Optional p7 As String = "1", _
                    Optional p8 As String = "1", _
                    Optional p9 As String = "1") As Boolean
    If p1 = "　" Then p1 = ""
    If p2 = "　" Then p1 = ""
    If p3 = "　" Then p1 = ""
    If p4 = "　" Then p1 = ""
    If p5 = "　" Then p1 = ""
    If p6 = "　" Then p1 = ""
    If p7 = "　" Then p1 = ""
    If p8 = "　" Then p1 = ""
    If p9 = "　" Then p1 = ""
    check = p1 <> "" And p2 <> "" And p3 <> "" And p4 <> "" And p5 <> "" _
            And p6 <> "" And p7 <> "" And p8 <> "" And p9 <> ""
End Function
                    
'   CPの取得
Public Function getCP(ByVal species As String, _
                        Optional ByVal PL As Double = 40, _
                        Optional ByVal indATK As Long = 15, _
                        Optional ByVal indDEF As Long = 15, _
                        Optional ByVal indHP As Long = 15) As Long
    Dim spec As Object
    Dim cpm As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    cpm = getCPM(PL)
    getCP = Fix(((indATK + spec("ATK")) * Sqr(indDEF + spec("DEF")) _
                * Sqr(indHP + spec("HP")) * cpm ^ 2) / 10)
    
End Function

Public Function getCPbyHP(ByVal species As String, _
                        ByVal HP As Double, _
                        Optional ByVal PL As Double = 40, _
                        Optional ByVal indATK As Long = 15, _
                        Optional ByVal indDEF As Long = 15) As Variant
    Dim spec As Object
    Dim cpm, CP, ihp As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    cpm = getCPM(PL)
    CP = Fix(((indATK + spec("ATK")) * Sqr(indDEF + spec("DEF")) _
                * Sqr(HP / cpm) * cpm ^ 2) / 10)
    ihp = HP / cpm - spec("HP")
    getCPbyHP = Array(CP, Int(ihp))
End Function

'   CPほかよりHPの取得
Public Function getHPbyCP(ByVal species As String, _
                        ByVal CP As Double, _
                        Optional ByVal PL As Double = 40, _
                        Optional ByVal indATK As Long = 15, _
                        Optional ByVal indDEF As Long = 15) As Variant
    Dim spec As Object
    Dim cpm, HP, ihp As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    cpm = getCPM(PL)
    HP = 100 * CP ^ 2 / (indATK + spec("ATK")) ^ 2 / (indDEF + spec("DEF")) / cpm ^ 3
    ihp = HP / cpm - spec("HP")
    getHPbyCP = Array(Int(HP), Int(ihp))
End Function

                    
'   CMPの取得
Public Function getCPM(ByVal PL As Double) As Double
    If PL = 0 Then Exit Function
    getCPM = WorksheetFunction.VLookup(PL, _
                shCpm.ListObjects(1).DataBodyRange, 2, False)
End Function

'   攻撃力、防御力、HPの計算
Public Function getPower(ByVal species As String, ByVal attr As String, ByVal ind As Long, ByVal PL As Double) As Double
    If species = "" Then Exit Function
    getPower = (ind + getSpcAttr(species, attr)) * getCPM(PL)
End Function

'   技の名前の配列を得る
Public Function getAtkNames(ByVal species As String, _
            Optional ByVal isCsv As Boolean = False, _
            Optional ByVal withLimited As Boolean = False) As Variant
    Dim i As Long
    Dim arr As Variant
    arr = getSpcAttrs(species, _
            Array(SPEC_NormalAttack, SPEC_SpecialAttack, _
            SPEC_NormalAttackLimited, SPEC_SpecialAttackLimited))
    If Not IsArray(arr) Then Exit Function
    If withLimited Then
        arr(0) = margeList(arr(0), arr(2), isCsv)
        arr(1) = margeList(arr(1), arr(3), isCsv)
        ReDim Preserve arr(1)
    Else
        ReDim Preserve arr(1)
        If Not isCsv Then
            arr(0) = Split(arr(0), ",")
            arr(1) = Split(arr(1), ",")
        End If
    End If
    getAtkNames = arr
End Function

'   コンマ区切りの文字列をマージ
'   もとはソート済み
Private Function margeList(ByRef l1 As Variant, ByRef l2 As Variant, _
        Optional ByVal isCsv As Boolean) As Variant
    Dim arr(1) As Variant
    Dim idx(2), lim(2), ri, riLim, si As Integer
    Dim ret() As String
    If l1 = "" Then l1 = l2: l2 = ""
    If l2 = "" Then
        If Not isCsv Then
            margeList = Split(l1, ",")
        Else
            margeList = l1
        End If
        Exit Function
    End If
    arr(0) = Split(l1, ",")
    lim(0) = UBound(arr(0))
    arr(1) = Split(l2, ",")
    lim(1) = UBound(arr(1))
    riLim = lim(0) + lim(1) + 1
    ReDim ret(riLim)
    While ri <= riLim
        If idx(0) > lim(0) Then
            si = 1
        ElseIf idx(1) > lim(1) Then
            si = 0
        ElseIf StrComp(arr(0)(idx(0)), arr(1)(idx(1)), vbTextCompare) <= 0 Then
            si = 0
        Else
            si = 1
        End If
        ret(ri) = arr(si)(idx(si))
        idx(si) = idx(si) + 1
        ri = ri + 1
    Wend
    If isCsv Then
        margeList = Join(ret, ",")
    Else
        margeList = ret
    End If
End Function
        

'   セルに複数タイプを書き込む。旧setTypeStr
Public Function setTypeToCell(ByVal types As Variant, ByVal Target As Excel.Range, _
        Optional ByVal synonym As Boolean = False)
    Dim i, j, ti, num As Long
    Dim str As String
    Dim items() As Variant
    num = UBound(types)
    ReDim items(num, 3)
    For i = 0 To UBound(types)
        ti = types(i)
        If Not IsNumeric(ti) Then ti = getTypeIndex(ti)
        If ti = 0 Then Exit For
        If synonym Then
            items(i, 0) = getTypeSynonym(ti)
            items(i, 1) = Len(str) + 1
            str = str & items(i, 0)
        Else
            items(i, 0) = getTypeName(ti)
            If str <> "" Then str = str & ","
            items(i, 1) = Len(str) + 1
            str = str & items(i, 0)
        End If
        items(i, 2) = Len(items(i, 0))
        items(i, 3) = getTypeColor(ti)
    Next
    Call writeColoredStr(Target, str, items)
End Function

'   セルに複数のわざ名を色付きで書き込む
Public Function setAtkNames(ByVal atkClass As Variant, ByVal atkNames As Variant, ByVal Target As Excel.Range)
    Dim items() As Variant
    Dim i, j, ti As Long
    Dim line, name As String
    If Not IsArray(atkNames) Then atkNames = Split(atkNames, ",")
    ReDim items(UBound(atkNames), 3)
    For i = 0 To UBound(atkNames)
        name = Trim(atkNames(i))
        ti = getTypeIndex(getAtkAttr(atkClass, name, C_TYPE))
        If line <> "" Then line = line & ","
        items(i, 1) = Len(line) + 1
        items(i, 2) = Len(name)
        items(i, 3) = getTypeColor(ti)
        line = line & name
    Next
    Call writeColoredStr(Target, line, items)
End Function

'   文字列を色付きで書き込む
Public Function writeColoredStr(ByVal Target As Excel.Range, _
                    ByVal str As String, ByVal info As Variant)
    Dim i As Long
    With Target
        .Value = str
        For i = 0 To UBound(info, 1)
            If info(i, 1) < 1 Then Exit For
            .Characters(start:=info(i, 1), Length:=info(i, 2)).Font.Color = info(i, 3)
        Next
    End With
End Function

'   種族より色の配列を得る
Function getColorForSpecies(ByVal species As String) As Variant
    Dim c(1), row As Long
    Dim cell As Range
    row = searchRow(species, SPEC_Name, shSpecies)
    If row = 0 Then
        getColorForSpecies = c
        Exit Function
    End If
    With shSpecies.ListObjects(1).DataBodyRange
        Set cell = getColumn(SPEC_Type1, .cells(row, 1))
        c(0) = cell.Font.Color
        c(1) = c(0)
        With cell.Offset(0, 1)
            If .Text <> "" Then c(1) = .Font.Color
        End With
    End With
    getColorForSpecies = c
End Function

'   区切り文字の中から指定の場所の値を得る
Function splitStr(ByVal txt As String, ByVal pos As Integer, _
                    Optional ByVal def As Variant = "", _
                    Optional ByVal delimiter As String = ",") As Variant
    Dim arr As Variant
    splitStr = def
    arr = Split(txt, delimiter)
    If UBound(arr) < pos - 1 Then Exit Function
    splitStr = arr(pos - 1)
    If splitStr = "" Then splitStr = def
End Function

'   全体設定の取得
Public Function getGlobalSettings()
    Set getGlobalSettings = getSettings(R_GlobalSettings, xlVertical)
End Function


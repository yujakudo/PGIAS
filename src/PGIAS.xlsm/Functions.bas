Attribute VB_Name = "Functions"


Option Explicit

'   CPM算出係数
Private Const k01 As Double = 0.018852250938
Private Const k00 As Double = -0.010016250938
Private Const k11 As Double = 0.01783805135
Private Const k10 As Double = 1.25744942000031E-04
Private Const k21 As Double = 0.017849811806
Private Const k20 As Double = -1.09464177999993E-04
Private Const k31 As Double = 0.00891892158
Private Const k30 As Double = 0.267817242602

'   CPM算出係数の配列
Dim CPMK As Variant

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
            getTypeName = .cells(no, 1).text
        End If
    End With
End Function

'   番号よりタイプの略号の取得
Public Function getTypeSynonym(ByVal no As Integer) As String
    getTypeSynonym = ""
    With Range(R_TypeSynonym)
        If 0 < no And no <= .rows.count Then
            getTypeSynonym = .cells(no, 1).text
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
        arr(row) = types.cells(row, 1).text
    Next
    getTypeArray = arr
End Function

'   天候の名前或は略号よりインデックスを得る
Public Function getWeatherIndex(ByVal str As String) As Integer
    Dim col As Long
    If Len(str) = 1 Then col = 2 Else col = 1
    On Error GoTo Err
    getWeatherIndex = WorksheetFunction.Match(str, Range(R_WeatherTable).columns(col), 0)
Err:
End Function

'   天候のインデックスより、名前或は略号を得る
Public Function getWeatherName(ByVal idx As Integer, _
                Optional ByVal isSynonym As Boolean = False, _
                Optional ByVal withColor As Boolean = False) As Variant
    Dim col, wc As Long
    Dim wname As String
    If idx = 0 Then
        If withColor Then
            getWeatherName = Array("", 0)
        Else
            getWeatherName = ""
        End If
        Exit Function
    End If
    If isSynonym Then col = 2 Else col = 1
    With Range(R_WeatherTable).cells(idx, col)
        wname = .text
        If withColor Then
            wc = .Font.Color
            getWeatherName = Array(wname, wc)
        Else
            getWeatherName = wname
        End If
    End With
Err:
End Function

'   天候の文字列またはインデックスから両方を得る。
Public Function getWeatherNameAndIndex(ByRef weather As Variant) As String
    If Not IsNumeric(weather) Then
        getWeatherNameAndIndex = weather
        weather = getWeatherIndex(weather)
    Else
        getWeatherNameAndIndex = getWeatherName(weather)
    End If
End Function

'   天候の全数
Public Function weathersNum() As Long
    weathersNum = Range(R_WeatherTable).rows.count
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
            str = .cells(row, col).text
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
                getSpcAttrs.item(.HeaderRowRange.cells(1, col).text) = _
                    .DataBodyRange.cells(row, col).value
            Next
        Else
            num = UBound(attrs)
            ReDim vals(num)
            For i = 0 To num
                vals(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).value
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
                getAtkAttrs.item(.HeaderRowRange.cells(1, col).text) = _
                    .DataBodyRange.cells(row, col).value
            Next
        Else
            num = UBound(attrs)
            ReDim vals(num)
            For i = 0 To num
                vals(i) = .ListColumns(attrs(i)).DataBodyRange.cells(row, 1).value
            Next
            getAtkAttrs = vals
        End If
    End With
End Function

'   天候ブーストの取得
Public Function weatherBoost(ByVal tp As String) As String
    Dim idx As Long
    idx = getTypeIndex(tp)
    If idx Then weatherBoost = Range(R_WeatherBoost).cells(idx, 1).text
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
            ByVal species As String, ByVal CP As Long, ByVal hp As Long, _
            ByVal iatk As Long, ByVal idef As Long, ByVal ihp As Long) As Double
    Dim thp, cpg, min, max, ref, PL, CPM, lPL, uPL As Double
    Dim row As Long
    Dim attrs As Variant
    If species = "" Or CP = 0 Or hp = 0 Then Exit Function
    attrs = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
    thp = ihp + attrs(2)
    cpg = (iatk + attrs(0)) * Sqr(idef + attrs(1)) * Sqr(thp) / 10
    min = hp / thp: ref = Sqr(CP / cpg)
    If min < ref Then min = ref
    max = (hp + 1) / thp: ref = Sqr((CP + 1) / cpg)
    If max > ref Then max = ref
    With shCpm.ListObjects(1)
        On Error GoTo Err
        row = WorksheetFunction.Match(min, .ListColumns("CPM").DataBodyRange, -1)
        getPL = .ListColumns("PL").DataBodyRange.cells(row, 1)
        row = WorksheetFunction.Match(max, .ListColumns("CPM").DataBodyRange, -1)
        On Error GoTo 0
        '上限のチェック
        uPL = .ListColumns("PL").DataBodyRange.cells(row, 1)
        If getPL > uPL Then getPL = 0
    End With
    Exit Function
Err:
    getPL = 0
End Function

'   目標PLの取得(下のgetPLbyCPに移行)
Public Function getTargetPL( _
            ByVal species As String, ByVal TCP As Long, _
            ByVal iatk As Long, ByVal idef As Long, ByVal ihp As Long) As Double
    Dim cpg, tcpm, PL, CPM, CP As Double
    Dim row As Long
    Dim attrs As Variant
    If species = "" Or TCP = 0 Then Exit Function
    attrs = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
    cpg = (iatk + attrs(0)) * Sqr(idef + attrs(1)) * Sqr(ihp + attrs(2)) / 10
    tcpm = Sqr((TCP + 1) / cpg)
    With shCpm.ListObjects(1)
        On Error GoTo Proc0
        row = WorksheetFunction.Match(tcpm, .ListColumns("CPM").DataBodyRange, -1)
        GoTo Proc1
Proc0:
        row = 1
Proc1:
        On Error GoTo 0
        row = row - 1
        Do
            row = row + 1
            CPM = .ListColumns("CPM").DataBodyRange.cells(row, 1)
            CP = WorksheetFunction.Floor(cpg * CPM * CPM, 1)
        Loop While CP > TCP
        getTargetPL = .ListColumns("PL").DataBodyRange.cells(row, 1)
    End With
    Exit Function
Err:
    getTargetPL = 0
End Function

'   目標PLの取得
'   atk, def,hpは現在のPower、PLは現在のPL
Public Function getPLbyCP(ByVal TCP As Long, ByVal PL As Double, _
            ByVal atk As Double, _
            ByVal def As Double, _
            ByVal hp As Double) As Double
    Dim cpg, CPMc As Double
    If TCP = 0 Or PL = 0 Then Exit Function
    CPMc = getCPM(PL)
    cpg = atk * Sqr(def) * Sqr(hp) / CPMc ^ 2 / 10
    getPLbyCP = getPLbyCpg(TCP, cpg)
End Function

'   目標CPよりPLを得る2
'   atk, def, hpは個体値
Public Function getPLbyCP2(ByVal TCP As Long, _
            ByVal species As String, _
            ByVal atk As Double, _
            ByVal def As Double, _
            ByVal hp As Double) As Double
    Dim attr As Variant
    Dim cpg As Double
    attr = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
    cpg = (attr(0) + atk) * Sqr(attr(1) + def) * Sqr(attr(2) + hp) / 10
    getPLbyCP2 = getPLbyCpg(TCP, cpg)
End Function

'   CPG((種族値＋個体値)の重み積/10)よりPLを得る
Public Function getPLbyCpg(ByVal TCP As Long, ByVal cpg As Double) As Double
    Dim CPM, tPL As Double
    CPM = Sqr((TCP + 1) / cpg)
    tPL = getPLbyCPM(CPM, True)
    If getCPM(tPL) >= CPM Then tPL = tPL - 0.5
    If tPL < 1 Then tPL = 1
    If tPL > 40 Then tPL = 40
    getPLbyCpg = tPL
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
    Dim CPM As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    CPM = getCPM(PL)
    getCP = Fix(((indATK + spec("ATK")) * Sqr(indDEF + spec("DEF")) _
                * Sqr(indHP + spec("HP")) * CPM ^ 2) / 10)
    
End Function

Public Function getCPbyHP(ByVal species As String, _
                        ByVal hp As Double, _
                        Optional ByVal PL As Double = 40, _
                        Optional ByVal indATK As Long = 15, _
                        Optional ByVal indDEF As Long = 15) As Variant
    Dim spec As Object
    Dim CPM, CP, ihp As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    CPM = getCPM(PL)
    CP = Fix(((indATK + spec("ATK")) * Sqr(indDEF + spec("DEF")) _
                * Sqr(hp / CPM) * CPM ^ 2) / 10)
    ihp = hp / CPM - spec("HP")
    getCPbyHP = Array(CP, Int(ihp))
End Function

'   CPほかよりHPの取得
Public Function getHPbyCP(ByVal species As String, _
                        ByVal CP As Double, _
                        Optional ByVal PL As Double = 40, _
                        Optional ByVal indATK As Long = 15, _
                        Optional ByVal indDEF As Long = 15) As Variant
    Dim spec As Object
    Dim CPM, hp, ihp As Double
    If species = "" Then Exit Function
    Set spec = getSpcAttrs(species)
    CPM = getCPM(PL)
    hp = 100 * CP ^ 2 / (indATK + spec("ATK")) ^ 2 / (indDEF + spec("DEF")) / CPM ^ 3
    ihp = hp / CPM - spec("HP")
    getHPbyCP = Array(Int(hp), Int(ihp))
End Function

                    
'   CMPの取得
Public Function getCPM(ByVal PL As Double) As Double
    Dim i As Integer
    If PL = 0 Then Exit Function
    If Not IsArray(CPMK) Then Call makeCPMKcache
    i = Int(PL / 10)
    If i > 3 Then i = 3
    getCPM = Sqr(CPMK(i)(1) * PL + CPMK(i)(0))
'    getCPM = WorksheetFunction.VLookup(PL, _
'                shCpm.ListObjects(1).DataBodyRange, 2, False)
End Function

Public Function getPLbyCPM(ByVal CPM As Double, _
                Optional align As Boolean = True) As Double
    Dim i As Integer
    If CPM = 0 Then Exit Function
    If Not IsArray(CPMK) Then Call makeCPMKcache
    If CPM < 0.422500009990532 Then
        i = 0
    ElseIf CPM < 0.597400009994978 Then
        i = 1
    ElseIf CPM < 0.731700000001367 Then
        i = 2
    Else
        i = 3
    End If
    getPLbyCPM = (CPM ^ 2 - CPMK(i)(0)) / CPMK(i)(1)
    If getPLbyCPM < 0 Then getPLbyCPM = 0
    If align Then
        getPLbyCPM = Int(getPLbyCPM * 2 + 0.5) / 2
    End If
End Function

Private Sub makeCPMKcache()
    CPMK = Array(Array(k00, k01), Array(k10, k11), _
                Array(k20, k21), Array(k30, k31))
End Sub

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
        ti = getTypeIndex(getAtkAttr(atkClass, name, C_Type))
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
        .value = str
        For i = 0 To UBound(info, 1)
            If info(i, 1) < 1 Then Exit For
            .Characters(start:=info(i, 1), Length:=info(i, 2)).Font.Color = info(i, 3)
        Next
        .Font.Size = 10
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
            If .text <> "" Then c(1) = .Font.Color
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

'   セルに天候のセット
Public Function setWeatherToCell(ByVal cel As Range, _
                Optional ByVal weather As Variant = -1) As Integer
    Dim sWeather As String
    Dim clr As Long
    If weather = -1 Then
        sWeather = cel.text
        weather = getWeatherIndex(sWeather)
    Else
        sWeather = getWeatherNameAndIndex(weather)
        cel.value = sWeather
    End If
    If weather > 0 Then
        cel.Font.Color = Range(R_WeatherTable).cells(weather, 1).Font.Color
    Else
        cel.Font.Color = 0
    End If
    setWeatherToCell = weather
End Function

'   自動目標（リーグ名）によるCP上限の取得
Public Function getCpUpper(ByVal league As String, _
                    Optional ByVal undefVal As Long = C_MaxLong) As Long
    getCpUpper = undefVal
    If league = C_League1 Then
        getCpUpper = C_UpperCPl1
    ElseIf league = C_League2 Then
        getCpUpper = C_UpperCPl2
    End If
End Function

'   自動目標に伴う設定の取得
Public Function getAutoTargetSettings(ByVal mode As String) As Object
    Dim subset As Object
    Set subset = CreateObject("Scripting.Dictionary")
    If mode = C_None Or mode = "" Then Exit Function
    If mode = C_League1 Or mode = C_League2 Or mode = C_League3 Then
        subset.item(C_YAxis) = left(IND_MtcSpecialAtk1CDPS, Len(IND_MtcSpecialAtk1CDPS) - 1) & "*"
        subset.item(C_YPrediction) = IND_prMtcCDPS
        subset.item(C_CpUpper) = getCpUpper(mode, 0)
        If subset(C_CpUpper) > 0 Then
            subset.item(C_PrCpLower) = subset(C_CpUpper) - 300
        Else    '   League3
            subset.item(C_CpUpper) = ""
            subset.item(C_PrCpLower) = 2800
        End If
        
        subset.item(C_SimMode) = C_Match
        subset.item(C_SelfAtkDelay) = 0
        subset.item(C_EnemyAtkDelay) = 0
        subset.item(CR_SetRankNum) = 5
        subset.item(CR_SetRankVar) = C_KTR
    Else
        subset.item(C_YAxis) = left(IND_GymSpecialAtk1CDPS, Len(IND_GymSpecialAtk1CDPS) - 1) & "*"
        subset.item(C_YPrediction) = IND_prGymCDPS
        subset.item(C_CpUpper) = ""
        subset.item(C_PrCpLower) = 2000
    
        subset.item(C_SimMode) = C_Gym
        subset.item(C_SelfAtkDelay) = 0.1
        subset.item(C_EnemyAtkDelay) = 2
        subset.item(CR_SetRankNum) = 6
        subset.item(CR_SetRankVar) = C_KT
    End If
    '   対策シート項目
    subset.item(CR_DefCpUpper) = subset.item(C_CpUpper)
    subset.item(CR_DefCpLower) = subset.item(C_PrCpLower)
    Set getAutoTargetSettings = subset
End Function

'   アメと星の砂の必要量
Public Function getResourceRequirment(ByVal curPL As Double, _
                ByVal prPL As Double) As Variant
    Dim candies, sands As Long
    Dim cel As Variant
    For Each cel In shCpm.ListObjects(1).ListColumns("PL").DataBodyRange
        If curPL < cel.value And cel.value <= prPL Then
            sands = sands + cel.Offset(0, 2).value
            candies = candies + cel.Offset(0, 3).value
        End If
    Next
    getResourceRequirment = Array(candies, sands)
End Function

'   耐久力
Public Function getEndurance(ByVal def As Double, ByVal hp As Double) As Double
    getEndurance = Fix(hp) / (1000 / def + 1)
End Function

Attribute VB_Name = "tableOperation"

'   種族表関連シートのコンボボックスメニュー設定
Public Sub setComboMenuOfSpeciesTable(ByRef sh As Worksheet, _
                                ByRef cmb As ComboBox)
    Dim shs, cmds As Variant
    shs = Array(shSpecies, shSpeciesAnalysis1, shNormalAttack, shSpecialAttack)
    shs = joinArray(shs, getSheetsByName(SBL_R_Settings))
    shs = joinArray(shs, getSheetsBySettingValue(SMAP_R_Settings, C_SheetName, sh.name))
    shs = copyArray(shs, sh)
    cmds = Array(cmdFilterReset, cmdSortReset)
    Call setComboMenu(cmb, shs, Nothing, cmds)
End Sub

'   コンボボックスのシート選択メニュー作成
Public Sub setComboMenu(ByRef cmb As ComboBox, _
                    Optional ByRef shs As Variant = Nothing, _
                    Optional ByRef names As Variant = Nothing, _
                    Optional ByRef commands As Variant = Nothing)
    Dim sh, nm, sheets As Variant
    With cmb
        .Clear
        If IsArray(shs) Then
            For Each sh In shs
                .AddItem sh.name
            Next
        End If
        If IsArray(names) Then
            For Each nm In names
                sheets = getSheetsByName(nm)
                For Each sh In sheets
                    .AddItem sh.name
                Next
            Next
        End If
        If IsArray(commands) Then
            For Each nm In commands
                .AddItem nm
            Next
        End If
    End With
End Sub

'   コンボボックスの値よりワークシートを得る
Public Function getSheetFromCombo(ByRef cmb As ComboBox) As Worksheet
    Dim sh As Variant
    With cmb
        For Each sh In Worksheets
            If .value = sh.name Then
                Set getSheetFromCombo = sh
                Exit For
            End If
        Next
    End With
End Function

'   種族、種族分析シートに移動
Public Sub jumpToSpeciesSheet(ByVal sh As Worksheet, Optional ByVal both As Boolean = True)
    Dim sha As Worksheet
    Dim species As String
    If sh Is shSpeciesAnalysis1 Then Set sha = shSpecies Else Set sha = shSpeciesAnalysis1
    species = getSpeciesFromCell()
    If species <> "" Then
        If both Then Call activateSpeciesSheet(sha, species, False)
        Call activateSpeciesSheet(sh, species, True)
    End If
End Sub

'   種族、種族分析シートに移動
Private Sub activateSpeciesSheet(ByVal sh As Worksheet, _
                    ByVal species As String, ByVal log As Boolean)
    Dim row As Long
    With sh
        row = searchRow(species, C_SpeciesName, .ListObjects(1))
        Call jumpTo(.ListObjects(1).DataBodyRange.cells(row, 1), log)
    End With
End Sub

'   わざシートに移動
Public Sub jumpToAttackSheet(ByVal sh As Worksheet)
    Call doMacro(msgSelectingAttack)
    Call selectSpeciesForAtkTable
    Call doMacro
    Call jumpTo(sh, True)
End Sub

'   個体マップに移動
Public Sub jumpToIndMap(ByVal sh As Worksheet, ByVal sameType As Boolean)
    Dim species As String
    Dim name As String
    '   絞り込み処理
    species = getSpeciesFromCell()
    name = getColumn(IND_Nickname, ActiveCell).text
    If species <> "" Then
        If sameType Then
            Call setSameTypeToMap(species, sh.Range(IMAP_R_TypeSelect))
        Else
            Call setSameTypeToMap("", sh.Range(IMAP_R_TypeSelect))
        End If
        sh.Range(IMAP_R_IndivSelect).value = name
    End If
    Call jumpTo(sh)
End Sub

'   種族マップに移動
Public Sub jumpToSpecMap(ByVal sh As Worksheet, ByVal sameType As Boolean)
    Dim species As String
    Dim stype As Variant
    species = getSpeciesFromCell()
    If species <> "" Then
        If sameType Then
            Call setSameTypeToMap(species, sh.Range(R_SpeciesMapTypeSelect))
        Else
            Call setSameTypeToMap("", sh.Range(R_SpeciesMapTypeSelect))
        End If
        sh.Range(R_SpeciesMapSpeciesSelect).value = species
    End If
    Call jumpTo(sh)
End Sub

'   コンボボックスのコマンドの実行
Public Sub execCombCommand(ByRef sh As Worksheet, ByRef cmb As ComboBox, _
                            ByVal sameType As Boolean, _
                        Optional ByVal flterInd As Range = Nothing)
    Dim shTo As Worksheet
    Dim species As String
    If cmb.value = cmdFilterReset Then
        Call resetTableFilter(sh)
        If Not flterInd Is Nothing Then
            flterInd.ClearContents
        End If
    ElseIf cmb.value = cmdSortReset Then
        Call CallByName(sh, "sortNormally", VbMethod)
    Else
        Set shTo = getSheetFromCombo(cmb)
        If Not shTo Is Nothing Then Call jumpConsideringSpecies(shTo, sameType)
    End If
    Call enableEvent(False)
    cmb.value = ""
    Call enableEvent(True)
End Sub

'   他のシートに移動
Private Sub jumpConsideringSpecies(ByVal sh As Object, ByVal sameType As Boolean)
    Dim sw As Boolean
    '   複数セル選択かテーブルデータ外なら単に移動
    If ActiveCell.CountLarge <> 1 Or _
        Application.Intersect(ActiveCell, ActiveSheet.ListObjects(1).DataBodyRange) Is Nothing Then
        Call jumpTo(sh)
    '   種族か種族分析か種族リーグ別
    ElseIf sh Is shSpecies Or sh Is shSpeciesAnalysis1 _
            Or checkNameInSheet(sh, SBL_R_Settings) Then
        If ActiveSheet Is shSpecies Or ActiveSheet Is shSpeciesAnalysis1 Then
            Call jumpToSpeciesSheet(sh, False)
        Else
            Call jumpToSpeciesSheet(sh, True)
        End If
    '   わざ
    ElseIf sh Is shNormalAttack Or sh Is shSpecialAttack Then
        If selectSpeciesForAtkTable() Then Call jumpTo(sh, True)
    '  種族マップ
    ElseIf checkNameInSheet(sh, SMAP_R_Settings) Then
        Call jumpToSpecMap(sh, sameType)
    '  個体マップ
    ElseIf checkNameInSheet(sh, IMAP_R_Settings) Then
        Call jumpToIndMap(sh, sameType)
    Else
        Call jumpTo(sh)
    End If
End Sub

'   種族リーグ別設定の変更（保留）
Public Function sblChangeSettings(ByVal target As Range, _
                ByVal rng As Range) As Boolean
    Dim key As String
    sblChangeSettings = False
    key = target.Offset(-1, 0).text
    If key = C_League Then  '   リーグ
        Call sblSetSettingsByLeague(rng)
        sblChangeSettings = True
    End If
End Function

'   リーグ別に仮想敵のパワーを書き換え（保留）
Private Sub sblSetSettingsByLeague(ByVal rng As Range)
    Dim settings As Object
    Set settings = getSettings(rng)
    If settings(C_League) = C_League1 Then
        settings.item(SBL_AtkPower) = 167
        settings.item(SBL_DefPower) = 205
        settings.item(SBL_HP2) = 205
    ElseIf settings(C_League) = C_League2 Then
        settings.item(SBL_AtkPower) = 209
        settings.item(SBL_DefPower) = 237
        settings.item(SBL_HP2) = 250
    ElseIf settings(C_League) = C_League3 Then
        settings.item(SBL_AtkPower) = 249
        settings.item(SBL_DefPower) = 256
        settings.item(SBL_HP2) = 251
    Else
        settings.item(SBL_AtkPower) = 100
        settings.item(SBL_DefPower) = 100
        settings.item(SBL_HP2) = 100
    End If
    Call enableEvent(False)
    Call setSettings(rng, settings)
    Call enableEvent(True)
End Sub

'   特殊なフィルタの設定
Public Sub sblSetFilter(ByVal target As Range)
    Dim tcol, fcol As String
    Dim col  As Long
    Dim crit As Variant
    Dim ope As XlAutoFilterOperator
    ope = xlOr
    With target.Parent.ListObjects(1)
        tcol = .HeaderRowRange.cells(1, target.column).text
        '   タイプでフィルタ
        If tcol = SBL_Type Then
            fcol = SBL_Type
            crit = "=*" & left(target.value, 1) & "*"
        '   通常わざ
        ElseIf tcol = SBL_NormalAtk Then
            fcol = SBL_FilterNormalAtk
            crit = target.value
        '   ゲージわざ
        ElseIf tcol = SBL_SpecialAtk1 Then
            fcol = SBL_FilterSpecialAtk
            crit = "=*" & target.value & "*"
        Else
            Exit Sub
        End If
        If fcol <> "" Then
            col = .ListColumns(fcol).DataBodyRange.column
            If target.text = "" Then
                .Range.AutoFilter Field:=col
            Else
                .Range.AutoFilter Field:=col, Criteria1:=crit, Operator:=ope
            End If
        End If
    End With
End Sub

'   種族リーグ別作成
Public Sub sblMakeTable(ByVal sh As Worksheet)
    Dim settings As Object
    Set settings = getSettings(sh.Range(SBL_R_Settings))
    Call sblCopySpecies(sh, settings(C_League))
    Call sblSetRows(sh, settings)
End Sub

'   パラメータの再計算
Public Sub sblRecalcParams(ByVal target As Range)
    Dim settings As Object
    Set settings = getSettings(target.Parent.Range(SBL_R_Settings))
    Call sblSetRows(target.Parent, settings, target)
End Sub

'   種族分析シートから最初の3列をコピー
Private Sub sblCopySpecies(ByVal sh As Worksheet, ByVal league As String)
    Dim rCel, wcel, acel As Range
    Dim colRef, lidx, i, num As Long
    Dim recInd, species As String
    Dim atk As Long
    Dim def As Long
    Dim hp As Long
    Dim tCP As Long
    Dim PL As Double
    Dim wcols As Variant
    Dim atks As Variant
    
    Call doMacro(msgstr(msgMakingTable, sh.name))
    tCP = getCpUpper(league, 0)
    lidx = getLeagueIndex(league)
    wcols = getColumnIndexes(sh, Array( _
            SBL_indATK, SBL_NormalAtk, SBL_PL))
    colRef = sh.ListObjects(1).ListColumns(SBL_Number).Range.column
    Set wcel = sh.ListObjects(1).HeaderRowRange.cells(1, colRef)
    Set wcel = sh.Range(wcel, wcel.Offset(0, 2))
    colRef = getColumnIndex(SA1_ReccIV & Trim(lidx), shSpeciesAnalysis1)
    num = WorksheetFunction.CountIf( _
            shSpeciesAnalysis1.ListObjects(1).ListColumns(colRef).DataBodyRange, _
            "<>")
    Call sblClearTable(sh, num)
    For Each rCel In shSpeciesAnalysis1.ListObjects(1).ListColumns(SA1_Number).DataBodyRange
        recInd = rCel.Offset(0, colRef - 1).text
        If recInd <> "" Then
            Set wcel = wcel.Offset(1, 0)
            shSpeciesAnalysis1.Range(rCel, rCel.Offset(0, 2)).copy
            wcel.PasteSpecial
            species = rCel.Offset(0, 1).value
            Call getIndValues(recInd, atk, def, hp)
            If tCP > 0 Then
                PL = getPLbyCP2(tCP, species, atk, def, hp)
            Else
                PL = 40
            End If
            With wcel.cells(1, 1).Offset(0, wcols(0) - 1)
                .value = atk
                .Offset(0, 1).value = def
                .Offset(0, 2).value = hp
            End With
            wcel.cells(1, 1).Offset(0, wcols(2) - 1).value = PL
            atks = sblGetReccAtkCell(rCel)
            Set acel = wcel.cells(1, 1).Offset(0, wcols(1) - 1)
            For i = 0 To 2
                If Not IsObject(atks(i)) Then
                    Call setAtkNames(1, atks(i), acel.Offset(0, i))
                ElseIf Not atks(i) Is Nothing Then
                    atks(i).copy
                    acel.Offset(0, i).PasteSpecial
                End If
            Next
        End If
    Next
    Call doMacro
End Sub

'   テーブルクリア
Private Sub sblClearTable(ByVal sh As Worksheet, Optional ByVal rows As Long = 1)
    With sh.ListObjects(1)
        If .DataBodyRange Is Nothing Then Exit Sub
        Call doMacro(msgstr(msgProcessing, Array(cmdClear, sh.name)))
        .DataBodyRange.ClearContents
        .Resize Range(.HeaderRowRange.cells(1, 1), _
            .DataBodyRange.cells(rows, .DataBodyRange.columns.count))
        Call doMacro
    End With
End Sub

'   cDPSランキングより技の取得
Private Function sblGetReccAtkCell(ByVal cel As Range) As Variant
    Dim atks(3) As Variant
    Dim i, j, rnkNum, normal, idx As Integer
    Dim itype(2) As Integer
    Dim allAtks, attr As Variant
    Dim str As String
    Dim max(1) As Double
    
    Set atks(0) = getColumn(SA1_CDST_NormalAtkName & "1", cel)
    Set atks(1) = getColumn(SA1_CDST_SpecialAtkName & "1", cel)
    itype(1) = getTypeIndex(getAtkAttr(C_IdSpecialAtk, atks(1), C_Type))
    rnkNum = shSpeciesAnalysis1.getAtkRankingNum
    '   cDPTランク順にタイプの異なるものを採用
    For i = 2 To rnkNum
        Set atks(2) = getColumn(SA1_CDST_SpecialAtkName & Trim(i), cel)
        If atks(2).text = "" Then Exit For
        itype(2) = getTypeIndex(getAtkAttr(C_IdSpecialAtk, atks(2), ATK_Type))
        If itype(1) <> itype(2) Then GoTo Continue
    Next
    '   見つからなかったら総当たりでさがす
    normal = getTypeIndex(C_Normal)
    allAtks = getAtkNames(getColumn(SA1_Name, cel).text, False, True)
    atks(2) = ""
    For i = 0 To UBound(allAtks(1))
        If allAtks(1)(i) <> "" Then
            attr = getAtkAttrs(1, allAtks(1)(i), Array(ATK_Type, ATK_DPE))
            itype(2) = getTypeIndex(attr(0))
            '   DPE最大のものを保持。ノーマルは別枠で
            If itype(2) = normal Then idx = 1 Else idx = 0
            If itype(1) <> itype(2) And attr(1) > max(idx) Then
                max(idx) = attr(1)
                atks(2 + idx) = allAtks(1)(i)
            End If
        End If
    Next
    '   ノーマルでないものがなかったらノーマル
    If atks(2) = "" And atks(3) <> "" Then atks(2) = atks(3)
    If atks(2) <> "" Then GoTo Continue
    '   見つからなかったら、cDPT二位
    Set atks(2) = getColumn(SA1_CDST_SpecialAtkName & "2", cel)
Continue:
    sblGetReccAtkCell = atks
End Function

'   行の設定
Private Sub sblSetRows(ByVal sh As Worksheet, _
                        ByRef settings As Object, _
                        Optional ByVal rng As Range = Nothing)
    Dim monCols, prmCols, atkCols As Variant
    Dim mon As Monster
    Dim enemy As Monster
    Dim cel, calcRng As Range
    
    Call doMacro(msgstr(msgMakingTable, sh.name))
    Set calcRng = sh.ListObjects(1).ListColumns(1).DataBodyRange
    If Not rng Is Nothing Then
        Set rng = sh.Range(sh.cells(rng.row, 1), _
                    sh.cells(rng.row + rng.rows.count - 1, 1))
        Set calcRng = Application.Intersect(calcRng, rng)
    End If
    Call getMonsterByPower(enemy, , settings(SBL_AtkPower), settings(SBL_DefPower), settings(SBL_HP2))
    For Each cel In calcRng
        Call sblGetMonsterByCell(mon, cel, monCols)
        Call sblSetParams(cel, mon, prmCols)
        Call sblSetAtkParams(cel, mon, enemy, atkCols)
    Next
    Call doMacro
End Sub

'   モンスターの生成
Private Sub sblGetMonsterByCell(ByRef mon As Monster, ByVal cel As Range, ByRef cols As Variant)
    Dim attr, atks As Variant
    attr = getRowValues(cel, Array(SBL_Species, SBL_PL, _
                        SBL_indATK, SBL_indDEF, SBL_indHP, _
                        SBL_NormalAtk, SBL_SpecialAtk1, SBL_SpecialAtk2))
    Call getMonster(mon, attr(0), attr(1), attr(2), attr(3), attr(4))
    If attr(7) = "" Then atks = Array(attr(6)) Else atks = Array(attr(6), attr(7))
    Call setAttacks(C_IdMtc, mon, attr(5), atks)
End Sub

'   パラメータのセット
Private Sub sblSetParams(ByVal cel As Range, ByRef mon As Monster, ByRef cols As Variant)
    Dim row As Long
    Dim scp As Variant
    If Not IsArray(cols) Then
        cols = getColumnIndexes(cel.Parent, Array(SBL_CP, SBL_HP, _
                    SBL_AtkPower, SBL_DefPower, SBL_HP2, _
                    SBL_SCP, SBL_DCP, SBL_Endurance))
    End If
    row = cel.row
    With cel.Parent
        .cells(row, cols(0)).value = mon.CP
        .cells(row, cols(1)).value = mon.fullHP
        .cells(row, cols(2)).value = mon.atkPower
        .cells(row, cols(3)).value = mon.defPower
        .cells(row, cols(4)).value = mon.hpPower
        scp = getSCP(mon)
        .cells(row, cols(5)).value = scp(0)
        .cells(row, cols(6)).value = scp(1)
        .cells(row, cols(7)).value = scp(2)
    End With
End Sub

'   わざのパラメータの書き込み
Private Sub sblSetAtkParams(ByVal cel As Range, ByRef mon As Monster, _
                        ByRef enemy As Monster, ByRef cols As Variant)
    Dim row As Long
    Dim sh As Worksheet
    Dim attr As Variant
    Dim ofs As Variant
    Dim i, idx As Integer
    Dim cdpss As CDpsSet
    Dim tcol As Long
    Dim max As Double
    Dim stype As String
    
    Set sh = cel.Parent
    If Not IsArray(cols) Then
        cols = getColumnIndexes(sh, Array(SBL_mTCP, _
                SBL_MtcNormalAtkDamage, SBL_MtcNormalAtkTDPS, SBL_MtcNormalAtkEPT, _
                SBL_MtcSpecialAtkDamage & "1", SBL_MtcSpecialAtkDPE & "1", _
                SBL_MtcSpecialAtkCDPS & "1", SBL_MtcSpecialAtkCycle & "1", _
                SBL_MtcSpecialAtkDamage & "2", _
                SBL_FilterNormalAtk, SBL_FilterSpecialAtk)) '   10
    End If
    row = cel.row
    ofs = Array(0, cols(5) - cols(4), cols(6) - cols(4), cols(7) - cols(4))
    Call calcADamage(C_IdNormalAtk, mon, enemy, True)
    With mon.attacks(mon.atkIndex(0).selected)
        attr = getAtkAttrs(C_IdNormalAtk, .name, Array(ATK_EPT, ATK_IdleTurnNum))
        sh.cells(row, cols(1)).value = .damage
        sh.cells(row, cols(2)).value = .damage / attr(1)
        sh.cells(row, cols(3)).value = attr(0)
        sh.cells(row, cols(9)).value = getTypeName(.itype)
    End With
    For i = 0 To 1
        idx = mon.atkIndex(1).lower + i
        If idx > mon.atkIndex(1).upper Then Exit For
        tcol = Array(cols(4), cols(8))(i)
        mon.atkIndex(1).selected = idx
        cdpss = calcCDPS(mon, enemy, True)
        With mon.attacks(mon.atkIndex(1).selected)
            attr = getAtkAttrs(C_IdSpecialAtk, .name, Array(ATK_GaugeVolume))
            sh.cells(row, tcol + ofs(0)).value = .damage
            sh.cells(row, tcol + ofs(1)).value = .damage / attr(0)
            sh.cells(row, tcol + ofs(2)).value = cdpss.cDPS
            sh.cells(row, tcol + ofs(3)).value = cdpss.Cycle
            If cdpss.cDPS > max Then max = cdpss.cDPS
            stype = stype & getTypeName(.itype) & ","
        End With
    Next
    sh.cells(row, cols(10)).value = stype
    sh.cells(row, cols(0)).value = max * getEndurance(mon.defPower, mon.hpPower)
End Sub


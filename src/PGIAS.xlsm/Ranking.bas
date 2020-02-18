Attribute VB_Name = "Ranking"


Option Explicit


'   計算するボタン
Public Function onCalcRankingClick(ByVal sh As Worksheet, ByVal isAll As Boolean, _
                ByRef settings As Object, ByRef dummySet As Object) _
                As Boolean
    onCalcRankingClick = False
'    If Not shIndividual.checkPL() Then Exit Function
    If Not isAll And (ActiveCell.CountLarge <> 1 Or _
        Application.Intersect(ActiveCell, sh.ListObjects(1).DataBodyRange) Is Nothing) Then Exit Function
    Call doMacro(msgstr(msgProcessing, Array(msgCalculate, msgRanking)))
    If isAll Then
        Call SetAllRanking(sh, settings, dummySet)
    Else
        Call SetRanking(Selection, settings, dummySet)
    End If
    Call doMacro
    onCalcRankingClick = True
End Function

'   クリア・削除ボタン
Public Function onRemoveRankingClick(ByVal sh As Worksheet, _
                ByVal isAll As Boolean, _
                Optional ByVal remove As Boolean = False, _
                Optional ByVal confirm As Boolean = True) As Boolean
    Dim cmd As String
    Dim confirmed As Boolean
    onRemoveRankingClick = False
    If Not isAll Then
        If ActiveCell.CountLarge <> 1 Or sh.ListObjects(1).DataBodyRange Is Nothing Then Exit Function
        If Application.Intersect(ActiveCell, sh.ListObjects(1).DataBodyRange) Is Nothing Then Exit Function
    End If
    If remove Then cmd = msgRemove Else cmd = msgClear
    If isAll Then
        confirmed = True
        If confirm Then
            confirmed = (MsgBox(msgstr(msgConfirm, Array(cmd)), vbYesNo) = vbYes)
        End If
        If confirmed Then
            Call doMacro(msgstr(msgProcessing, Array(cmd, msgRanking)))
            Call ClearAllRanking(sh, remove)
            Call doMacro
            onRemoveRankingClick = True
        End If
    Else
        Call doMacro(msgstr(msgProcessing, Array(cmd, msgRanking)))
        Call ClearCalcedRank(Selection, remove)
        Call doMacro
        onRemoveRankingClick = True
    End If
End Function

'   ランクすべてを計算
Public Sub SetAllRanking(ByVal table As Variant, ByRef settings As Object, ByRef dummySet As Object)
    Dim row As Long
    Dim cel As Range
    
    Call ClearAllRanking(table)
    With getListObject(table).DataBodyRange
        row = 1
        Do While row <= .rows.count
            Set cel = .cells(row, 1)
            If cel.Text <> "" Then
                Call SetARanking(cel, settings, dummySet)
            End If
            row = row + 1
        Loop
    End With
End Sub

Private Sub SetRanking(ByVal rng As Range, ByRef settings As Object, ByRef dummySet As Object)
    Dim i As Long
    Dim lo As ListObject
    Dim rows As Variant
    
    Set lo = rng.ListObject
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    rows = getExecRows(rng)
    If IsArray(rows) Then
        For i = 0 To UBound(rows)
            Call ClearDataRange(lo, Array(rows(i)))
            Call SetARanking(lo.DataBodyRange.cells(rows(i)(0), 1), settings, dummySet)
        Next
    End If
End Sub

'   ランク一つを計算
Private Sub SetARanking(ByVal Target As Range, _
            ByRef settings As Object, ByRef dummySet As Object)
    Dim stime As Date
    Dim ssbj, memo As String
    
    stime = Now
    ssbj = Target.Text
    memo = Target.Offset(0, 1).Text
    If memo <> "" Then ssbj = ssbj & "(" & memo & ")"
    Call dspProgress(msgstr(msgCalcRank, _
            Array(ssbj)), _
            2 * (Range(R_WeatherTable).rows.count + 1) _
            * shIndividual.ListObjects(1).DataBodyRange.rows.count)
    
    Call fillRank(Target, settings, dummySet)
    
    Call dspProgress("", 0)
    getColumn(BE_CalcTime, Target).Value = DateDiff("s", stime, Now)
End Sub

'   すべてクリア
Private Sub ClearAllRanking(ByVal table As Variant, _
        Optional ByVal remove As Boolean = False)
    Dim cel As Range
    Dim row As Long
    Dim lo As ListObject
    Set lo = getListObject(table)
    If remove Then
        Call ClearDataRange(lo, Array( _
            Array(1, lo.DataBodyRange.rows.count)), remove)
    Else
        Call ClearCalcedRank(lo.DataBodyRange, remove)
    End If
End Sub

'   範囲でクリア
Private Sub ClearCalcedRank(ByVal rng As Range, _
            Optional ByVal remove As Boolean = False)
    Dim rows As Variant
    rows = getExecRows(rng)
    If IsArray(rows) Then
        Call ClearDataRange(rng.ListObject, rows, remove)
    End If
End Sub

'   範囲より、処理する行を取得する
Private Function getExecRows(ByVal rng As Range) As Variant
    Dim rows(), num As Long
    Dim col, row, srow As Long
    
    ReDim rows(rng.rows.count)
    If rng.ListObject Is Nothing Then Exit Function
    If rng.ListObject.DataBodyRange Is Nothing Then Exit Function
    With rng.ListObject.DataBodyRange
        col = getColumnIndex(BE_Species, .Parent)
        srow = rng.row - .row + 1
        If srow < 1 Then srow = 1
        row = rng.row - .row + rng.rows.count
        If row > .rows.count Then row = .rows.count
        Do
            '   種族名が書いてあったら開始行
            If .cells(row, col).Text <> "" Then
                rows(num) = Array(row, 0)
                If num > 0 Then rows(num)(1) = rows(num - 1)(0) - 1
                num = num + 1
            End If
            row = row - 1
        Loop Until row < 1 Or (row < srow And .cells(row + 1, col).Text <> "")
        row = rows(0)(0) + 1
        Do While row <= .rows.count
            If .cells(row, col).Text <> "" Then
                rows(0)(1) = row - 1
                Exit Do
            End If
            row = row + 1
        Loop
        If rows(0)(1) = 0 Then rows(0)(1) = .rows.count
    End With
    ReDim Preserve rows(num - 1)
    getExecRows = rows
End Function

'   rowsの行範囲をクリアする
Private Sub ClearDataRange(ByVal lo As ListObject, ByVal rows As Variant, _
            Optional ByVal remove As Boolean = False)
    Dim cols As Variant
    Dim i As Long
    cols = getColumnIndexes(lo, Array(BE_Species, BE_RankBase, BE_CalcTime))
    With lo.DataBodyRange
        For i = 0 To UBound(rows)
            '   削除フラグがないか、最後のデータなら2行残してクリア
            If Not remove Or (rows(i)(1) - rows(i)(0)) + 1 = .rows.count Then
                If rows(i)(1) > rows(i)(0) + 1 Then
                    With Range(.cells(rows(i)(0) + 2, 1), .cells(rows(i)(1), 1)).EntireRow
                        .Delete
                    End With
                End If
                If Not remove Then
                    Range(.cells(rows(i)(0), cols(1)), .cells(rows(i)(0) + 1, cols(2))).ClearContents
                Else
                    .cells(rows(i)(0), 1).Value = "?"
                    .cells(rows(i)(0) + 1, 1).Value = ""
                    Range(.cells(rows(i)(0), 2), .cells(rows(i)(0) + 1, cols(2))).ClearContents
                End If
                Call setBorders(.Parent, Array(rows(i)(0), rows(i)(0) + 1), True)
            Else    '   削除フラグがあって、最後のデータでない
                With Range(.cells(rows(i)(0), 1), .cells(rows(i)(1), 1)).EntireRow
                    .Delete
                End With
                Call setBorders(.Parent, Array(rows(i)(0), rows(i)(0)), True)
            End If
        Next
    End With
End Sub

'   枠線で囲む
Private Sub setBorders(ByVal table As Variant, ByVal rrow As Variant, _
                Optional ByVal draw As Boolean = True)
    With getListObject(table).DataBodyRange
        With Range(.cells(rrow(0), 1), .cells(rrow(1), .columns.count))
            If draw Then
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 16
                    .Weight = xlThin
                End With
                If .rows.count > 1 Then
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 16
                        .Weight = xlThin
                    End With
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                End If
            Else
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        End With
    End With
End Sub
'   ランクを計算してシートに書き込む
Private Sub fillRank(ByVal cel As Range, ByRef settings As Object, ByRef dummySet As Object)
    Dim defs As Object
    Dim ranks(1) As Variant
    Dim r2col As Variant
    Dim key As String
    Dim col, lcol, rankNum As Long
    Dim i, j, lim, maxLim As Integer
    
    '   行データの取得
    lcol = getColumnIndex(BE_Rank & BE_SuffixBase, cel) - 1
    Set defs = getRowValues(cel, Nothing, Array(1, lcol))
    r2col = getColumnIndexes(cel, Array(BE_PL, BE_CPHP))
    For col = r2col(0) To r2col(1)
        key = cel.ListObject.HeaderRowRange.cells(1, col).Text
        defs.item(key & "2") = cel.Offset(1, col - 1).Value
    Next
    If defs(BE_Species) = "" Then Exit Sub
    rankNum = settings(BE_SetRankNum)
    '   現在/予測のループ
    For i = 0 To 1
        '   データの取得
        ranks(i) = getRanks(settings, defs, dummySet, (i = 1))
        If Not IsArray(ranks(i)) Then Err
        '   書き込む形式に修正
        ranks(i) = alignRanks(ranks(i), rankNum)
        For j = 0 To 1
            If IsArray(ranks(i)(j)) Then
                lim = UBound(ranks(i)(j))
                If maxLim < lim Then maxLim = lim
            End If
        Next
    Next
    Call insertRows(cel, maxLim + 1)
    '   シートに書き込む
    For i = 0 To 1
        Call insertAlignedRanks(cel, ranks(i), (i = 1))
    Next
    Exit Sub
Err:
End Sub

'   必要数の行を挿入する
Private Sub insertRows(ByVal cel As Range, ByVal rowNum As Long)
    Dim lcol, srow, rows As Long
    If cel.Offset(1, 0).Text = "" Then rows = 2 Else rows = 1
    If rowNum > rows Then
        Range(cel.Offset(rows, 0), cel.Offset(rowNum - 1, 0)).EntireRow.Insert
    End If
    srow = cel.ListObject.DataBodyRange.cells(1, 1).row
    srow = cel.row - srow + 1
    Call setBorders(cel, Array(srow, srow + rowNum - 1), True)
End Sub

'   成形済みランクデータの挿入
Private Sub insertAlignedRanks(ByVal cel As Range, ByRef aligned As Variant, _
                                Optional ByVal isPrediction As Boolean = False)
    Dim i, j As Long
    Dim scol As Variant
    Dim wcel As Range
    Dim placed As String
    Static exists As String
    
    '   先頭列番号の取得
    If Not isPrediction Then
        exists = ""
        scol = Array(BE_Rank & BE_SuffixBase, BE_Weather & BE_SuffixWeather)
    Else
        scol = Array(BE_Rank & BE_SuffixPredictBase, BE_Weather & BE_SuffixPredictWeather)
    End If
    scol = getColumnIndexes(cel, scol)
    
'    Call enableEvent(False)
    '   基本ランク。placed、existsは登場個体の文字列
    placed = writeAData(cel.Offset(0, scol(0) - 1), aligned(0), True, exists)
    '   天候での差分
    placed = placed & writeAData(cel.Offset(0, scol(1) - 1), aligned(1), False, exists)
'    Call enableEvent(True)
    exists = placed
End Sub

'   一つのデータをセルに書き込む
Private Function writeAData(ByVal scel As Range, ByVal data As Variant, _
            ByVal isBase As Boolean, Optional ByVal exists As String = "") As String
    Dim sidx, lidx, i, idx As Integer
    Dim cel As Range
    Dim name, placed As String
    
    If Not IsArray(data) Then Exit Function
    sidx = 0
    lidx = UBound(data(0))
    If isBase Then
        sidx = 1
        lidx = lidx - 1
    End If
    If exists <> "" Then exists = exists & "|"
    For i = 0 To UBound(data)
        For idx = sidx To lidx
            Set cel = scel.Offset(i, idx - sidx)
            cel.Value = data(i)(idx)
            If idx = 0 Then '   天候
                Call WeatherChange(cel)
            ElseIf idx = 3 Or idx = 5 Then  '   技
                Call AtkChange(cel, False)
            ElseIf idx = 2 Then '   ニックネーム
                name = "|" & data(i)(idx)
                placed = placed & name
                '  現在のランクになく予測のみにある名前は強調する
                If exists <> "" And InStr(exists & "|", name & "|") < 1 Then
                    If InStr(placed, name & "|") < 1 Then
                        cel.Font.ColorIndex = BE_NewEntryColorIndex
                    Else ' 再登場
                        cel.Font.ColorIndex = BE_ReEntryColorIndex
                    End If
                Else
                    cel.Font.ColorIndex = 0
                End If
            End If
        Next
    Next
    writeAData = placed
End Function

'   ランクより書き込み順のデータ生成
Private Function alignRanks(ByRef ranks As Variant, ByVal rankNum As Long) As Variant
    Dim sbj, ret(1) As Variant
    Dim wthNum, idx, sidx, ridx, ri, wi, rinum, i, j As Long
    Dim isBase, isSame() As Boolean
    Dim weather, out As String
    
'    wthNum = Range(R_WeatherTable).rows.count
    wthNum = UBound(ranks)
    '   天候なしの基本ランク
    ReDim data(rankNum - 1), isSame(rankNum - 1)
    For ri = 0 To rankNum - 1
        data(ri) = getAlignedRankData(ranks(0)(ri), ri + 1)
    Next
    ret(0) = data
    '   天候ありランクを縦にならべる
    ReDim data(wthNum * rankNum)
    idx = 0 '作成するデータのインデックス
    wi = 1
    Do While wi <= wthNum   '天候ごと
        '   基本ランクと同名のフラグをクリア
        For i = 0 To rankNum - 1
            isSame(i) = False
        Next
        rinum = 0
        sidx = idx
        weather = Range(R_WeatherTable).cells(wi, 1).Text
        For i = 0 To rankNum - 1  '順位ごと
            sbj = ranks(wi)(i)
            isBase = False
            out = ""
            '   基本ランク内に、参照中の個体・技があれば記録しない
            For j = 0 To rankNum - 1
                If ranks(0)(j)(2) = sbj(2) Then
                    isSame(j) = True
                    '   名前が同じで技が異なる場合は、記録くするが、同一行のランクアクトに同名を記録する
                    If ranks(0)(j)(5) <> sbj(5) Then
                        isBase = True
                        j = rankNum
                    End If
                    Exit For
                End If
            Next
            '   基本ランクに同一個体・技がなかったので記録する
            If j >= rankNum Then
                rinum = rinum + 1
                out = ""
                '   個体だけ一致する場合は、ランクアウトのセルを、同じ個体の名でうめる
                If isBase Then out = sbj(2): rinum = rinum - 1
                data(idx) = getAlignedRankData(sbj, i + 1, weather, out)
                idx = idx + 1
            End If
        Next
        '   空いているランクアウトのセルを、基本順位の下から埋める
        If idx > sidx Then
            i = rankNum - 1
            ridx = idx - 1
            While ridx >= sidx
                If data(ridx)(10) = "" Then
                    While isSame(i)
                        i = i - 1
                    Wend
                   data(ridx)(10) = ret(0)(i)(2)
                   i = i - 1
                End If
                ridx = ridx - 1
            Wend
        End If
        wi = wi + 1
    Loop
    If idx > 0 Then
        ReDim Preserve data(idx - 1)
        ret(1) = data
    Else
        ret(1) = 0
    End If
    alignRanks = ret
End Function

'   一行分のランクデータを列順にならべる
Private Function getAlignedRankData(ByRef ktrs As Variant, ByVal rank As Long, _
                        Optional ByVal weather As String = "", _
                        Optional ByVal rankout As String = "") As Variant
    '                           天候, 順位, 名前, 通常わざ, tDPS
    '                       ゲージわざ, tDPS, cDPS, KT, KTR, ランク落ち
    getAlignedRankData = Array(weather, rank, ktrs(2), ktrs(3), ktrs(4), _
                            ktrs(5), ktrs(6), ktrs(7), ktrs(1), ktrs(0), rankout)
End Function

'   ランク計算
Private Function getRanks(ByRef settings As Object, ByRef defs As Object, _
                    ByRef enemySet As Object, _
                    Optional ByVal predict As Boolean = False) As Variant
    Dim enemy As Monster
    Dim species As String
    Dim rank() As Variant
    Dim lcol, wthNum, wi As Long
    Dim mode As Integer
    
    If settings(BE_SetMode) = C_Gym Then mode = C_IdGym Else mode = C_IdMtc
    '   敵の設定
    If Not setEnemy(enemy, defs, enemySet, mode, _
            settings(BE_SetEnemyAtkDelay), settings(BE_SetWithLimit)) Then
        Exit Function
    End If
    '   天候の数。ランクの数は天候＋１
    If enemy.mode = C_IdGym Then
        wthNum = Range(R_WeatherTable).rows.count
    Else    '   トレーナーバトルは天候に変わりがない
        wthNum = 0
    End If
    ReDim rank(wthNum)
    For wi = 0 To wthNum
        rank(wi) = getKtRank(settings(BE_SetRankNum), enemy, _
            (settings(BE_SetRankVar) = "KTR"), defs(BE_UpperCP), defs(BE_LowerCP), _
            predict, wi, settings(BE_SetSelfAtkDelay))
    Next
    getRanks = rank
End Function

'   敵の設定
Private Function setEnemy(ByRef enemy As Monster, ByRef defs As Object, _
                ByRef dummySet As Object, _
                Optional ByVal mode As Integer = 0, _
                Optional ByVal atkDelay As Double = 0, _
                Optional ByVal withLimited As Boolean = False) As Boolean
    Dim atks, atkDef, types As Variant
    Dim species As String
    setEnemy = False
    species = defs(BE_Species)
    '   ？の場合はダミーの敵
    If species = "?" Or species = "？" Then
        types = Split(defs(BE_Memo), ",")
        If UBound(types) < 1 Then ReDim Preserve types(1)
        Call getDummyMonster(enemy, dummySet, mode, atkDelay, types)
        setEnemy = True
        Exit Function
    End If
    If Not speciesExists(species) Then Exit Function
    '  敵の設定
    '   PL、個体値を最後に弄った場合
    If defs(BE_PL & "2") = 0 Then
        Call getMonster(enemy, species, defs(BE_PL), _
                defs(BE_ATK), defs(BE_DEF), defs(BE_HP))
    Else
        Call getMonsterByPower(enemy, species, _
                defs(BE_ATK & "2"), defs(BE_DEF & "2"), defs(BE_HP & "2"))
    End If
    '   わざのセット
    atkDef = Array(defs(BE_NormalAttack), defs(BE_SpecialAttack))
    If "" = atkDef(0) Or "" = atkDef(1) Then
        atks = getAtkNames(species, False, withLimited)
        If atkDef(0) <> "" Then atks(0) = Array(atkDef(0))
        If atkDef(1) <> "" Then atks(1) = Array(atkDef(1))
    Else
        atks = Array(Array(atkDef(0)), Array(atkDef(1)))
    End If
    Call setAttacks(mode, enemy, atks(0), atks(1), atkDelay, (mode = C_IdMtc))
    setEnemy = True
End Function

'   セルの選択・変更

'   セルの選択
Public Function onRankingSelectionChange(ByVal Target As Range) As Boolean
    Dim s As String
    onRankingSelectionChange = False
    With Target.Parent.ListObjects(1)
        If Target.CountLarge <> 1 Or _
            Application.Intersect(Target, .DataBodyRange) Is Nothing Then Exit Function
        Call setInputList
        Select Case .HeaderRowRange.cells(1, Target.column).Text
            Case BE_Memo
                s = Target.Offset(0, -1).Text
                If s <> "?" And s <> "？" Then Target.Font.ColorIndex = 0
            Case BE_NormalAttack
                Call AtkSelected(Target)
            Case BE_SpecialAttack
                Call AtkSelected(Target)
        End Select
    End With
    onRankingSelectionChange = True
End Function


'   セル値の変更
Public Function onRankingSheetChange(ByVal Target As Range, ByRef settings As Object) As Boolean
    Dim title As String
    onRankingSheetChange = False
    
    With Target.Parent.ListObjects(1)
        If Target.CountLarge <> 1 Or _
            Application.Intersect(Target, .DataBodyRange) Is Nothing Then Exit Function
        '   全角スペースのみはクリア
        If Target.Text = "" Or Target.Text = "　" Then
            Call enableEvent(False)
            Target.ClearContents
            Call enableEvent(True)
            Target.Validation.Delete
            Exit Function
        End If
        title = .HeaderRowRange.cells(1, Target.column).Text
    End With
    If title = BE_Species Then
        Call speciesChange(Target, settings)
    ElseIf title = BE_Memo Then
        Call memoChange(Target)
    ElseIf title = BE_CPHP Then
        '   2行目ならHPチェック変換
        If Target.Offset(0, 1 - Target.column).Text = "" Then
            Call checkHPValue(Target)
        End If
        Call calcParams(Target)
    ElseIf title = BE_PL Or title = BE_ATK _
            Or title = BE_DEF Or title = BE_HP Or title = BE_CPHP Then
        Call calcParams(Target)
    ElseIf title = BE_NormalAttack Then
        Call AtkChange(Target)
    ElseIf title = BE_SpecialAttack Then
        Call AtkChange(Target)
    End If
    onRankingSheetChange = True
End Function

'   HPのセルでレベルをHPに変換
Public Sub checkHPValue(ByVal Target As Range)
    Dim txt, h As String
    Dim lvl As Integer
    txt = Target.Text
    If IsNumeric(txt) Then Exit Sub
    txt = StrConv(txt, vbNarrow)
    txt = StrConv(txt, vbLowerCase)
    h = left(txt, 1)
    If h = "l" Or h = "s" Then
        lvl = val(Mid(txt, 2))
        If lvl < 1 And 5 < lvl Then Exit Sub
        Target.Value = Array(0, 600, 1800, 3600, 9000, 15000)(lvl)
    End If
End Sub

'   種族名の変更
Private Sub speciesChange(ByVal Target As Range, ByRef settings As Object)
    Dim species, natk, satk As String
    Dim row As Long
    If Target.Offset(1, 0) <> "" Then Call insertRows(Target, 2)
    If speciesExpectation(Target) Then
        species = getColumn(BE_Species, Target).Text
        Call setEnemyParams(Target, settings, species)
    Else
        Call setEnemyParams(Target, settings)
    End If
End Sub

'   備考の変更
Private Sub memoChange(ByVal Target As Range)
    Dim species As String
    Dim str As String
    species = getColumn(BE_Species, Target).Text
    If species = "?" Or species = "？" Then
        str = Replace(Target.Text, "、", ",")
        str = Replace(str, "，", ",")
        Call enableEvent(False)
        Target.Value = Replace(str, "・", ",")
        Call setTypeColorsOnCell(Target, , True)
        Call enableEvent(True)
    End If
End Sub

'   敵のパラメータをセット
Private Sub setEnemyParams(ByVal Target As Range, ByRef settings As Object, _
            Optional ByVal species As String = "", _
            Optional ByVal PL As Double = 40, _
            Optional ByVal atk As Long = 15, _
            Optional ByVal def As Long = 15, _
            Optional ByVal hp As Long = 15)
    Dim ncols, ecols, nvals, evals As Variant
    Dim atkNames, types As Variant
    Dim i As Long
    
    ncols = Array(BE_PL, BE_ATK, BE_DEF, BE_UpperCP, BE_LowerCP)
    ecols = Array(BE_HP, BE_NormalAttack, BE_SpecialAttack)
    If species = "" Then
        nvals = Array("", "", "", settings(BE_DefCpUpper), settings(BE_DefCpLower))
        evals = Array("", "", "")
    Else
        nvals = Array(PL, atk, def, settings(BE_DefCpUpper), settings(BE_DefCpLower))
        atkNames = seachAndGetValues(species, SA1_Name, shSpeciesAnalysis1, _
                            Array(SA1_CDSP_NormalAtkName & "1", SA1_CDSP_SpecialAtkName & "1"))
        evals = Array(hp, atkNames(0), atkNames(1))
        types = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    End If
    '   イベントを発生させない列
    Call enableEvent(False)
        '   備考の2行目
        If IsArray(types) Then
            Call setTypeToCell(types, getColumn(BE_Memo, Target).Offset(1, 0))
        Else
            getColumn(BE_Memo, Target).Offset(1, 0).Value = " "
        End If
        For i = 0 To UBound(nvals)
            getColumn(ncols(i), Target).Value = nvals(i)
        Next
    Call enableEvent(True)
    '   イベントを発生させる列
    For i = 0 To UBound(evals)
        getColumn(ecols(i), Target).Value = evals(i)
    Next
End Sub

Private Sub calcParams(ByVal Target As Range)
    Dim tcel(1) As Range
    Dim cat As Integer
    Dim mon As Monster
    Dim cols, vals(1) As Variant
    Dim cphpCol As Long
    Dim i, j As Integer
    
    cols = Array(BE_Species, BE_PL, BE_ATK, BE_DEF, BE_HP, BE_CPHP)
    cphpCol = getColumnIndex(BE_CPHP, Target)
    '   領域の上から2行のセルを取得。2行を選択していなければ終了
    Set tcel(0) = Target.Parent.cells(Target.row, 1)
    If tcel(0).Text <> "" Then
        Set tcel(1) = tcel(0).Offset(1, 0)
        cat = 0
    ElseIf tcel(0).Offset(-1, 0).Text <> "" Then
        Set tcel(1) = tcel(0)
        Set tcel(0) = tcel(1).Offset(-1, 0)
        cat = 1
    Else
        Exit Sub
    End If
    
    '   セルの値の取得
    vals(0) = getRowValues(tcel(0), cols)
    If Not speciesExists(vals(0)(0)) Then Exit Sub
    vals(1) = getRowValues(tcel(1), cols)
    '   変更位置によって個体生成
    If Target.column = cphpCol Then
        Call getMonsterByCpHp(mon, vals(0)(0), vals(0)(5), vals(1)(5))
        cat = 2
    ElseIf cat = 0 Then
        Call getMonster(mon, vals(0)(0), vals(0)(1), vals(0)(2), vals(0)(3), vals(0)(4))
    Else
        Call getMonsterByPower(mon, vals(0)(0), vals(1)(2), vals(1)(3), vals(1)(4))
    End If
    '値を戻す
    vals(0)(1) = mon.PL: vals(0)(2) = mon.indATK: vals(0)(3) = mon.indDEF
    vals(0)(4) = mon.indHP: vals(0)(5) = mon.CP
    vals(1)(1) = cat: vals(1)(2) = mon.atkPower: vals(1)(3) = mon.defPower
    vals(1)(4) = mon.hpPower: vals(1)(5) = mon.fullHP
    '   書き込み
    cols = getColumnIndexes(Target, cols)
    Call enableEvent(False)
    For i = 0 To 1
        For j = 1 To 5
            tcel(i).Offset(0, cols(j) - 1).Value = vals(i)(j)
        Next
    Next
    Call enableEvent(True)
End Sub

'  全てワイルドカードに設定
Public Sub setWildCardAll(ByVal sh As Worksheet, ByRef settings As Object)
    Dim rcel As Range
    Dim row As Long
    Dim cols As Variant
    
    Call ClearAllRanking(sh, True)
    Call doMacro(msgSetWildCard)
    With sh.ListObjects(1).DataBodyRange
        cols = getColumnIndexes(.Parent, _
                Array(BE_Species, BE_Memo, BE_UpperCP, BE_LowerCP, BE_CalcTime))
        .cells(1, cols(0)).Value = "?"
        .cells(1, cols(2)).Value = settings(BE_DefCpUpper)
        .cells(1, cols(3)).Value = settings(BE_DefCpLower)
        .cells(2, cols(1)).Value = " "
        Call setBorders(.Parent, Array(1, 2), True)
        row = 3
        For Each rcel In shClassifiedByType.ListObjects(1) _
                    .ListColumns(CBT_Type).DataBodyRange
            .cells(row, cols(0)) = "?"
            .cells(row, cols(1)) = rcel.Text
            Call setTypeColorsOnCell(.cells(row, cols(1)), , True)
            .cells(row, cols(2)) = settings(BE_DefCpUpper)
            .cells(row, cols(3)) = settings(BE_DefCpLower)
            .cells(row + 1, cols(1)) = "."
            Call setBorders(.Parent, Array(row, row + 1), True)
            row = row + 2
        Next
    End With
    Call doMacro
End Sub

Public Sub doBattleSimOnSheet(ByVal sh As Worksheet, _
            ByVal rSetting As String, ByVal mode As Integer)
    Dim settings As Object
    Dim self As Monster
    Dim enemy As Monster
    Dim rrow As Variant
    
    If ActiveCell.CountLarge <> 1 Or _
        Application.Intersect(ActiveCell, sh.ListObjects(1).DataBodyRange) Is Nothing Then Exit Sub
    Set settings = getSettings(rSetting)
    If Not getSelfFromSheet(mode, self, settings) Then Exit Sub
    '   相手の作成
    rrow = getExecRows(ActiveCell)
    
End Sub

'   自分の個体の生成
Private Function getSelfFromSheet(ByVal mode As Integer, _
                ByRef self As Monster, ByRef settings As Object) As Boolean
    Dim title, suffix, nickname As String
    Dim cols, rrow As Variant
    Dim i As Long
    Dim sh As Worksheet
    
    title = sh.ListObjects(1).HeaderRowRange.cells(1, ActiveCell.column).Text
    nickname = ActiveCell.Text
    If nickname = "" Or InStr(title, BE_CtrName) < 1 Then Exit Function
    '   ランキングの部分より個体作成
    suffix = Mid(title, InStr(title, "_"))
    If InStr(suffix, "w") > 0 Then
        cols = Array(BE_NormalAttack, BE_SpecialAttack, "")
    Else
        cols = Array(BE_NormalAttack, BE_SpecialAttack, BE_Weather)
    End If
    For i = 0 To UBound(cols)
        If cols(i) <> "" Then
            cols(i) = getColumnIndex(cols(i) & suffix, sh)
            cols(i) = cells(ActiveCell.row, cols(i)).Value
        End If
    Next
    Call getIndividual(nickname, self, (InStr(suffix, "p") > 0))
    Call setAttacks(mode, self, cols(0), cols(1), settings(BE_SetSelfAtkDelay), (mode = C_IdMtc))
    
End Function

Private Function getEnemyFromSheet(ByVal mode As Integer, _
                ByRef enemy As Monster, ByRef settings As Object) As Boolean

End Function

'   ランキングのモードが変わったので、タイトルを書き変える
Public Sub changeRankingMode(ByVal Target As Range, ByVal mode As Variant)
    Dim cel As Variant
    Dim words As Variant
    Dim idx, rnum, i As Integer
    Dim wcels(1, 2) As Range
    Dim inWthr As Boolean
    
    If Not IsNumeric(mode) Then
        If mode = C_Gym Then idx = C_IdGym Else idx = C_IdMtc
    End If
    words = Array("PS_", "PT_")
    Call enableEvent(False)
    inWthr = False: rnum = 0
    For Each cel In Target.Parent.ListObjects(1).HeaderRowRange
        cel.Value = Replace(cel.Text, words(1 - idx), words(idx))
        If InStr(cel.Text, BE_SuffixWeather) > 0 _
                Or InStr(cel.Text, BE_SuffixPredictWeather) > 0 Then
            If inWthr Then
                Set wcels(rnum, 1) = cel
            Else
                Set wcels(rnum, 0) = cel
                Set wcels(rnum, 1) = cel
                inWthr = True
            End If
        ElseIf inWthr Then
            inWthr = False
            rnum = rnum + 1
        End If
    Next
    If inWthr Then rnum = rnum + 1
    Call enableEvent(True)
    For i = 0 To rnum - 1
        Range(wcels(i, 0), wcels(i, 1)).EntireColumn.Hidden = (idx = C_IdMtc)
    Next
End Sub


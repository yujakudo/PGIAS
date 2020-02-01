Attribute VB_Name = "Ranking"


Option Explicit


'   計算するボタン
Public Function onCalcRankingClick(ByVal sh As Worksheet, ByVal isAll As Boolean, _
                ByVal settingsRangeStr As String, ByVal mode As Integer) _
                As Boolean
    onCalcRankingClick = False
'    If Not shIndividual.checkPL() Then Exit Function
    If Not isAll And (ActiveCell.CountLarge <> 1 Or _
        Application.Intersect(ActiveCell, sh.ListObjects(1).DataBodyRange) Is Nothing) Then Exit Function
    Call doMacro(msgstr(msgProcessing, Array(msgCalculate, msgRanking)))
    If isAll Then
        Call SetAllRanking(sh, settingsRangeStr, mode)
    Else
        Call SetRanking(Selection, settingsRangeStr, mode)
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
Public Sub SetAllRanking(ByVal table As Variant, ByVal settingsRange As Variant, _
                    Optional ByVal mode As Integer = 0)
    Dim settings As Object
    Dim row As Long
    Dim cel As Range
    
    Call ClearAllRanking(table)
    Set settings = getSettings(settingsRange)
    With getListObject(table).DataBodyRange
        row = 1
        Do While row <= .rows.count
            Set cel = .cells(row, 1)
            If cel.Text <> "" Then
                Call SetARanking(cel, settings, mode)
            End If
            row = row + 1
        Loop
    End With
End Sub

Private Sub SetRanking(ByVal rng As Range, ByVal settingsRange As Variant, _
                    Optional ByVal mode As Integer = 0)
    Dim settings As Object
    Dim i As Long
    Dim lo As ListObject
    Dim rows As Variant
    
    Set lo = rng.ListObject
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set settings = getSettings(settingsRange)
    rows = getExecRows(rng)
    If IsArray(rows) Then
        For i = 0 To UBound(rows)
            Call ClearDataRange(lo, Array(rows(i)))
            Call SetARanking(lo.DataBodyRange.cells(rows(i)(0), 1), settings, mode)
        Next
    End If
End Sub

'   ランク一つを計算
Private Sub SetARanking(ByVal Target As Range, ByVal settings As Object, _
                    Optional ByVal mode As Integer = 0)
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
    
    Call fillRank(mode, Target, settings)
    
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
            '   削除フラグがないか、最後のデータなら1行残してクリア
            If Not remove Or (rows(i)(1) - rows(i)(0)) + 1 = .rows.count Then
                If rows(i)(1) > rows(i)(0) Then
                    With Range(.cells(rows(i)(0) + 1, 1), .cells(rows(i)(1), 1)).EntireRow
    '                    .Borders(xlEdgeBottom).LineStyle = xlNone
                        .Delete
                    End With
                End If
                If Not remove Then
                    Range(.cells(rows(i)(0), cols(1)), .cells(rows(i)(0), cols(2))).ClearContents
                Else
                    .cells(rows(i)(0), 1).Value = "?"
                    Range(.cells(rows(i)(0), 2), .cells(rows(i)(0), cols(2))).ClearContents
                End If
                With Range(.cells(rows(i)(0), 1), .cells(rows(i)(0), cols(2)))
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                End With
            Else    '   削除フラグがあって、最後のデータでない
                With Range(.cells(rows(i)(0), 1), .cells(rows(i)(1), 1)).EntireRow
    '                .Borders(xlEdgeTop).LineStyle = xlNone
    '                .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Delete
                End With
                Range(.cells(rows(i)(0), 1), .cells(rows(i)(0), cols(2))).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
        Next
    End With
End Sub


'   ランクを計算してシートに書き込む
Private Sub fillRank(ByVal mode As Integer, _
        ByVal cel As Range, ByRef settings As Object)
    Dim defs As Object
    Dim ranks(1) As Variant
    Dim lcol, rankNum As Long
    Dim i, j, lim, maxLim As Integer
    
    '   行データの取得
    lcol = getColumnIndex(BE_Rank & BE_SuffixBase, cel) - 1
    Set defs = getRowValues(cel, Nothing, Array(1, lcol))
    If defs(BE_Species) = "" Then Exit Sub
    rankNum = settings(BE_SetRankNum)
    For i = 0 To 1
        '   データの取得
        ranks(i) = getRanks(mode, settings, defs, (i = 1))
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
    Dim lcol As Long
'    Call enableEvent(False)
    If rowNum > 1 Then
        Range(cel.Offset(1, 0), cel.Offset(rowNum - 1, 0)).EntireRow.Insert
    End If
    lcol = cel.ListObject.DataBodyRange.columns.count - 1
    With Range(cel, cel.Offset(rowNum - 1, lcol))
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 16
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 16
            .Weight = xlThin
        End With
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
'    Call enableEvent(True)
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
Private Function getRanks(ByVal mode As Integer, _
                    ByRef settings As Object, ByRef defs As Object, _
                    Optional ByVal predict As Boolean = False) As Variant
    Dim enemy As Monster
    Dim species As String
    Dim rank() As Variant
    Dim lcol, wthNum, wi As Long
    
    '   敵の設定
    If Not setEnemy(enemy, defs, mode, settings(BE_SetAtkDelay), _
            settings(BE_SetWithLimit)) Then
        Exit Function
    End If
    '   天候の数。ランクの数は天候＋１
    If mode = C_IdGym Then
        wthNum = Range(R_WeatherTable).rows.count
    Else    '   トレーナーバトルは天候に変わりがない
        wthNum = 0
    End If
    ReDim rank(wthNum)
    For wi = 0 To wthNum
        rank(wi) = getKtRank(settings(BE_SetRankNum), enemy, _
            (settings(BE_SetRankVar) = "KTR"), defs(BE_UpperCP), defs(BE_LowerCP), _
            predict, wi, settings(BE_SetAtkDelay))
    Next
    getRanks = rank
End Function

'   敵の設定
Private Function setEnemy(ByRef enemy As Monster, ByRef defs As Object, _
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
        Call getDummyMonster(enemy, mode, atkDelay, types)
        setEnemy = True
        Exit Function
    End If
    If Not speciesExists(species) Then Exit Function
    '  敵の設定
    Call getMonster(enemy, species, defs(BE_PL), _
            defs(BE_ATK), defs(BE_Def), defs(BE_IHP), defs(BE_HP))
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
Public Function onRankingSheetChange(ByVal Target As Range, ByVal rSetting As String) As Boolean
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
        Select Case .HeaderRowRange.cells(1, Target.column).Text
            Case BE_Species
                Call speciesChange(Target, rSetting)
            Case BE_Memo
                Call memoChange(Target)
            Case BE_PL
                Call calcCp(Target)
            Case BE_ATK
                Call calcCp(Target)
            Case BE_Def
                Call calcCp(Target)
            Case BE_IHP
                Call calcCp(Target)
            Case BE_HP
                Call hpChange(Target)
            Case BE_CP
                Call cpChange(Target)
            Case BE_NormalAttack
                Call AtkChange(Target)
            Case BE_SpecialAttack
                Call AtkChange(Target)
        End Select
    End With
    onRankingSheetChange = True
End Function

'   種族名の変更
Public Sub speciesChange(ByVal Target As Range, ByVal rSetting As String)
    Dim species, natk, satk As String
    Dim row As Long
    If speciesExpectation(Target) Then
        species = getColumn(BE_Species, Target).Text
        Call setEnemyParams(Target, rSetting, species)
    Else
            Call setEnemyParams(Target, rSetting)
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
Private Sub setEnemyParams(ByVal Target As Range, ByVal rSetting As String, _
            Optional ByVal species As String = "", _
            Optional ByVal PL As Double = 40, _
            Optional ByVal atk As Long = 15, _
            Optional ByVal def As Long = 15, _
            Optional ByVal HP As Long = 15)
    Dim ncols, ecols, nvals, evals As Variant
    Dim atkNames As Variant
    Dim i As Long
    Dim settings As Object
    
    Set settings = getSettings(rSetting)
    ncols = Array(BE_PL, BE_ATK, BE_Def, BE_UpperCP, BE_LowerCP)
    ecols = Array(BE_IHP, BE_NormalAttack, BE_SpecialAttack)
    If species = "" Then
        nvals = Array("", "", "", settings(BE_DefCpUpper), settings(BE_DefCpLower))
        evals = Array("", "", "")
    Else
        nvals = Array(PL, atk, def, settings(BE_DefCpUpper), settings(BE_DefCpLower))
        atkNames = seachAndGetValues(species, SA1_Name, shSpeciesAnalysis1, _
                            Array(SA1_CDSP_NormalAtkName & "1", SA1_CDSP_SpecialAtkName & "1"))
        evals = Array(HP, atkNames(0), atkNames(1))
    End If
    '   イベントを発生させない列
    Call enableEvent(False)
    For i = 0 To UBound(nvals)
        getColumn(ncols(i), Target).Value = nvals(i)
    Next
    Call enableEvent(True)
    '   イベントを発生させる列
    For i = 0 To UBound(evals)
        getColumn(ecols(i), Target).Value = evals(i)
    Next
End Sub

'   CPを計算する
Private Sub calcCp(ByVal Target As Range)
    Dim HP As Double
    Dim CP As Long
    Dim enAttrs As Variant
    
    enAttrs = getRowValues(Target, Array( _
                BE_Species, BE_PL, BE_ATK, BE_Def, BE_IHP))
    Call enableEvent(False)
    If enAttrs(0) <> "" And enAttrs(1) > 0 Then ' Check Species and PL
        CP = getCP(enAttrs(0), enAttrs(1), enAttrs(2), enAttrs(3), enAttrs(4))
        HP = getPower(enAttrs(0), "HP", enAttrs(4), enAttrs(1))
        getColumn(BE_CP, Target).Value = CP
        getColumn(BE_HP, Target).Value = Fix(HP)
    Else
        getColumn(BE_CP, Target).Value = ""
        getColumn(BE_HP, Target).Value = ""
    End If
    Call enableEvent(True)
End Sub

'   CPが変わったので、HP、iHPを再計算する。
Private Sub cpChange(ByVal Target As Range)
    Dim hps As Variant
    Dim enAttrs As Variant
    enAttrs = getRowValues(Target, Array( _
                BE_Species, BE_PL, BE_ATK, BE_Def, BE_CP))
    Call enableEvent(False)
    If enAttrs(0) <> "" And enAttrs(4) > 0 Then
        hps = getHPbyCP(enAttrs(0), enAttrs(4), enAttrs(1), enAttrs(2), enAttrs(3))
        getColumn(BE_HP, Target).Value = hps(0)
        getColumn(BE_IHP, Target).Value = hps(1)
    Else
        getColumn(BE_HP, Target).Value = ""
        getColumn(BE_IHP, Target).Value = ""
    End If
    Call enableEvent(True)
End Sub

'   HPが変わったので、CP、iHPを再計算する。
Private Sub hpChange(ByVal Target As Range)
    Dim cps As Variant
    Dim enAttrs As Variant
    enAttrs = getRowValues(Target, Array( _
                BE_Species, BE_PL, BE_ATK, BE_Def, BE_HP))
    Call enableEvent(False)
    If enAttrs(0) <> "" And enAttrs(4) > 0 Then
        cps = getCPbyHP(enAttrs(0), enAttrs(4), enAttrs(1), enAttrs(2), enAttrs(3))
        getColumn(BE_CP, Target).Value = cps(0)
        getColumn(BE_IHP, Target).Value = cps(1)
    Else
        getColumn(BE_CP, Target).Value = ""
        getColumn(BE_IHP, Target).Value = ""
    End If
    Call enableEvent(True)
End Sub

'  全てワイルドカードに設定
Public Sub setWildCardAll(ByVal sh As Worksheet, ByVal rSetting As String)
    Dim rcel As Range
    Dim settings As Object
    Dim row As Long
    Dim cols As Variant
    
    Set settings = getSettings(rSetting)
    Call ClearAllRanking(sh)
    Call doMacro(msgSetWildCard)
    With sh.ListObjects(1).DataBodyRange
        cols = getColumnIndexes(.Parent, _
                Array(BE_Species, BE_Memo, BE_UpperCP, BE_LowerCP))
        .cells(1, cols(0)) = "?"
        .cells(1, cols(2)) = settings(BE_DefCpUpper)
        .cells(1, cols(3)) = settings(BE_DefCpLower)
        row = 2
        For Each rcel In shClassifiedByType.ListObjects(1) _
                    .ListColumns(CBT_Type).DataBodyRange
            .cells(row, cols(0)) = "?"
            .cells(row, cols(1)) = rcel.Text
            Call setTypeColorsOnCell(.cells(row, cols(1)), , True)
            .cells(row, cols(2)) = settings(BE_DefCpUpper)
            .cells(row, cols(3)) = settings(BE_DefCpLower)
            row = row + 1
        Next
    End With
    Call doMacro
End Sub

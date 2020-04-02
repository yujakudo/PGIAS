Attribute VB_Name = "Ranking"


Option Explicit

'   �_�~�[�ݒ�i�[�p
Public Type DummySet
    Power As Object
    Attack(1) As Object
End Type

'   �����L���O�̃t���O
Enum RankFlag
    FRNK_REGULAR = 1
    FRNK_NEWENTRY = 2
    FRNK_NEWENTRY2 = 4
    FRNK_DROPENTRY = 8
End Enum

'   setDefParams�̃t���O
Enum DefParamFlag
    FDP_UNSET_ATK = 1
    FDP_AUTO = 2
End Enum

'   �V��̐ݒ�
Public Function onSetWeatherClick(ByVal sh As Worksheet, _
                ByVal isAll As Boolean) As Boolean
    Dim sWth As String
    Dim rrow As Variant
    
    onSetWeatherClick = False
    Call doMacro(msgCopyingRegion)
    If isAll Then
        rrow = False
    Else
        rrow = getExecRows(Selection)
        If Not IsArray(rrow) Then
            Exit Function
            Call doMacro
        End If
        rrow = Array(rrow(UBound(rrow))(0), rrow(0)(1))
    End If
    Call setWeather(sh, , rrow)
    Call doMacro
    onSetWeatherClick = True
End Function

'   �v�Z����{�^��
Public Function onCalcRankingClick(ByVal sh As Worksheet, _
                ByVal isAll As Boolean) As Boolean
    onCalcRankingClick = False
'    If Not shIndividual.checkPL() Then Exit Function
    If Not isAll And (ActiveCell.CountLarge <> 1 Or _
        Application.Intersect(ActiveCell, sh.ListObjects(1).DataBodyRange) Is Nothing) Then Exit Function
    Call doMacro(msgstr(msgProcessing, Array(cmdCalculate, msgRanking)))
    If isAll Then
        Call SetAllRanking(sh)
    Else
        Call SetRanking(Selection)
    End If
    Call doMacro
    onCalcRankingClick = True
End Function

'   �N���A�E�폜�{�^��
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
    If remove Then cmd = cmdRemove Else cmd = cmdClear
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

'   �����N���ׂĂ��v�Z
Public Sub SetAllRanking(ByVal sh As Worksheet)
    Dim row, col As Long
    Dim cel As Range
    Dim settings As Object
    Dim dmySet As DummySet
    Dim sWth As String
    Dim wth As Integer
    Dim stime As Double
    
    stime = Timer
    sWth = sh.Range(CR_R_WeatherGuess).text
    wth = getWeatherIndex(sWth)
    Set settings = getSettings(sh.Range(CR_R_Settngs))
    dmySet = getDummySettings(sh)
    Call ClearAllRanking(sh)
    With sh.ListObjects(1).DataBodyRange
        col = getColumnIndex(CR_Species, .Parent)
        row = 1
        Do While row <= .rows.count
            Set cel = .cells(row, col)
            If cel.text <> "" Then
                Call SetARanking(cel, settings, dmySet)
                Call setWeatherToCell(getColumn(CR_Weather, cel), wth)
            End If
            row = row + 1
        Loop
    End With
    Call copyRegion(cel, , wth)
    Call setTimeAndDate(sh.Range(CR_R_AllCalcTime), stime)
End Sub

Private Sub SetRanking(ByVal rng As Range)
    Dim i, col As Long
    Dim lo As ListObject
    Dim rows As Variant
    Dim settings As Object
    Dim dmySet As DummySet
    Dim cel As Range
    Dim rnkNum As Integer
    
    Set settings = getSettings(rng.Parent.Range(CR_R_Settngs))
    dmySet = getDummySettings(rng.Parent)
    rnkNum = settings(CR_SetRankNum)
    Set lo = rng.ListObject
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    rows = getExecRows(rng)
    col = getColumnIndex(CR_Species, lo)
    If IsArray(rows) Then
        For i = 0 To UBound(rows)
            Call ClearDataRange(lo, Array(rows(i)))
            Set cel = lo.DataBodyRange.cells(rows(i)(0), col)
            Call SetARanking(cel, settings, dmySet)
            rows(i)(1) = rows(i)(0) + rnkNum - 1
            Call copyRegion(cel, Array(rows(i)(0), rows(i)(1)))
        Next
    End If
End Sub

'   �����N����v�Z
Private Sub SetARanking(ByVal Target As Range, _
            ByRef settings As Object, ByRef dmySet As DummySet)
    Dim stime As Double
    Dim ssbj, memo As String
    Dim cel As Range
    Dim wthNum As Integer
    
    stime = Timer
    ssbj = Target.text
    memo = getColumn(CR_Memo, Target).text
    If memo <> "" Then ssbj = ssbj & "(" & memo & ")"
    If settings(CR_SetMode) = C_Gym Then wthNum = weathersNum() + 1 Else wthNum = 1
    Call dspProgress(msgstr(msgCalcRank, Array(ssbj)), _
            2 * wthNum * shIndividual.ListObjects(1).DataBodyRange.rows.count)
    Call fillRank(Target, settings, dmySet)
    
    Call dspProgress("", 0)
    Set cel = getColumn(CR_Time, Target)
    cel.value = Timer - stime
    cel.NumberFormatLocal = "0.0_ "
    cel.Offset(1, 0).value = Now
    cel.Offset(1, 0).NumberFormatLocal = "m/d h:mm"
End Sub

'   ���ׂăN���A
Private Sub ClearAllRanking(ByVal sh As Worksheet, _
        Optional ByVal remove As Boolean = False)
    Dim cel As Range
    Dim row As Long
    Dim lo As ListObject
    Set lo = getListObject(sh)
    If remove Then
        Call ClearDataRange(lo, Array( _
            Array(1, lo.DataBodyRange.rows.count)), remove)
    Else
        Call ClearCalcedRank(lo.DataBodyRange, remove)
    End If
    sh.Range(CR_R_AllCalcTime).ClearContents
    Application.Goto lo.DataBodyRange.cells(1, 1)
End Sub

'   �͈͂ŃN���A
Private Sub ClearCalcedRank(ByVal rng As Range, _
            Optional ByVal remove As Boolean = False)
    Dim rows As Variant
    rows = getExecRows(rng)
    If IsArray(rows) Then
        Call ClearDataRange(rng.ListObject, rows, remove)
    End If
End Sub

'   �͈͂��A��������s���擾����
Private Function getExecRows(ByVal rng As Range) As Variant
    Dim rows(), num As Long
    Dim col, row, srow As Long
    
    ReDim rows(rng.rows.count)
    If rng.ListObject Is Nothing Then Exit Function
    If rng.ListObject.DataBodyRange Is Nothing Then Exit Function
    With rng.ListObject.DataBodyRange
        col = getColumnIndex(CR_Species, .Parent)
        srow = rng.row - .row + 1
        If srow < 1 Then srow = 1
        row = rng.row - .row + rng.rows.count
        If row > .rows.count Then row = .rows.count
        Do
            '   �푰���������Ă�������J�n�s
            If .cells(row, col).text <> "" Then
                rows(num) = Array(row, 0)
                If num > 0 Then rows(num)(1) = rows(num - 1)(0) - 1
                num = num + 1
            End If
            row = row - 1
        Loop Until row < 1 Or (row < srow And .cells(row + 1, col).text <> "")
        If num = 0 Then Exit Function
        row = rows(0)(0) + 1
        Do While row <= .rows.count
            If .cells(row, col).text <> "" Then
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

'   rows�̍s�͈͂��N���A����
Private Sub ClearDataRange(ByVal lo As ListObject, ByVal rows As Variant, _
            Optional ByVal remove As Boolean = False)
    Dim cols As Variant
    Dim i, col As Long
    With lo.DataBodyRange
        cols = Array(1, getColumnIndex(CR_Time, lo), .columns.count)
        For i = 0 To UBound(rows)
            '   �폜�t���O���Ȃ����A�Ō�̃f�[�^�Ȃ�2�s�c���ăN���A
            If Not remove Or (rows(i)(1) - rows(i)(0)) + 1 = .rows.count Then
                If rows(i)(1) > rows(i)(0) + 1 Then
                    With Range(.cells(rows(i)(0) + 2, 1), .cells(rows(i)(1), 1)).EntireRow
                        .Delete
                    End With
                End If
                If Not remove Then
                    Range(.cells(rows(i)(0), cols(1)), .cells(rows(i)(0) + 1, cols(2))).ClearContents
                Else
'                    col = getColumnIndex(CR_Species, lo)
'                    .cells(rows(i)(0), col).Value = ""
'                    .cells(rows(i)(0) + 1, col).Value = ""
                    Range(.cells(rows(i)(0), 2), .cells(rows(i)(0) + 1, cols(2))).ClearContents
                End If
                Call setBorders(.Parent, Array(rows(i)(0), rows(i)(0) + 1), True)
            Else    '   �폜�t���O�������āA�Ō�̃f�[�^�łȂ�
                With Range(.cells(rows(i)(0), 1), .cells(rows(i)(1), 1)).EntireRow
                    .Delete
                End With
                Call setBorders(.Parent, Array(rows(i)(0), rows(i)(0)), True)
            End If
        Next
    End With
End Sub

'   �g���ň͂�
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

'   �����N���v�Z���ăV�[�g�ɏ�������
Private Sub fillRank(ByVal cel As Range, ByRef settings As Object, _
                ByRef dmySet As DummySet)
    Dim curRanks() As KtRank
    Dim prdRanks() As KtRank
    Dim i As Integer
    Dim cnt As Object
    Dim autoTarget As String
    
    Call insertRows(cel, settings(CR_SetRankNum))
    curRanks = getRanks(cel, settings, dmySet, False)
    autoTarget = settings(C_AutoTarget)
    If autoTarget <> "" And autoTarget <> C_None Then
        Call shIndividual.SetAutoTargetPL(autoTarget, settings(C_Level))
        shIndividual.Calculate
        prdRanks = getRanks(cel, settings, dmySet, True)
        Call shIndividual.SetAutoTargetPL
        shIndividual.Calculate
    Else
        prdRanks = getRanks(cel, settings, dmySet, True)
    End If
'    Set cnt = setRegularFlag(curRanks)
'    Call setRegularFlag(prdRanks, cnt)
    Call setRegularFlag(curRanks)
    Call setRegularFlag(prdRanks)
    Call setNewEntryFlag(curRanks, prdRanks)
    Call writeRanks(cel, curRanks, False)
    Call writeRanks(cel, prdRanks, True)
End Sub

'   �K�v���̍s��}������
Private Sub insertRows(ByVal cel As Range, ByVal rowNum As Long)
    Dim lcol, srow, rows As Long
    If cel.Offset(1, 0).text = "" Then rows = 2 Else rows = 1
    If rowNum > rows Then
        Range(cel.Offset(rows, 0), cel.Offset(rowNum - 1, 0)).EntireRow.Insert
    End If
    srow = cel.ListObject.DataBodyRange.cells(1, 1).row
    srow = cel.row - srow + 1
    Call setBorders(cel, Array(srow, srow + rowNum - 1), True)
End Sub

'   �V���\�������w�肵�āA��C���f�N�X�̃I�t�Z�b�g�𓾂�
Private Function getColumnOffset(ByRef cel As Range, _
                            ByRef colNames As Variant, _
                            ByVal isPrediction As Boolean, _
                    Optional ByVal weather As Integer = 0, _
                    Optional ByVal isIndex As Boolean = True) As Variant
    Dim suffix As String
    Dim cols As Variant
    Dim i As Long
    '   �\���A�V����T�t�B�b�N�X�𓾂�
    If isPrediction Then
        suffix = CR_SuffixPredict & CR_SuffixWeather
    Else
        suffix = CR_SuffixBase & CR_SuffixWeather
    End If
    For i = 0 To UBound(colNames)
        colNames(i) = colNames(i) & suffix & Trim(weather)
    Next
    If isIndex Then
        '   �C���f�b�N�X�𓾂ăI�t�Z�b�g�ɕϊ�
        cols = getColumnIndexes(cel, colNames)
        For i = 1 To UBound(cols)
            cols(i) = cols(i) - cols(0)
        Next
        cols(0) = 0
    End If
    getColumnOffset = cols
End Function

'   �����N�f�[�^�̏�������
Private Sub writeRanks(ByVal cel As Range, ByRef ranks() As KtRank, _
                                Optional ByVal isPrediction As Boolean = False)
    Dim i, j, wth As Long
    Dim cols, colNames As Variant
    Dim wcel As Range
    Dim tcolName As String
    Dim placed As String
    Static exists As String
    colNames = Array(CR_Rank, CR_CtrName, CR_CtrNormalAttack, _
                    CR_CtrSpecialAttack, CR_KT, CR_KTR)
    cols = getColumnOffset(cel, colNames, isPrediction, 0)
    tcolName = left(colNames(0), Len(colNames(0)) - 1)
    For wth = 0 To UBound(ranks)
        Set wcel = getColumn(tcolName & Trim(wth), cel)
        For i = 0 To UBound(ranks(wth).rank)
            With ranks(wth).rank(i)
                wcel.value = i + 1
                wcel.Font.Bold = (.flag And FRNK_REGULAR)
                wcel.NumberFormatLocal = "G/�W��"
                wcel.Offset(0, cols(1)).value = .nickname
                If .flag And (FRNK_NEWENTRY Or FRNK_NEWENTRY2) Then
                    wcel.Offset(0, cols(1)).Font.ColorIndex = CR_NewEntryColorIndex
                ElseIf .flag And FRNK_DROPENTRY Then
                    wcel.Offset(0, cols(1)).Font.ColorIndex = CR_DropEntryColorIndex
                Else
                    wcel.Offset(0, cols(1)).Font.ColorIndex = 0
                End If
                wcel.Offset(0, cols(1) + 1).value = .PL
                For j = 0 To 1
                    Call setAtkNames(j, .Attack(j).name, _
                            wcel.Offset(0, cols(2 + j)))
                    wcel.Offset(0, cols(2 + j) + 1).value = .Attack(j).stDPS
                    wcel.Offset(0, cols(2 + j) + 2).value = _
                        (.Attack(j).stDPS - .Attack(j).tDPS) / .Attack(j).tDPS
                Next
                wcel.Offset(0, cols(3) + 3).value = .Cycle
                wcel.Offset(0, cols(3) + 4).value = .scDPS
                wcel.Offset(0, cols(3) + 5).value = (.scDPS - .cDPS) / .cDPS
                wcel.Offset(0, cols(4)).value = .KT
                wcel.Offset(0, cols(5)).value = .KTR
            End With
            Set wcel = wcel.Offset(1, 0)
        Next
    Next
End Sub

'   �����N�v�Z
Private Function getRanks(ByVal cel As Range, ByRef settings As Object, _
                    ByRef dmySet As DummySet, _
                    Optional ByVal predict As Boolean = False) As KtRank()
    Dim enemy As Monster
    Dim CPlimit As Variant
    Dim isKTR As Boolean
    Dim wthNum, wi As Long
    Dim ranks() As KtRank
    
    '   �G��CP�㉺���̎擾
    enemy = getEnemySettings(cel, settings, dmySet, CPlimit)
    '   �V��̐��B�����N�̐��͓V��{�P
    If enemy.mode = C_IdGym Then
        wthNum = weathersNum()
    Else    '   �g���[�i�[�o�g���͓V��ɕς�肪�Ȃ�
        wthNum = 0
    End If
    ReDim ranks(wthNum)
    isKTR = (settings(CR_SetRankVar) = "KTR")
    For wi = 0 To wthNum
        ranks(wi) = getKtRank(settings(CR_SetRankNum), enemy, isKTR, _
                CPlimit, predict, wi, settings(CR_SetSelfAtkDelay))
    Next
    getRanks = ranks
End Function

'   �V��ʑS�ĂɃ����N�C���������̃`�F�b�N
Private Function setRegularFlag(ByRef ranks() As KtRank, _
                    Optional ByRef prevCount As Object = Nothing) As Object
    Dim wth, i, j As Integer
    Dim rlim, wrim As Integer
    Dim appr, hash2 As Object
    
    rlim = UBound(ranks(0).rank)
    wrim = UBound(ranks)
    Set appr = CreateObject("Scripting.Dictionary")
    Set hash2 = CreateObject("Scripting.Dictionary")
    '   ���O���J�E���g
    For wth = 0 To wrim
        For i = 0 To rlim
            With ranks(wth).rank(i)
                If appr.exists(.nickname) Then
                    appr.item(.nickname) = appr.item(.nickname) + 1
                Else
                    appr.item(.nickname) = 1
                End If
            End With
        Next
    Next
    '   �t���O�𗧂Ă�
    For wth = 0 To wrim
        For i = 0 To rlim
            With ranks(wth).rank(i)
                '   �o�ꐔ���V��+1�������烌�M�����[
                If appr.item(.nickname) > wrim Then
                    .flag = .flag Or FRNK_REGULAR
                End If
                '   ���݃J�E���g���w�肳�ꂽ�ꍇ
                If Not prevCount Is Nothing Then
                    '   ���݃J�E���g�ɂȂ�������V�G���g���[
                    If Not prevCount.exists(.nickname) Then
                        If Not hash2.exists(.nickname) Then
                            .flag = .flag Or FRNK_NEWENTRY
                            hash2.item(.nickname) = 1
                        Else
                            .flag = .flag Or FRNK_NEWENTRY2
                        End If
                    End If
                End If
            End With
        Next
    Next
    Set setRegularFlag = appr
End Function


'   �\���ɐV�o��̃t���O�𗧂Ă�
Private Sub setNewEntryFlag(ByRef curRanks() As KtRank, _
                            ByRef prdRanks() As KtRank)
    Dim wth, i, j As Integer
    Dim hash, hash2 As Object
    Dim rlim As Integer
    
    rlim = UBound(curRanks(0).rank)
    Set hash = CreateObject("Scripting.Dictionary")
    Set hash2 = CreateObject("Scripting.Dictionary")
    For wth = 0 To UBound(curRanks)
        hash.RemoveAll
        hash2.RemoveAll
        For i = 0 To rlim
            hash.item(curRanks(wth).rank(i).nickname) = 1
        Next
        For i = 0 To rlim
            With prdRanks(wth).rank(i)
                hash2.item(.nickname) = 1
                If Not hash.exists(.nickname) Then
                    .flag = .flag Or FRNK_NEWENTRY
                End If
            End With
        Next
        For i = 0 To rlim
            With curRanks(wth).rank(i)
                If Not hash2.exists(.nickname) Then
                    .flag = .flag Or FRNK_DROPENTRY
                End If
            End With
        Next
    Next
End Sub

'   �Z���̑I���E�ύX

'   �Z���̑I��
Public Function onRankingSelectionChange(ByVal Target As Range) As Boolean
    Dim ridx As Variant
    onRankingSelectionChange = False
    If Target.CountLarge <> 1 Then Exit Function
    If Not Application.Intersect(Target, Target.Parent.Range(CR_R_ListSelect)) Is Nothing Then
        Call setListList(Target)
        onRankingSelectionChange = True
        Exit Function
    End If
    With Target.Parent.ListObjects(1)
        If Application.Intersect(Target, .DataBodyRange) Is Nothing Then Exit Function
        Call setInputList
        ridx = getRowIndex(Target, True)
        Select Case .HeaderRowRange.cells(1, Target.column).text
            Case CR_Weather
                If ridx(0) = 0 Then Call weatherSelected(Target)
            Case CR_Memo
                If ridx(0) = 0 And ridx(1) <> "?" Then Target.Font.ColorIndex = 0
            Case CR_Attacks
                If ridx(0) >= 0 Then Call AtkSelected(Target, ridx(0), ridx(1))
        End Select
    End With
    onRankingSelectionChange = True
End Function

'   ���X�g�̃v���_�E�����X�g�̐ݒ�
Private Function setListList(ByVal Target As Range)
    With Target.Validation
        .Delete
        .Add xlValidateList, Formula1:=Join(shList.getListNames(), ",")
    End With
End Function

'   ��C���f�N�X�̎擾�B
'   0:��s��(�푰���Ɠ��s)�A1:2�s�ځA-1:���̑�
Private Function getRowIndex(ByVal Target As Range, _
                    Optional ByVal needSpecies As Boolean = False) As Variant
    Dim spCel As Range
    Dim species As String
    Set spCel = getColumn(CR_Species, Target)
    getRowIndex = -1
    species = spCel.text
    If species <> "" Then
        getRowIndex = 0
    Else
        species = spCel.Offset(-1, 0).text
        If species <> "" Then getRowIndex = 1
    End If
    If needSpecies Then
        If species = "�H" Then species = "?"
        getRowIndex = Array(getRowIndex, species)
    End If
End Function

'   �Z���l�̕ύX
Public Function onRankingSheetChange(ByVal Target As Range, _
                        Optional ByVal nullAtk As Boolean = False) As Boolean
    Dim title As String
    Dim ridx As Variant
    Dim rSettings As Range
    Dim cel As Range
    
    onRankingSheetChange = False
    Set rSettings = Target.Parent.Range(CR_R_Settngs)
    If Target.row = 1 Then
        '   ����̓V��ݒ�
        If Not Application.Intersect(Target, Target.Parent.Range(CR_R_WeatherGuess)) Is Nothing Then
            Call setWeatherToCell(Target)
            With Target.Parent
                onRankingSheetChange = True
                Call enableEvent(False)
                If Target.text <> "" Then
                    .chkAll.value = True
                    .cmbCommand.value = cmdSetWeather
                Else
                    .chkAll.value = False
                    .cmbCommand.value = cmdCalculate
                End If
                Call enableEvent(True)
            End With
        End If
        Exit Function
    End If
    With Target.Parent.ListObjects(1)
        If Target.row <= .HeaderRowRange.row Then
            '   �ݒ�̃��[�h���ύX���ꂽ�ꍇ
            If Not Application.Intersect(Target, rSettings) Is Nothing Then
                onRankingSheetChange = settingsChange(Target, rSettings)
            End If
            Exit Function
        End If
        '   �e�[�u���̃f�[�^�̈�łȂ��Ȃ�I��
        If Target.CountLarge <> 1 Or _
            Application.Intersect(Target, .DataBodyRange) Is Nothing Then Exit Function
        '   �ύX�Z���̗�^�C�g��
        title = .HeaderRowRange.cells(1, Target.column).text
        '   �S�p�X�y�[�X�݂̂̓N���A���ďI��
        If title <> CR_Weather And (Target.text = "" Or Target.text = "�@") Then
            Call enableEvent(False)
            Target.ClearContents
            Target.Validation.Delete
            Call enableEvent(True)
            Exit Function
        End If
    End With
    '   ��C���f�b�N�X�Ǝ푰��
    ridx = getRowIndex(Target, True)
    If ridx(0) = 0 And title = CR_Weather Then
        Call weatherChange(Target)
    ElseIf ridx(0) = 0 And title = CR_Species Then
        Call speciesChange(Target, rSettings, nullAtk)
    ElseIf ridx(0) = 0 And title = CR_Memo Then
        Call memoChange(Target)
    ElseIf ridx(0) = 1 And title = CR_CPHP Then
        '   CP/HP��2�s�ڂ�HP�̒l�`�F�b�N
        Call checkHPValue(Target)
        Call calcParams(Target)
    ElseIf title = CR_PL Or title = CR_ATK _
            Or title = CR_DEF Or title = CR_HP Or title = CR_CPHP Then
        Call calcParams(Target)
    ElseIf title = CR_Attacks Then
        Call AtkChange(Target, True, ridx(0), ridx(1))
    End If
    onRankingSheetChange = True
End Function

'   �ݒ肪�ύX���ꂽ
Private Function settingsChange(ByVal Target As Range, _
                    ByVal rngSettings As Range) As Boolean
    Dim key As String
    Dim idx As Long
    
    settingsChange = False
    key = Target.Offset(-1, 0).text
    If key = CR_SetMode Then
        Call doMacro(msgChaingingSettings)
        Call changeRankingMode(Target)
        Call doMacro
        settingsChange = True
    ElseIf key = C_AutoTarget Then
        If Target.text <> C_None And Target.text <> "" Then
            Call doMacro(msgChaingingSettings)
            Call setSettings(rngSettings, getAutoTargetSettings(Target.text))
            idx = WorksheetFunction.Match(CR_SetMode, rngSettings.rows(1), 0)
            Call changeRankingMode(rngSettings.cells(2, idx))
            Call doMacro
        End If
        settingsChange = True
    End If
End Function

'   �V��Z���̑I���B���͋K���̐ݒ�
Private Sub weatherSelected(ByVal Target As Range)
    Dim lst As String
    Dim ri As Long
    lst = "�@"
    With Range(R_WeatherTable)
        For ri = 1 To .rows.count
            lst = lst & "," & .cells(ri, 1)
        Next
    End With
    Call setInputList(Target, lst)
End Sub

'   �V��Z���̕ύX�B�F��t���ė̈���R�s�[
Private Sub weatherChange(ByVal Target As Range)
    Dim wth As Integer
    Dim rrows As Variant
    Call setWeatherToCell(Target)
    rrows = getExecRows(Target)
    Call doMacro(msgCopyingRegion)
    Call copyRegion(Target, rrows(0))
    Call doMacro
End Sub

'   HP�̃Z���Ń��x����HP�ɕϊ�
Public Sub checkHPValue(ByVal Target As Range)
    Dim txt, h As String
    Dim lvl As Integer
    txt = Target.text
    If IsNumeric(txt) Then Exit Sub
    txt = StrConv(txt, vbNarrow)
    txt = StrConv(txt, vbLowerCase)
    h = left(txt, 1)
    If h = "l" Or h = "s" Then
        lvl = val(Mid(txt, 2))
        If lvl < 1 And 5 < lvl Then Exit Sub
        Call enableEvent(False)
        Target.value = Array(0, 600, 1800, 3600, 9000, 15000)(lvl)
        Call enableEvent(True)
    End If
End Sub

'   �푰���̕ύX
Private Sub speciesChange(ByVal Target As Range, _
                    ByVal settingRng As Range, _
            Optional ByVal nullAtk As Boolean = False)
    Dim species, natk, satk As String
    Dim row As Long
    Dim exists As Boolean
    Dim settings As Object
    Dim dmySet As DummySet
    Dim flag As Integer
    
    If Target.Offset(1, 0) <> "" Then
        Call insertRows(Target, 2)
    Else    '   �e�[�u���g��
        Target.Offset(1, 1).value = " "
    End If
    exists = speciesExpectation(Target)
    species = getColumn(CR_Species, Target).text
    Set settings = getSettings(settingRng)
    Call enableEvent(False)
    '   �푰�������݂��Ȃ��Ƃ�
    If Not exists Then
        '   �_�~�[�̐ݒ�
        If species = "?" Or species = "�H" Then
            dmySet = getDummySettings(Target.Parent)
            Call setDummyParam(Target, settings, dmySet.Power, True)
        Else    '   �����Ȃ̂ŃN���A
            Call clearEnemySettings(Target, True)
        End If
    Else
        '   �L���Ȏ푰���̂Ƃ�
        flag = FDP_AUTO
        If nullAtk Then flag = flag Or FDP_UNSET_ATK
        Call setDefParams(Target, settings, flag, species)
    End If
    '   �����ƓV��̃N���A
    With getColumn(CR_Time, Target)
        .ClearContents
        .Offset(1, 0).ClearContents
    End With
    getColumn(CR_Weather, Target).ClearContents
    Call enableEvent(True)
End Sub

'   ���l�̕ύX
Private Sub memoChange(ByVal Target As Range)
    Dim species As String
    Dim str As String
    species = getColumn(CR_Species, Target).text
    If species = "?" Or species = "�H" Then
        str = Replace(Target.text, "�A", ",")
        str = Replace(str, "�C", ",")
        Call enableEvent(False)
        Target.value = Replace(str, "�E", ",")
        Call setTypeColorsOnCell(Target, , True)
        Call enableEvent(True)
    End If
End Sub

'   ����̐ݒ�̎擾
'   �߂�l�́A0:Monster, 1:CP�㉺���̔z��A2:����
Private Function getEnemySettings(ByVal Target As Range, _
                ByRef settings As Object, _
                ByRef dmySet As DummySet, _
                ByRef CPlimit As Variant) As Monster
    Dim mode, i As Integer
    Dim tcel As Range
    Dim limit, memo, atks(1) As Variant
    Dim atkStr As String
    '   "getEnemySettings"�́AMonster�^�̕ϐ��Ƃ��ėp���Ă���
    Set tcel = getEnemyMon(Target, getEnemySettings)
    If tcel Is Nothing Then Exit Function
    If settings(CR_SetMode) = C_Gym Then mode = C_IdGym Else mode = C_IdMtc
    '   �Z���N���X���擾
    For i = 0 To 1
        atks(i) = getEnemySettingValue(tcel, i, CR_Attacks)
        '   �Z������̏ꍇ
        If atks(i) = "" Then
            '   �_�~�[�Ȃ�_�~�[�̐ݒ�I�u�W�F�N�g
            If getEnemySettings.species = "" Then
                Set atks(i) = dmySet.Attack(i)
            Else    '   �푰�̎w�肪����Ȃ�A�푰�V�[�g���\�ȋZ���擾
                atkStr = getSpcAttr(getEnemySettings.species, _
                        Array(SPEC_NormalAttack, SPEC_SpecialAttack)(i))
                If settings(CR_SetWithLimit_b) Then
                    atkStr = atkStr & "," & _
                            getSpcAttr(getEnemySettings.species, _
                            Array(SPEC_NormalAttack, SPEC_SpecialAttack)(i))
                End If
                atks(i) = Split(atkStr, ",")
            End If
        End If
    Next
    Call setAttacks(mode, getEnemySettings, atks(0), atks(1), _
                    settings(CR_SetEnemyAtkDelay), True)
    CPlimit = Array( _
                getEnemySettingValue(tcel, 0, CR_CPLimit), _
                getEnemySettingValue(tcel, 1, CR_CPLimit))
    If CPlimit(0) = "" Then CPlimit(0) = 0
    If CPlimit(1) = "" Then CPlimit(1) = 0
End Function

'   �G�{�̂̎擾
'   cat�i�ݒ�J�e�S���j���w�莞��PL2�s�ڂ̒l��p����
'   �߂�l�͈�s�ڋZ�̃Z��
Public Function getEnemyMon(ByVal Target As Range, _
                ByRef mon As Monster, _
                Optional ByVal cat As Integer = -1) As Range
    Dim ridx, types As Variant
    Dim acol As Long
    Dim tcel As Range
    Dim species As String
    
    ridx = getRowIndex(Target, True)
    If ridx(0) < 0 Then Exit Function
    If ridx(1) = "?" Then species = "" Else species = ridx(1)
    acol = getColumnIndex(CR_Attacks, Target)
    Set tcel = Target.Parent.cells(Target.row - ridx(0), acol)
    If cat < 0 Then cat = getEnemySettingValue(tcel, 1, CR_PL)
    If cat = 0 Then '   PL, �̒l
        Call getMonster(mon, species, _
                getEnemySettingValue(tcel, 0, CR_PL), _
                getEnemySettingValue(tcel, 0, CR_ATK), _
                getEnemySettingValue(tcel, 0, CR_DEF), _
                getEnemySettingValue(tcel, 0, CR_HP))
    ElseIf cat = 1 Then '   �p���[
        Call getMonsterByPower(mon, species, _
                getEnemySettingValue(tcel, 1, CR_ATK), _
                getEnemySettingValue(tcel, 1, CR_DEF), _
                getEnemySettingValue(tcel, 1, CR_HP))
    Else  '   CP/HP
        Call getMonsterByCpHp(mon, species, _
                getEnemySettingValue(tcel, 0, CR_CPHP), _
                getEnemySettingValue(tcel, 1, CR_CPHP))
    End If
    If ridx(1) = "?" Then
        types = Split(getEnemySettingValue(tcel, 0, CR_Memo), ",")
        If UBound(types) >= 0 Then mon.itype(0) = getTypeIndex(types(0))
        If UBound(types) > 0 Then mon.itype(1) = getTypeIndex(types(1))
    End If
    Set getEnemyMon = tcel
End Function

'   �G�ݒ�l�̎擾
Private Function getEnemySettingValue(ByVal tcel As Range, ByVal ridx As Integer, _
                    ByVal col As Variant) As Variant
    If Not IsNumeric(col) Then col = getColumnIndex(col, tcel)
    getEnemySettingValue = tcel.Offset(ridx, col - tcel.column).value
End Function

'   �G�ݒ�l�̐ݒ�
Private Sub setEnemySettingValue(ByVal tcel As Range, ByVal ridx As Integer, _
                    ByVal col As Variant, ByRef val As Variant)
    If Not IsNumeric(col) Then col = getColumnIndex(col, tcel)
    tcel.Offset(ridx, col - tcel.column).value = val
End Sub

'   ����̐ݒ�̏�������
'   �C�x���gONOFF�͌Ăяo�����ōs��
Private Function setEnemySettings(ByVal Target As Range, ByVal cat As Integer, _
                ByRef mon As Monster, _
                Optional ByRef limit As Variant = False, _
                Optional ByRef attacks As Variant = False) As Range
    Dim ridx As Variant
    Dim acol As Long
    Dim tcel As Range
    
    ridx = getRowIndex(Target, True)
    If ridx(0) < 0 Then Exit Function
    acol = getColumnIndex(CR_Attacks, Target)
    Set tcel = Target.Parent.cells(Target.row - ridx(0), acol)
    Call setEnemySettingValue(tcel, 0, CR_PL, mon.PL)
    Call setEnemySettingValue(tcel, 1, CR_PL, cat)
    Call setEnemySettingValue(tcel, 0, CR_ATK, mon.indATK)
    Call setEnemySettingValue(tcel, 0, CR_DEF, mon.indDEF)
    Call setEnemySettingValue(tcel, 0, CR_HP, mon.indHP)
    Call setEnemySettingValue(tcel, 1, CR_ATK, mon.atkPower)
    Call setEnemySettingValue(tcel, 1, CR_DEF, mon.defPower)
    Call setEnemySettingValue(tcel, 1, CR_HP, mon.hpPower)
    Call setEnemySettingValue(tcel, 0, CR_CPHP, mon.CP)
    Call setEnemySettingValue(tcel, 1, CR_CPHP, mon.fullHP)
    If IsArray(limit) Then
        Call setEnemySettingValue(tcel, 0, CR_CPLimit, limit(0))
        Call setEnemySettingValue(tcel, 1, CR_CPLimit, limit(1))
    End If
    tcel.ClearContents: tcel.Offset(1, 0).ClearContents
    If IsArray(attacks) Then
        Call setAtkNames(C_IdNormalAtk, attacks(0), tcel)
        Call setAtkNames(C_IdSpecialAtk, attacks(1), tcel.Offset(1, 0))
    End If
    Set setEnemySettings = tcel
End Function

'   �G�̐ݒ�̃N���A
'   �C�x���gONOFF�͌Ăяo�����ōs��
Private Sub clearEnemySettings(ByVal Target As Range, Optional ByVal withCpLim As Boolean = False)
    Dim ridx As Integer
    Dim rcol As Variant
    Dim lcol As String
    Dim row As Long
    
    ridx = getRowIndex(Target)
    If ridx < 0 Then Exit Sub
    If withCpLim Then lcol = CR_CPLimit Else lcol = CR_CPHP
    rcol = getColumnIndexes(Target, Array(CR_Attacks, lcol))
    row = Target.row - ridx
    With Target.Parent
        Range(.cells(row, rcol(0)), .cells(row + 1, rcol(1))).ClearContents
    End With
End Sub

'   �f�t�H���g�̃p�����[�^���Z�b�g
'   �C�x���g�}���͌Ăяo������
Private Sub setDefParams(ByVal Target As Range, _
            ByRef settings As Object, _
            Optional ByVal flag As Integer = 0, _
            Optional ByVal species As String = "", _
            Optional ByVal PL As Double = 40, _
            Optional ByVal atk As Long = 15, _
            Optional ByVal def As Long = 15, _
            Optional ByVal hp As Long = 15, _
            Optional ByVal CP As Long = 0)
    Dim mon As Monster
    Dim atks, types As Variant
    Dim tcel As Range
    Dim cat As Integer
    Dim CpUpper As Long
    
    If (flag And FDP_AUTO) Then
        Call setIVasAutoTarget(settings, species, atk, def, hp, CpUpper)
    End If
    cat = getMonBySettingVal(mon, species, PL, atk, def, hp, CP, CpUpper)
    '   �킴���w��t���O���Ȃ�������Z���Z�b�g
    If 0 = (flag And FDP_UNSET_ATK) Then
        If settings(CR_SetMode) = C_Gym Then
            atks = Array(SA1_CDSP_NormalAtkName & "1", SA1_CDSP_SpecialAtkName & "1")
        Else
            atks = Array(SA1_CDST_NormalAtkName & "1", SA1_CDST_SpecialAtkName & "1")
        End If
        atks = seachAndGetValues(species, SA1_Name, shSpeciesAnalysis1, atks)
    End If
    Set tcel = setEnemySettings(Target, cat, mon, _
            Array(settings(CR_DefCpUpper), settings(CR_DefCpLower)), _
            atks)
    types = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    Call setTypeToCell(types, getColumn(CR_Memo, tcel.Offset(1, 0)))
End Sub

'   �����ڕW�ɂ���Č̒l��ݒ�
Private Function setIVasAutoTarget(ByRef settings As Object, _
            ByVal species As String, _
            ByRef atk As Long, ByRef def As Long, ByRef hp As Long, _
            ByRef CpUpper As Long) As Integer
    Dim amode, cname As String
    Dim rivs As Variant
    Dim pos(2) As Integer

    setIVasAutoTarget = 0
    rivs = seachAndGetValues(species, SA1_Name, shSpeciesAnalysis1, _
                            Array(SA1_LeagueIV1, SA1_LeagueIV2))
    amode = settings(C_AutoTarget)
    If amode = C_League1 Then
        CpUpper = C_UpperCPl1
        rivs = rivs(0)
        setIVasAutoTarget = 1
    ElseIf amode = C_League2 Then
        CpUpper = C_UpperCPl2
        rivs = rivs(1)
        setIVasAutoTarget = 2
    Else
            Exit Function
    End If
    pos(0) = InStr(rivs, vbCrLf)
    If pos(0) > 0 Then
        rivs = left(rivs, pos(0) - 1)
    End If
    pos(0) = InStr(rivs, ":")
    If pos(0) < 1 Then Exit Function
    pos(1) = InStr(rivs, ",")
    If pos(1) < 1 Then pos(1) = Len(rivs) + 1
    rivs = Trim(Mid(rivs, pos(0) + 1, pos(1) - pos(0) - 1))
    atk = val("&H" + Mid(rivs, 1, 1))
    def = val("&H" + Mid(rivs, 2, 1))
    hp = val("&H" + Mid(rivs, 3, 1))
End Function


'   �ݒ�l��胂���X�^�𓾂�
'   �߂�̓J�e�S���ԍ�
Private Function getMonBySettingVal(ByRef mon As Monster, _
            ByVal species As String, _
            ByVal PL As Double, ByVal atk As Long, _
            ByVal def As Long, ByVal hp As Long, _
            ByVal CP As Long, _
            Optional ByVal CpUpper As Long = 0) As Integer
    Dim cat As Integer
    Dim nPL, cCP, k As Double
    If PL > 0 Then
        getMonBySettingVal = 0
        If CpUpper > 0 Then
            PL = getPLbyCP2(CpUpper, species, atk, def, hp)
        End If
        Call getMonster(mon, species, PL, atk, def, hp)
    ElseIf CP > 0 Then
        getMonBySettingVal = 2
        If CpUpper > 0 And CP > CpUpper Then
            hp = hp * Sqr((CpUpper + 0.5) / (CP + 0.5))
            CP = CpUpper
        End If
        Call getMonsterByCpHp(mon, species, CP, hp)
    Else
        getMonBySettingVal = 1
        If CpUpper > 0 Then
            cCP = atk * Sqr(def * hp) / 10
            If Fix(cCP) > CpUpper Then
                k = Sqr((CpUpper + 0.5) / cCP)
                atk = atk * k
                def = def * k
                hp = hp * k
            End If
        End If
        Call getMonsterByPower(mon, species, atk, def, hp)
    End If
End Function

'   �ύX�ʒu�J�e�S���̎擾
Private Function getChangeCategory(ByVal Target As Range) As Integer
    Dim ridx As Variant
    ridx = getRowIndex(Target, True)
'    If ridx(0) < 0 Or Len(ridx(1)) < 2 Then
    If ridx(0) < 0 Then
        getChangeCategory = -1
    ElseIf Target.column = getColumnIndex(CR_CPHP, Target) Then
        getChangeCategory = 2
    ElseIf ridx(0) = 0 Then
        getChangeCategory = 0
    Else
        getChangeCategory = 1
    End If
End Function

'   �p�����[�^���v�Z���ď�������
Private Sub calcParams(ByVal Target As Range)
    Dim cat As Integer
    Dim mon As Monster
        cat = getChangeCategory(Target)
    If cat >= 0 Then
        Call getEnemyMon(Target, mon, cat)
        Call enableEvent(False)
        Call setEnemySettings(Target, cat, mon)
        Call enableEvent(True)
    End If
End Sub

'   �_�~�[�p�����[�^�̃Z�b�g
'   �C�x���gONOFF�͌Ăяo�����ōs��
Private Sub setDummyParam(ByVal Target As Range, _
            ByRef settings As Object, ByRef dmyPow As Object, _
            Optional ByVal withLimit As Boolean = False)
    Dim ridx, cat As Integer
    Dim mon As Monster
    Dim lim As Variant
    Dim CpUpper As Long
    
    ridx = getRowIndex(Target)
    If ridx < 0 Then Exit Sub
    CpUpper = getCpUpper(settings(C_AutoTarget), 0)
    cat = getMonBySettingVal(mon, "", 0, _
                dmyPow(CR_DmyAtkPower), dmyPow(CR_DmyDefPower), _
                dmyPow(CR_DmyHP), dmyPow(CR_DmyCP), CpUpper)
    If withLimit Then lim = Array(settings(CR_DefCpUpper), settings(CR_DefCpLower))
    Call setEnemySettings(Target, cat, mon, lim)
End Sub

'  �S�ă_�~�[�ɐݒ�
Public Sub setDummyAll(ByVal sh As Worksheet)
    Dim cel As Range
    Dim rcel As Variant
    Dim settings As Object
    Dim dmySet As DummySet
    Dim cols As Variant
    Dim row As Long
    
    Call ClearAllRanking(sh, True)
    Call doMacro(msgSetWildCard)
    Set settings = getSettings(sh.Range(CR_R_Settngs))
    dmySet = getDummySettings(sh)
    cols = getColumnIndexes(sh, Array(CR_Species, CR_Memo))
    Set cel = sh.ListObjects(1).DataBodyRange.cells(1, cols(1))
    cols(0) = cols(0) - cols(1) '   �I�t�Z�b�g�ɂ���
    cel.Offset(0, cols(0)).value = "?"
    Call setDummyParam(cel, settings, dmySet.Power, True)
    row = 1
    Call setBorders(sh, Array(row, row + 1), True)
    '   �^�C�v�ʕ\�̃^�C�v�̍s�Ń��[�v
    For Each rcel In shClassifiedByType.ListObjects(1) _
                .ListColumns(CBT_Type).DataBodyRange
        Set cel = cel.Offset(2, 0)
        row = row + 2
        cel.Offset(0, cols(0)).value = "?"
        Call setDummyParam(cel, settings, dmySet.Power, True)
        cel.value = rcel.text
        Call setTypeColorsOnCell(cel, , True)
        Call setBorders(sh, Array(row, row + 1), True)
    Next
    Call doMacro
End Sub

'   �ݒ�̎擾
Private Function getBattleSettings() As Object
    Dim settings As Object
    Set settings = getSettings(CR_R_Settngs)
End Function

'   �_�~�[�ݒ�̎擾
Private Function getDummySettings(ByVal sh As Worksheet) As DummySet
    Dim col, i As Long
    Dim obj(2) As Object
    
    i = 0
    Set obj(i) = CreateObject("Scripting.Dictionary")
    With sh.Range(CR_R_DummyEnemy)
        For col = 1 To .columns.count
            If "" = .cells(1, col).text Then
                i = i + 1
                Set obj(i) = CreateObject("Scripting.Dictionary")
            Else
                obj(i).item(.cells(1, col).text) = .cells(2, col).value
            End If
        Next
    End With
    With getDummySettings
        Set .Power = obj(0)
        For i = 0 To 1
            Set .Attack(i) = obj(i + 1)
            .Attack(i).item("name") _
                = Array(C_DummyNormalAttack, C_DummySpecialAttack)(i)
        Next
    End With
End Function

'   �V���ݒ�B
Public Sub setWeather(ByVal sh As Worksheet, _
            Optional ByVal weather As Variant = -1, _
            Optional ByVal rrow As Variant = False)
    Dim sWeather As String
    Dim cols As Variant
    Dim row As Long
    
    If weather < 0 Then
        sWeather = sh.Range(CR_R_WeatherGuess).text
        If sWeather = "" Then Exit Sub
        If sWeather = C_NotSet Then sWeather = ""
        weather = getWeatherIndex(sWeather)
    Else
        sWeather = getWeatherNameAndIndex(weather)
    End If
    With sh.ListObjects(1).DataBodyRange
        If Not IsArray(rrow) Then
            rrow = Array(1, .rows.count)
        End If
        cols = getColumnIndexes(.Parent, Array(CR_Weather, CR_Species))
        For row = rrow(0) To rrow(1)
            If .cells(row, cols(1)).text <> "" Then
                Call setWeatherToCell(.cells(row, cols(0)), weather)
            End If
        Next
    End With
    Call copyRegion(sh, rrow, weather)
End Sub

'   �̈�R�s�[
'   �C�x���g�Ǘ��͌Ăяo�����ōs��
Private Sub copyRegion(ByVal table As Variant, _
                    Optional ByVal rrow As Variant = False, _
                    Optional ByVal weather As Variant = -1)
    Dim sWeather As String
    Dim cols As Variant
    Dim width, tcol As Long
    Dim rfrom, rto As Range
    cols = getColumnIndexes(table, Array( _
            CR_Weather, CR_Rank & CR_SuffixBase, _
            CR_Rank & CR_SuffixBase & CR_SuffixWeather & "0"))
    With getListObject(table).DataBodyRange
        If weather < 0 Then
            sWeather = .cells(rrow(0), cols(0)).text
            weather = getWeatherIndex(sWeather)
        Else
            sWeather = getWeatherNameAndIndex(weather)
        End If
        If Not IsArray(rrow) Then
            rrow = Array(1, .rows.count)
        End If
        width = cols(2) - cols(1)
        tcol = cols(2) + width * weather
        Range(.cells(rrow(0), tcol), .cells(rrow(1), tcol + width - 1)).copy
'        On Error GoTo Err
        Range(.cells(rrow(0), cols(1)), .cells(rrow(1), cols(2) - 1)) _
                .PasteSpecial xlPasteAll
'        On Error GoTo 0
        Application.CutCopyMode = False
        .cells(rrow(0), cols(0)).Activate
        Exit Sub
Err:
        Call copyValues(.cells(rrow(0), tcol), .cells(rrow(0), cols(1)), _
                        rrow(1) - rrow(0) + 1, width)
    End With
End Sub

'   �i�e�[�u�����R�s�y�͓{����H�j
Private Sub copyValues(ByVal fromCel As Range, ByVal toCel As Range, _
                        ByVal rh As Long, ByVal cw As Long)
    Dim row, col As Long
    For row = 0 To rh - 1
        For col = 0 To cw - 1
            toCel.Offset(row, col).value = fromCel.Offset(row, col).value
            toCel.Offset(row, col).Font.Color = fromCel.Offset(row, col).Font.Color
        Next
    Next
End Sub

'   ���X�g����ǉ�
Public Sub addFromList(ByVal sh As Worksheet, _
            Optional ByVal nullAtk As Boolean = False)
    Dim rcel As Range
    Dim wcel As Range
    Dim lname As String
    Dim lo As ListObject
    Dim settings As Object
    Dim flag As Integer
    
    Set settings = getSettings(sh.Range(CR_R_Settngs))
    lname = sh.Range(CR_R_ListSelect).text
    Set lo = shList.getEnemyList(lname)
    If lo Is Nothing Then Exit Sub
    If nullAtk Then flag = FDP_UNSET_ATK Else flag = 0
    Call doMacro(msgstr(msgAddingListItems, lname))
    With sh.ListObjects(1).ListColumns(CR_Species).DataBodyRange
        Set wcel = .cells(1, 1)
        If .rows.count > 2 Or wcel.text <> "" Then
            Set wcel = .cells(.rows.count, 1).Offset(1, 0)
        End If
    End With
    For Each rcel In lo.ListColumns(LI_Species).DataBodyRange
        Call setListItem(settings, rcel, wcel, flag)
    Next
    Call doMacro
End Sub

'   ���X�g���ڂ̒ǉ�
Private Function setListItem(ByRef settings As Object, _
                            ByRef rcel As Range, ByRef wcel As Range, _
                    Optional ByVal flag As Integer = 0, _
                    Optional ByVal CpUpper As Long = 0) As Boolean
    Dim attr As Object
    Dim cel As Range
    Dim row As Long
    Dim spl As String
    
    setListItem = False
    Set attr = getRowValues(rcel)
    If speciesExists(attr(LI_Species)) Then
        wcel.value = attr(LI_Species)
        wcel.Offset(1, 1).value = " "
        If attr(LI_Category) <> "" And attr(LI_Note) <> "" Then spl = ": " Else spl = ""
        wcel.Offset(0, 1).value = attr(LI_Category) & spl & attr(LI_Note)
        If attr(LI_HP) > 0 Then
            Call setDefParams(wcel, settings, flag, _
                    attr(LI_Species), attr(LI_PL), _
                        attr(LI_ATK), attr(LI_DEF), attr(LI_HP), attr(LI_CP))
        Else
            flag = flag Or FDP_AUTO
            Call setDefParams(wcel, settings, flag, attr(LI_Species))
        End If
        row = wcel.row - wcel.ListObject.HeaderRowRange.row
        Call setBorders(wcel, Array(row, row + 1))
        Set wcel = wcel.Offset(2, 0)
        setListItem = True
    End If
End Function


'   �����L���O�̃��[�h���ς�����̂ŁA�^�C�g���������ς���
Public Sub changeRankingMode(ByVal Target As Range)
    Dim cel As Variant
    Dim words As Variant
    Dim mode, curMode As Integer
    Dim sh As Worksheet
    Dim rng As Range
    
    If Target.text = C_Gym Then
        mode = C_IdGym
    ElseIf Target.text = C_Match Then
        mode = C_IdMtc
    Else
        Exit Sub
    End If
    Set sh = Target.Parent
    If mode = currentSimMode(sh) Then Exit Sub
    Set rng = sh.ListObjects(1).DataBodyRange
    If rng.rows.count > 2 Or rng.cells(1, 2) <> "" Then
        If MsgBox(msgAskChangeBattleMode, vbOKCancel) <> vbOK Then
            If Target.Parent.chkShowColumns.visible Then
                Target.value = C_Gym
            Else
                Target.value = C_Match
            End If
            Exit Sub
        End If
    End If
    Call ClearAllRanking(sh, True)
    words = Array("PS_", "PT_")
    For Each cel In Target.Parent.ListObjects(1).HeaderRowRange
        cel.value = Replace(cel.text, words(1 - mode), words(mode))
    Next
    sh.ListObjects(1).ListColumns(CR_Weather).DataBodyRange _
                .EntireColumn.Hidden = Not (mode = C_IdGym)
    Call weatherRankingVisible(sh, False)
    Target.Parent.chkShowColumns.visible = (mode = C_IdGym)
    sh.Range(CR_R_WeatherGuess).ClearContents
End Sub

'   �V��ʕ\��
Public Sub weatherRankingVisible(ByVal sh As Worksheet, _
            Optional ByVal visible As Boolean = True)
    Dim rcol As Variant
    rcol = getColumnIndexes(sh, _
            Array(CR_Rank & CR_SuffixBase & CR_SuffixWeather & "0", _
                  CR_KTR & CR_SuffixPredict & CR_SuffixWeather & Trim(weathersNum())))
    Range(sh.cells(1, rcol(0)), sh.cells(1, rcol(1))).EntireColumn.Hidden _
            = (Not visible)
End Sub

'   �W�v����Ă΂��o�ꐔ�̃J�E���g
'   �߂�l�̓j�b�N�l�[�����L�[�Ƃ����n�b�V���e�[�u���ŁA
'   ���ڂ́APL, �J�E���g, �t���O�t���J�E���g
Public Function getCountOfRanked(ByVal sh As Worksheet, _
                        ByVal isPrediction As Boolean) As Object
    Dim suffix As String
    Dim wth, wthNum As Integer
    Dim i, rankLower As Integer
    Dim row, col As Long
    Dim cel As Range
    Dim colName As String
    Dim tmp As Variant
    Dim settings As Object
    
    Set settings = getSettings(sh.Range(CR_R_Settngs))
    If isPrediction Then
        suffix = CR_SuffixPredict & CR_SuffixWeather
        rankLower = settings(CR_CountRankPr)
    Else
        suffix = CR_SuffixBase & CR_SuffixWeather
        rankLower = settings(CR_CountRankCur)
    End If
    If rankLower < 1 Then rankLower = 3
    If currentSimMode(sh) = C_IdMtc Then wthNum = weathersNum()
    
    Set getCountOfRanked = CreateObject("Scripting.dictionary")
    '   �V��+1���[�v
    For wth = 0 To wthNum
        '   ���ʂ̗�Ń��[�v
        colName = CR_Rank & suffix & Trim(wth)
        For Each cel In sh.ListObjects(1).ListColumns(colName).DataBodyRange
            '   �w��̃����N���ł�������J�E���g
            If 0 < cel.value And cel.value <= rankLower Then
                With cel.Offset(0, 1)
                    If getCountOfRanked.exists(.text) Then
                        tmp = getCountOfRanked.item(.text)
                        tmp(1) = tmp(1) + 1
                    Else
                        tmp = Array(.Offset(0, 1).value, 1, 0)
                    End If
                    If (isPrediction And .Font.ColorIndex = CR_NewEntryColorIndex) Or _
                        (Not isPrediction And .Font.ColorIndex <> CR_DropEntryColorIndex) Then
                        tmp(2) = tmp(2) + 1
                    End If
                    getCountOfRanked.item(.text) = tmp
                End With
            End If
        Next
    Next
End Function

'   ���݂̃V�~�����[�V�������[�h
Private Function currentSimMode(ByVal sh As Worksheet) As Integer
    If sh.ListObjects(1).ListColumns(CR_Weather).DataBodyRange _
                .EntireColumn.Hidden Then
        currentSimMode = C_IdMtc
    Else
        currentSimMode = C_IdGym
    End If
End Function

'   �w�b�_�̓V��ʗ̈�̃Z�b�g
Public Sub setCounterHeader(ByVal sh As Worksheet)
    Dim row, col, i, j As Long
    Dim rcol As Variant
    Dim wnc As Variant
    Dim org, torg, icel, tcel, ncel As Range
    Dim str As String
    
    Call enableEvent(False)
    rcol = getColumnIndexes(sh, _
            Array(CR_Rank & CR_SuffixBase, CR_KTR & CR_SuffixPredict))
    row = sh.ListObjects(1).HeaderRowRange.row - 1
    Set org = Range(sh.cells(row, rcol(0)), sh.cells(row + 3, rcol(1)))
    Set torg = org.cells(1, 1)
    Set icel = org.cells(1, 1).Offset(0, org.columns.count)
    For col = 0 To org.columns.count / 2 - 1
        With torg.Offset(0, org.columns.count / 2 + col)
            .ColumnWidth = torg.Offset(0, col).ColumnWidth
            str = torg.Offset(1, col).text
            .Offset(1, 0).value = Replace(str, CR_SuffixBase, CR_SuffixPredict)
            .Offset(2, 0).NumberFormatLocal = torg.Offset(2, col).NumberFormatLocal
            .Offset(3, 0).NumberFormatLocal = torg.Offset(3, col).NumberFormatLocal
        End With
    Next
    For i = 0 To weathersNum()
        If icel.text = "" Then
            org.EntireColumn.copy
            icel.EntireColumn.Insert shift:=xlToRight
            Set icel = icel.Offset(0, -org.columns.count)
        Else
            org.copy
            icel.Offset(0, 0).Select
            ActiveSheet.Paste
            For col = 0 To org.columns.count - 1
                icel.Offset(0, col).ColumnWidth = torg.Offset(0, col).ColumnWidth
            Next
        End If
        Application.CutCopyMode = False
        For j = 0 To org.columns.count - 1
            Set tcel = icel.Offset(1, j)
            tcel.value = torg.Offset(1, j).text & CR_SuffixWeather & Trim(i)
        Next
        wnc = getWeatherName(i, False, True)
        If wnc(0) = "" Then wnc(0) = "���ݒ�"
        With icel.Offset(0, 1)
            .value = wnc(0)
            .Font.Color = wnc(1)
        End With
        With icel.Offset(0, org.columns.count / 2 + 1)
            .value = wnc(0)
            .Font.Color = wnc(1)
        End With
        Range(sh.cells(1, icel.column), sh.cells(1, icel.column + org.columns.count - 1)).Clear
        Set icel = icel.Offset(0, org.columns.count)
    Next
    Call enableEvent(True)
End Sub

Private Function getBodyName(ByVal str As String) As String
    Dim pos As Long
    pos = InStr(str, "_")
    If pos < 1 Then
        getBodyName = str
        Exit Function
    End If
    getBodyName = left(str, pos - 1)
End Function


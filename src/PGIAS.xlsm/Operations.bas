Attribute VB_Name = "Operations"

Option Explicit

'   Win32
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long

'   ���͋K���̃��X�g���Z���ɐݒ肷��
Sub setInputList(Optional ByVal Target As Range = Nothing, _
                        Optional ByVal lst As String = "", _
                        Optional ByVal showList = False)
    Static lastTarget As Range
    '   ��̋K�����N���A
    If Not lastTarget Is Nothing Then
        On Error GoTo ValidationClearSkip
        lastTarget.Validation.Delete
ValidationClearSkip:
        On Error GoTo 0
        Set lastTarget = Nothing
    End If
    If Target Is Nothing Then Exit Sub
    '   �K���̐ݒ�
    With Target
        .Validation.Delete
        If lst <> "" Then
            .Validation.Add Type:=xlValidateList, Formula1:=lst
            If .text <> "" And Not InStr(lst, .text) > 0 Then .value = ""
            Set lastTarget = Target
            If showList Then SendKeys "%{Down}"
        End If
    End With
End Sub

'   �푰���̗\���ϊ����̓��͋K���ݒ�
Public Function speciesExpectation(ByVal Target As Range) As Boolean
    Dim txt As String
    Dim rng As Range
    If IsNull(Target.text) Then Exit Function
    txt = StrConv(Target.text, vbKatakana)
    Set rng = shSpecies.ListObjects(1).ListColumns(SPEC_Name).Range
    speciesExpectation = rangeExpectation(Target, rng, txt)
End Function

'   �̗̂\���ϊ����̓��͋K���ݒ�
Public Function individualExpectation(ByVal Target As Range) As Boolean
    Dim txt As String
    Dim rng As Range
    Set rng = shIndividual.ListObjects(1).ListColumns(IND_Nickname).Range
    individualExpectation = rangeExpectation(Target, rng)
    If Not individualExpectation Then
        txt = StrConv(Target.text, vbKatakana)
        individualExpectation = rangeExpectation(Target, rng, txt)
    End If
End Function

'   range�w��ł̓��͋K���ݒ�
Private Function rangeExpectation(ByVal Target As Range, _
                ByVal rng As Range, _
        Optional ByVal txt As String = "") As Boolean
    Dim first, cell As Range
    Dim cand(), stmp As String
    Dim clim, cnum, i As Integer
    
    rangeExpectation = False
    If txt = "" Then txt = Target.text
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
    '   �ŏ��̌��ɐݒ�
    Call enableEvent(False)
    Target.value = cand(0)
    '   ������₪����Ȃ烊�X�g�ɐݒ�
    If cnum > 1 Then
        ReDim Preserve cand(cnum - 1)
        Call setInputList(Target, Join(cand, ","))
        On Error GoTo Continue
        Target.Select
        SendKeys "%{Down}"
Continue:
        On Error GoTo 0
    End If
    Call enableEvent(True)
    rangeExpectation = True
End Function

'   �킴�̑I���B���͋K���̐ݒ�
Sub AtkSelected(ByVal Target As Range, _
                Optional ByVal atkClass As Integer = -1, _
                Optional ByVal species As String = "")
    Dim lo As ListObject
    Dim cname As String
    Dim lst As Variant
    
    If GetAsyncKeyState(vbKeyEscape) Then
        Call setInputList(Target)
        Exit Sub
    End If
    Set lo = Target.Parent.ListObjects(1)
    cname = lo.HeaderRowRange.cells(Target.column).text
    If atkClass < 0 Then atkClass = getAtkClassIndex(cname)
    If species = "" Then species = getColumn(C_SpeciesName, Target).text
    If Not speciesExists(species) Then Exit Sub
    lst = getAtkNames(species, True, True)
    '   ���炩�̓��͂�����A���X�g�ɂȂ��Ȃ�ݒ肵�Ȃ��ŏI��
    If Target.text <> "" And InStr(lst(atkClass), Target.text) < 1 Then Exit Sub
    ' �Q�[�W2�ƖڕW�Z�ɂ́A���I���ɖ߂��󔒂�ǉ�
    If cname = IND_SpecialAtk2 Or cname = IND_TargetNormalAtk _
        Or cname = IND_TargetSpecialAtk Then lst(atkClass) = "�@," & lst(atkClass)
    Call setInputList(Target, lst(atkClass))
End Sub

'   �킴������{�^���̏���
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

'   �푰���I������킴�\�����H����
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

'   �Z�����푰���𓾂�
Function getSpeciesFromCell(Optional ByVal cel As Range = Nothing) As String
    Dim row, col, lrow As Long
    getSpeciesFromCell = ""
    '   �s�A�푰���̎擾�ƁA����
    If cel Is Nothing Then Set cel = ActiveCell
    If IsEmpty(cel.ListObject) Then Exit Function
    getSpeciesFromCell = getColumn(C_SpeciesName, cel).text
End Function

'   �푰�̑I�������Z�b�g
Sub deselectSpecies()
    Call setSpeciesOnAtk("")
    Call filterAtkReset
    Call ShowCorrColumns(False)
End Sub

'   �킴�̕\���푰�Ńt�B���^
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

'   �푰�ł̃t�B���^�̃��Z�b�g
Sub filterAtkReset()
    Range(TBL_NormalAtk).AutoFilter Field:=1
    Range(TBL_SpecialAtk).AutoFilter Field:=1
End Sub

'   �␳�̗�̕\��/��\��
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

'   �킴�̃V�[�g�Ɏ푰������������
Private Sub setSpeciesOnAtk(ByVal species As String)
    Dim dspcol, stype As Variant
    Dim i, j, gcol As Long
    Dim cur As Worksheet
    
    Set cur = ActiveSheet
    dspcol = Array(R_NormalAtkSpeciesSelect, R_SpecialAtkSpeciesSelect)
    stype = Array("", "")
    If species <> "" Then
        stype = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    End If
    For i = 0 To 1
        With Range(dspcol(i))
            .value = species
            .Offset(0, 1).value = stype(0)
            .Offset(0, 1).Font.Color = getTypeColor(stype(0))
            .Offset(0, 2).value = stype(1)
            .Offset(0, 2).Font.Color = getTypeColor(stype(1))
            '   �킴�̃V�[�g���A�N�e�B�u�ɂ��āA�J�[�\����������悤�Ɉړ�
            With .Parent
                .Activate
                gcol = 2
'                If species <> "" Then
'                    gcol = WorksheetFunction.Match(ATK_typeMatch, _
'                            .ListObjects(1).HeaderRowRange, 0)
'                    Application.Calculate '�v����
'                End If
                Application.Goto .cells(1, gcol), True
            End With
        End With
    Next
    cur.Activate
End Sub

'   �^�C�v�A�킴�̐F�t��
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

'   �킴�̕ύX�B�F��t����
Sub AtkChange(ByVal Target As Range, _
        Optional ByVal isInput As Boolean = True, _
        Optional ByVal atkClass As Variant = -1, _
        Optional ByVal species As String = "")
    Dim typeColor As Long
    Dim atkType, atk As String
    Dim cel As Range
    
    '   �N���X�i�ʏ�/�Q�[�W�j�̎擾
    If atkClass < 0 Then atkClass = getAttackClassByHeader(Target)
    typeColor = 0
    atk = Target.text
    If atk <> "" Then
        On Error GoTo unknownAttack
        atkType = getAtkAttr(atkClass, atk, C_Type)
        On Error GoTo 0
        typeColor = getTypeColor(atkType)
    End If
    If typeColor Then
        Target.Font.Color = typeColor
        If isInput Then
            On Error GoTo addAttack
            If Target.Validation.Type = xlValidateList Then Exit Sub
addAttack:
            On Error GoTo 0
            If species = "" Then species = getSpeciesFromCell(Target)
            Call shSpecies.addAttackToSpecies(atkClass, atk, species)
        End If
    Else
        Target.Font.ColorIndex = 1
    End If
    Exit Sub
unknownAttack:
    MsgBox msgstr(msgUnknownAttackName, Array(atkClass, atk))
    Target.value = ""
End Sub

'   �w�b�_�^�C�g�����A�Z�N���X�i�ʏ�/�Q�[�W�j�̎擾
Private Function getAttackClassByHeader(ByVal Target As Range) As Integer
    Dim cel As Range
    Dim atkClass As String
    If Target.ListObject Is Nothing Then
        '   Target���e�[�u���̊O�Ȃ̂ŁA�͈͓��܂ők��
        Set cel = Target.Offset(-1, 0)
        While cel.ListObject Is Nothing And cel.row > 1
            Set cel = cel.Offset(-1, 0)
        Wend
        If cel.row = 1 Then Exit Function
        atkClass = cel.ListObject.HeaderRowRange(1, Target.column).text
    Else
        atkClass = Target.ListObject.HeaderRowRange(1, Target.column).text
    End If
    If InStr(atkClass, C_SpecialAttack) > 0 Then
        getAttackClassByHeader = C_IdSpecialAtk
    Else
        getAttackClassByHeader = C_IdNormalAtk
    End If
End Function

'   �V��̕ύX�B�F��ς���
Public Sub weatherChange(ByVal Target As Range)
    Dim idx As Integer
    On Error GoTo Err
    idx = WorksheetFunction.Match(Target.text, _
            Range(R_WeatherTable).columns(1), 0)
    On Error GoTo 0
    Target.Font.Color = Range(R_WeatherTable).cells(idx, 1).Font.Color
    Exit Sub
Err:
    Target.Font.ColorIndex = 1
End Sub

'   �e�[�u���̃w�b�_�^�C�g���̃T�t�B�b�N�X�̕\���ؑ�
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

'   �\�[�g
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

'   ���ԂƓ��t�̏�������
Public Sub setTimeAndDate(ByVal rng As Range, ByVal start As Double)
    Dim stime As Double
    stime = Timer - start
    rng.Areas(1).value = getTimeStr(stime, "'")
    rng.Areas(2).value = Now
End Sub

'   �̒l��16�i����10�i���ɕϊ�����
Public Sub decimalizeIndivValue(ByVal Target As Range)
    Dim c As Integer
    c = Asc(UCase(left(Target.text, 1))) - 55
    If 9 < c And c < 16 Then
        Target.value = c
    ElseIf Target.value = 0 And Target.text <> "0" Then
        Target.ClearContents
    End If
End Sub

'   �R���{�{�b�N�X�̃V�[�g�I�����j���[�쐬
Public Sub setComboMenu(ByRef cmb As ComboBox, _
                    Optional ByRef shs As Variant = Nothing, _
                    Optional ByRef names As Variant = Nothing)
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
    End With
End Sub

'   �R���{�{�b�N�X�̒l��胏�[�N�V�[�g�𓾂�
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

'   �푰�A�푰���̓V�[�g�Ɉړ�
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

'   �푰�A�푰���̓V�[�g�Ɉړ�
Private Sub activateSpeciesSheet(ByVal sh As Worksheet, _
                    ByVal species As String, ByVal log As Boolean)
    Dim row As Long
    With sh
        row = searchRow(species, C_SpeciesName, .ListObjects(1))
        Call jumpTo(.ListObjects(1).DataBodyRange.cells(row, 1), log)
    End With
End Sub

'   �킴�V�[�g�Ɉړ�
Public Sub jumpToAttackSheet(ByVal sh As Worksheet)
    Call doMacro(msgSelectingAttack)
    Call selectSpeciesForAtkTable
    Call doMacro
    Call jumpTo(sh, True)
End Sub

'   �̃}�b�v�Ɉړ�
Public Sub jumpToIndMap(ByVal sh As Worksheet, ByVal sameType As Boolean)
    Dim species As String
    Dim name As String
    '   �i�荞�ݏ���
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

'   �푰�}�b�v�Ɉړ�
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


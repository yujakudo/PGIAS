Attribute VB_Name = "Operations"

Option Explicit

'   Win32
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long

'   入力規則のリストをセルに設定する
Sub setInputList(Optional ByVal Target As Range = Nothing, _
                        Optional ByVal lst As String = "", _
                        Optional ByVal showList = False)
    Static lastTarget As Range
    '   先の規則をクリア
    If Not lastTarget Is Nothing Then
        lastTarget.Validation.Delete
        Set lastTarget = Nothing
    End If
    If Target Is Nothing Then Exit Sub
    '   規則の設定
    With Target
        .Validation.Delete
        If lst <> "" Then
            .Validation.Add Type:=xlValidateList, Formula1:=lst
            If .Text <> "" And Not InStr(lst, .Text) > 0 Then .Value = ""
            Set lastTarget = Target
            If showList Then SendKeys "%{Down}"
        End If
    End With
End Sub

'   種族名の予測変換候補の入力規則設定
Public Function speciesExpectation(ByVal Target As Range) As Boolean
    Dim txt, lst, fcand As String
    Dim rng, first, cell As Range
    
    speciesExpectation = False
    txt = StrConv(Target.Text, vbKatakana)
    Set rng = shSpecies.cells(2, 1).ListObject.ListColumns(SPEC_Name).Range
    Set first = rng.Find(txt, LookAt:=xlPart)
    If first Is Nothing Then Exit Function
    Set cell = first
    Do
        If InStr(cell.Text, txt) = 1 Then
            If lst <> "" Then lst = lst & ","
            lst = lst & cell.Text
            If fcand = "" Then fcand = cell.Text
        End If
        Set cell = rng.FindNext(cell)
    Loop While cell <> first
    If fcand = "" Then Exit Function
    '   最初の候補に設定
    Call enableEvent(False)
    Target.Value = fcand
    '   複数候補があるならリストに設定
    If fcand <> lst Then
        Call setInputList(Target, lst)
        Target.Select
        SendKeys "%{Down}"
    End If
    Call enableEvent(True)
    speciesExpectation = True
End Function

'   わざの選択。入力規則の設定
Sub AtkSelected(ByVal Target As Range)
    Dim lo As ListObject
    Dim ati As Long
    Dim cname, species As String
    Dim lst As Variant
    
    If GetAsyncKeyState(vbKeyEscape) Then
        Exit Sub
        Call setInputList(Target)
    End If
    Set lo = Target.Parent.ListObjects(1)
    cname = lo.HeaderRowRange.cells(Target.column).Text
    species = getColumn(C_SpeciesName, Target).Text
    If Not speciesExists(species) Then Exit Sub
    lst = getAtkNames(species, True, True)
    ati = getAtkClassIndex(cname)
    '   何らかの入力があり、リストにないなら設定しないで終了
    If Target.Text <> "" And InStr(lst(ati), Target.Text) < 1 Then Exit Sub
    ' ゲージ2と目標技には、未選択に戻す空白を追加
    If cname = IND_SpecialAtk2 Or cname = IND_TargetNormalAtk _
        Or cname = IND_TargetSpecialAtk Then lst(ati) = "　," & lst(ati)
    Call setInputList(Target, lst(ati))
End Sub

'   わざを見るボタンの処理
Sub ClickShowAttack()
    Dim colTitle As String
    Dim sh As Worksheet
    
    With ActiveCell
        If .ListObject Is Nothing Then Exit Sub
        If .row < .ListObject.DataBodyRange.row Then Exit Sub
        doMacro ("わざを選択しています。")
        If selectSpeciesForAtkTable() Then
            colTitle = .ListObject.HeaderRowRange.cells(1, .column).Text
            Set sh = shNormalAttack
            If getAtkClassIndex(colTitle) = C_IdSpecialAtk Then
                Set sh = shSpecialAttack
            End If
            sh.Activate
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
    getSpeciesFromCell = getColumn(C_SpeciesName, cel).Text
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
            .Value = species
            .Offset(0, 1).Value = stype(0)
            .Offset(0, 1).Font.Color = getTypeColor(stype(0))
            .Offset(0, 2).Value = stype(1)
            .Offset(0, 2).Font.Color = getTypeColor(stype(1))
            '   わざのシートをアクティブにして、カーソルを見えるように移動
            With .Parent
                .Activate
                gcol = 2
                If species <> "" Then
                    gcol = WorksheetFunction.Match(ATK_typeMatch, _
                            .ListObjects(1).HeaderRowRange, 0)
                    Application.Calculate '要検討
                End If
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
    
    If cell.Text = "" Then Exit Sub
    If isCsv Then
        cell.Font.Color = rgbBlack
        val = cell.Text
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
        stp = cell.Text
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
Sub AtkChange(ByVal Target As Range, Optional ByVal isInput As Boolean = True)
    Dim tc As Long
    Dim tp, atype, atk As String
    Dim species As String
    Dim cel As Range
    
    If Target.ListObject Is Nothing Then
        Set cel = Target.Offset(-1, 0)
        While cel.ListObject Is Nothing And cel.row > 1
            Set cel = cel.Offset(-1, 0)
        Wend
        If cel.row = 1 Then Exit Sub
        atype = cel.ListObject.HeaderRowRange(1, Target.column).Text
    Else
        atype = Target.ListObject.HeaderRowRange(1, Target.column).Text
    End If
    If InStr(atype, C_SpecialAttack) > 0 Then
        atype = C_SpecialAttack
    Else
        atype = C_NormalAttack
    End If
    tc = 0
    atk = Target.Text
    If atk <> "" Then
        On Error GoTo unknownAttack
        tp = getAtkAttr(atype, atk, C_TYPE)
        On Error GoTo 0
        tc = getTypeColor(tp)
    End If
    If tc Then
        Target.Font.Color = tc
        If isInput Then
            On Error GoTo addAttack
            If Target.Validation.Type = xlValidateList Then Exit Sub
addAttack:
            On Error GoTo 0
            species = getSpeciesFromCell(Target)
            Call shSpecies.addAttackToSpecies(atype, atk, species)
        End If
    Else
        Target.Font.ColorIndex = 1
    End If
    Exit Sub
unknownAttack:
    MsgBox msgstr(msgUnknownAttackName, Array(atype, atk))
    Target.Value = ""
End Sub

'   天候の変更。色を変える
Public Sub WeatherChange(ByVal Target As Range)
    Dim idx As Integer
    On Error GoTo Err
    idx = WorksheetFunction.Match(Target.Text, _
            Range(R_WeatherTable).columns(1), 0)
    On Error GoTo 0
    Target.Font.Color = Range(R_WeatherTable).cells(idx, 1).Font.Color
    Exit Sub
Err:
    Target.Font.ColorIndex = 1
End Sub

'   テーブルのヘッダタイトルのサフィックスの表示切替
Public Sub switchHeaderSuffixes(Optional ByVal show As Boolean = False)
    Dim shi As Long
    For shi = 1 To Worksheets.count
        If Worksheets(shi).ListObjects.count > 0 Then
            Call switchHeaderSuffixesATable(Worksheets(shi), show)
        End If
    Next
End Sub

Private Sub switchHeaderSuffixesATable(ByVal sh As Worksheet, _
                Optional ByVal show As Boolean = False)
    Dim col, cc, defc, pos As Long
    Debug.Print sh.name
    With sh.ListObjects(1).HeaderRowRange
        defc = .cells(1, 1).Font.Color
        For col = 1 To .columns.count
            With .cells(1, col)
                cc = defc
                If Not show Then cc = .Interior.Color
                pos = InStr(.Text, "_")
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



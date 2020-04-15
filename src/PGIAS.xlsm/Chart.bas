Attribute VB_Name = "Chart"

'   ���W
Private Type xy
    x As Double
    y As Double
End Type

'   ��`
Private Type RectEdge
    top As Double
    bottom As Double
    left As Double
    right As Double
End Type

'   �v���b�g
Private Type Plot
    label As RectEdge
    point As xy
End Type

'   makeIndivMap�̃t���O
Public Enum FMAP_FLAG
    FMAP_FIRST = 1
    FMAP_LAST = 2
    FMAP_ALL = 3
End Enum

Const LabelSize As Long = 8
Const DefMarkerSize As Long = 5
Const msgSetMarkerLabel As String = "�}�[�J�[�ƃ��x���̐ݒ�"
Const msgSetLabelsHorizontally As String = "���x���ʒu�̒����i�������j"
Const msgSetLabelsVertically As String = "���x���ʒu�̒����i�c�����j"
Const msgSetAlignment As String = "������"

'   �}�b�v�ݒ�̕ύX
Public Function onChangeMapSettings(ByVal Target As Range, _
                ByVal rng As Variant) As Boolean
    Dim key As String
    onChangeMapSettings = False
    If Not IsObject(rng) Then Set rng = Range(rng)
    key = Target.Offset(-1, 0).text
    If key = C_AutoTarget Then
        Call setSettings(rng, getAutoTargetSettings(Target.text))
        onChangeMapSettings = True
    End If
End Function

'   �����x���̐ݒ�
Public Sub setAxisLabel(ByRef obj As ChartObject, ByVal xLabel As String, _
                ByVal yLabel As String)
    With obj.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.text = removeSuffix(xLabel)
    End With
    With obj.Chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.text = removeSuffix(yLabel)
    End With
End Sub

'   �T�t�B�b�N�X�̍폜
Private Function removeSuffix(ByVal label As String)
    Dim pos As Integer
    pos = InStr(label, "_")
    If pos > 0 Then
        removeSuffix = left(label, pos - 1)
    Else
        removeSuffix = label
    End If
End Function

'   �͈͂��^�C�v�����̎擾
Public Function getTypesFromRange(ByVal rng As Variant) As Variant
    Dim stype() As String
    Dim i As Integer
    Dim cel As Variant
    
    If Not IsObject(rng) Then Set rng = Range(rng)
    ReDim stype(rng.count - 1)
    For Each cel In rng
        If cel.text <> "" And cel.text <> "�@" Then
            stype(i) = cel.text
            i = i + 1
        End If
    Next
    If i > 0 Then
        ReDim Preserve stype(i - 1)
        getTypesFromRange = stype
    Else
        getTypesFromRange = Array("")
    End If
End Function

'   �^�C�v�I���E�푰�I���̃Z���̕ύX
Sub CheckEmpasis(ByVal Target As Range, ByRef cho As ChartObject, _
                ByVal typeCell As String, ByVal empasisCell As String, _
                Optional showArrow As Boolean = True)
    Dim stype(), emph, species As String
    Dim i As Integer
    Dim cel, rng As Range
    '   �ύX�Z�����A�^�C�v�E�����E�푰�I���Z���łȂ�������I��
    Set rng = Union(Range(typeCell), Range(empasisCell))
    If Application.Intersect(Target, rng) Is Nothing Then Exit Sub
    '   �^�C�v��z��œ���
    i = 0
    If typeCell <> "" Then
        ReDim stype(Range(typeCell).count - 1)
        For Each cel In Range(typeCell)
            If cel.text <> "" And cel.text <> "�@" Then
                stype(i) = cel.text
                i = i + 1
            End If
        Next
        If i > 0 Then
            ReDim Preserve stype(i - 1)
        Else
            stype(0) = ""
        End If
    End If
    Call setMarker(cho, stype, emph, showArrow)
End Sub


'   ���f�[�^�����ׂăN���A
Public Sub DeleteSourceData(ByRef cho As ChartObject)
    With cho.Chart
        While .SeriesCollection.count
            .SeriesCollection(1).Delete
        Wend
    End With
End Sub

'   ���f�[�^���Z�b�g����
Public Sub SetSourceData(ByRef cho As ChartObject, src As Range)
    Dim r As Range
    Dim arr As Variant
    With cho.Chart
        While .SeriesCollection.count
            .SeriesCollection(1).Delete
        Wend
        .ChartType = xlXYScatter
        .SetSourceData src
    End With
End Sub

'   �F���������i�푰���������j����B�ǂ��炩0�Ȃ画��̕K�v���Ȃ�
Private Function isMatchColor(ByRef clrs1 As Variant, ByRef clrs2 As Variant) As Boolean
    Dim lim1, lim2, i, j As Integer
    isMatchColor = False
    If clrs1(0) = 0 Or clrs2(0) = 0 Then
        isMatchColor = True
        Exit Function
    End If
    For i = 0 To UBound(clrs1)
        For j = 0 To UBound(clrs2)
            If clrs1(i) = clrs2(j) Then
                isMatchColor = True
                Exit Function
            End If
        Next
    Next
End Function

'   �}�[�J�[��ݒ肷��
Public Sub setMarker(ByRef obj As ChartObject, _
                Optional ByVal stype As Variant = "", _
                Optional ByVal emphasis As String = "", _
                Optional ByVal showArrow As Boolean = True)
    Dim i, j, flag, posc(1), row, ct() As Long
    Dim sname As String
    Dim isArrowHead, isEmphasis As Boolean
    Dim prevPoint As point
    Dim part As Variant
    Dim ctype() As Long
    
    If obj.Chart.SeriesCollection.count = 0 Then Exit Sub
    ReDim ctype(0)
    ctype(0) = 0
    If IsArray(stype) Then
        If stype(0) <> "" Then
            ReDim ctype(UBound(stype))
            For i = 0 To UBound(stype)
                ctype(i) = getTypeColor(stype(i))
            Next
        End If
    End If
    With obj.Chart.SeriesCollection(1)
        Call dspProgress("", .points.count)
        For i = 1 To .points.count
            Call dspProgress
            With .points(i)
                '   ���̐�ł͂Ȃ�
                If msoTrue <> .Format.line.visible Then
                    isArrowHead = False
                    posc(0) = .MarkerForegroundColor
                    posc(1) = .MarkerBackgroundColor
                Else    '   ���̐�
                    isArrowHead = True
                    Set prevPoint = .Parent.points(i - 1)
                    posc(0) = prevPoint.MarkerForegroundColor
                    posc(1) = prevPoint.MarkerBackgroundColor
                End If
                sname = .DataLabel.text
                part = Split(sname, " ")
                isEmphasis = (part(0) = emphasis)
                '   �������x���łȂ��A�^�C�v�w�肪���肻��ƐF����v����Ƃ�
                '   �܂��́A���i�\���j��\���Ŗ��̐�̃v���b�g�̎��́c
                If ((Not isEmphasis) And Not isMatchColor(ctype, posc)) _
                        Or (showArrow = False And isArrowHead) Then
                    '   �v���b�g������
                    .DataLabel.height = 0
                    .MarkerStyle = xlMarkerStyleNone
                    If isArrowHead Then
                        .Format.line.Transparency = 1
                    End If
                Else
                    '   �v���b�g��\������
                    .DataLabel.height = LabelSize
                    .MarkerStyle = xlMarkerStyleCircle
                    If isArrowHead Then
                        .Format.line.Transparency = 0.5
                        .Parent.points(i - 1).DataLabel.height = 0
                    End If
                    If isEmphasis Then
                        .MarkerSize = Fix(DefMarkerSize * 2)
                    Else
                        .MarkerSize = DefMarkerSize
                    End If
                End If
            End With
        Next i
    End With
   Call dspProgress("", 0)
End Sub

Sub setMarkerLabels(ByRef obj As ChartObject, ByVal celLabel As Range, _
                    Optional ByVal celSpecies As Range = Nothing, _
                    Optional ByVal celArrow As Range = Nothing, _
                    Optional ByVal alignSteps As Long = 2)
    Dim sro As Series
    Dim cnt As Long
    Dim plots() As Plot
    
    Set sro = obj.Chart.SeriesCollection(1)
    cnt = sro.points.count
    ReDim plots(cnt)
    cnt = cnt * (2 + alignSteps * 2)
    Call dspProgress(msgSetMarkerLabel, cnt)
    Call setMarkerLabelShape(obj, celLabel, plots, celSpecies, celArrow)
    Call alignLabels(plots, alignSteps)
    Call refrectLabels(sro, plots)
    Call dspProgress("", 0)
End Sub


'   �}�[�J�[�̃��x���ƐF��ݒ肷��
Private Sub setMarkerLabelShape(ByRef obj As ChartObject, _
                    ByVal celLabel As Range, _
                    ByRef plots() As Plot, _
                    Optional ByVal celSpecies As Range = Nothing, _
                    Optional ByVal celArrow As Range = Nothing)
    Dim prog, row As Long
    Dim lstr As String
    Dim cc As Variant
    Dim lh, lw As Double
    Dim x, y As Variant
    Dim avobe As Boolean
    
    If celSpecies Is Nothing Then Set celSpecies = celLabel
    With obj.Chart.SeriesCollection(1)
        .ClearFormats
        .HasDataLabels = True
        .HasLeaderLines = True
        y = .Values
        x = .XValues
        With .LeaderLines.Format.line
            .visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Transparency = 0.7
        End With
        With .DataLabels.Format.TextFrame2.TextRange
            .ParagraphFormat.Alignment = msoAlignLeft
            .Font.Size = LabelSize
            .Font.Fill.visible = msoTrue
            .Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
            .Font.Fill.ForeColor.Brightness = 0.4
        End With
        For i = 1 To .points.count
            Call dspProgress
            Set celLabel = celLabel.Offset(1, 0)
            Set celSpecies = celSpecies.Offset(1, 0)
            With .points(i)
                .HasDataLabel = True
                lstr = celLabel.text
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = DefMarkerSize
                lh = LabelSize
                lw = Len(lstr) * LabelSize
                plots(i).point.x = .left
                plots(i).point.y = .top
                plots(i).label.left = .left + DefMarkerSize / 2
                plots(i).label.right = plots(i).label.left + lw
                plots(i).label.top = .top - lh / 2
                plots(i).label.bottom = .top + lh / 2
                With .DataLabel
                    With .Format.TextFrame2
                       .TextRange.Font.Size = LabelSize
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .MarginBottom = 0
                    End With
                    .text = lstr
                    .height = lh
                    .width = lw
'                    .top = .Parent.top - LabelSize / 2
                End With
                cc = getColorForSpecies(celSpecies.text)
                With .Format.line
                    .Weight = 2
                    .visible = msoFalse ' msoTrue�ɂ���Ɛ��܂ŕ\�������
                End With
                With .Format.Fill
                    .visible = msoTrue
                    .ForeColor.RGB = cc(0)
                    .BackColor.RGB = cc(1)
                End With
                .MarkerForegroundColor = cc(0)
                .MarkerBackgroundColor = cc(1)
                '   ����̎w�肪����ꍇ
                If Not celArrow Is Nothing Then
                    Set celArrow = celArrow.Offset(1, 0)
                    If celArrow.value Then
                        '   �i���݂̃|�C���g����̐�Ƃ��āj���̕`��
                        With .Format.line
                            .visible = msoTrue
                            .ForeColor.RGB = cc(0)
                            .Transparency = 0.5
                            .EndArrowheadStyle = msoArrowheadTriangle
                            .EndArrowheadLength = msoArrowheadLong
                        End With
                        With .Format.Fill
                            .ForeColor.RGB = cc(1)
                            .Transparency = 0.5
                        End With
                        '   �O�̃v���b�g�̃��x���͏���
                        .Parent.points(i - 1).DataLabel.height = 0
                    End If
                End If
            End With
        Next i
    End With
End Sub

'   �v���b�g�̍��E�Ƀ��x���̔z�u
Private Sub setLabelPos(ByRef dl As DataLabel, Optional ByVal isLeft As Boolean = False)
    With dl
        If isLeft Then
            .left = .Parent.left - .width - LabelSize / 2
            .HorizontalAlignment = xlRight
        Else
            .left = .Parent.left + LabelSize / 2
            .HorizontalAlignment = xlLeft
        End If
    End With
End Sub

Private Function newRectEdge(ByRef obj As Object) As RectEdge
    With obj
        newRectEdge.top = .top
        newRectEdge.left = .left
        newRectEdge.bottom = .top + .height
        newRectEdge.right = .left + .width
    End With
End Function

Private Sub setRectEdge(ByRef sbj As RectEdge, _
                        Optional ByVal left As Double = 0, _
                        Optional ByVal right As Double = 0, _
                        Optional ByVal top As Double = 0, _
                        Optional ByVal bottom As Double = 0)
    With sbj
        .top = top
        .left = left
        .bottom = bottom
        .right = right
    End With
End Sub

'   ���x���ʒu�̒���
Private Function alignLabels(ByRef plots() As Plot, _
                    Optional ByVal steps As Long = 2) As Double
    Dim idx4h As Variant
    Dim idx4v As Variant
    Dim margin As xy
    Dim center As xy
    Dim considerPlot As Boolean
    
    If steps <= 0 Then Exit Function
    If steps > 100 Then steps = 100
    center = getCenter(plots)
    Call flipLabels(plots, center)
    idx4h = getOrderByXY(plots, center, 0)  '   ���S���獶�E��
    idx4v = getOrderByXY(plots, center, 1)  '   ���S����㉺��
    margin.x = DefMarkerSize / 2
    margin.y = LabelSize * 1.5
    Call dspProgress(msgSetLabelsHorizontally)
    Call alignLabelsHorizontally(idx4h, plots, margin, False)
    Call dspProgress(msgSetLabelsVertically)
    Call alignLabelsVertically(idx4h, plots, margin, False)
    steps = steps - 1
    margin.x = LabelSize * 2
    margin.y = LabelSize * 2
    considerPlot = False
    While steps > 0
        If steps = 1 Then considerPlot = Not considerPlot
        Call dspProgress(msgSetLabelsHorizontally)
        Call alignLabelsHorizontally(idx4h, plots, margin, considerPlot)
        Call dspProgress(msgSetLabelsVertically)
        Call alignLabelsVertically(idx4h, plots, margin, considerPlot)
        Call dspProgress(msgSetAlignment)
        steps = steps - 1
        margin.x = margin.x + LabelSize * 2
        margin.y = margin.y + LabelSize
    Wend
End Function

Private Sub flipLabels(ByRef plots() As Plot, ByRef center As xy)
    Dim i As Long
    Dim width As Double
    For i = 1 To UBound(plots)
        With plots(i)
            If .point.x < center.x Then
                width = .label.right - .label.left
                .label.left = .point.x - width - DefMarkerSize / 2
                .label.right = .label.left + width
            End If
        End With
    Next
End Sub

'   ���������̒���
Function alignLabelsHorizontally(ByRef idx As Variant, ByRef plots() As Plot, _
                                ByRef margin As xy, _
                                Optional ByVal expand As Boolean = True) As Double
    Dim olv As RectEdge
    Dim space As RectEdge
    Dim free As RectEdge
    Dim pos As xy
    Dim prog, i, loops As Long
    Dim shift, width As Double
    
    loops = UBound(idx)
    For i = 1 To loops
        Call dspProgress
        pos = plots(idx(i)).point
        Call getOverlap(idx(i), plots, olv, space, free, margin)
        With plots(idx(i)).label
            width = .right - .left
            shift = 0
            '   ���E�Ƃ��d�Ȃ肪��������A���ԂɈړ�
            If olv.left > 0 And olv.right > 0 Then
                shift = (olv.left - olv.right) / 2
                If .left + shift < free.left Then shift = free.left - .left
                If .right + shift > free.right Then shift = free.right - .right
            '   �������d�Ȃ��Ă���̂ŉE�Ɉړ�
            ElseIf olv.left > 0 Then
                shift = olv.left
                ajs = pos.x + DefMarkerSize / 2 - .left '�E�����̂��傤�Ǘǂ��ʒu
                If shift < ajs And space.right >= ajs Then
                    shift = ajs
                ElseIf shift > space.right Then
'                    If expand Then shift = space.right Else shift = 0
                    If expand Then shift = (shift + space.right) / 2 Else shift = 0
                End If
            '   �E�����d�Ȃ��Ă���̂ō��Ɉړ�
            ElseIf olv.right > 0 Then
                shift = olv.right
                ajs = .left - (pos.x - DefMarkerSize / 2 - width) '�������̂��傤�Ǘǂ��ʒu
                If shift < ajs And space.left >= ajs Then
                    shift = ajs
                ElseIf shift > space.left Then
'                    If expand Then shift = space.left Else shift = 0
                    If expand Then shift = (shift + space.left) / 2 Else shift = 0
                End If
                shift = -shift
            End If
            .left = .left + shift
            .right = .right + shift
        End With
    Next
End Function

'   ���������̒���
Function alignLabelsVertically(ByRef idx As Variant, ByRef plots() As Plot, _
                                ByRef margin As xy, _
                                Optional ByVal considerPlot As Boolean = False) As Double
    Dim olv As RectEdge
    Dim space As RectEdge
    Dim free As RectEdge
    Dim pos As xy
    Dim prog, loops, i As Long
    Dim shift As Double
    loops = UBound(idx)
    For i = 1 To loops
        Call dspProgress
        pos = plots(idx(i)).point
        Call getOverlap(idx(i), plots, olv, space, free, margin, considerPlot)
        With plots(idx(i)).label
            shift = 0
            '   �㉺�Ƃ��d�Ȃ肪��������A���ԂɈړ�
            If olv.top > 0 And olv.bottom > 0 Then
                shift = (olv.top - olv.bottom) / 2
                If .top + shift < free.top Then shift = free.top - .top
                If .bottom + shift > free.bottom Then shift = free.bottom - .bottom
            '   �ゾ���d�Ȃ��Ă���̂ŉ��Ɉړ�
            ElseIf olv.top > 0 Then
                shift = olv.top
'                If shift > space.bottom Then shift = space.bottom
                If shift > space.bottom Then shift = (shift + space.bottom) / 2
            '   �������d�Ȃ��Ă���̂ŏ�Ɉړ�
            ElseIf olv.bottom > 0 Then
                shift = olv.bottom
'                If shift > space.top Then shift = space.top
                If shift > space.top Then shift = (shift + space.top) / 2
                shift = -shift
            End If
            .top = .top + shift
            .bottom = .bottom + shift
        End With
    Next
End Function

'   �񂹂̐ݒ�
Private Sub refrectLabels(ByRef sro As Series, plots() As Plot)
    Dim i As Long
    Dim pos As point
    Dim dl As DataLabel
    Dim width As Double
    
    For i = 1 To sro.points.count
        Call dspProgress
        Set dl = sro.points(i).DataLabel
        With plots(i)
            width = .label.right - .label.left
            dl.width = width
'            dl.height = .label.bottom - .label.top
            dl.left = .label.left
            dl.top = .label.top
            If .label.left + width * 2 / 3 <= .point.x Then
                dl.HorizontalAlignment = xlRight
            ElseIf .label.left + width / 3 >= .point.x Then
                dl.HorizontalAlignment = xlLeft
            Else
                dl.HorizontalAlignment = xlCenter
            End If
        End With
    Next
End Sub

'   �d�Ȃ�̗ʂƁA���Ԃ̗ʂ𓾂�
Private Function getOverlap(ByVal idx As Long, ByRef plots() As Plot, _
                        ByRef olv As RectEdge, ByRef space As RectEdge, _
                        ByRef free As RectEdge, ByRef margin As xy, _
                        Optional ByVal considerPlot As Boolean = False)
    Dim dl As RectEdge
    Dim ref As RectEdge
    Dim mark As Double
    
    '   Clear olv and space
    Call setRectEdge(olv)
    Call initFreeSpace(plots(idx), free, space, margin)
    For i = 1 To UBound(plots)
        If i <> idx Then
            Call updateOverlap( _
                    plots(idx).label, plots(i).label, free, olv, space)
        End If
        If considerPlot Then
            With plots(i).point
                mark = DefMarkerSize / 2
                Call setRectEdge(ref, _
                    .x - mark / 2, .x + mark / 2, _
                    .y - mark, .y + mark)
            End With
            Call updateOverlap(plots(idx).label, ref, free, olv, space)
        End If
    Next
End Function

'   ���x���̉���Ɖ��\�����̏�����
Private Sub initFreeSpace(ByRef plt As Plot, _
                    ByRef free As RectEdge, ByRef space As RectEdge, _
                    ByRef margin As xy)
    Dim width As Double
    With plt
        width = .label.right - .label.left
        free.left = .point.x - width - margin.x
        free.right = .point.x + width + margin.x
        free.top = .point.y - margin.y
        free.bottom = .point.y + margin.y
        space.left = .label.left - free.left
        space.right = free.right - .label.right
        space.top = .label.top - free.top
        space.bottom = free.bottom - .label.bottom
        If space.left < 0 Then space.left = 0
        If space.right < 0 Then space.right = 0
        If space.top < 0 Then space.top = 0
        If space.bottom < 0 Then space.bottom = 0
    End With
End Sub

'   �d�Ȃ�̗ʂƉ��\�����̍X�V
Private Sub updateOverlap(ByRef dl As RectEdge, ByRef ref As RectEdge, _
                        ByRef free As RectEdge, _
                        ByRef olv As RectEdge, ByRef space As RectEdge)
    Dim sp, ho, vo As Double
    Dim hOver, vOver, refOnLeft, refOnTop As Boolean
    
    hOver = dl.right > ref.left And ref.right > dl.left
    vOver = dl.bottom > ref.top And ref.bottom > dl.top
    
    '   �d�Ȃ���
    If hOver And vOver Then
        '   �d�Ȃ����������v�Z���āAovl�ɂ��̏㉺���E�����̏d�Ȃ�̍ő�l���X�g�A����
        '   ��������
        If ref.left < dl.left And ref.right < dl.right Then
            refOnLeft = True
        ElseIf dl.right < ref.right And dl.left < ref.left Then
            refOnLeft = False
        Else
            refOnLeft = (ref.left + ref.right) < (free.left + free.right)
        End If
        If refOnLeft Then
            ho = ref.right - dl.left
            If olv.left < ho Then olv.left = ho: space.left = 0
        Else
            ho = dl.right - ref.left
            If olv.right < ho Then olv.right = ho: space.right = 0
        End If
        '   ��������
        If ref.top < dl.top And ref.bottom < dl.bottom Then
            refOnTop = True
        ElseIf dl.bottom < ref.bottom And dl.top < ref.top Then
            refOnTop = False
        Else
            refOnTop = (ref.top + ref.bottom) < (free.top + free.bottom)
        End If
        If refOnTop Then
            ho = ref.bottom - dl.top
            If olv.top < ho Then olv.top = ho: space.top = 0
        Else
            ho = dl.bottom - ref.top
            If olv.bottom < ho Then olv.bottom = ho: space.bottom = 0
        End If
    Else
        '   X�̂ݏd�Ȃ��Ă���΁A�㉺�̌��Ԃ̍ŏ��l���X�g�A
        If hOver Then
            If ref.top < dl.top Then
                sp = dl.top - ref.bottom
                If space.top > sp Then space.top = sp
            Else
                sp = ref.top - dl.bottom
                If space.bottom > sp Then space.bottom = sp
            End If
        End If
        '   Y�̂ݏd�Ȃ��Ă���΁A���E�̌��Ԃ̍ŏ��l���X�g�A
        If vOver Then
            If ref.left < dl.left Then
                sp = dl.left - ref.right
                If space.left > sp Then space.left = sp
            Else
                sp = ref.left - dl.right
                If space.right > sp Then space.right = sp
            End If
        End If
    End If
End Sub

'   ���ς𓾂�
Private Function getCenter(ByRef plots() As Plot) As xy
    Dim c As xy
    Dim num As Long
    num = UBound(plots)
    For i = 1 To num
        With plots(i).point
            c.x = c.x + .x
            c.y = c.y + .y
        End With
    Next i
    c.x = c.x / num
    c.y = c.y / num
    getCenter = c
End Function

'   X�܂���Y�̒��S����̋����̏���
Private Function getOrderByXY(ByRef plots() As Plot, ByRef center As xy, _
                Optional flag As Long = 0) As Variant
    Dim i, num As Long
    Dim vals() As Double
    num = UBound(plots)
    ReDim vals(num)
    For i = 1 To num
        With plots(i).point
            If flag = 0 Then
                vals(i) = Abs(.x - center.x)
            ElseIf flag = 1 Then
                vals(i) = Abs(.y - center.y)
            End If
        End With
    Next i
    getOrderByXY = getSortedIndex(vals)
End Function

'   �\�[�g
Private Function getSortedIndex(ByVal vals As Variant) As Variant
    Dim idx(), num, i, j, tmp As Long
    num = UBound(vals)
    ReDim idx(num)
    For i = 0 To num
        idx(i) = i
    Next
    For i = 1 To num - 1
        For j = num - 1 To i Step -1
            If j = 0 Then
                j = j
            End If
            If vals(idx(j)) > vals(idx(j + 1)) Then
                tmp = idx(j)
                idx(j) = idx(j + 1)
                idx(j + 1) = tmp
            End If
        Next
    Next
    getSortedIndex = idx
End Function

Public Sub setSameTypeToMap(ByVal species As String, ByVal rng As Range)
    Dim stype As Variant
    stype = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
    enableEvent False
    With rng
        .cells(1, 1).value = stype(0)
        .cells(1, 2).value = stype(1)
    End With
    enableEvent True
End Sub


'   ��C���f�b�N�X�̔z����擾����
'   (0)�͈����̗�C���f�b�N�X�A(1)��X���̌��݂Ɨ\���A(2)��Y���̌��݂Ɨ\��
Private Function getAxisColIndex(ByRef settings As Object, _
                                ByVal paramCol As Variant)
    Dim colSet, colRet, colNames As Variant
    Dim colName As String
    Dim xy, np, i As Integer
    Dim lo As ListObject
    Dim colIdx() As Long
    
    Set lo = shIndividual.ListObjects(1)
    colSet = Array( _
        Array(settings(C_XAxis), settings(C_XPrediction)), _
        Array(settings(C_YAxis), settings(C_YPrediction)) _
    )
    colRet = Array(Array(), Array(Array(), Array()), Array(Array(), Array()))
    ReDim colIdx(UBound(paramCol))
    For i = 0 To UBound(paramCol)
        colIdx(i) = getColumnIndex(paramCol(i), lo)
    Next
    
    colRet(0) = colIdx
    For xy = 1 To 2
        For np = 0 To 1
            colName = colSet(xy - 1)(np)
            If InStr(colName, "*") < 1 Then
                colRet(xy)(np) = getColumnIndex(colName, lo)
            Else
                colRet(xy)(np) = Array( _
                    getColumnIndex(Replace(colName, "*", "1"), lo), _
                    getColumnIndex(Replace(colName, "*", "2"), lo) _
                )
            End If
        Next
    Next
    getAxisColIndex = colRet
End Function


'   �̃}�b�v�����
Public Sub makeIndivMap(ByRef rngTbl As Range, _
                        ByRef settings As Object, _
                Optional ByVal sequence As Integer = FMAP_ALL)
    Dim cho As ChartObject
    Dim league As Integer
    Dim stime As Double
    
    stime = Timer
    Call doMacro(msgstr(msgMaking, rngTbl.Parent.name))
    league = shIndividual.SetAutoTargetPL(settings(C_AutoTarget), settings(C_Level))
    If sequence And FMAP_FIRST Then
        Call shIndividual.calcAllIndividualTable(F_FORCEALL)
    Else
        Call shIndividual.calcAllIndividualTable(F_PREDICTION)
    End If
    Call makeOrgTable(rngTbl, settings)
    Set cho = rngTbl.Parent.ChartObjects(1)
    With rngTbl
        Call SetSourceData(cho, Range( _
            .cells(1, 3), .cells(.rows.count, 4)))
    End With
    With rngTbl.Offset(-1, 0)
        Call setMarkerLabels(cho, _
            .cells(1, 1), .cells(1, 2), .cells(1, 5), settings(C_LabelAlign))
    End With
    Call setAxisLabel(cho, settings(C_XAxis), settings(C_YAxis))
    If (settings(C_AutoTarget) <> "" And settings(C_AutoTarget) <> C_None) _
            Or (sequence And FMAP_LAST) Then
        Call shIndividual.SetAutoTargetPL
        Call shIndividual.calcAllIndividualTable(F_PREDICTION)
    End If
    Call setTimeAndDate(rngTbl.Parent.Range(IMAP_R_MakingTime), stime)
    Call doMacro
End Sub

'   ���\�����
Public Sub makeOrgTable(ByRef rng As Range, _
                        ByRef settings As Object)
    Dim srow, i As Long
    Dim celMap1, celMap As Range
    Dim limCP, colIdx, val, mval As Variant
'    Dim addr, maddr As Variant
    Dim xy, np As Integer
    
    rng.value = ""
    Set celMap1 = rng.cells(1, 1).Offset(-1, 0)
    colIdx = getAxisColIndex(settings, _
            Array(IND_Nickname, IND_Species, IND_PL, IND_prPL, IND_CP, IND_prCP))
    '   CP����l
    limCP = Array(settings(C_CpUpper), settings(C_PrCpLower))
    If limCP(0) = 0 Then limCP(0) = C_MaxLong
    If limCP(1) = 0 Then limCP(1) = 0
    With shIndividual
        srow = .ListObjects(1).DataBodyRange.row
        Set celMap = celMap1
        While .cells(srow, 1).text <> ""
            '   ��C���f�b�N�X�ɂ�����l�ƃA�h���X�̎擾
'            Call getValAndAddr(srow, colIdx, val, addr)
            val = getValueRecursive(srow, colIdx)
            If IsError(val(0)(5)) Then val(0)(5) = 0
            If val(0)(0) = "" Or val(0)(2) = 0 _
                Or val(0)(4) > limCP(0) _
                Or (val(0)(5) > 0 And val(0)(5) < limCP(1)) Then
                GoTo Continue
            End If
            '   XY���A���݁E�\���̉��̒l���z��Ȃ�A�ő�l�擾
            For xy = 1 To 2
                For np = 0 To 1
                    If IsArray(val(xy)(np)) Then
                        mval = 0: maddr = ""
                        For i = 0 To UBound(val(xy)(np))
                            If i = 0 Or (val(xy)(np)(i) <> "" And mval < val(xy)(np)(i)) Then
                                mval = val(xy)(np)(i)
 '                               maddr = addr(xy)(np)(i)
                            End If
                        Next
                        val(xy)(np) = mval
'                        addr(xy)(np) = maddr
                    End If
                Next
            Next
            '   ���ݒl�̏�������
            Set celMap = celMap.Offset(1, 0)
            With celMap
                .value = val(0)(0) & " l." & val(0)(2)
                .Offset(0, 1).value = val(0)(1)
                .Offset(0, 2).value = val(1)(0)
                .Offset(0, 3).value = val(2)(0)
                .Offset(0, 4).value = 0
            End With
            '   ����Η\���l�̏�������
            If val(1)(1) <> "" And val(2)(1) <> "" And val(0)(5) <= limCP(0) Then
                Set celMap = celMap.Offset(1, 0)
                With celMap
                    .value = val(0)(0) & " l." & val(0)(3)
                    .Offset(0, 1).value = val(0)(1)
                    .Offset(0, 2).value = val(1)(1)
                    .Offset(0, 3).value = val(2)(1)
                    .Offset(0, 4).value = 1
                End With
            End If
Continue:
            srow = srow + 1
        Wend
    End With
    Set rng = Range(celMap1.Offset(1, 0), celMap.Offset(0, 4))
End Sub


'   �ē����ăc���[�\���̗�C���f�b�N�X���l�ƃA�h���X���擾
Private Sub getValAndAddr(ByVal row As Long, ByRef colIdx As Variant, _
            ByRef val As Variant, ByRef addr As Variant)
    Dim arrVal() As Variant
    Dim arrAddr() As Variant
    Dim lim As Long
    If IsArray(colIdx) Then
        lim = UBound(colIdx)
        ReDim arrVal(lim), arrAddr(lim)
        For i = 0 To lim
            Call getValAndAddr(row, colIdx(i), arrVal(i), arrAddr(i))
        Next
        val = arrVal: addr = arrAddr
    Else
        With shIndividual.cells(row, colIdx)
            val = .value
            addr = "=" & shIndividual.name & "!" & Replace(.Address, shName, "")
        End With
    End If
End Sub

'   �ē����ăc���[�\���̗�C���f�b�N�X���l���擾
Private Function getValueRecursive(ByVal row As Long, ByRef colIdx As Variant) As Variant
    Dim arrVal() As Variant
    Dim lim As Long
    If IsArray(colIdx) Then
        lim = UBound(colIdx)
        ReDim arrVal(lim)
        For i = 0 To lim
            arrVal(i) = getValueRecursive(row, colIdx(i))
        Next
        getValueRecursive = arrVal
    Else
        getValueRecursive = shIndividual.cells(row, colIdx).value
    End If
End Function

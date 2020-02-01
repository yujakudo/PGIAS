Attribute VB_Name = "Chart"
Private Type xy
    x As Double
    y As Double
End Type

Private Type RectEdge
    top As Double
    bottom As Double
    left As Double
    right As Double
End Type

Private Type Plot
    label As RectEdge
    point As xy
End Type


Const LabelSize As Long = 8
Const DefMarkerSize As Long = 5
Const msgSetMarkerLabel As String = "マーカーとラベルの設定"
Const msgSetLabelsHorizontally As String = "ラベル位置の調整（横方向）"
Const msgSetLabelsVertically As String = "ラベル位置の調整（縦方向）"
Const msgSetAlignment As String = "文字寄せ"

Sub setAxisLabel(ByRef obj As ChartObject, ByVal xLabel As String, _
                ByVal yLabel As String)
    With obj.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = removeSuffix(xLabel)
    End With
    With obj.Chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = removeSuffix(yLabel)
    End With
End Sub

Private Function removeSuffix(ByVal label As String)
    Dim pos As Integer
    pos = InStr(label, "_")
    If pos > 0 Then
        removeSuffix = left(label, pos - 1)
    Else
        removeSuffix = label
    End If
End Function

'   タイプ選択・種族選択のセルの変更
Sub CheckEmpasis(ByVal Target As Range, ByRef cho As ChartObject, _
                ByVal typeCell As String, ByVal empasisCell As String, _
                Optional showArrow As Boolean = True)
    Dim stype(), emph, species As String
    Dim i As Integer
    Dim cel, rng As Range
    '   変更セルが、タイプ・強調・種族選択セルでなかったら終了
    Set rng = Union(Range(typeCell), Range(empasisCell))
    If Application.Intersect(Target, rng) Is Nothing Then Exit Sub
    '   タイプを配列で得る
    i = 0
    If typeCell <> "" Then
        ReDim stype(Range(typeCell).count - 1)
        For Each cel In Range(typeCell)
            If cel.Text <> "" And cel.Text <> "　" Then
                stype(i) = cel.Text
                i = i + 1
            End If
        Next
        If i > 0 Then
            ReDim Preserve stype(i - 1)
        Else
            stype(0) = ""
        End If
    End If
    emph = Range(empasisCell).Text
    Call setMarker(cho, stype, emph, showArrow)
End Sub


'   元データをすべてクリア
Public Sub DeleteSourceData(ByRef cho As ChartObject)
    With cho.Chart
        While .SeriesCollection.count
            .SeriesCollection(1).Delete
        Wend
    End With
End Sub

'   元データをセットする
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

'   色が同じか（種族が同じか）判定。どちらか0なら判定の必要がない
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

'   マーカーを設定する
Sub setMarker(ByRef obj As ChartObject, _
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
                '   矢印の先ではない
                If msoTrue <> .Format.line.Visible Then
                    isArrowHead = False
                    posc(0) = .MarkerForegroundColor
                    posc(1) = .MarkerBackgroundColor
                Else    '   矢印の先
                    isArrowHead = True
                    Set prevPoint = .Parent.points(i - 1)
                    posc(0) = prevPoint.MarkerForegroundColor
                    posc(1) = prevPoint.MarkerBackgroundColor
                End If
                sname = .DataLabel.Text
                part = Split(sname, " ")
                isEmphasis = (part(0) = emphasis)
                '   強調ラベルでなく、タイプ指定がありそれと色が一致するとき
                '   または、矢印（予測）非表示で矢印の先のプロットの時は…
                If ((Not isEmphasis) And Not isMatchColor(ctype, posc)) _
                        Or (showArrow = False And isArrowHead) Then
                    '   プロットを消す
                    .DataLabel.height = 0
                    .MarkerStyle = xlMarkerStyleNone
                    If isArrowHead Then
                        .Format.line.Transparency = 1
                    End If
                Else
                    '   プロットを表示する
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


'   マーカーのラベルと色を設定する
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
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Transparency = 0.7
        End With
        With .DataLabels.Format.TextFrame2.TextRange
            .ParagraphFormat.Alignment = msoAlignLeft
            .Font.Size = LabelSize
            .Font.Fill.Visible = msoTrue
            .Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
            .Font.Fill.ForeColor.Brightness = 0.4
        End With
        For i = 1 To .points.count
            Call dspProgress
            Set celLabel = celLabel.Offset(1, 0)
            Set celSpecies = celSpecies.Offset(1, 0)
            With .points(i)
                .HasDataLabel = True
                lstr = celLabel.Text
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
                    .Text = lstr
                    .height = lh
                    .width = lw
'                    .top = .Parent.top - LabelSize / 2
                End With
                cc = getColorForSpecies(celSpecies.Text)
                With .Format.line
                    .Weight = 2
                    .Visible = msoFalse ' msoTrueにすると線まで表示される
                End With
                With .Format.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = cc(0)
                    .BackColor.RGB = cc(1)
                End With
                .MarkerForegroundColor = cc(0)
                .MarkerBackgroundColor = cc(1)
                '   矢印列の指定がある場合
                If Not celArrow Is Nothing Then
                    Set celArrow = celArrow.Offset(1, 0)
                    If celArrow.Value Then
                        '   （現在のポイントを矢印の先として）矢印の描画
                        With .Format.line
                            .Visible = msoTrue
                            .ForeColor.RGB = cc(0)
                            .Transparency = 0.5
                            .EndArrowheadStyle = msoArrowheadTriangle
                            .EndArrowheadLength = msoArrowheadLong
                        End With
                        With .Format.Fill
                            .ForeColor.RGB = cc(1)
                            .Transparency = 0.5
                        End With
                        '   前のプロットのラベルは消す
                        .Parent.points(i - 1).DataLabel.height = 0
                    End If
                End If
            End With
        Next i
    End With
End Sub

'   プロットの左右にラベルの配置
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

'   ラベル位置の調整
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
    idx4h = getOrderByXY(plots, center, 0)  '   中心から左右へ
    idx4v = getOrderByXY(plots, center, 1)  '   中心から上下へ
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

'   水平方向の調整
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
            '   左右とも重なりがあったら、中間に移動
            If olv.left > 0 And olv.right > 0 Then
                shift = (olv.left - olv.right) / 2
                If .left + shift < free.left Then shift = free.left - .left
                If .right + shift > free.right Then shift = free.right - .right
            '   左だけ重なっているので右に移動
            ElseIf olv.left > 0 Then
                shift = olv.left
                ajs = pos.x + DefMarkerSize / 2 - .left '右方向のちょうど良い位置
                If shift < ajs And space.right >= ajs Then
                    shift = ajs
                ElseIf shift > space.right Then
'                    If expand Then shift = space.right Else shift = 0
                    If expand Then shift = (shift + space.right) / 2 Else shift = 0
                End If
            '   右だけ重なっているので左に移動
            ElseIf olv.right > 0 Then
                shift = olv.right
                ajs = .left - (pos.x - DefMarkerSize / 2 - width) '左方向のちょうど良い位置
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

'   垂直方向の調整
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
            '   上下とも重なりがあったら、中間に移動
            If olv.top > 0 And olv.bottom > 0 Then
                shift = (olv.top - olv.bottom) / 2
                If .top + shift < free.top Then shift = free.top - .top
                If .bottom + shift > free.bottom Then shift = free.bottom - .bottom
            '   上だけ重なっているので下に移動
            ElseIf olv.top > 0 Then
                shift = olv.top
'                If shift > space.bottom Then shift = space.bottom
                If shift > space.bottom Then shift = (shift + space.bottom) / 2
            '   下だけ重なっているので上に移動
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

'   寄せの設定
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

'   重なりの量と、隙間の量を得る
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

'   ラベルの可動域と可動可能距離の初期化
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

'   重なりの量と可動可能距離の更新
Private Sub updateOverlap(ByRef dl As RectEdge, ByRef ref As RectEdge, _
                        ByRef free As RectEdge, _
                        ByRef olv As RectEdge, ByRef space As RectEdge)
    Dim sp, ho, vo As Double
    Dim hOver, vOver, refOnLeft, refOnTop As Boolean
    
    hOver = dl.right > ref.left And ref.right > dl.left
    vOver = dl.bottom > ref.top And ref.bottom > dl.top
    
    '   重なった
    If hOver And vOver Then
        '   重なった部分を計算して、ovlにその上下左右方向の重なりの最大値をストアする
        '   水平方向
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
        '   垂直方向
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
        '   Xのみ重なっていれば、上下の隙間の最小値をストア
        If hOver Then
            If ref.top < dl.top Then
                sp = dl.top - ref.bottom
                If space.top > sp Then space.top = sp
            Else
                sp = ref.top - dl.bottom
                If space.bottom > sp Then space.bottom = sp
            End If
        End If
        '   Yのみ重なっていれば、左右の隙間の最小値をストア
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

'   平均を得る
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

'   XまたはYの中心からの距離の昇順
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

'   ソート
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
        .cells(1, 1).Value = stype(0)
        .cells(1, 2).Value = stype(1)
    End With
    enableEvent True
End Sub

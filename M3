Sub ArWell(ByVal WS As String, ByVal Row As Long, ByVal Col As Long, ByRef DynArrayWell())
    Dim LBoundRow As Long 'Нижняя граница изменяющегося массива
    Dim UBoundRow As Long 'Верхняя граница изменяющегося массива
    LBoundRow = Row 'Строка с которой начинаем массив
    UBoundRow = Module0.LastRow(WS, Col)
    ReDim DynArrayWell(LBoundRow To UBoundRow) 'перезаписываем значения массива
End Sub
Sub ArDate(ByVal WS As String, ByVal Row As Long, ByVal Col As Long, ByRef DynArrayDate())
    Dim LBoundColumn As Long 'Нижняя граница изменяющегося массива
    Dim UBoundColumn As Long 'Верхняя граница изменяющегося массива
    LBoundColumn = Col
    UBoundColumn = Module0.LastColumn(WS, Row)
    ReDim DynArrayDate(LBoundColumn To UBoundColumn) 'перезаписываем значения массива
End Sub
Sub ArSource(ByVal WSSource As String, ByVal Row As Long, ByVal Col As Long, ByRef DynArraySource())
    Dim LBoundRowSource As Long 'Нижняя граница изменяющегося массива
    Dim UBoundRowSource As Long 'Верхняя граница изменяющегося массива
    LBoundRowSource = Row
    UBoundRowSource = Module0.LastRow(WSSource, Col)
    ReDim DynArraySource(LBoundRowSource To UBoundRowSource) 'перезаписываем значения массива
End Sub
Sub ArPipe(ByVal WSPipe As String, ByVal Row As Long, ByVal Col As Long, ByRef rangeToSearchDataPipe As Range)
    Set WSP = ThisWorkbook.Worksheets(WSPipe)
    Set rangeToSearchDataPipe = WSP.Range("A2:A" & Module0.LastRow(WSPipe, Col))
End Sub
Sub ArCon(ByVal WSConnect As String, ByVal Row As Long, ByVal Col As Long, ByRef rangeToSearchDatAndObjCon As Range)
    Set WSC = ThisWorkbook.Worksheets(WSConnect)
    Set rangeToSearchDatAndObjCon = WSC.Range("A2:A" & Module0.LastRow(WSConnect, Col))
End Sub
Sub ArObj(ByVal WSObject As String, ByVal Row As Long, ByVal Col As Long, ByRef rangeToSearchObject As Range)
    Set WSO = ThisWorkbook.Worksheets(WSObject)
    Set rangeToSearchObject = WSO.Range("A2:A" & Module0.LastRow(WSObject, Col))
End Sub
Sub ArConSource(ByVal WSConSource As String, ByVal Row As Long, ByVal Col As Long, ByRef rangeToSearchDatAndSourceCon As Range)
    Set WSCS = ThisWorkbook.Worksheets(WSConSource)
    Set rangeToSearchDatAndSourceCon = WSCS.Range("A2:A" & Module0.LastRow(WSConSource, Col))
End Sub

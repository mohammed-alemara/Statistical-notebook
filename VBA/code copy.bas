Sub CopyMappedColumns_SurveyToSheet2

    Dim oDoc As Object
    Dim oSrc As Object
    Dim oDst As Object
    Dim oCursor As Object

    Dim map As Variant
    Dim i As Long
    Dim srcCol As Long, dstCol As Long
    Dim srcRow As Long, dstRow As Long
    Dim lastRow As Long

    Dim fixedCols As Variant
    Dim fixedCol As Long
    Dim fixedValue As Variant
    Dim r As Long, c As Long

    oDoc = ThisComponent
    oSrc = oDoc.Sheets.getByName("survey")
    oDst = oDoc.Sheets.getByName("sheet2")

    ' ===== مسح محتويات sheet2 =====
    oDst.getCellRangeByPosition( _
        0, 0, _
        oDst.Columns.Count - 1, _
        oDst.Rows.Count - 1 _
    ).clearContents(7)

    ' ===== جدول التعيين (أعمدة المصدر) =====
    map = Array( _
Array(2),_
Array(3),_
Array(4),_
Array(5),_
Array(6),_
Array(7),_
Array(8),_
Array(9),_
Array(10),_
Array(11),_
Array(12),_
Array(13),_
Array(14),_
Array(15),_
Array(16),_
Array(17),_
Array(18),_
Array(19),_
Array(20),_
Array(21),_
Array(22),_
Array(23),_
Array(24),_
Array(25),_
Array(26),_
Array(27),_
Array(28),_
Array(29),_
Array(30),_
Array(31),_
Array(32),_
Array(33),_
Array(34),_
Array(35),_
Array(36),_
Array(37),_
Array(38),_
Array(39),_
Array(40),_
Array(45),_
Array(46),_
Array(47),_
Array(49),_
Array(50),_
Array(51),_
Array(53),_
Array(54),_
Array(55),_
Array(57),_
Array(58),_
Array(59),_
Array(61),_
Array(62),_
Array(63),_
Array(65),_
Array(66),_
Array(67),_
Array(68),_
Array(69),_
Array(74),_
Array(75),_
Array(76),_
Array(78),_
Array(79),_
Array(80),_
Array(82),_
Array(83),_
Array(84),_
Array(86),_
Array(87),_
Array(88),_
Array(90),_
Array(91),_
Array(92),_
Array(94),_
Array(95),_
Array(96),_
Array(97)_
    )

    ' ===== تحديد آخر صف مستخدم في survey =====
    oCursor = oSrc.createCursor()
    oCursor.gotoEndOfUsedArea(True)
    lastRow = oCursor.RangeAddress.EndRow

    ' ===== النسخ =====
    For i = 0 To UBound(map)

        srcCol = map(i)(0)
        dstCol = i

        srcRow = 1          ' الصف الثاني في survey
        dstRow = 3          ' الصف الرابع في sheet2

        Do While srcRow <= lastRow
            oDst.getCellByPosition(dstCol, dstRow).Formula = _
                oSrc.getCellByPosition(srcCol, srcRow).Formula
            srcRow = srcRow + 1
            dstRow = dstRow + 1
        Loop

    Next i

    ' ===== تكرار الأعمدة 39 و40 و45 بناءً على أول قيمة =====
    fixedCols = Array(38, 39,42,45,48,51,54,_
                                      58,59,62,65,68,71,74_
    )

    For c = 0 To UBound(fixedCols)

        fixedCol = fixedCols(c)
        fixedValue = oDst.getCellByPosition(fixedCol, 3).Formula

        For r = 3 To dstRow - 1
            oDst.getCellByPosition(fixedCol, r).Formula = fixedValue
        Next r

    Next c

End Sub


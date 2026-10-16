Attribute VB_Name = "Module1"
Option Explicit

' A:E on "Спец_ТКП" is the generated five-column specification preview.
' F is a hidden group marker ("H" = section header, "I" = item) used when transferring formatting to Word.
' Source columns: C = description, D = article, F = unit, G = quantity.
' Prices, totals and works are deliberately not included in Appendix 3.
Public Sub ОбновитьСпецификациюТКП()
    Dim itemCount As Long
    On Error GoTo Failed
    Application.Calculate
    itemCount = СобратьСпецификацию()
    MsgBox "Лист 'Спец_ТКП' обновлен. Позиций: " & itemCount & ".", vbInformation
    Exit Sub
Failed:
    MsgBox "Не удалось обновить спецификацию: " & Err.Description, vbExclamation
End Sub

Public Sub СформироватьТКП()
    Dim wd As Object, doc As Object, tbl As Object, specRange As Object
    Dim spec As Worksheet, info As Worksheet
    Dim templatePath As String, itemCount As Long, lastRow As Long
    Dim r As Long, col As Long
    Dim isGroup As Boolean
    Dim number As String, subjectText As String, headers As Variant, errorText As String

    On Error GoTo Failed

    Application.Calculate
    itemCount = СобратьСпецификацию()
    If itemCount = 0 Then
        MsgBox "Нет позиций с количеством больше нуля. Заполните столбец G на листах оборудования, ЗИП или ПО.", vbExclamation
        Exit Sub
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 1000, "СформироватьТКП", "Сначала сохраните генератор на диск."
    End If
    templatePath = НайтиШаблонТКП(ThisWorkbook.Path)
    If Len(templatePath) = 0 Then
        Err.Raise vbObjectError + 1001, "СформироватьТКП", _
                  "Шаблон ТКП не найден рядом с генератором: " & _
                  ThisWorkbook.Path & Application.PathSeparator & "ТКП_шаблон.dotx"
    End If

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo Failed
    If wd Is Nothing Then Set wd = CreateObject("Word.Application")
    wd.Visible = True
    Set doc = wd.Documents.Add(templatePath)

    Set info = НайтиЛист("Наценки-Скидки")
    number = ЧистыйТекст(info.Range("C2").Value2)
    doc.Bookmarks("Num").Range.Text = number
    doc.Bookmarks("Customer").Range.Text = ЧистыйТекст(info.Range("C3").Value2)
    doc.Bookmarks("Object").Range.Text = ЧистыйТекст(info.Range("C4").Value2)
    subjectText = ЧистыйТекст(info.Range("C5").Value2)
    If doc.Bookmarks.Exists("Subject") And Len(subjectText) > 0 Then
        doc.Bookmarks("Subject").Range.Text = subjectText
    End If
    doc.Bookmarks("HeaderNum").Range.Text = number

    Set spec = НайтиЛист("Спец_ТКП")
    lastRow = spec.Cells(spec.Rows.Count, "A").End(xlUp).Row
    Set specRange = doc.Bookmarks("Spec_Table").Range
    If specRange.Tables.Count > 0 Then
        Set tbl = specRange.Tables(1)
        Set specRange = tbl.Range
        tbl.Delete
    End If

    ' Two header rows followed by the already numbered preview rows.
    Set tbl = doc.Tables.Add(specRange, lastRow + 1, 5)
    headers = Array("№ п/п", "Наименование и техническая характеристика", _
                    "Артикул", "Ед. изм.", "Кол-во")
    For col = 1 To 5
        tbl.Cell(1, col).Range.Text = headers(col - 1)
        tbl.Cell(2, col).Range.Text = CStr(col)
    Next col

    For r = 2 To lastRow
        isGroup = (ЧистыйТекст(spec.Cells(r, "F").Value2) = "H")
        For col = 1 To 5
            tbl.Cell(r + 1, col).Range.Text = ЧистыйТекст(spec.Cells(r, col).Value2)
        Next col
        tbl.Cell(r + 1, 1).Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
        tbl.Cell(r + 1, 4).Range.ParagraphFormat.Alignment = 1
        tbl.Cell(r + 1, 5).Range.ParagraphFormat.Alignment = 1
        If isGroup Then
            tbl.Rows(r + 1).Range.Font.Bold = True
            tbl.Rows(r + 1).Shading.BackgroundPatternColor = RGB(217, 217, 217) ' #D9D9D9
        End If
    Next r

    tbl.Borders.Enable = True
    tbl.Borders.InsideColor = RGB(191, 191, 191)  ' #BFBFBF
    tbl.Borders.OutsideColor = RGB(191, 191, 191) ' #BFBFBF
    tbl.Range.Font.Name = "Arial"
    tbl.Range.Font.Size = 10
    tbl.Range.Cells.VerticalAlignment = 1 ' wdCellAlignVerticalCenter

    tbl.Rows(1).HeadingFormat = True
    tbl.Rows(2).HeadingFormat = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(2).Range.Font.Bold = True
    tbl.Rows(1).Range.Font.Color = RGB(255, 255, 255)
    tbl.Rows(2).Range.Font.Color = RGB(255, 255, 255)
    tbl.Rows(1).Shading.BackgroundPatternColor = RGB(192, 0, 0) ' #C00000
    tbl.Rows(2).Shading.BackgroundPatternColor = RGB(192, 0, 0) ' #C00000
    tbl.Rows(1).Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    tbl.Rows(2).Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    tbl.AutoFitBehavior 2 ' wdAutoFitWindow: fit the five columns to the page.
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    On Error GoTo 0
    MsgBox "Не удалось сформировать ТКП: " & errorText, vbExclamation
End Sub

Private Function НайтиШаблонТКП(ByVal baseFolder As String) As String
    Dim candidates As Variant, i As Long, fullPath As String
    candidates = Array("ТКП_шаблон.dotx", "ТКП_шаблон.docx", _
                       "Коммерческое предложение_шаблон.docx", _
                       "Коммерческое предложение.docx")
    For i = LBound(candidates) To UBound(candidates)
        fullPath = baseFolder & Application.PathSeparator & CStr(candidates(i))
        If Len(Dir$(fullPath)) > 0 Then
            НайтиШаблонТКП = fullPath
            Exit Function
        End If
    Next i
End Function

Private Function НайтиЛист(ByVal primaryName As String, Optional ByVal altName As String = "") As Worksheet
    On Error Resume Next
    Set НайтиЛист = ThisWorkbook.Worksheets(primaryName)
    If НайтиЛист Is Nothing And Len(altName) > 0 Then
        Set НайтиЛист = ThisWorkbook.Worksheets(altName)
    End If
    On Error GoTo 0
End Function

Private Function СобратьСпецификацию() As Long
    Dim dst As Worksheet, src As Worksheet
    Dim oldLastRow As Long, outRow As Long, headingNo As Long
    Dim sheetNo As Long, totalItems As Long

    Set dst = НайтиЛист("Спец_ТКП")
    oldLastRow = dst.Cells(dst.Rows.Count, "A").End(xlUp).Row
    With dst.Range("A1:F" & CStr(Application.Max(oldLastRow, 1)))
        .ClearContents
        .Font.Bold = False
    End With
    dst.Range("A1:E1").Value = Array("№ п/п", "Наименование и техническая характеристика", _
                                      "Артикул", "Ед. изм.", "Кол-во")
    dst.Range("A1:E1").Font.Bold = True

    outRow = 2
    headingNo = 0
    totalItems = 0

    ' 1. Программное обеспечение — all positions with G > 0 from SW(w.o.VAT)
    Set src = НайтиЛист("SW(w.o.VAT)", "ПО")
    If Not src Is Nothing Then
        totalItems = totalItems + ОбработатьЛист(src, 0, "Программное обеспечение", _
                                                 dst, outRow, headingNo)
    End If

    ' 2. Cabinets 1..17 — single pass per cabinet sheet (blocks 1..3)
    For sheetNo = 1 To 17
        Set src = НайтиЛист(CStr(sheetNo), "Шкаф " & CStr(sheetNo))
        If Not src Is Nothing Then
            totalItems = totalItems + ОбработатьЛист(src, sheetNo, "", _
                                                     dst, outRow, headingNo)
        End If
    Next sheetNo

    ' 3. ЗИП — all positions with G > 0 from Spare parts
    Set src = НайтиЛист("Spare parts", "ЗИП")
    If Not src Is Nothing Then
        totalItems = totalItems + ОбработатьЛист(src, 0, "ЗИП", _
                                                 dst, outRow, headingNo)
    End If

    dst.Columns("A").ColumnWidth = 7
    dst.Columns("B").ColumnWidth = 65
    dst.Columns("C").ColumnWidth = 27
    dst.Columns("D").ColumnWidth = 13
    dst.Columns("E").ColumnWidth = 12
    dst.Columns("F").Hidden = True
    dst.Range("B1:B" & CStr(Application.Max(outRow - 1, 1))).WrapText = True
    СобратьСпецификацию = totalItems
End Function

Private Function ОбработатьЛист(ByVal src As Worksheet, ByVal cabNo As Long, _
                                ByVal fixedTitle As String, ByVal dst As Worksheet, _
                                ByRef outRow As Long, ByRef headingNo As Long) As Long
    Dim lastRow As Long, r As Long
    Dim currentBlock As Long, effBlock As Long, prevBlock As Long
    Dim itemNo As Long, sheetItems As Long
    Dim hasCaptions As Boolean
    Dim caption As String, sectionTitle As String
    Dim description As String, article As String, unit As String
    Dim rawQty As Variant, qty As Double

    lastRow = src.Cells(src.Rows.Count, "B").End(xlUp).Row
    If lastRow < 48 Then lastRow = 48

    currentBlock = 0
    prevBlock = 0
    hasCaptions = False
    itemNo = 0
    sheetItems = 0

    For r = 1 To lastRow
        caption = ЧистыйТекст(src.Cells(r, "B").Value2)

        If StrComp(Left$(caption, Len("Спецификация ")), "Спецификация ", vbTextCompare) = 0 Then
            currentBlock = НомерБлока(caption)
            hasCaptions = True
            If cabNo > 0 Then itemNo = 0
        ElseIf StrComp(Left$(caption, Len("Итого")), "Итого", vbTextCompare) = 0 Then
            currentBlock = 0
        ElseIf StrComp(caption, "№ п/п", vbTextCompare) = 0 Then
            ' Table column header row — skip.
        Else
            effBlock = currentBlock
            If effBlock = 0 And Not hasCaptions Then
                effBlock = БлокПоСтроке(r)
                If cabNo > 0 And effBlock <> prevBlock Then
                    itemNo = 0
                    prevBlock = effBlock
                End If
            End If

            If effBlock > 0 Then
                rawQty = src.Cells(r, "G").Value2
                If Not IsError(rawQty) Then
                    If IsNumeric(rawQty) And Not IsEmpty(rawQty) Then
                        qty = CDbl(rawQty)
                        If qty > 0 Then
                            description = ЧистыйТекст(src.Cells(r, "C").Value2)
                            article = ЧистыйТекст(src.Cells(r, "D").Value2)
                            unit = ЧистыйТекст(src.Cells(r, "F").Value2)

                            ' Some NVT descriptions still have broken #REF! price-list lookups.
                            ' Keep the position in the specification by using its article.
                            If Len(description) = 0 Or description = "0" Or Left$(description, 1) = "#" Then
                                description = article
                            End If
                            If Len(unit) = 0 Then unit = "шт."

                            If Len(description) > 0 Then
                                If itemNo = 0 Then
                                    If cabNo > 0 Then
                                        sectionTitle = ЗаголовокБлокаШкафа(cabNo, effBlock)
                                    Else
                                        sectionTitle = fixedTitle
                                    End If
                                    headingNo = headingNo + 1
                                    dst.Cells(outRow, "A").NumberFormat = "@"
                                    dst.Cells(outRow, "A").Value2 = CStr(headingNo) & "."
                                    dst.Cells(outRow, "B").Value2 = sectionTitle
                                    dst.Cells(outRow, "C").Value2 = ""
                                    dst.Cells(outRow, "D").Value2 = "Компл."
                                    dst.Cells(outRow, "E").Value2 = 1
                                    dst.Cells(outRow, "F").Value2 = "H"
                                    dst.Range("A" & outRow & ":E" & outRow).Font.Bold = True
                                    outRow = outRow + 1
                                End If

                                itemNo = itemNo + 1
                                sheetItems = sheetItems + 1
                                dst.Cells(outRow, "A").NumberFormat = "@"
                                dst.Cells(outRow, "A").Value2 = CStr(itemNo) & "."
                                dst.Cells(outRow, "B").Value2 = description
                                dst.Cells(outRow, "C").Value2 = article
                                dst.Cells(outRow, "D").Value2 = unit
                                dst.Cells(outRow, "E").Value2 = qty
                                dst.Cells(outRow, "F").Value2 = "I"
                                outRow = outRow + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    ОбработатьЛист = sheetItems
End Function

Private Function ЗаголовокБлокаШкафа(ByVal cabNo As Long, ByVal blockNo As Long) As String
    Select Case blockNo
        Case 1
            ЗаголовокБлокаШкафа = "Шкаф " & CStr(cabNo) & " ПТК"
        Case 2
            ЗаголовокБлокаШкафа = "Шкаф " & CStr(cabNo) & " телекоммуникационное оборудование"
        Case 3
            ЗаголовокБлокаШкафа = "Шкаф " & CStr(cabNo) & " шкафное оборудование"
    End Select
End Function

Private Function НомерБлока(ByVal caption As String) As Long
    If InStr(1, caption, "НВТ", vbTextCompare) > 0 Then
        НомерБлока = 1
    ElseIf InStr(1, caption, "Эксара", vbTextCompare) > 0 Or _
           InStr(1, caption, "ИнСАТ", vbTextCompare) > 0 Then
        НомерБлока = 2
    ElseIf InStr(1, caption, "3-х поставщиков", vbTextCompare) > 0 Then
        НомерБлока = 3
    End If
End Function

Private Function БлокПоСтроке(ByVal r As Long) As Long
    If r >= 3 And r <= 9 Then
        БлокПоСтроке = 1
    ElseIf r >= 14 And r <= 15 Then
        БлокПоСтроке = 2
    ElseIf r >= 20 And r <= 47 Then
        БлокПоСтроке = 3
    End If
End Function

Private Function ЧистыйТекст(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    ЧистыйТекст = Trim$(Replace(Replace(Replace(Replace(CStr(value), _
                   vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " "))
End Function

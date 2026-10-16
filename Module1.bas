Attribute VB_Name = "Module1"
Option Explicit

' A:E on "Спец_ТКП" is the generated five-column specification preview.
' F is a hidden group marker used when transferring section formatting to Word.
' Source: C = description, D = article, F = unit, G = quantity.
' Prices and work are deliberately not included in Appendix 3.
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
    Dim wd As Object, doc As Object, tbl As Object
    Dim spec As Worksheet, info As Worksheet
    Dim templatePath As String, itemCount As Long, lastRow As Long
    Dim r As Long, col As Long
    Dim isGroup As Boolean
    Dim number As String, headers As Variant, errorText As String

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
    templatePath = ThisWorkbook.Path & Application.PathSeparator & "ТКП_шаблон.dotx"
    If Len(Dir$(templatePath)) = 0 Then
        Err.Raise vbObjectError + 1001, "СформироватьТКП", "Шаблон ТКП не найден рядом с генератором: " & templatePath
    End If

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo Failed
    If wd Is Nothing Then Set wd = CreateObject("Word.Application")
    wd.Visible = True
    Set doc = wd.Documents.Add(templatePath)

    Set info = ThisWorkbook.Worksheets("Наценки-Скидки")
    number = ЧистыйТекст(info.Range("C2").Value2)
    doc.Bookmarks("Num").Range.Text = number
    doc.Bookmarks("Customer").Range.Text = ЧистыйТекст(info.Range("C3").Value2)
    doc.Bookmarks("Object").Range.Text = ЧистыйТекст(info.Range("C4").Value2)
    doc.Bookmarks("HeaderNum").Range.Text = number

    Set spec = ThisWorkbook.Worksheets("Спец_ТКП")
    lastRow = spec.Cells(spec.Rows.Count, "A").End(xlUp).Row
    ' Two header rows followed by the already numbered preview rows.
    Set tbl = doc.Tables.Add(doc.Bookmarks("Spec_Table").Range, lastRow + 1, 5)
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
        If isGroup Then tbl.Rows(r + 1).Range.Font.Bold = True
    Next r

    tbl.Borders.Enable = True
    tbl.Rows(1).HeadingFormat = True
    tbl.Rows(2).HeadingFormat = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(2).Range.Font.Bold = True
    tbl.AutoFitBehavior 2 ' wdAutoFitWindow: fit the five columns to the page.
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    On Error GoTo 0
    MsgBox "Не удалось сформировать ТКП: " & errorText, vbExclamation
End Sub

Private Function СобратьСпецификацию() As Long
    Dim items() As Variant, titles As Variant
    Dim capacity As Long, itemCount As Long
    Dim sheetNo As Long, groupNo As Long, i As Long, outRow As Long
    Dim itemNo As Long, headingNo As Long
    Dim groupHasItems As Boolean, oldLastRow As Long
    Dim dst As Worksheet

    capacity = 128
    ReDim items(1 To 5, 1 To capacity)

    For sheetNo = 1 To 17
        ЧитатьЛист ThisWorkbook.Worksheets(CStr(sheetNo)), items, itemCount, capacity
    Next sheetNo
    ЧитатьЛист ThisWorkbook.Worksheets("Spare parts"), items, itemCount, capacity
    ЧитатьЛист ThisWorkbook.Worksheets("SW(w.o.VAT)"), items, itemCount, capacity

    titles = Array("Оборудование НВТ", "Оборудование Эксара", _
                   "Оборудование 3-х поставщиков", "ЗИП НВТ", _
                   "ЗИП Эксара", "ЗИП 3-х поставщиков", _
                   "ПО НВТ", "ПО ИнСАТ", "ПО 3-х поставщиков")

    Set dst = ThisWorkbook.Worksheets("Спец_ТКП")
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

    For groupNo = 1 To 9
        groupHasItems = False
        For i = 1 To itemCount
            If CLng(items(1, i)) = groupNo Then
                If Not groupHasItems Then
                    headingNo = headingNo + 1
                    dst.Cells(outRow, "A").NumberFormat = "@"
                    dst.Cells(outRow, "A").Value2 = CStr(headingNo) & "."
                    dst.Cells(outRow, "B").Value2 = titles(groupNo - 1)
                    dst.Cells(outRow, "F").Value2 = "H"
                    dst.Range("A" & outRow & ":E" & outRow).Font.Bold = True
                    outRow = outRow + 1
                    groupHasItems = True
                End If
                dst.Cells(outRow, "A").NumberFormat = "@"
                dst.Cells(outRow, "A").Value2 = CStr(itemNo + 1) & "."
                dst.Cells(outRow, "B").Value2 = items(2, i)
                dst.Cells(outRow, "C").Value2 = items(3, i)
                dst.Cells(outRow, "D").Value2 = items(4, i)
                dst.Cells(outRow, "E").Value2 = items(5, i)
                dst.Cells(outRow, "F").Value2 = "I"
                itemNo = itemNo + 1
                outRow = outRow + 1
            End If
        Next i
        itemNo = 0
    Next groupNo

    dst.Columns("A").ColumnWidth = 7
    dst.Columns("B").ColumnWidth = 65
    dst.Columns("C").ColumnWidth = 27
    dst.Columns("D").ColumnWidth = 13
    dst.Columns("E").ColumnWidth = 12
    dst.Columns("F").Hidden = True
    dst.Range("B1:B" & CStr(outRow - 1)).WrapText = True
    СобратьСпецификацию = itemCount
End Function

Private Sub ЧитатьЛист(ByVal src As Worksheet, ByRef items() As Variant, _
                      ByRef itemCount As Long, ByRef capacity As Long)
    Dim lastRow As Long, r As Long, groupNo As Long
    Dim caption As String, description As String, article As String, unit As String
    Dim rawQty As Variant, qty As Double

    lastRow = src.Cells(src.Rows.Count, "B").End(xlUp).Row
    For r = 1 To lastRow
        caption = ЧистыйТекст(src.Cells(r, "B").Value2)
        If StrComp(Left$(caption, Len("Спецификация ")), "Спецификация ", vbTextCompare) = 0 Then
            groupNo = НомерРаздела(caption, src.Name)
        ElseIf StrComp(Left$(caption, Len("Итого ")), "Итого ", vbTextCompare) = 0 Then
            groupNo = 0
        ElseIf groupNo > 0 Then
            rawQty = src.Cells(r, "G").Value2
            If Not IsError(rawQty) Then
                If IsNumeric(rawQty) Then
                    qty = CDbl(rawQty)
                    If qty > 0 Then
                        description = ЧистыйТекст(src.Cells(r, "C").Value2)
                        article = ЧистыйТекст(src.Cells(r, "D").Value2)
                        unit = ЧистыйТекст(src.Cells(r, "F").Value2)

                        ' Some NVT descriptions still have broken #REF! price-list lookups.
                        ' Keep the position in the specification by using its article.
                        If Len(description) = 0 Or description = "0" Then description = article
                        If Len(description) = 0 Or Len(unit) = 0 Then
                            Err.Raise vbObjectError + 1002, "Спец_ТКП", _
                                      "Лист '" & src.Name & "', строка " & r & _
                                      ": при положительном количестве нужны наименование/артикул (C/D) и единица измерения (F)."
                        End If

                        ДобавитьПозицию items, itemCount, capacity, groupNo, _
                                         description, article, unit, qty
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Function НомерРаздела(ByVal caption As String, ByVal sheetName As String) As Long
    Dim base As Long, suffix As Long

    If InStr(1, caption, "Спецификация оборудования", vbTextCompare) = 1 Then
        If sheetName = "Spare parts" Then base = 3
    ElseIf InStr(1, caption, "Спецификация ПО", vbTextCompare) = 1 Then
        base = 6
    Else
        Exit Function
    End If

    If InStr(1, caption, "НВТ", vbTextCompare) > 0 Then
        suffix = 1
    ElseIf InStr(1, caption, "Эксара", vbTextCompare) > 0 Or _
           InStr(1, caption, "ИнСАТ", vbTextCompare) > 0 Then
        suffix = 2
    ElseIf InStr(1, caption, "3-х поставщиков", vbTextCompare) > 0 Then
        suffix = 3
    End If
    If suffix = 0 Then Exit Function
    НомерРаздела = base + suffix
End Function

Private Sub ДобавитьПозицию(ByRef items() As Variant, ByRef itemCount As Long, _
                           ByRef capacity As Long, ByVal groupNo As Long, _
                           ByVal description As String, ByVal article As String, _
                           ByVal unit As String, ByVal qty As Double)
    Dim i As Long

    ' Consolidate identical items across sheets 1..17; never mix different
    ' groups, articles, units or technical descriptions.
    For i = 1 To itemCount
        If CLng(items(1, i)) = groupNo Then
            If StrComp(items(2, i), description, vbTextCompare) = 0 And _
               StrComp(items(3, i), article, vbTextCompare) = 0 And _
               StrComp(items(4, i), unit, vbTextCompare) = 0 Then
                items(5, i) = CDbl(items(5, i)) + qty
                Exit Sub
            End If
        End If
    Next i

    If itemCount = capacity Then
        capacity = capacity * 2
        ReDim Preserve items(1 To 5, 1 To capacity)
    End If

    itemCount = itemCount + 1
    items(1, itemCount) = groupNo
    items(2, itemCount) = description
    items(3, itemCount) = article
    items(4, itemCount) = unit
    items(5, itemCount) = qty
End Sub

Private Function ЧистыйТекст(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    ЧистыйТекст = Trim$(Replace(Replace(Replace(Replace(CStr(value), _
                   vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " "))
End Function

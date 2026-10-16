Attribute VB_Name = "Module1"
Sub СформироватьТКП()
    Dim ws As Worksheet, d As Variant, i As Long, g As Long, p As Long
    Dim s As String, wd As Object, doc As Object, rng As Object, t As Object

    Set ws = ThisWorkbook.Worksheets("Спец_ТКП")
    d = ws.Range("A1").CurrentRegion.Value

    ' Собираем текст таблицы с нумерацией в стиле Приложения 3
    s = "№ п/п" & vbTab & "Наименование и техническая характеристика" & vbTab & _
        "Артикул" & vbTab & "Ед. изм." & vbTab & "Кол-во" & vbCr & _
        "1" & vbTab & "2" & vbTab & "3" & vbTab & "4" & vbTab & "5"
    For i = 2 To UBound(d, 1)
        If d(i, 1) = "H" Then g = g + 1: p = 0: s = s & vbCr & g & "." _
            Else p = p + 1: s = s & vbCr & p & "."
        s = s & vbTab & Replace(d(i, 2), vbLf, " ") & vbTab & d(i, 3) & _
                vbTab & d(i, 4) & vbTab & d(i, 5)
    Next i

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then Set wd = CreateObject("Word.Application")
    wd.Visible = True
    Set doc = wd.Documents.Add(ThisWorkbook.Path & "\ТКП_шаблон.dotx")

    ' Шапка ТКП
    With ThisWorkbook.Worksheets("Наценки-Скидки")
        doc.Bookmarks("Num").Range.Text = .Range("C2").Value
        doc.Bookmarks("Customer").Range.Text = .Range("C3").Value
        doc.Bookmarks("Object").Range.Text = .Range("C4").Value
    End With

    ' Таблица спецификации
    Set rng = doc.Bookmarks("Spec_Table").Range
    rng.Text = s
    Set t = rng.ConvertToTable(Separator:=1, NumColumns:=5)   ' 1 = табуляция
    t.Borders.Enable = True
    t.Rows(1).HeadingFormat = True                            ' повтор шапки на каждой странице
    t.Rows(1).Range.Font.Bold = True
    For i = 2 To UBound(d, 1)
        If d(i, 1) = "H" Then t.Rows(i + 1).Range.Font.Bold = True
    Next i
End Sub


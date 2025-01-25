Private Sub Document_Open()
    FocusOnSecondRow
End Sub

Private Sub karekod_click()
    ParseKarekod
End Sub

Private Sub temizle_click()
    ClearData
End Sub

Sub ClearData()
    Dim doc As Document
    Dim rng As Range
    Dim para As paragraph
    Dim i As Long

    Set doc = ActiveDocument

    ' İkinci paragraftan itibaren tüm verileri temizle
    For i = 2 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        Set rng = para.Range
        rng.Text = ""
    Next i
End Sub
Sub ParseKarekod()
    Dim doc As Document
    Dim newDoc As Document
    Dim newTbl As Table
    Dim i As Long
    Dim karekod As String
    Dim pieces(1 To 8) As String
    Dim paragraph As paragraph

    Set doc = ActiveDocument

    ' Yeni belge oluştur
    Set newDoc = Documents.Add

    ' Yeni tablo oluştur ve başlıkları ekle
    Set newTbl = newDoc.Tables.Add(Range:=newDoc.Content, NumRows:=1, NumColumns:=8)
    newTbl.Cell(1, 1).Range.Text = "BARKOD KOD"
    newTbl.Cell(1, 2).Range.Text = "BARKOD"
    newTbl.Cell(1, 3).Range.Text = "SERİNO KOD"
    newTbl.Cell(1, 4).Range.Text = "SERİNO"
    newTbl.Cell(1, 5).Range.Text = "SKT KOD"
    newTbl.Cell(1, 6).Range.Text = "SKT"
    newTbl.Cell(1, 7).Range.Text = "LOT KOD"
    newTbl.Cell(1, 8).Range.Text = "LOT"
    newTbl.Rows(1).HeadingFormat = True

    ' Kenarlık ekle ve verileri hizala
    With newTbl
        .Borders.Enable = True
        For i = 1 To .Rows.Count
            For j = 1 To .Columns.Count
                .Cell(i, j).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cell(i, j).VerticalAlignment = wdCellAlignVerticalCenter
            Next j
        Next i
    End With

    ' Belgedeki paragrafları gezerek karekodları parçala ve yeni tabloya ekle
    For i = 2 To doc.Paragraphs.Count ' İkinci paragraftan itibaren başla
        karekod = Trim(doc.Paragraphs(i).Range.Text)
        If Len(karekod) >= 50 Then
            pieces(1) = Mid(karekod, 1, 2)
            pieces(2) = Mid(karekod, 3, 14)
            pieces(3) = Mid(karekod, 17, 2)
            pieces(4) = Mid(karekod, 19, 16)
            pieces(5) = Mid(karekod, 35, 2)
            pieces(6) = Mid(karekod, 37, 6)
            pieces(7) = Mid(karekod, 43, 2)
            pieces(8) = Mid(karekod, 45, 6)

            newTbl.Rows.Add
            For j = 1 To 8
                newTbl.Cell(newTbl.Rows.Count, j).Range.Text = pieces(j)
                newTbl.Cell(newTbl.Rows.Count, j).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                newTbl.Cell(newTbl.Rows.Count, j).VerticalAlignment = wdCellAlignVerticalCenter
            Next j
        End If
    Next i

    ' Sütun genişliklerini otomatik ayarla
    newTbl.AutoFitBehavior wdAutoFitContent

    ' İlk satırın wrap text özelliğini aktif et
    For j = 1 To 8
        newTbl.Cell(1, j).Range.ParagraphFormat.WordWrap = True
    Next j
End Sub


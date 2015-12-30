Public Class wordMethod_2

    Public WordApplication
    Public WordBook

    Public Sub New(filepath As String, Optional ifvis As Boolean = False)

        WordApplication = CreateObject("Word.Application")

        WordApplication.Visible = ifvis
        If (WordApplication Is Nothing) Then
            MsgBox("Unable to start Word", vbOKOnly + vbExclamation, "Error")
            Exit Sub
        Else
            'Open up the workbook
            On Error Resume Next
            WordBook = WordApplication.Workbooks.Open(filepath)

            If (WordBook Is Nothing) Then
                MsgBox("Unable to find workbook", vbOKOnly + vbExclamation, "Error")
                Exit Sub

            End If
        End If


    End Sub
    Public Sub New(Optional ifvis As Boolean = False)

        WordApplication = CreateObject("Word.Application")

        WordApplication.Visible = ifvis
        If (WordApplication Is Nothing) Then
            MsgBox("Unable to start Word", vbOKOnly + vbExclamation, "Error")
            Exit Sub
        Else
            'Open up the workbook
            On Error Resume Next
            WordBook = WordApplication.Workbooks.Add

            If (WordBook Is Nothing) Then
                MsgBox("Unable to creat workbook", vbOKOnly + vbExclamation, "Error")
                Exit Sub

            End If
        End If


    End Sub


    Public Function Get_CellValue(rowindex As Integer, colindex As Integer, Optional sheetid As Integer = 1) As String

        Dim wSheet = WordBook.Worksheets.Item(sheetid)

        Return Strings.Trim(wSheet.Cells(rowindex, colindex).text)

    End Function

    Public Sub Set_CellValue(rowindex As Integer, colindex As Integer, content As String, Optional sheetid As Integer = 1)

        Dim wSheet = WordBook.Worksheets.Item(sheetid)

        wSheet.Cells(rowindex, colindex) = Strings.Trim(content)

    End Sub

    Public Sub SaveAs(filepath As String)
        WordBook.SaveAs(filepath)
    End Sub
    Public Sub Save()
        WordBook.Save()
    End Sub
    Public Sub Quit()
        WordApplication.quit()

    End Sub

    Public Shared Function SaveDataTableToWord(extbDic As Dictionary(Of String, DataTable), Optional filepath As String = "") As Object

        Dim app = CreateObject("Word.Application")

        app.Visible = True


        Dim wBook = app.Workbooks.Add

        For Each item In extbDic
            Dim WordTable = item.Value


            wBook.Sheets.Add()
            Dim wSheet = wBook.Worksheets.Item(1)
            Dim row = WordTable.Rows.Count
            Dim col = WordTable.Columns.Count
            If row > 0 Then

                For i = 0 To row - 1
                    wSheet.Cells(i + 2, 1) = (i + 1).ToString()
                    For j = 0 To col - 1


                        wSheet.Cells(i + 2, j + 2) = WordTable.Rows(i)(j).ToString()

                    Next

                Next



            End If
            wSheet.Cells(1, 1) = "SEQ"
            For k = 0 To col - 1

                wSheet.Cells(1, 2 + k) = WordTable.Columns(k).ColumnName

            Next

            wSheet.Columns.AutoFit()
            wSheet.Rows.AutoFit()
        Next



        If filepath <> "" Then

            app.DisplayAlerts = False
            app.AlertBeforeOverwriting = False
            wBook.SaveAs(filepath)

            app.Quit()
            app = Nothing
            Return Nothing
        Else

            Return wBook

        End If


    End Function


    Public Shared Function SaveDataTableToWord(WordTable As DataTable, Optional filepath As String = "") As Object

        Dim app = CreateObject("Word.Application")

        app.Visible = True


        Dim wBook = app.Workbooks.Add
        Do Until wBook.Sheets.count >= 1
            wBook.Sheets.Add()
        Loop

        Dim wSheet = wBook.Worksheets.Item(1)
        Dim row = WordTable.Rows.Count
        Dim col = WordTable.Columns.Count
        If row > 0 Then

            For i = 0 To row - 1
                wSheet.Cells(i + 2, 1) = (i + 1).ToString()
                For j = 0 To col - 1


                    wSheet.Cells(i + 2, j + 2) = WordTable.Rows(i)(j).ToString()

                Next

            Next



        End If
        wSheet.Cells(1, 1) = "SEQ"
        For k = 0 To col - 1

            wSheet.Cells(1, 2 + k) = WordTable.Columns(k).ColumnName

        Next

        wSheet.Columns.AutoFit()
        wSheet.Rows.AutoFit()
        If filepath <> "" Then

            app.DisplayAlerts = False
            app.AlertBeforeOverwriting = False
            wBook.SaveAs(filepath)

            app.Quit()
            app = Nothing
            Return Nothing
        Else

            Return wSheet

        End If

    End Function


    Public Shared Function LoadDataFromWord(FilePath As String) As DataTable


        'Create an Word worksheet object
        Dim WordApplication = CreateObject("Word.Application")
        WordApplication.Visible = False
        Dim WordBook
        Dim WordSheet
        'Open up the file
        If (WordApplication Is Nothing) Then
            MsgBox("Unable to start Word", vbOKOnly + vbExclamation, "Error")
            Exit Function
        Else
            'Open up the workbook
            On Error Resume Next
            WordBook = WordApplication.Workbooks.Open(FilePath)

            If (WordBook Is Nothing) Then
                MsgBox("Unable to find workbook", vbOKOnly + vbExclamation, "Error")
                Exit Function
            Else

                'Open up the worksheet
                WordSheet = WordApplication.Worksheets.Item(1)

                If (WordSheet Is Nothing) Then
                    MsgBox("Unable to find worksheet", vbOKOnly + vbExclamation, "Error")
                    Exit Function

                End If
            End If
        End If

        Dim RowIndex
        RowIndex = 2
        Dim ColIndex
        ColIndex = 1
        Dim MyName As String

        MyName = Strings.Trim(WordSheet.Cells(RowIndex, 1).text)

        Dim MyNameCol = Strings.Trim(WordSheet.Cells(1, ColIndex).text)
        Dim dt = New DataTable()
        'dt.Columns.Add(Strings.Trim(WordSheet.Cells(1, 1).text), GetType(String))
        While (Len(MyNameCol) > 0)

            dt.Columns.Add(MyNameCol, GetType(String))


            ColIndex = ColIndex + 1
            MyNameCol = Strings.Trim(WordSheet.Cells(1, ColIndex).text)


        End While



        While (Len(MyName) > 0)

            Dim dr = dt.NewRow()
            For i = 1 To ColIndex - 1

                dr(i - 1) = Strings.Trim(WordSheet.Cells(RowIndex, i).text)


            Next




            RowIndex = RowIndex + 1
            MyName = Strings.Trim(WordSheet.Cells(RowIndex, 1).text)


            dt.Rows.Add(dr)
        End While

        Return dt



    End Function


    Public Shared Function SaveDataTableToWordTran(extbDic As Dictionary(Of String, DataTable), Optional filepath As String = "") As Object


        Dim app = CreateObject("Word.Application")

        app.Visible = True


        Dim wBook = app.Workbooks.Add

        For Each item In extbDic
            Dim WordTable = item.Value


            wBook.Sheets.Add()
            Dim wSheet = wBook.Worksheets.Item(1)
            wSheet.Name = item.Key
            Dim row = WordTable.Rows.Count
            Dim col = WordTable.Columns.Count
            If row > 0 Then

                For i = 0 To row - 1
                    ' wSheet.Cells(2, i + 2) = (i + 1).ToString()
                    For j = 0 To col - 1


                        wSheet.Cells(j + 2, i + 2) = WordTable.Rows(i)(j).ToString()

                    Next

                Next



            End If
            wSheet.Cells(1, 1) = item.Key
            For k = 0 To col - 1

                wSheet.Cells(2 + k, 1) = WordTable.Columns(k).ColumnName

            Next

            wSheet.Columns.AutoFit()
            wSheet.Rows.AutoFit()




            Dim allColumn = wSheet.Columns
            Dim allColumn3 = wSheet.Range(wSheet.Cells(2, 1), wSheet.Cells(WordTable.Columns.Count + 1, WordTable.Rows.Count + 1))

            allColumn.HorizontalAlignment = 2

            allColumn.ColumnWidth = 15
            allColumn.WrapText = True
            Dim allColumn2 = wSheet.Columns(1)
            allColumn2.ColumnWidth = 18.5
            allColumn2.RowHeight = 27
            Dim allColumn4 = wSheet.Rows(5)
            allColumn4.RowHeight = 50
            allColumn.AutoFit()
            ' wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells(1, Wordsheet.Rows.Count + 1)).Merge()
            allColumn3.Borders.LineStyle = 1
            allColumn3.BorderAround(1, 2, 1)


            'wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, Wordsheet.Rows.Count + 1]).Merge();
            '      //allColumn3.Borders.LineStyle=1;
            '   allColumn3.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());
            '   // wSheet.Name = abcname[k];










        Next



        If filepath <> "" Then

            app.DisplayAlerts = False
            app.AlertBeforeOverwriting = False
            wBook.SaveAs(filepath)

            app.Quit()
            app = Nothing
            Return Nothing
        Else

            Return wBook

        End If




    End Function







    Public Shared Sub setNumberFormat(wsheet As Object, colnum As Integer, NumberFormat As String)
        wsheet.Columns(colnum).NumberFormatLocal = NumberFormat
    End Sub

    Public Shared Sub SaveAs(wsheet As Object, filepath As String)
        wsheet.SaveAs(filepath)
    End Sub



End Class

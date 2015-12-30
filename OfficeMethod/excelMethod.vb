Public Class excelMethod

    Public ExcelApplication
    Public ExcelBook

    Public Sub New(filepath As String, Optional ifvis As Boolean = False)

        ExcelApplication = CreateObject("Excel.Application")

        ExcelApplication.Visible = ifvis
        If (ExcelApplication Is Nothing) Then
            MsgBox("Unable to start Excel", vbOKOnly + vbExclamation, "Error")
            Exit Sub
        Else
            'Open up the workbook
            On Error Resume Next
            ExcelBook = ExcelApplication.Workbooks.Open(filepath)

            If (ExcelBook Is Nothing) Then
                MsgBox("Unable to find workbook", vbOKOnly + vbExclamation, "Error")
                Exit Sub

            End If
        End If


    End Sub
    Public Sub New(Optional ifvis As Boolean = False)

        ExcelApplication = CreateObject("Excel.Application")

        ExcelApplication.Visible = ifvis
        If (ExcelApplication Is Nothing) Then
            MsgBox("Unable to start Excel", vbOKOnly + vbExclamation, "Error")
            Exit Sub
        Else
            'Open up the workbook
            On Error Resume Next
            ExcelBook = ExcelApplication.Workbooks.Add

            If (ExcelBook Is Nothing) Then
                MsgBox("Unable to creat workbook", vbOKOnly + vbExclamation, "Error")
                Exit Sub

            End If
        End If


    End Sub


    Public Function Get_CellValue(rowindex As Integer, colindex As Integer, Optional sheetid As Integer = 1) As String

        Dim wSheet = ExcelBook.Worksheets.Item(sheetid)

        Return Strings.Trim(wSheet.Cells(rowindex, colindex).text)

    End Function

    Public Sub Set_CellValue(rowindex As Integer, colindex As Integer, content As String, Optional sheetid As Integer = 1)

        Dim wSheet = ExcelBook.Worksheets.Item(sheetid)

        wSheet.Cells(rowindex, colindex) = Strings.Trim(content)

    End Sub

    Public Sub SaveAs(filepath As String)
        ExcelBook.SaveAs(filepath)
    End Sub
    Public Sub Save()
        ExcelBook.Save()
    End Sub
    Public Sub Quit()
        ExcelApplication.quit()

    End Sub

    Public Shared Function SaveDataTableToExcel(extbDic As Dictionary(Of String, DataTable), Optional filepath As String = "") As Object

        Dim app = CreateObject("Excel.Application")

        app.Visible = True


        Dim wBook = app.Workbooks.Add

        For Each item In extbDic
            Dim excelTable = item.Value


            wBook.Sheets.Add()
            Dim wSheet = wBook.Worksheets.Item(1)
            Dim row = excelTable.Rows.Count
            Dim col = excelTable.Columns.Count
            If row > 0 Then

                For i = 0 To row - 1
                    wSheet.Cells(i + 2, 1) = (i + 1).ToString()
                    For j = 0 To col - 1


                        wSheet.Cells(i + 2, j + 2) = excelTable.Rows(i)(j).ToString()

                    Next

                Next



            End If
            wSheet.Cells(1, 1) = "SEQ"
            For k = 0 To col - 1

                wSheet.Cells(1, 2 + k) = excelTable.Columns(k).ColumnName

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


    Public Shared Function SaveDataTableToExcel(excelTable As DataTable, Optional filepath As String = "", Optional wBook As Object = Nothing, Optional iftext As Object = False) As Object
        Dim app
        If wBook Is Nothing Then

            app = CreateObject("Excel.Application")
            wBook = app.Workbooks.Add
            Do Until wBook.Sheets.count >= 1
                wBook.Sheets.Add()
            Loop
        Else
            app = wBook.Application
        End If


        app.Visible = True




        Dim wSheet = wBook.Worksheets.Item(1)
        If (iftext) Then
            wSheet.Columns.NumberFormatLocal = "@"
        End If

        Dim row = excelTable.Rows.Count
        Dim col = excelTable.Columns.Count
        If row > 0 Then

            For i = 0 To row - 1
                wSheet.Cells(i + 2, 1) = (i + 1).ToString()
                For j = 0 To col - 1


                    wSheet.Cells(i + 2, j + 2) = excelTable.Rows(i)(j).ToString()

                Next

            Next



        End If
        wSheet.Cells(1, 1) = "SEQ"
        For k = 0 To col - 1

            wSheet.Cells(1, 2 + k) = excelTable.Columns(k).ColumnName

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


    Public Shared Function LoadDataFromExcel(FilePath As String) As DataTable


        'Create an Excel worksheet object
        Dim ExcelApplication = CreateObject("Excel.Application")
        ExcelApplication.Visible = False
        Dim ExcelBook
        Dim ExcelSheet
        'Open up the file
        If (ExcelApplication Is Nothing) Then
            MsgBox("Unable to start Excel", vbOKOnly + vbExclamation, "Error")
            Exit Function
        Else
            'Open up the workbook
            On Error Resume Next
            ExcelBook = ExcelApplication.Workbooks.Open(FilePath)

            If (ExcelBook Is Nothing) Then
                MsgBox("Unable to find workbook", vbOKOnly + vbExclamation, "Error")
                Exit Function
            Else

                'Open up the worksheet
                ExcelSheet = ExcelApplication.Worksheets.Item(1)

                If (ExcelSheet Is Nothing) Then
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

        MyName = Strings.Trim(ExcelSheet.Cells(RowIndex, 1).text)

        Dim MyNameCol = Strings.Trim(ExcelSheet.Cells(1, ColIndex).text)
        Dim dt = New DataTable()
        'dt.Columns.Add(Strings.Trim(ExcelSheet.Cells(1, 1).text), GetType(String))
        While (Len(MyNameCol) > 0)

            dt.Columns.Add(MyNameCol, GetType(String))


            ColIndex = ColIndex + 1
            MyNameCol = Strings.Trim(ExcelSheet.Cells(1, ColIndex).text)


        End While



        While (Len(MyName) > 0)

            Dim dr = dt.NewRow()
            For i = 1 To ColIndex - 1

                dr(i - 1) = Strings.Trim(ExcelSheet.Cells(RowIndex, i).text)


            Next




            RowIndex = RowIndex + 1
            MyName = Strings.Trim(ExcelSheet.Cells(RowIndex, 1).text)


            dt.Rows.Add(dr)
        End While

        Return dt



    End Function


    Public Shared Function SaveDataTableToExcelTran(extbDic As Dictionary(Of String, DataTable), Optional filepath As String = "") As Object


        Dim app = CreateObject("Excel.Application")

        app.Visible = True


        Dim wBook = app.Workbooks.Add

        For Each item In extbDic
            Dim excelTable = item.Value


            wBook.Sheets.Add()
            Dim wSheet = wBook.Worksheets.Item(1)
            wSheet.Name = item.Key
            Dim row = excelTable.Rows.Count
            Dim col = excelTable.Columns.Count
            If row > 0 Then

                For i = 0 To row - 1
                    ' wSheet.Cells(2, i + 2) = (i + 1).ToString()
                    For j = 0 To col - 1


                        wSheet.Cells(j + 2, i + 2) = excelTable.Rows(i)(j).ToString()

                    Next

                Next



            End If
            wSheet.Cells(1, 1) = item.Key
            For k = 0 To col - 1

                wSheet.Cells(2 + k, 1) = excelTable.Columns(k).ColumnName

            Next

            wSheet.Columns.AutoFit()
            wSheet.Rows.AutoFit()




            Dim allColumn = wSheet.Columns
            Dim allColumn3 = wSheet.Range(wSheet.Cells(2, 1), wSheet.Cells(excelTable.Columns.Count + 1, excelTable.Rows.Count + 1))

            allColumn.HorizontalAlignment = 2

            allColumn.ColumnWidth = 15
            allColumn.WrapText = True
            Dim allColumn2 = wSheet.Columns(1)
            allColumn2.ColumnWidth = 18.5
            allColumn2.RowHeight = 27
            Dim allColumn4 = wSheet.Rows(5)
            allColumn4.RowHeight = 50
            allColumn.AutoFit()
            ' wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells(1, excelsheet.Rows.Count + 1)).Merge()
            allColumn3.Borders.LineStyle = 1
            allColumn3.BorderAround(1, 2, 1)


            'wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, excelsheet.Rows.Count + 1]).Merge();
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

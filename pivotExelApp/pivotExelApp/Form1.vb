Imports Microsoft.Office.Interop

Public Class Form1

    Private fileNameReference As String
    Private pathReference As String
    Private referencePath As String

#Region "global Variable"
    Private Division As String
    Private Insurer As String
    Private PolClass As String
    Private InvoiceClass As String
#End Region

#Region "Buttons"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSelectRaw.Click
        Using dialog As New OpenFileDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            txtRawData.Text = dialog.FileName
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnSelectReference.Click
        Using dialog As New OpenFileDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            txtReferenceData.Text = dialog.FileName
            fileNameReference = System.IO.Path.GetFileName(dialog.FileName)
            pathReference = System.IO.Path.GetDirectoryName(dialog.FileName)

            referencePath = pathReference + "\" + "[" + fileNameReference + "]"
        End Using
    End Sub
#End Region
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnCreatePivot.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim newWorkBook As Excel.Workbook
        Dim newWorkSheet As Excel.Worksheet

        xlApp = New Excel.Application

        Dim filenameRaw As String = txtRawData.Text
        If txtRawData.Text.Length > 0 Then
            filenameRaw = txtRawData.Text
        Else
            MsgBox("Please select raw File")
            Exit Sub
        End If

        Dim filenameReference As String = txtReferenceData.Text
        If txtReferenceData.Text.Length > 0 Then
            filenameReference = txtReferenceData.Text

            'get the last cell for reference sheets start
            Division = lastCell(filenameReference, "Division")
            Insurer = lastCell(filenameReference, "Insurer")
            PolClass = lastCell(filenameReference, "PolClass")
            InvoiceClass = lastCell(filenameReference, "InvoiceClass")
            'get the last cell for reference sheets end

        Else
            MsgBox("Please select reference File")
            Exit Sub
        End If

        xlWorkBook = xlApp.Workbooks.Open(filenameRaw)
        xlWorkSheet = xlWorkBook.Worksheets(1)

        Dim xllastcell As String
        Dim splt As String()

        xllastcell = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address
        splt = xllastcell.Split("$")

        'Add new Workbook
        newWorkBook = xlApp.Workbooks.Add
        newWorkSheet = newWorkBook.Worksheets("Sheet1")

        xlWorkSheet.Copy(newWorkSheet)
        'xlApp.Visible = True

        newWorkSheet = newWorkBook.Worksheets(1)
        newWorkSheet.Activate()

        With newWorkSheet
            'devicion
            .Columns("E:E").Insert(Microsoft.Office.Interop.Excel.Constants.xlLeft)
            .Cells(1, 5).value = "Division"
            .Cells(2, 5).value = "=VLOOKUP(D2,'" + referencePath + "Division'!$A$1:" + Division + ",2,)"
            .Cells(2, 5).Copy()
            .Range("E3:E" + splt(2)).PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats)
            'Creditor Name
            .Columns("I:I").Insert(Microsoft.Office.Interop.Excel.Constants.xlLeft)
            .Cells(1, 9).value = "Creditor Name"
            .Cells(2, 9).value = "=VLOOKUP(H2,'" + referencePath + "Insurer'!$A$1:" + Insurer + ",2)"
            .Cells(2, 9).Copy()
            .Range("I3:I" + splt(2)).PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats)
            'Policy Class Name 
            .Columns("N:N").Insert(Microsoft.Office.Interop.Excel.Constants.xlLeft)
            .Cells(1, 14).value = "Policy Class Name"
            .Cells(2, 14).value = "=CONCATENATE(M2,""-"",VLOOKUP(M2,'" + referencePath + "PolClass'!$A$2:" + PolClass + ",2))"
            .Cells(2, 14).Copy()
            .Range("N3:N" + splt(2)).PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats)
            'Howden 
            .Columns("O:O").Insert(Microsoft.Office.Interop.Excel.Constants.xlLeft)
            .Cells(1, 15).value = "Howden Class"
            .Cells(2, 15).value = "=VLOOKUP(M2,'" + referencePath + "PolClass'!$A$2:" + PolClass + ",2)"
            .Cells(2, 15).Copy()
            .Range("O3:O" + splt(2)).PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats)
            'Invoice Class
            .Columns("AT:AT").Insert(Microsoft.Office.Interop.Excel.Constants.xlLeft)
            .Cells(1, 46).value = "Invoice Class"
            .Cells(2, 46).value = "=VLOOKUP(AS2,'" + referencePath + "InvoiceClass'!$A$2:" + InvoiceClass + ",2)"
            .Cells(2, 46).Copy()
            .Range("AT3:AT" + splt(2)).PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats)



        End With

        xlWorkBook.Close()

        'start for pivot part
        CreatePivot(xlApp, newWorkBook, newWorkSheet)

    End Sub


#Region "Create Pivot"
    'create pivot
    Private Sub CreatePivot(ByVal xlApp As Excel.Application, ByVal xlWorkBook As Excel.Workbook, ByVal xlWorkSheet As Excel.Worksheet)
        Dim xlWorkBookPivot As Excel.Workbook

        Dim xllastcell As String
        xllastcell = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address

        xlWorkBookPivot = xlApp.Workbooks.Add()

        With xlWorkBookPivot

            'a.	Per Line Net Comm
            .Sheets.Add().name = "Per Line Net Comm"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Invoice Class").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'row
            .ActiveSheet.PivotTables(1).PivotFields("Howden Class").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'value
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With

            'b.	Mgt Report 2
            .Sheets.Add().name = "Mgt Report 2"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Executive").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Invoice Class").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Client Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With

            'b.	Mgt Report 1
            .Sheets.Add().name = "Mgt Report 1"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Executive").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With

            'd.	AE New-Renew
            .Sheets.Add().name = "AE New-Renew"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Executive").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Premium")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Premium]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            'column
            .ActiveSheet.PivotTables(1).PivotFields("Data").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            'e.	Howden Class
            .Sheets.Add().name = "Howden Class"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Invoice Class").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Howden Class").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Premium")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Premium]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            'column
            .ActiveSheet.PivotTables(1).PivotFields("Data").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            'f.	Invoice Class
            .Sheets.Add().name = "Invoice Class"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Invoice Class").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Client Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Premium")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Premium]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            'column
            .ActiveSheet.PivotTables(1).PivotFields("Data").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            'g.	Top Client
            .Sheets.Add().name = "Top Client"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Client Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Premium")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Premium]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            'column
            .ActiveSheet.PivotTables(1).PivotFields("Data").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            'h.	Top Insurer
            .Sheets.Add().name = "Top Insurer"
            .ActiveSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlWorkSheet.Range("A1:" & xllastcell))
            'Filter
            .ActiveSheet.PivotTables(1).PivotFields("Financial Period").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .ActiveSheet.PivotTables(1).PivotFields("Division").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'Row
            .ActiveSheet.PivotTables(1).PivotFields("Creditor Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            'Value
            With .ActiveSheet.PivotTables(1).PivotFields("Premium")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Premium]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            With .ActiveSheet.PivotTables(1).PivotFields("Nett Commission")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Caption = "[Sum of Nett Commission]"
                .Function = Excel.XlConsolidationFunction.xlSum
            End With
            'column
            .ActiveSheet.PivotTables(1).PivotFields("Data").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

        End With
        xlWorkBook.Close(False)
        xlApp.Visible = True


    End Sub
#End Region

    Private Function lastCell(ByVal fileNameTable As String, ByVal sheetName As String) As String
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim ret As String

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileNameTable)
        xlWorkSheet = xlWorkBook.Worksheets(sheetName)


        ret = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address
        Return ret
        xlApp.Quit()
    End Function


End Class

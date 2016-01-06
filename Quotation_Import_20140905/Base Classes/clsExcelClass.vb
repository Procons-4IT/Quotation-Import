Imports Microsoft.Office.Interop.Excel

Public Class clsExcelClass
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Function getExcelTemplate(ByVal strPath As String, ByRef dtExcel As System.Data.DataTable) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim dr As DataRow
            Dim excel As Application = New Application
            Dim w As Workbook = excel.Workbooks.Open(strPath)

            ' Loop over all sheets.
            For i As Integer = 1 To w.Sheets.Count
                Dim sheet As Worksheet = w.Sheets(i)
                Dim r As Range = sheet.UsedRange
                Dim array(,) As Object = r.Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault)
                If array IsNot Nothing Then
                    Dim strRefQuotation As String
                    Dim strDocEntry As String
                    strRefQuotation = array(3, 11)
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If strRefQuotation.Length > 0 Then
                        oRecordSet.DoQuery("Select DocEntry From OPQT Where DocNum = '" + strRefQuotation + "'")
                        If Not oRecordSet.EoF Then
                            strDocEntry = oRecordSet.Fields.Item(0).Value

                            Dim intURow As Integer = array.GetUpperBound(0)
                            Dim intUCol As Integer = array.GetUpperBound(1)

                            For intRow As Integer = 17 To intURow

                                If Not IsNothing(array(intRow, 2)) Then
                                    dr = dtExcel.NewRow()
                                    If array(intRow, 2) <> "" Then

                                        dr(0) = strRefQuotation  'DocNum
                                        dr(1) = array(intRow, 1) 'LineNum
                                        dr(2) = array(intRow, 2) 'ItemCode
                                        dr(3) = array(intRow, 3) 'Description

                                        If IsNothing(array(intRow, 4)) Then ' Required Qty

                                        Else
                                            dr(4) = array(intRow, 4) 'Best Posssible Date
                                        End If

                                        dr(5) = array(intRow, 5) 'UOM

                                        If IsNothing(array(intRow, 6)) Then

                                        Else
                                            dr(6) = array(intRow, 6) 'Best Posssible Date
                                        End If

                                        dr(7) = array(intRow, 7) 'Curr
                                        dr(8) = array(intRow, 8) 'Terms

                                        If IsNothing(array(intRow, 9)) Then ' Available Qty

                                        Else
                                            dr(9) = array(intRow, 9)
                                        End If

                                        If IsNothing(array(intRow, 10)) Then
                                            'dr(10) = System.DBNull
                                        Else
                                            dr(10) = array(intRow, 10) 'Best Posssible Date
                                        End If

                                        dr(11) = array(intRow, 11) 'Remarks
                                        dr(12) = "Ready for Import" 'Import Remarks

                                        dtExcel.Rows.Add(dr)
                                    End If
                                End If
                            Next
                        Else
                            _retVal = False
                            w.Close()
                            Throw New Exception("Imported Purchase Quotation No : " & strRefQuotation & " does not Exists...")
                        End If
                    End If
                End If
            Next
            w.Close()
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class

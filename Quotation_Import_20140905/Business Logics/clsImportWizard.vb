Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsImportWizard
    Inherits clsBase
    Private strQuery As String
    Private oGrid As SAPbouiCOM.Grid

    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable

    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ImpWiz, frm_ImpWiz)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_ImpWiz
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ImpWiz Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "9" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "17")
                                ElseIf (pVal.ItemUID = "7") Then 'Next
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "17") Then
                                            If oApplication.Utilities.GetExcelData(oForm, "17") Then
                                                loadData(oForm)
                                                oForm.Items.Item("4").Enabled = True
                                                oApplication.Utilities.Message("Quotation Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "3") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 1
                                    oForm.Items.Item("3").Enabled = False
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "4") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Items.Item("3").Enabled = True
                                    oForm.Items.Item("5").Enabled = True
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "5") Then
                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you Want to Import Data to Quotation?", 2, "Yes", "No", "")
                                    If _retVal = 1 Then
                                        oGrid = oForm.Items.Item("8").Specific
                                        Dim strDocNum, strDocEntry As String
                                        If oGrid.Rows.Count > 0 Then
                                            strDocNum = oGrid.DataTable.GetValue("DocNum", 0)
                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            If strDocNum.Length > 0 Then
                                                oRecordSet.DoQuery("Select DocEntry From OPQT Where DocNum = '" + strDocNum + "'")
                                                If Not oRecordSet.EoF Then
                                                    strDocEntry = oRecordSet.Fields.Item(0).Value
                                                    If oApplication.Utilities.importPQData(oForm, strDocEntry) Then
                                                        oApplication.Utilities.Message("Purchase Quotation Updated Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                        oForm.Close()
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "12") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "13") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 3
                                    oForm.Freeze(False)
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ImpWiz Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim sQuery As String
            oForm.DataSources.DataTables.Add("Dt_Import")
            oForm.DataSources.DataTables.Add("Dt_ErrorLog")

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            sQuery = " Select DocNum,LineNum,ItemCode,Dscription,Price,OpenQty,UnitMsr,DocCur,Terms,Quantity,ShipDate,FreeTxt From Z_PQIM "
            sQuery += "  Where 1 = 2 "
            oDt_Import.ExecuteQuery(sQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import
            formatGrid(oForm)

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            oDt_ErrorLog.ExecuteQuery("Select Convert(VarChar(250),'') As 'Error'")
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_ErrorLog

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("8").Specific
            formatAll(oForm, oGrid)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadData(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select T0.DocNum,T0.LineNum,T0.ItemCode,T0.Dscription,ISNULL(T0.OpenQty,0) As OpenQty,T0.UnitMsr,ISNULL(T0.Price,0) As Price,T0.DocCur,T0.Terms,ISNULL(T0.Quantity,0) As Quantity,T0.ShipDate,T0.FreeTxt From Z_PQIM T0 "
            strQuery += " JOIN OPQT T1 On T0.DocNum = T1.DocNum "
            strQuery += " JOIN PQT1 T2 On T1.DocEntry = T2.DocEntry "
            strQuery += " AND (T0.LineNum) = (T2.VisOrder+1) "
            strQuery += " And T2.LineStatus = 'O' "
            oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import
            formatAll(oForm, oGrid)

            For index As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            strQuery = " Select 'InValid Line No : ' + T0.LineNum As 'Error' From Z_PQIM T0 "
            strQuery += " JOIN OPQT T1 On T0.DocNum = T1.DocNum "
            strQuery += " LEFT OUTER JOIN PQT1 T2 On T1.DocEntry = T2.DocEntry "
            strQuery += " And T0.LineNum = (T2.VisOrder+1) "
            strQuery += " Where T2.VisOrder Is Null "
            strQuery += " Union All "
            strQuery += " Select 'Row Status Closed Row No : ' + T0.LineNum As 'Error' From Z_PQIM T0 "
            strQuery += " JOIN OPQT T1 On T0.DocNum = T1.DocNum "
            strQuery += " JOIN PQT1 T2 On T1.DocEntry = T2.DocEntry "
            strQuery += " And T0.LineNum = (T2.VisOrder+1) "
            strQuery += " And T2.LineStatus = 'C' "

            oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_ErrorLog
            For index As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid)
        Try
            oForm.Freeze(True)
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Document No"
            oGrid.Columns.Item("LineNum").TitleObject.Caption = "Line"
            oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Stock Code"
            oGrid.Columns.Item("Dscription").TitleObject.Caption = "Item Desc"
            oGrid.Columns.Item("OpenQty").TitleObject.Caption = "Requiered Quantity"
            oGrid.Columns.Item("UnitMsr").TitleObject.Caption = " Unit of measure"
            oGrid.Columns.Item("Price").TitleObject.Caption = "Price"
            oGrid.Columns.Item("DocCur").TitleObject.Caption = "Currency"
            oGrid.Columns.Item("Terms").TitleObject.Caption = "INCO terms "
            oGrid.Columns.Item("Quantity").TitleObject.Caption = "Available Qty"
            oGrid.Columns.Item("ShipDate").TitleObject.Caption = "Best Possible Date"
            oGrid.Columns.Item("FreeTxt").TitleObject.Caption = "Remarks"
            'oGrid.Columns.Item("Remarks").TitleObject.Caption = "Import Remarks"

            oGrid.Columns.Item("ItemCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("ItemCode")
            oEditColumn.LinkedObjectType = "4"

            'Currency 
            oGrid.Columns.Item("DocCur").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComboColumn = oGrid.Columns.Item("DocCur")
            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oComboColumn.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("OpenQty").RightJustified = True
            oGrid.Columns.Item("Quantity").RightJustified = True
            oGrid.Columns.Item("Price").RightJustified = True

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

End Class

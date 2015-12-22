﻿Public Class clshrLoginSetup
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1 As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oOption, oOption1 As SAPbouiCOM.OptionBtn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_hr_Logsetup, frm_hr_LoginSetup)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "26"
        oForm.EnableMenu(mnu_ADD, True)
        oForm.EnableMenu(mnu_FIND, True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Freeze(False)
    End Sub
    Private Sub AddMode(ByVal aform As SAPbouiCOM.Form)
        Dim strCode As String
        strCode = oApplication.Utilities.getMaxCode("@Z_HR_LOGIN", "DocEntry")
        aform.Items.Item("26").Enabled = True
        oApplication.Utilities.setEdittextvalue(aform, "26", strCode)
        aform.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        aform.Items.Item("26").Enabled = False
        aform.Items.Item("12").Enabled = True
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oRec As SAPbobsCOM.Recordset
        Dim strLoginPassword, strSAPPassword As String
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oApplication.Utilities.getEdittextvalue(aForm, "12") = "" Then
                oApplication.Utilities.Message("UserId missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "14") = "" Then
                oApplication.Utilities.Message("Password missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            Else
                strLoginPassword = oApplication.Utilities.getEdittextvalue(aForm, "14")
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Employee ID missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "8") <> "" Then
                If oApplication.Utilities.getEdittextvalue(aForm, "10") = "" Then
                    oApplication.Utilities.Message("SAP Password missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Return False
                Else
                    strSAPPassword = oApplication.Utilities.getEdittextvalue(aForm, "10")
                End If
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRec.DoQuery("Select * from [@Z_HR_LOGIN] where U_Z_EMPID='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "'")
            Else
                oRec.DoQuery("Select * from [@Z_HR_LOGIN] where U_Z_EMPID='" & oApplication.Utilities.getEdittextvalue(aForm, "4") & "' and DocEntry <> '" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "'")
            End If
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.Message("Record already exists for this employee...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRec.DoQuery("Select * from [@Z_HR_LOGIN] where Upper(U_Z_UID)='" & oApplication.Utilities.getEdittextvalue(aForm, "12").ToUpper() & "'")
            Else
                oRec.DoQuery("Select * from [@Z_HR_LOGIN] where Upper(U_Z_UID)='" & oApplication.Utilities.getEdittextvalue(aForm, "12").ToUpper() & "' and DocEntry <> '" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "'")
            End If
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.Message("Record already exists for this ESS User...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim strEncryptText As String = oApplication.Utilities.Encrypt(strLoginPassword, oApplication.Utilities.key)
            oApplication.Utilities.setEdittextvalue(aForm, "14", strEncryptText) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

            Dim strEncryptText1 As String = oApplication.Utilities.Encrypt(strSAPPassword, oApplication.Utilities.key)
            oApplication.Utilities.setEdittextvalue(aForm, "10", strEncryptText1) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_hr_LoginSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oRec As SAPbobsCOM.Recordset
                                Dim val1, val2 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If IsNothing(oDataTable) Then
                                            Exit Sub
                                        End If
                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val2 = oDataTable.GetValue("userId", 0)
                                            val1 = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("middleName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val2)
                                            oRec.DoQuery("Select isnull(USER_CODE,'') from OUSR where INTERNAL_K='" & val2 & "'")
                                            If oRec.RecordCount > 0 Then
                                                Try
                                                    oApplication.Utilities.setEdittextvalue(oForm, "8", oRec.Fields.Item(0).Value)
                                                Catch ex As Exception
                                                End Try
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        End If
                                        If pVal.ItemUID = "8" Then
                                            val = oDataTable.GetValue("USER_CODE", 0)
                                            val1 = oDataTable.GetValue("INTERNAL_K", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "8", val)
                                        End If
                                      
                                        oForm.Freeze(False)
                                    End If

                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_hr_Logsetup
                    LoadForm()
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_hr_LoginSetup Then
                        If pVal.BeforeAction = False Then
                            AddMode(oForm)
                        End If

                    End If

                Case mnu_InvSO
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm.Items.Item("12").Enabled = False
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.TypeEx = frm_hr_LoginSetup Then



                    Dim strEncryptText As String = oApplication.Utilities.getLoginPassword(oApplication.Utilities.getEdittextvalue(oForm, "14"))
                    oApplication.Utilities.setEdittextvalue(oForm, "14", strEncryptText) ' oApplication.Utilities.getEdittextvalue(aForm, "8")

                    Dim strEncryptText1 As String = oApplication.Utilities.getLoginPassword(oApplication.Utilities.getEdittextvalue(oForm, "10"))
                    oApplication.Utilities.setEdittextvalue(oForm, "10", strEncryptText1) ' oApplication.Utilities.getEdittextvalue(aForm, "8")
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class

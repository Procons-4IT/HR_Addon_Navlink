Imports System.IO
Imports System.Net.Mail

Public Class clsActivity
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText, oEditFDate, oEditTDate As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1, oCombobox2, oCombobox3, oComboLevel As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboBoxcolumn As SAPbouiCOM.ComboBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private oDtAppraisal As SAPbouiCOM.DataTable
    Private oRecordSet As SAPbobsCOM.Recordset
    Private sQuery As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_hr_AppEmail) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        oForm = oApplication.Utilities.LoadForm(xml_hr_AppraisalEmail, frm_hr_AppEmail)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        FillDepartment(oForm)
        FillPosition(oForm)
        FillPeriod(oForm)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("Desc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "36", "Desc")
        oForm.DataSources.UserDataSources.Add("empno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "20", "empno")
        oEditText = oForm.Items.Item("20").Specific

        'oEditFDate = oForm.Items.Item("34").Specific
        'oEditTDate = oForm.Items.Item("36").Specific
        'oForm.DataSources.UserDataSources.Add("fdate", SAPbouiCOM.BoDataType.dt_DATE)
        'oApplication.Utilities.setUserDatabind(oForm, "34", "fdate")
        'oForm.DataSources.UserDataSources.Add("tdate", SAPbouiCOM.BoDataType.dt_DATE)
        'oApplication.Utilities.setUserDatabind(oForm, "36", "tdate")

        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "empId"
        oForm.DataSources.UserDataSources.Add("empno1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "23", "empno1")
        oEditText = oForm.Items.Item("23").Specific
        oEditText.ChooseFromListUID = "CFL2"
        oEditText.ChooseFromListAlias = "empId"
        oForm.PaneLevel = 1

        Dim osta As SAPbouiCOM.StaticText
        osta = oForm.Items.Item("19").Specific
        osta.Caption = "Step " & oForm.PaneLevel & " of 4"
        oForm.Items.Item("19").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
        InitializeAppTable()
        oForm.Freeze(False)
    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCombobox = objForm.Items.Item("7").Specific
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "171"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "171"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FillDepartment(ByVal sform As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCombobox = sform.Items.Item("1000007").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        oTempRec.DoQuery("Select Code,Remarks from OUDP")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            Try
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("Code").Value, oTempRec.Fields.Item("Remarks").Value)
            Catch ex As Exception

            End Try

            oTempRec.MoveNext()
        Next
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        sform.Items.Item("1000007").DisplayDesc = True
    End Sub

    Private Sub FillPeriod(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCombobox = aForm.Items.Item("25").Specific
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        'oTempRec.DoQuery("Select Code,Name from OFPR order by Code desc")
        oTempRec.DoQuery("Select ""U_Z_PerCode"" as ""Code"",""U_Z_PerDesc"" AS ""Name"" from ""@Z_HR_PERAPP"" order by Code desc")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            oCombobox.ValidValues.Add(oTempRec.Fields.Item("Code").Value, oTempRec.Fields.Item("Name").Value)
            oTempRec.MoveNext()
        Next
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        aForm.Items.Item("25").DisplayDesc = True
    End Sub

    Private Sub FillPosition(ByVal sform As SAPbouiCOM.Form)
        oCombobox = sform.Items.Item("1000008").Specific
        oCombobox1 = sform.Items.Item("29").Specific
        Dim oSlpRS As SAPbobsCOM.Recordset
        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        For intRow As Integer = oCombobox1.ValidValues.Count - 1 To 0 Step -1

            Try
                oCombobox1.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try

        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox1.ValidValues.Add("", "")
        oSlpRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oSlpRS.DoQuery("Select U_Z_PosCode,U_Z_PosName from [@Z_HR_OPOSIN] order by DocEntry")
        For intRow As Integer = 0 To oSlpRS.RecordCount - 1
            Try
                oCombobox.ValidValues.Add(oSlpRS.Fields.Item(0).Value, oSlpRS.Fields.Item(1).Value)
            Catch ex As Exception

            End Try
            Try
                oCombobox1.ValidValues.Add(oSlpRS.Fields.Item(0).Value, oSlpRS.Fields.Item(1).Value)
            Catch ex As Exception

            End Try

            oSlpRS.MoveNext()
        Next
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oCombobox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        sform.Items.Item("1000008").DisplayDesc = True
        sform.Items.Item("29").DisplayDesc = True
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strDept, strPeriod As String

            oCombobox = aForm.Items.Item("1000007").Specific

            strDept = oCombobox.Selected.Description
            oCombobox1 = aForm.Items.Item("25").Specific
            strPeriod = oCombobox1.Selected.Value

            If strPeriod = "" Then
                oApplication.Utilities.Message("Enter Period...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Function
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Dim strFromEMP, strToEMP, strDept, strFromPos, strToPos, strPeriod, strqry As String
        Dim strEMPCondition As String = ""
        Dim strPositionCondition As String = ""
        Dim strdeptcondition As String = ""
        Dim strPeriodcondition As String = ""
        Dim strDatecondition As String = ""
        Dim strFDate As String = ""
        Dim strTDate As String = ""
        Dim strLevelStartFrom As String = ""

        oCombobox = aform.Items.Item("1000007").Specific
        oCombobox1 = aform.Items.Item("1000008").Specific

        oCombobox2 = aform.Items.Item("29").Specific
        oCombobox3 = aform.Items.Item("25").Specific
        strFromEMP = oApplication.Utilities.getEdittextvalue(aform, "20")
        strToEMP = oApplication.Utilities.getEdittextvalue(aform, "23")

        strDept = oCombobox.Selected.Value
        strFromPos = oCombobox1.Selected.Value
        strToPos = oCombobox2.Selected.Value
        strPeriod = oCombobox3.Selected.Value


        If strFromEMP <> "" And strToEMP <> "" Then
            strEMPCondition = " Convert(Decimal,T1.empID) between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)
        ElseIf strFromEMP <> "" And strToEMP = "" Then
            strEMPCondition = " Convert(Decimal,T1.empID) >= " & CDbl(strFromEMP)
        ElseIf strFromEMP = "" And strToEMP <> "" Then
            strEMPCondition = " Convert(Decimal,T1.empID) <= " & CDbl(strToEMP)
        Else
            strEMPCondition = " 1=1"
        End If

        If strDept <> "" Then
            strdeptcondition = " Convert(Decimal,T1.dept) = " & CDbl(strDept)
        Else
            strdeptcondition = " 1=1"
        End If

        If strPeriod <> "" Then
            strPeriodcondition = "T0.U_Z_Period = '" & strPeriod & "'"
        Else
            strPeriodcondition = " 1=1"
        End If

        If strFromPos <> "" And strToPos <> "" Then
            strPositionCondition = "T1.U_Z_HR_PosiCode between '" & strFromPos & "' and '" & strToPos & "'"
        ElseIf strFromPos <> "" And strToPos = "" Then
            strPositionCondition = "T1.U_Z_HR_PosiCode >= '" & strFromPos & "'"
        ElseIf strFromPos = "" And strToPos <> "" Then
            strPositionCondition = "T1.U_Z_HR_PosiCode <= '" & strToPos & "'"
        Else
            strPositionCondition = " 1 = 1"
        End If


        Dim strcondition, strqry1 As String


        Dim strPeriod1 As String
        oCombobox = aform.Items.Item("25").Specific
        strPeriod1 = oCombobox.Selected.Description
        strcondition = strEMPCondition & " and " & strdeptcondition & " and " & strPeriodcondition & " and " & strPositionCondition & "  Order by T1.empID"
        strqry1 = "Select  'Y' as 'Select',T0.U_Z_EmpId,T0.DocEntry,T1.firstName,T1.lastName,T1.email as 'ccID',T2.email as 'toID' from [@Z_HR_OSEAPP] T0 JOIN OHEM T1 ON T0.U_Z_EmpID = T1.EmpID JOIN OHEM T2 On T1.Manager = T2.EmpID where " & strcondition

        'strqry = "select 'Y' as 'Select', empID,firstName,lastName,email ,T1.Name as 'Department',U_Z_HR_PosiCode, '" & strPeriod1 & "' 'Period',  U_Z_HR_PosiName,'" & strLevelStartGrid & "' as 'Level Start From'  from OHEM T0 inner join OUDP T1 "
        'strqry = strqry & " on T0.dept=T1.Code where empID not in (" & strqry1 & ") and  " & strcondition

        oGrid = aform.Items.Item("10").Specific
        oGrid.DataTable.ExecuteQuery(strqry1)
        Try
            aform.Freeze(True)
            FormatGrid(aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub FormatGrid(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("10").Specific
        oGrid.Columns.Item("Select").TitleObject.Caption = "Select"
        oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGrid.Columns.Item("Select").Editable = True
        oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee Id"
        oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
        oEditTextColumn.LinkedObjectType = "171"
        oGrid.Columns.Item("U_Z_EmpId").Editable = False
        oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Appraisal No"
        oGrid.Columns.Item("DocEntry").Editable = False
        oGrid.Columns.Item("firstName").TitleObject.Caption = "First Name"
        oGrid.Columns.Item("firstName").Editable = False
        oGrid.Columns.Item("lastName").TitleObject.Caption = "Last Name"
        oGrid.Columns.Item("lastName").Editable = False
        oGrid.Columns.Item("toID").TitleObject.Caption = "Manager Email ID"
        oGrid.Columns.Item("toID").Editable = False
        oGrid.Columns.Item("ccID").TitleObject.Caption = "Employee Email ID"
        oGrid.Columns.Item("ccID").Editable = False
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub

    Private Sub InitializeAppTable()
        oForm.DataSources.DataTables.Add("dtActivity")
        oDtAppraisal = oForm.DataSources.DataTables.Item("dtActivity")
        oDtAppraisal.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("EmpID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("toID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("ccID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("Type", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        oDtAppraisal.Columns.Add("Path", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
    End Sub

    Private Sub SendMail(aDocEntry As String)
        If oApplication.Utilities.checkmailconfiguration() = False Then
            oApplication.Utilities.Message("Email configuration not availble...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        Dim ore, ore1 As SAPbobsCOM.Recordset
        Dim strFilename As String
        Dim mailServer As String
        Dim mailPort As String
        Dim mailId As String
        Dim mailUser As String
        Dim mailPwd As String
        Dim mailSSL As String
        Dim toID As String
        Dim ccID As String
        Dim mType As String
        Dim path As String
        Dim sQuery As String
        Dim strEmpName As String
        ore = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ore1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ore.DoQuery("Select * from OCLG where clgCode=" & aDocEntry)
        If ore.RecordCount > 0 Then
            If ore.Fields.Item("AtcEntry").Value > 0 Then
                ore1.DoQuery("Select * from ATC1 where AbsEntry=" & ore.Fields.Item("AtcEntry").Value)
                For intRow As Integer = 0 To ore1.RecordCount - 1
                    strFilename = ""
                    strFilename = ore1.Fields.Item("trgtPath").Value & "\" & ore1.Fields.Item("FileName").Value & "." & ore1.Fields.Item("FileExt").Value
                    path = strFilename
                    ore1.MoveNext()
                Next
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("Select U_Z_SMTPSERV,U_Z_SMTPPORT,U_Z_SMTPUSER,U_Z_SMTPPWD,U_Z_SSL From [@Z_HR_OMAIL]")
                If Not oRecordSet.EoF Then
                    mailServer = oRecordSet.Fields.Item("U_Z_SMTPSERV").Value
                    mailPort = oRecordSet.Fields.Item("U_Z_SMTPPORT").Value
                    mailId = oRecordSet.Fields.Item("U_Z_SMTPUSER").Value
                    mailPwd = oRecordSet.Fields.Item("U_Z_SMTPPWD").Value
                    mailSSL = oRecordSet.Fields.Item("U_Z_SSL").Value
                    mType = "AC"
                    SendMailforUsers(mailServer, mailPort, mailId, mailPwd, mailSSL, toID, ccID, mType, path, aDocEntry, aDocEntry)
                Else
                    oApplication.Utilities.Message("Mail Server Details Not Configured...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            End If
        End If

        'If Not IsNothing(oDtAppraisal) And oDtAppraisal.Rows.Count > 0 Then
        '    oApplication.SBO_Application.StatusBar.SetText("Generating Report Started....", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    '  oApplication.Utilities.generateReport(oDtAppraisal)
        '    oApplication.SBO_Application.StatusBar.SetText("Process Sending Mail....", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    oApplication.Utilities.SendMail(oDtAppraisal, "Activity")
        '    oApplication.SBO_Application.StatusBar.SetText("Mail Process Completed Sucessfully....", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        'End If


    End Sub


    Private Sub SendMailforUsers(ByVal mailServer As String, ByVal mailPort As String, ByVal mailId As String, ByVal mailpwd As String, ByVal mailSSL As String, ByVal toId As String, ByVal ccId As String, ByVal mType As String, ByVal path As String, ByVal DocEntry As String, ByVal Name As String, Optional ByVal Period As String = "")
        Dim mail As New Net.Mail.MailMessage
        Try
            'Dim strRptPath As String = System.Windows.Forms.Application.StartupPath.Trim() & "\Report.pdf"
            Dim strMessage, strQuery As String
            Dim SmtpServer As New Net.Mail.SmtpClient()

            Dim sQuery As String
            Dim strEmpName As String
            Dim oTest, oTemp As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            SmtpServer.Credentials = New Net.NetworkCredential(mailId, mailpwd)
            SmtpServer.Port = mailPort
            SmtpServer.EnableSsl = mailSSL
            SmtpServer.Host = mailServer
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(mailId, "HRMS")
            mail.IsBodyHtml = True
            mail.Priority = MailPriority.High
            If mType = "AI" Then
            ElseIf mType = "AC" Then 'Activity
                strQuery = "SELECT T1.email,isnull(T1.firstName,'') +' '+ isnull(T1.lastName,'') as 'EmpName',T1.userId from OCLG T0 JOIN OHEM T1 ON T0.U_Z_HREmpID=T1.empID where T0.ClgCode='" & DocEntry & "'"
                oTemp.DoQuery(strQuery)
                If oTemp.RecordCount > 0 Then
                    mail.To.Add(oTemp.Fields.Item(0).Value)
                    oTest.DoQuery("Select * from [@Z_HR_OWEB]")
                    Dim strESSLink As String = ""
                    If oTest.RecordCount > 0 Then
                        strESSLink = oTest.Fields.Item("U_Z_WebPath").Value
                    End If
                    strMessage = "Appraisal Document No : " & DocEntry & ", Employee Name : " & oTemp.Fields.Item("EmpName").Value ' & "," & strPeriod & ""
                    mail.Subject = "Document From HR"
                    mail.Body = BuildHtmBody(DocEntry)
                    mail.Attachments.Add(New Net.Mail.Attachment(path))
                End If
            End If

            SmtpServer.Send(mail)

        Catch ex As Exception

        Finally
            mail.Dispose()
        End Try
    End Sub
    Private Function BuildHtmBody(aDocEntry As Integer) As String

        Dim AMessage, strQuery As String
        Dim OTemp As SAPbobsCOM.Recordset
        OTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQuery = "SELECT T1.email,isnull(T1.firstName,'') +' '+ isnull(T1.lastName,'') as 'EmpName',T1.userId from OCLG T0 JOIN OHEM T1 ON T0.U_Z_HREmpID=T1.empID where T0.ClgCode='" & aDocEntry & "'"
        oTemp.DoQuery(strQuery)
        AMessage = "<table width-'100%'>"
        AMessage += "<tr><td>Dear : " & OTemp.Fields.Item("EmpName").Value & "</td><td>"
        AMessage += "</td></tr><tr><td></td></tr> "
        AMessage += "<tr><td>Please find attached your requested Document</td><td> <tr><td></td></tr> "
        AMessage += "<tr><td>Best Regards</td><td>"
        AMessage += "</td></tr>"

        strQuery = "SELECT T1.email,isnull(T1.firstName,'') +' '+ isnull(T1.lastName,'') as 'EmpName',T1.userId from OCLG T0 JOIN OHEM T1 ON T0.U_Z_AssEmpID=T1.empID where T0.ClgCode='" & aDocEntry & "'"
        OTemp.DoQuery(strQuery)
        AMessage += "<tr><td> " & OTemp.Fields.Item("EmpName").Value & " </td><td>"
        AMessage += "</td></tr>"
        AMessage += "</table>"
        Return AMessage
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = "1234" Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2 Or oForm.PaneLevel = 3) Then
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
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "25" Then
                                    oCombobox = oForm.Items.Item("25").Specific
                                    Dim strdesc As String = oCombobox.Selected.Description
                                    oApplication.Utilities.setEdittextvalue(oForm, "36", oCombobox.Selected.Description)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 3 Then
                                            Databind(oForm)
                                        End If
                                        Dim osta As SAPbouiCOM.StaticText
                                        osta = oForm.Items.Item("19").Specific
                                        osta.Caption = "Step " & oForm.PaneLevel & " of 4"
                                        oForm.Items.Item("19").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                    Case "4"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        Dim osta As SAPbouiCOM.StaticText
                                        osta = oForm.Items.Item("19").Specific
                                        osta.Caption = "Step " & oForm.PaneLevel & " of 4"
                                        oForm.Items.Item("19").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                    Case "5"
                                        oGrid = oForm.Items.Item("10").Specific
                                        'SendMail(oForm)
                                        oCombobox3 = oForm.Items.Item("25").Specific
                                        Dim ostatic As SAPbouiCOM.StaticText
                                        ostatic = oForm.Items.Item("30").Specific
                                        ostatic.Caption = "The Appraisal Mail Generated Sucessfully...."
                                        oForm.Items.Item("30").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                        ostatic = oForm.Items.Item("31").Specific
                                        ostatic.Caption = "Perid : " & oCombobox3.Selected.Value & "(" & oCombobox.Selected.Description & ")"
                                        oForm.Items.Item("31").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                        ostatic = oForm.Items.Item("32").Specific
                                        ostatic.Caption = "Number of Employee : " & oGrid.DataTable.Rows.Count
                                        oForm.Items.Item("32").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                        Dim osta As SAPbouiCOM.StaticText
                                        osta = oForm.Items.Item("19").Specific
                                        oForm.Items.Item("19").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
                                        oForm.PaneLevel = 4
                                        osta.Caption = "Step " & oForm.PaneLevel & " of 4"
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2, val3 As String
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
                                        intChoice = 0
                                        oForm.Freeze(True)

                                        If pVal.ItemUID = "20" Then
                                            val1 = oDataTable.GetValue("empID", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "20", val1)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                        If pVal.ItemUID = "23" Then
                                            val1 = oDataTable.GetValue("empID", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "23", val1)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    ' oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                'Case mnu_hr_AppraisalEmail
                '    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.TypeEx = frm_Activity Then
                    Dim oAct As SAPbobsCOM.Contacts
                    oAct = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)
                    If oAct.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oAct.Status = -3 Then
                            SendMail(oAct.ContactCode.ToString)
                        End If
                    End If

                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class

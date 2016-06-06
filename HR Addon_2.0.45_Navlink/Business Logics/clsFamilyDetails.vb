Public Class clsFamilyDetails1
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oComboboxColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal aCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_OFMD1, frm_LEB_OFMD1)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataSources.UserDataSources.Add("Name1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oEditText = oForm.Items.Item("9").Specific
        oEditText.DataBind.SetBound(True, "", "Name1")
        oApplication.Utilities.setEdittextvalue(oForm, "7", aCode)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select isnull(firstName,'') + ' ' + isnull(middleName,'') +' ' + isnull(lastName,'') from OHEM where empid=" & CInt(aCode))
        oApplication.Utilities.setEdittextvalue(oForm, "9", oTemp.Fields.Item(0).Value)
        Databind(oForm, aCode)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
    End Sub

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form, ByVal intCode As Integer)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            ' dtTemp.ExecuteQuery("Select * from [@Z_PAY_EMPFAMILY] where U_Z_EmpID='" & intCode & "' order by Code")
            dtTemp.ExecuteQuery("SELECT T0.[Code], T0.[Name], T0.[U_Z_EmpID], T0.[U_Z_MemCode], T0.[U_Z_MemName], T0.[U_Z_Gender] , T0.[U_Z_DOB], T0.[U_Z_DOM], T0.[U_Z_STUD], T0.[U_Z_Emp], T0.[U_Z_Married], T0.[U_Z_DOJ], T0.[U_Z_DOT],T0.[U_Z_MRC] , T0.[U_Z_BCR],T0.[U_Z_INS],T0.[U_Z_NSSF],T0.[U_Z_StopAllowance] FROM [dbo].[@Z_PAY_EMPFAMILY]  T0  where U_Z_EmpID='" & intCode & "' order by Code")

            oGrid.DataTable = dtTemp
            Formatgrid(oGrid)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)

        agrid.Columns.Item("Code").Visible = False
        agrid.Columns.Item("Name").Visible = False
        agrid.Columns.Item("U_Z_EmpID").Visible = False
        agrid.Columns.Item("U_Z_MemCode").TitleObject.Caption = "Member Type"
        agrid.Columns.Item("U_Z_MemCode").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox

        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select U_Z_Code,U_Z_Name from [@Z_PAY_OFAM] order by convert(numeric,code)")

        oComboboxColumn = agrid.Columns.Item("U_Z_MemCode")
        ' oComboboxColumn.ValidValues.Add("", "")
        For intRow As Integer = oComboboxColumn.ValidValues.Count - 1 To 0 Step -1
            oComboboxColumn.ValidValues.Remove(intRow)
        Next

        For intRow As Integer = 0 To oTemp.RecordCount - 1
            oComboboxColumn.ValidValues.Add(oTemp.Fields.Item(0).Value, oTemp.Fields.Item(1).Value)
            oTemp.MoveNext()
        Next
        oComboboxColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("U_Z_MemName").TitleObject.Caption = "Member Name"
        agrid.Columns.Item("U_Z_Gender").TitleObject.Caption = "Gender"
        agrid.Columns.Item("U_Z_Gender").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oComboboxColumn = agrid.Columns.Item(5)
        ' oComboboxColumn.ValidValues.Add("", "")
        For intRow As Integer = oComboboxColumn.ValidValues.Count - 1 To 0 Step -1
            oComboboxColumn.ValidValues.Remove(intRow)
        Next
        oComboboxColumn.ValidValues.Add("B", "Male")
        oComboboxColumn.ValidValues.Add("G", "Female")
        oComboboxColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both


        'dtTemp.ExecuteQuery(" T0.[U_Z_DOB], T0.[U_Z_DOM], T0.[U_Z_STUD], T0.[U_Z_Emp], T0.[U_Z_Married], T0.[U_Z_DOJ], T0.[U_Z_DOT],T0.[U_Z_NSSF],T0.[U_Z_StopAllowance] FROM [dbo].[@Z_PAY_EMPFAMILY]  T0  where U_Z_EmpID='" & intCode & "' order by Code")

        agrid.Columns.Item("U_Z_DOB").TitleObject.Caption = "Date of Birth"
        agrid.Columns.Item("U_Z_DOM").TitleObject.Caption = "Date of Marriage"
        agrid.Columns.Item("U_Z_STUD").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_STUD").TitleObject.Caption = "Is Student"

        agrid.Columns.Item("U_Z_Emp").TitleObject.Caption = "Employement Status"

        agrid.Columns.Item("U_Z_Emp").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oComboboxColumn = agrid.Columns.Item(9)
        For intRow As Integer = oComboboxColumn.ValidValues.Count - 1 To 0 Step -1
            oComboboxColumn.ValidValues.Remove(intRow)
        Next
        oComboboxColumn.ValidValues.Add("Y", "Employed")
        oComboboxColumn.ValidValues.Add("N", "Not Employed")
        oComboboxColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("U_Z_Emp").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox



        agrid.Columns.Item("U_Z_Married").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_Married").TitleObject.Caption = "Is Married"


        agrid.Columns.Item("U_Z_DOJ").TitleObject.Caption = "Date of Joining"
        agrid.Columns.Item("U_Z_DOT").TitleObject.Caption = "Date of Resignation"
        '  agrid.Columns.Item(13).TitleObject.Caption = "NSSF Declaration "
        agrid.Columns.Item("U_Z_NSSF").TitleObject.Caption = "NSSF Declaration"
        agrid.Columns.Item("U_Z_NSSF").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_StopAllowance").TitleObject.Caption = "Stop Allowance"
        agrid.Columns.Item("U_Z_StopAllowance").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

        agrid.Columns.Item("U_Z_MRC").TitleObject.Caption = "Marriage Certificate Received"
        agrid.Columns.Item("U_Z_MRC").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

        agrid.Columns.Item("U_Z_BCR").TitleObject.Caption = "Birth Certificate Received"
        agrid.Columns.Item("U_Z_BCR").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

        agrid.Columns.Item("U_Z_INS").TitleObject.Caption = "Insurance"
        agrid.Columns.Item("U_Z_INS").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        If aGrid.DataTable.GetValue("U_Z_MemName", aGrid.DataTable.Rows.Count - 1) <> "" Then
            aGrid.DataTable.Rows.Add()
            aGrid.Columns.Item("U_Z_MemCode").Click(aGrid.DataTable.Rows.Count - 1, False)
        End If
    End Sub
#End Region

#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strChoice = "Cancel" Then
            oTemprec.DoQuery("Update [@Z_PAY_EMPFAMILY] set NAME=CODE where Name Like '%DX'")
        Else
            oTemprec.DoQuery("Delete from  [@Z_PAY_EMPFAMILY]  where NAME Like '%DX'")
        End If

    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc, strNSSF As String
        Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
        oGrid = aform.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1

            If oGrid.DataTable.GetValue("U_Z_MemName", intRow) <> "" Then

                strCode = oGrid.DataTable.GetValue(0, intRow)
                strECode = oGrid.DataTable.GetValue(1, intRow)
                oUserTable = oApplication.Company.UserTables.Item("Z_PAY_EMPFAMILY")
                If oUserTable.GetByKey(strCode) = False Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_PAY_EMPFAMILY", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_EmpID").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    oUserTable.UserFields.Fields.Item("U_Z_MemCode").Value = (oGrid.DataTable.GetValue("U_Z_MemCode", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_MemName").Value = (oGrid.DataTable.GetValue("U_Z_MemName", intRow))
                    If oGrid.DataTable.GetValue("U_Z_Gender", intRow) = "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_Gender").Value = "B"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_Gender").Value = oGrid.DataTable.GetValue("U_Z_Gender", intRow)
                    End If


                    oUserTable.UserFields.Fields.Item("U_Z_DOB").Value = (oGrid.DataTable.GetValue("U_Z_DOB", intRow))
                    Dim dtdate As Date
                    dtdate = oGrid.DataTable.GetValue("U_Z_DOM", intRow)
                    If Year(dtdate) <> 1 Then
                        oUserTable.UserFields.Fields.Item("U_Z_DOM").Value = (oGrid.DataTable.GetValue("U_Z_DOM", intRow))
                    End If
                    ' MsgBox(oGrid.DataTable.GetValue(7, intRow))


                    If oGrid.DataTable.GetValue("U_Z_STUD", intRow) = "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_STUD").Value = "N"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_STUD").Value = oGrid.DataTable.GetValue("U_Z_STUD", intRow)

                    End If

                    If oGrid.DataTable.GetValue("U_Z_Emp", intRow) = "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = "N"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = oGrid.DataTable.GetValue("U_Z_Emp", intRow)

                    End If


                    If oGrid.DataTable.GetValue("U_Z_Married", intRow) = "" Then
                        oUserTable.UserFields.Fields.Item("U_Z_Married").Value = "N"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_Married").Value = oGrid.DataTable.GetValue("U_Z_Married", intRow)

                    End If
                    '  oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = (oGrid.DataTable.GetValue(7, intRow))
                    dtdate = oGrid.DataTable.GetValue("U_Z_DOJ", intRow)
                    If Year(dtdate) <> 1 Then
                        oUserTable.UserFields.Fields.Item("U_Z_DOJ").Value = (oGrid.DataTable.GetValue("U_Z_DOJ", intRow))
                    End If
                    dtdate = oGrid.DataTable.GetValue("U_Z_DOT", intRow)
                    ' MsgBox(Year(dtdate))
                    If Year(dtdate) <> 1 Then
                        oUserTable.UserFields.Fields.Item("U_Z_DOT").Value = (oGrid.DataTable.GetValue(12, intRow))
                    End If

                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_NSSF")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        strNSSF = "Y"
                    Else
                        strNSSF = "N"
                    End If
                    oUserTable.UserFields.Fields.Item("U_Z_NSSF").Value = strNSSF

                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_StopAllowance")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        strNSSF = "Y"
                    Else
                        strNSSF = "N"
                    End If

                    oUserTable.UserFields.Fields.Item("U_Z_StopAllowance").Value = strNSSF


                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_MRC")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        strNSSF = "Y"
                    Else
                        strNSSF = "N"
                    End If

                    oUserTable.UserFields.Fields.Item("U_Z_MRC").Value = strNSSF


                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_BCR")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        strNSSF = "Y"
                    Else
                        strNSSF = "N"
                    End If

                    oUserTable.UserFields.Fields.Item("U_Z_BCR").Value = strNSSF

                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_INS")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        strNSSF = "Y"
                    Else
                        strNSSF = "N"
                    End If

                    oUserTable.UserFields.Fields.Item("U_Z_INS").Value = strNSSF

                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Committrans("Cancel")
                        Return False
                    End If

                Else
                    strCode = oGrid.DataTable.GetValue(0, intRow)
                    If oUserTable.GetByKey(strCode) Then
                        oUserTable.Code = strCode
                        oUserTable.Name = strECode
                        oUserTable.UserFields.Fields.Item("U_Z_EmpID").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                        oUserTable.UserFields.Fields.Item("U_Z_MemCode").Value = (oGrid.DataTable.GetValue("U_Z_MemCode", intRow))
                        oUserTable.UserFields.Fields.Item("U_Z_MemName").Value = (oGrid.DataTable.GetValue("U_Z_MemName", intRow))
                        If oGrid.DataTable.GetValue("U_Z_Gender", intRow) = "" Then
                            oUserTable.UserFields.Fields.Item("U_Z_Gender").Value = "B"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_Gender").Value = oGrid.DataTable.GetValue("U_Z_Gender", intRow)
                        End If


                        oUserTable.UserFields.Fields.Item("U_Z_DOB").Value = (oGrid.DataTable.GetValue("U_Z_DOB", intRow))
                        Dim dtdate As Date
                        dtdate = oGrid.DataTable.GetValue("U_Z_DOM", intRow)
                        If Year(dtdate) <> 1 Then
                            oUserTable.UserFields.Fields.Item("U_Z_DOM").Value = (oGrid.DataTable.GetValue("U_Z_DOM", intRow))
                        End If
                        ' MsgBox(oGrid.DataTable.GetValue(7, intRow))


                        If oGrid.DataTable.GetValue("U_Z_STUD", intRow) = "" Then
                            oUserTable.UserFields.Fields.Item("U_Z_STUD").Value = "N"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_STUD").Value = oGrid.DataTable.GetValue("U_Z_STUD", intRow)

                        End If

                        If oGrid.DataTable.GetValue("U_Z_Emp", intRow) = "" Then
                            oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = "N"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = oGrid.DataTable.GetValue("U_Z_Emp", intRow)

                        End If


                        If oGrid.DataTable.GetValue("U_Z_Married", intRow) = "" Then
                            oUserTable.UserFields.Fields.Item("U_Z_Married").Value = "N"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_Married").Value = oGrid.DataTable.GetValue("U_Z_Married", intRow)

                        End If
                        '  oUserTable.UserFields.Fields.Item("U_Z_Emp").Value = (oGrid.DataTable.GetValue(7, intRow))
                        dtdate = oGrid.DataTable.GetValue("U_Z_DOJ", intRow)
                        If Year(dtdate) <> 1 Then
                            oUserTable.UserFields.Fields.Item("U_Z_DOJ").Value = (oGrid.DataTable.GetValue("U_Z_DOJ", intRow))
                        End If
                        dtdate = oGrid.DataTable.GetValue("U_Z_DOT", intRow)
                        ' MsgBox(Year(dtdate))
                        If Year(dtdate) <> 1 Then
                            oUserTable.UserFields.Fields.Item("U_Z_DOT").Value = (oGrid.DataTable.GetValue(12, intRow))
                        End If

                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_NSSF")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            strNSSF = "Y"
                        Else
                            strNSSF = "N"
                        End If
                        oUserTable.UserFields.Fields.Item("U_Z_NSSF").Value = strNSSF

                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_StopAllowance")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            strNSSF = "Y"
                        Else
                            strNSSF = "N"
                        End If

                        oUserTable.UserFields.Fields.Item("U_Z_StopAllowance").Value = strNSSF



                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_MRC")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            strNSSF = "Y"
                        Else
                            strNSSF = "N"
                        End If

                        oUserTable.UserFields.Fields.Item("U_Z_MRC").Value = strNSSF


                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_BCR")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            strNSSF = "Y"
                        Else
                            strNSSF = "N"
                        End If

                        oUserTable.UserFields.Fields.Item("U_Z_BCR").Value = strNSSF

                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_INS")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            strNSSF = "Y"
                        Else
                            strNSSF = "N"
                        End If

                        oUserTable.UserFields.Fields.Item("U_Z_INS").Value = strNSSF
                        If oUserTable.Update() <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Committrans("Cancel")
                            Return False
                        End If
                    End If
                End If
            End If
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Committrans("Add")
        Dim ORe As SAPbobsCOM.Recordset
        ORe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ORe.DoQuery("Select * from [@Z_PAY_EMPFAMILY] where U_Z_EmpID='" & oApplication.Utilities.getEdittextvalue(aform, "7") & "' and U_Z_MemCode Like 'CH%'")
        Dim intChile As Integer = 0
        If ORe.RecordCount > 0 Then
            intChile = ORe.RecordCount
        End If
        ORe.DoQuery("Update OHEM set nChildren=" & intChile & " where empid=" & oApplication.Utilities.getEdittextvalue(aform, "7"))


        Databind(aform, oApplication.Utilities.getEdittextvalue(aform, "7"))
    End Function
#End Region

#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Dim strCode, strname As String
        Dim otemprec As SAPbobsCOM.Recordset
        For intRow = 0 To agrid.DataTable.Rows.Count - 1
            If agrid.Rows.IsSelected(intRow) Then
                strCode = agrid.DataTable.GetValue(0, intRow)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oApplication.Utilities.ExecuteSQL(oTemp, "update [@Z_PAY_EMPFAMILY] set  NAME =NAME +'DX'  where CODE='" & strCode & "'")
                agrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region


#Region "Validate Grid details"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strECode, strECode1, strEname, strEname1 As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            oComboboxColumn = aGrid.Columns.Item(3)
            Try
                strECode = oComboboxColumn.GetSelectedValue(intRow).Value
            Catch ex As Exception
                strECode = ""
            End Try

            If strECode <> "" Then
                strEname = aGrid.DataTable.GetValue(4, intRow)
                If strEname = "" Then
                    oApplication.Utilities.Message("Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                Dim strdate As String
                strdate = (aGrid.DataTable.GetValue(6, intRow))
                If strdate = "" Then
                    oApplication.Utilities.Message("Date of Birth is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If strECode.StartsWith("C") Then
                    Dim strGender As String
                    Try
                        strGender = oGrid.DataTable.GetValue(5, intRow)
                    Catch ex As Exception
                        strGender = ""

                    End Try
                    If strGender = "" Then
                        oApplication.Utilities.Message("Gender is missing... for Child", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
                Dim isStudent As String
                Dim isEmp As String
                If strECode.StartsWith("C") = True Then
                    Dim ocheckboxcolumn As SAPbouiCOM.CheckBoxColumn
                    ocheckboxcolumn = aGrid.Columns.Item("U_Z_STUD")
                    If ocheckboxcolumn.IsChecked(intRow) = True Then
                        oComboboxColumn = aGrid.Columns.Item("U_Z_Emp")
                        If oComboboxColumn.GetSelectedValue(intRow).Value = "Y" Then
                            oApplication.Utilities.Message("Student should not be employeed", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If

                        oComboboxColumn = aGrid.Columns.Item(3)
                        strECode1 = oComboboxColumn.GetSelectedValue(intRow).Value
                        If strECode1.StartsWith("C") Then
                            ocheckboxcolumn = aGrid.Columns.Item("U_Z_Married")
                            If ocheckboxcolumn.IsChecked(intRow) = True Then
                                oApplication.Utilities.Message("Student should not be Married Status", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If

                    End If
                End If
                For intInnerLoop As Integer = intRow To aGrid.DataTable.Rows.Count - 1
                    oComboboxColumn = aGrid.Columns.Item(3)
                    strECode1 = oComboboxColumn.GetSelectedValue(intInnerLoop).Value
                    ' strECode1 = aGrid.DataTable.GetValue(0, intInnerLoop)
                    strEname1 = aGrid.DataTable.GetValue(4, intInnerLoop)
                    If strECode1 <> "" And strEname1 = "" Then
                        oApplication.Utilities.Message("Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode1 = "" And strEname1 <> "" Then
                        oApplication.Utilities.Message("Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode = strECode1 And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("This entry  already exists. Memeber Code : " & strECode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aGrid.Columns.Item(0).Click(intInnerLoop, , 1)
                        Return False
                    End If
                Next
            End If
        Next
        Return True
    End Function

#End Region

#Region "Populate SelfDetails"
    Private Sub populateSelfDetails(ByVal aRow As Integer, ByVal aForm As SAPbouiCOM.Form)
        Dim strType, strEmpID As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            aForm.Freeze(True)
            strEmpID = oApplication.Utilities.getEdittextvalue(aForm, "7")
            oRec.DoQuery("Select isnull(sex,'M') 'Gender', isnull(firstName,'') + ' ' + isnull(middleName,'') + ' ' +isnull(lastName,'') 'Name',* from OHEM where empid=" & strEmpID)
            oGrid = aForm.Items.Item("5").Specific
            oComboboxColumn = oGrid.Columns.Item("U_Z_MemCode")
            strType = oComboboxColumn.GetSelectedValue(aRow).Value
            If strType = "S" Then
                If oRec.Fields.Item("Gender").Value = "M" Then
                    oGrid.DataTable.SetValue("U_Z_Gender", aRow, "B")
                Else
                    oGrid.DataTable.SetValue("U_Z_Gender", aRow, "G")
                End If
                oGrid.DataTable.SetValue("U_Z_MemName", aRow, oRec.Fields.Item("Name").Value)
                oGrid.DataTable.SetValue("U_Z_DOB", aRow, oRec.Fields.Item("birthDate").Value)
                oGrid.DataTable.SetValue("U_Z_DOJ", aRow, oRec.Fields.Item("startDate").Value)
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_LEB_OFMD1 Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "2" Then
                                    Committrans("Cancel")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("5").Specific
                                If pVal.ItemUID = "5" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                    If oGrid.DataTable.GetValue("Ref", pVal.Row) <> "" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "5" And pVal.ColUID = "U_Z_MemCode" Then
                                    populateSelfDetails(pVal.Row, oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    If validation(oGrid) = True Then
                                        AddtoUDT1(oForm)
                                    End If
                                End If
                                If pVal.ItemUID = "3" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    AddEmptyRow(oGrid)
                                End If
                                If pVal.ItemUID = "4" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    RemoveRow(pVal.Row, oGrid)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        'If pVal.ItemUID = "5" Then
                                        '    oGrid = oForm.Items.Item("5").Specific
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    Try

                                        '        oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        '    Catch ex As Exception
                                        '    End Try
                                        'End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
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
                'Ca'se mnu_CardType
                '    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = False Then
                        AddEmptyRow(oGrid)
                    End If

                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = True Then
                        RemoveRow(1, oGrid)
                        BubbleEvent = False
                        Exit Sub
                    End If

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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    'Case mnu_LeaveMaster
                    '    oMenuobject = New clsEarning
                    '    oMenuobject.MenuEvent(pVal, BubbleEvent)
                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class

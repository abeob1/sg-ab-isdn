Imports System.Data

Module modPurchaseOrder

    Private objForm As SAPbouiCOM.Form
    Private oStatic As SAPbouiCOM.StaticText
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oCombo As SAPbouiCOM.ComboBox
    Private oRecordSet As SAPbobsCOM.Recordset
    Private sSQL As String
    Private blnAdd, blnItemCheck As Boolean

#Region "Form Modification"
    Private Sub FormModification(ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "FormModification"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Dim objItem, oItem As SAPbouiCOM.Item
            objItem = objForm.Items.Add("edEntity", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem = objForm.Items.Item("222")
            objItem.Top = oItem.Top + 17
            objItem.Width = oItem.Width
            objItem.Left = oItem.Left
            objItem.Width = oItem.Width

            oEdit = objItem.Specific
            oEdit.DataBind.SetBound(True, "OPOR", "U_ENTITY")

            objItem = objForm.Items.Add("stEntity", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem = objForm.Items.Item("230")
            objItem.Left = oItem.Left
            objItem.Height = oItem.Height
            objItem.Top = oItem.Top + 17
            objItem.Width = oItem.Width
            objItem.LinkTo = "edEntity"

            oStatic = objItem.Specific
            oStatic.Caption = "Entity"

            objItem = objForm.Items.Add("edDraft", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem = objForm.Items.Item("edEntity")
            objItem.Top = oItem.Top + 17
            objItem.Width = oItem.Width
            objItem.Left = oItem.Left
            objItem.Width = oItem.Width

            oEdit = objItem.Specific
            oEdit.DataBind.SetBound(True, "OPOR", "U_DRAFTKEY")

            objItem = objForm.Items.Add("stDraft", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem = objForm.Items.Item("stEntity")
            objItem.Left = oItem.Left
            objItem.Height = oItem.Height
            objItem.Top = oItem.Top + 17
            objItem.Width = oItem.Width
            objItem.LinkTo = "edDraft"

            oStatic = objItem.Specific
            oStatic.Caption = "Target Draft Key"

            objForm.Items.Item("edEntity").Enabled = False
            objForm.Items.Item("edDraft").Enabled = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub
#End Region
#Region "Check Fields"
    Private Function CheckFields(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim bcheck As Boolean
        bcheck = True

        Dim sSuppCode As String = String.Empty
        Dim sCustCode As String = String.Empty
        Dim sDummyItem As String = String.Empty
        Dim sTargetDB As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDs As DataSet

        oEdit = objForm.Items.Item("4").Specific
        sSuppCode = oEdit.Value

        sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSQL)
        If oRecordSet.RecordCount > 0 Then

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("System Initialization", sFuncName)
            If GetCompanyInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value
            sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value

            If sCustCode = "" Then
                sErrDesc = "Customer code is mandatory\Check mapping table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                bcheck = False
                Return bcheck
                Exit Function
            End If

            sSQL = "SELECT IFNULL(""CardCode"",'') AS ""CardCode"" FROM " & sTargetDB & ".""OCRD"" WHERE ""CardCode"" = '" & sCustCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)

            Dim sTargetCardCode As String = String.Empty
            oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
            If oDs.Tables(0).Rows.Count > 0 Then
                sTargetCardCode = oDs.Tables(0).Rows(0).Item(0).ToString
            End If
            If sTargetCardCode = "" Then
                sErrDesc = "Customer code : " & sCustCode & " does not exists in the target database"
                Call WriteToLogFile(sErrDesc, sFuncName)
                bcheck = False
                Return bcheck
                Exit Function
            End If

            sSQL = "SELECT IFNULL(""ItemCode"",'') AS ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemCode"" = '" & p_oCompDef.sDummyItemCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)

            oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
            If oDs.Tables(0).Rows.Count > 0 Then
                sDummyItem = oDs.Tables(0).Rows(0).Item(0).ToString
            End If
            If sDummyItem = "" Then
                sErrDesc = "Item code : " & p_oCompDef.sDummyItemCode & " does not exists in the target database"
                Call WriteToLogFile(sErrDesc, sFuncName)
                bcheck = False
                Return bcheck
                Exit Function
            End If

            If blnItemCheck = False Then
                sErrDesc = String.Empty
                If CheckItemCodeinTargetDB(objForm, sTargetDB, sErrDesc) <> "" Then
                    If sErrDesc <> "" Then
                        sErrDesc = sErrDesc & ". Do you want to proceed?"
                        Dim RetCode As Integer
                        RetCode = p_oSBOApplication.MessageBox(sErrDesc, 1, "Yes", "No")
                        If RetCode = 1 Then
                            blnItemCheck = True
                        Else
                            blnItemCheck = False
                        End If
                    End If
                End If
            End If

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                Dim sTrgtDraftKey, sDraftStatus As String
                sDraftStatus = String.Empty

                oEdit = objForm.Items.Item("edEntity").Specific
                sTargetDB = oEdit.Value
                oEdit = objForm.Items.Item("edDraft").Specific
                sTrgtDraftKey = oEdit.Value

                sSQL = "SELECT ""WddStatus"" FROM " & sTargetDB & ".""ODRF"" WHERE ""DocEntry"" = '" & sTrgtDraftKey & "' AND ""DocStatus"" = 'O' "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                If oDs.Tables(0).Rows.Count > 0 Then
                    sDraftStatus = oDs.Tables(0).Rows(0).Item(0).ToString
                End If
                If sDraftStatus = "-" Then

                Else
                    sErrDesc = "Cannot Update Purchase Order. Sales order Draft Status in Target company is not Draft"
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    bcheck = False
                    Return bcheck
                    Exit Function
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking the target company connection", sFuncName)
            If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then
                Call WriteToLogFile(sErrDesc, sFuncName)
                sErrDesc = "Error while connecting to the target company/Please check credentials in Company datatable"
                Call WriteToLogFile(sErrDesc, sFuncName)
                bcheck = False
                Return bcheck
                Exit Function
            End If

        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        Return bcheck
    End Function
#End Region
#Region "Create SO Draft in the Target Company"
    Private Function CreateSODraft(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateSODraft"
        Dim sSuppCode As String = String.Empty
        Dim sCustCode As String = String.Empty
        Dim sSapUser As String = String.Empty
        Dim sSapPass As String = String.Empty
        Dim oDs As DataSet
        Dim sItemCode As String = String.Empty
        Dim sItemName As String = String.Empty
        Dim sTargetDB As String = String.Empty
        Dim sTrgtItemCode As String = String.Empty
        Dim sTrgtDBItemMsg As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim dPostDate, dDelivDate, dDocDate As Date
            Dim blnLineAdded As Boolean

            oEdit = objForm.Items.Item("4").Specific
            sSuppCode = oEdit.Value
            oEdit = objForm.Items.Item("10").Specific
            dPostDate = GetDateTimeValue(oEdit.String)
            oEdit = objForm.Items.Item("12").Specific
            dDelivDate = GetDateTimeValue(oEdit.String)
            oEdit = objForm.Items.Item("46").Specific
            dDocDate = GetDateTimeValue(oEdit.String)

            oCombo = objForm.Items.Item("3").Specific
            If oCombo.Selected.Value = "I" Then

                sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value
                    sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking itemcode in target database " & sTargetDB, sFuncName)
                    CheckItemCodeinTargetDB(objForm, sTargetDB, sTrgtDBItemMsg)

                    If p_oTargetCompany.Connected = False Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sFuncName)
                        If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If p_oTargetCompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sFuncName)

                        If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Dim oSODraft As SAPbobsCOM.Documents
                        oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                        oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                        oSODraft.CardCode = sCustCode
                        oSODraft.DocDate = dPostDate
                        oSODraft.DocDueDate = dDelivDate
                        oSODraft.TaxDate = dDocDate
                        oSODraft.UserFields.Fields.Item("U_ERRMSG").Value = sTrgtDBItemMsg
                        Dim iCount As Integer = 1

                        oMatrix = objForm.Items.Item("38").Specific
                        For i As Integer = 1 To oMatrix.RowCount - 1
                            sItemName = oMatrix.Columns.Item("3").Cells.Item(i).Specific.value
                            oCombo = oMatrix.Columns.Item("257").Cells.Item(i).Specific
                            Dim sType As String
                            sType = oCombo.Selected.Value
                            If sType = "T" Then
                                oSODraft.SpecialLines.Add()
                                oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                If i - 1 = 0 Then
                                    oSODraft.SpecialLines.AfterLineNumber = -1
                                End If
                                oSODraft.SpecialLines.AfterLineNumber = i - 1
                                oSODraft.SpecialLines.LineText = sItemName
                                blnLineAdded = True
                            Else
                                If oMatrix.Columns.Item("1").Cells.Item(i).Specific.value <> "" Then
                                    sTrgtItemCode = String.Empty
                                    sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                    oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                    If oDs.Tables(0).Rows.Count > 0 Then
                                        sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                    End If
                                    If sTrgtItemCode = "" Then
                                        sTrgtItemCode = p_oCompDef.sDummyItemCode
                                    End If

                                    If iCount > 1 Then
                                        oSODraft.Lines.Add()
                                    End If
                                    oSODraft.Lines.ItemCode = sTrgtItemCode


                                    oSODraft.Lines.Quantity = oMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                                    oSODraft.Lines.DiscountPercent = oMatrix.Columns.Item("15").Cells.Item(i).Specific.value
                                    Dim sUnitPrice As String = String.Empty
                                    Dim arUnitPrice As Array
                                    sUnitPrice = oMatrix.Columns.Item("14").Cells.Item(i).Specific.value
                                    arUnitPrice = sUnitPrice.Split(" ")
                                    oSODraft.Lines.UnitPrice = arUnitPrice(1)
                                    iCount = iCount + 1
                                    blnLineAdded = True
                                End If

                            End If

                        Next

                        If blnLineAdded = True Then
                            If oSODraft.Add() <> 0 Then
                                sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                Throw New ArgumentException(sErrDesc)
                            Else
                                Dim sSoDraftKey As String = String.Empty
                                sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                objForm.Items.Item("edEntity").Enabled = True
                                objForm.Items.Item("edDraft").Enabled = True

                                objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objForm.Items.Item("edEntity").Enabled = False
                                objForm.Items.Item("edDraft").Enabled = False
                            End If
                        End If

                    End If

                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateSODraft = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateSODraft = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Update So Draft"
    Private Function UpdateSODraft(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "UpdateSODraft"
        Dim sTrgtDraftKey As String = String.Empty
        Dim sTargetDB As String = String.Empty
        Dim sSuppCode As String = String.Empty
        Dim oDs As DataSet

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oEdit = objForm.Items.Item("4").Specific
            sSuppCode = oEdit.Value
            oEdit = objForm.Items.Item("edEntity").Specific
            sTargetDB = oEdit.Value
            oEdit = objForm.Items.Item("edDraft").Specific
            sTrgtDraftKey = oEdit.Value

            If sTrgtDraftKey <> "" Then
                Dim dPostDate, dDelivDate, dDocDate As Date
                Dim blnLineAdded As Boolean
                Dim sDraftStatus As String = String.Empty
                Dim sItemName As String = String.Empty
                Dim sCustCode As String = String.Empty
                Dim sTrgtItemCode As String = String.Empty

                oEdit = objForm.Items.Item("4").Specific
                sSuppCode = oEdit.Value
                oEdit = objForm.Items.Item("10").Specific
                dPostDate = GetDateTimeValue(oEdit.String)
                oEdit = objForm.Items.Item("12").Specific
                dDelivDate = GetDateTimeValue(oEdit.String)
                oEdit = objForm.Items.Item("46").Specific
                dDocDate = GetDateTimeValue(oEdit.String)


                sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value
                    If sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value Then

                        sSQL = "SELECT ""WddStatus"" FROM " & sTargetDB & ".""ODRF"" WHERE ""DocEntry"" = '" & sTrgtDraftKey & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                        If oDs.Tables(0).Rows.Count > 0 Then
                            sDraftStatus = oDs.Tables(0).Rows(0).Item(0).ToString
                        End If
                        If sDraftStatus = "-" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sFuncName)
                            If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            If p_oTargetCompany.Connected Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sFuncName)

                                If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                Dim oSODraft As SAPbobsCOM.Documents
                                oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                If oSODraft.GetByKey(sTrgtDraftKey) Then
                                    If oSODraft.Remove() <> 0 Then
                                        sErrDesc = p_oTargetCompany.GetLastErrorDescription
                                        Call WriteToLogFile("Error while removing Draft", sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                        oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                        oSODraft.CardCode = sCustCode
                                        oSODraft.DocDate = dPostDate
                                        oSODraft.DocDueDate = dDelivDate
                                        oSODraft.TaxDate = dDocDate

                                        Dim iCount As Integer = 1

                                        oMatrix = objForm.Items.Item("38").Specific
                                        For i As Integer = 1 To oMatrix.RowCount - 1
                                            sItemName = oMatrix.Columns.Item("1").Cells.Item(i).Specific.value
                                            oCombo = oMatrix.Columns.Item("257").Cells.Item(i).Specific
                                            Dim sType As String
                                            sType = oCombo.Selected.Value
                                            If sType = "T" Then
                                                oSODraft.SpecialLines.Add()
                                                oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                                If i - 1 = 0 Then
                                                    oSODraft.SpecialLines.AfterLineNumber = -1
                                                End If
                                                oSODraft.SpecialLines.AfterLineNumber = i - 1
                                                oSODraft.SpecialLines.LineText = sItemName


                                            Else
                                                sTrgtItemCode = String.Empty
                                                sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                                oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                                If oDs.Tables(0).Rows.Count > 0 Then
                                                    sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                                End If
                                                If sTrgtItemCode = "" Then
                                                    sTrgtItemCode = p_oCompDef.sDummyItemCode
                                                End If

                                                If iCount > 1 Then
                                                    oSODraft.Lines.Add()
                                                End If
                                                oSODraft.Lines.ItemCode = sTrgtItemCode


                                                oSODraft.Lines.Quantity = oMatrix.Columns.Item("11").Cells.Item(i).Specific.value
                                                oSODraft.Lines.DiscountPercent = oMatrix.Columns.Item("15").Cells.Item(i).Specific.value
                                                Dim sUnitPrice As String = String.Empty
                                                Dim arUnitPrice As Array
                                                sUnitPrice = oMatrix.Columns.Item("14").Cells.Item(i).Specific.value
                                                arUnitPrice = sUnitPrice.Split(" ")
                                                oSODraft.Lines.UnitPrice = arUnitPrice(1)


                                            End If
                                            iCount = iCount + 1
                                            blnLineAdded = True
                                        Next

                                        If blnLineAdded = True Then
                                            If oSODraft.Add() <> 0 Then
                                                sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                                Throw New ArgumentException(sErrDesc)
                                            Else
                                                Dim sSoDraftKey As String = String.Empty
                                                sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                                objForm.Items.Item("edEntity").Enabled = True
                                                objForm.Items.Item("edDraft").Enabled = True

                                                objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                                objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                                objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                objForm.Items.Item("edEntity").Enabled = False
                                                objForm.Items.Item("edDraft").Enabled = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            sErrDesc = "Cannot Update Purchase Order. Sales order Draft Status in Target company is not Draft"
                            Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            UpdateSODraft = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            UpdateSODraft = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Check the itemCode"
    Private Function CheckItemCodeinTargetDB(ByVal objForm As SAPbouiCOM.Form, ByVal sTargetDB As String, ByRef sMessage As String) As String
        Dim bCheck As Boolean
        bCheck = True

        Dim sItemName, sTrgtItemCode As String
        Dim sItemList As String = String.Empty
        Dim oDs As DataSet

        oMatrix = objForm.Items.Item("38").Specific
        For i As Integer = 1 To oMatrix.RowCount
            If oMatrix.Columns.Item("1").Cells.Item(i).Specific.value <> "" Then
                sTrgtItemCode = String.Empty
                sItemName = oMatrix.Columns.Item("3").Cells.Item(i).Specific.value
                sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                If oDs.Tables(0).Rows.Count > 0 Then
                    sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                End If
                If sTrgtItemCode = "" Then
                    If sItemList = "" Then
                        sItemList = sItemName
                    Else
                        sItemList = sItemList & "," & sItemName
                    End If
                End If
            End If
        Next
        If sItemList <> "" Then
            sMessage = "Part number " & sItemList & " does not exists in target database " & sTargetDB
        End If

        Return sMessage
    End Function
#End Region

#Region "Item Event"
    Public Sub PurchaseOrder_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "PurchaseOrder_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            p_oTargetCompany = New SAPbobsCOM.Company
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If CheckFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    oEdit = objForm.Items.Item("edDraft").Specific
                                    If oEdit.Value = "" Then
                                        If CreateSODraft(objForm, sErrDesc) <> RTN_SUCCESS Then
                                            RollbackTransaction(p_oTargetCompany, sErrDesc)
                                            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        blnAdd = True
                                    End If
                                End If
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If CheckFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        ElseIf pval.ItemUID = "2" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Try
                                    If p_oTargetCompany.Connected Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                        RollbackTransaction(p_oTargetCompany, sErrDesc)
                                        p_oTargetCompany.Disconnect()
                                    End If
                                Catch ex As Exception
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company not connected", sFuncName)
                                End Try
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        objForm.Items.Item("edEntity").Enabled = False
                        objForm.Items.Item("edDraft").Enabled = False
                        '    If pval.CharPressed = "9" Then
                        '        If pval.ItemUID = "16" Then
                        '            oMatrix = objForm.Items.Item("38").Specific
                        '            If oMatrix.RowCount > 1 Then
                        '                Dim oCombo As SAPbouiCOM.ComboBox
                        '                oCombo = oMatrix.Columns.Item("257").Cells.Item(1).Specific
                        '                Dim s As String
                        '                s = oCombo.Selected.Value
                        '                MsgBox(s)
                        '            End If
                        '        End If
                        '    End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        FormModification(objForm)

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If pval.Action_Success = True Then
                                If p_oTargetCompany.Connected Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                    CommitTransaction(p_oTargetCompany, sErrDesc)
                                    p_oTargetCompany.Disconnect()
                                End If
                                blnAdd = False
                                blnItemCheck = False
                            End If
                            objForm.Items.Item("edEntity").Enabled = False
                            objForm.Items.Item("edDraft").Enabled = False
                        End If

                End Select
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Form Data Event"
    Public Sub PO_SBO_Application_FormDataEvent(ByRef busObj As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sfuncName As String = "PO_SBO_Application_FormDataEvent"
        Dim sErrDesc As String = String.Empty

        Try
            If (busObj.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And busObj.ActionSuccess = True) Then
                Dim sPODocentry As Integer
                Dim sTargetDB As String = String.Empty
                Dim sSuppCode As String = String.Empty
                Dim sCustCode As String = String.Empty
                Dim oDs As DataSet
                Dim sDraftStatus As String = String.Empty
                Dim sItemName As String = String.Empty
                Dim sTrgtItemCode As String = String.Empty
                Dim dPostDate, dDelivDate, dDocDate As Date
                Dim sPOType As String = String.Empty
                Dim sItemCode As String = String.Empty
                Dim sTrgtDraftKey As String = String.Empty
                Dim sTrgtDBItemMsg As String = String.Empty

                sPODocentry = objForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).Trim()
                sSuppCode = objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim()
                sPOType = objForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).Trim()
                sTargetDB = objForm.DataSources.DBDataSources.Item(0).GetValue("U_ENTITY", 0).Trim()
                sTrgtDraftKey = objForm.DataSources.DBDataSources.Item(0).GetValue("U_DRAFTKEY", 0).Trim()
               
                If sPOType = "I" Then

                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dPostDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDueDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDelivDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("TaxDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

                    Dim oPOR As SAPbobsCOM.Documents = Nothing
                    oPOR = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    If oPOR.GetByKey(sPODocentry) Then
                        sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value
                            sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking itemcode in target database " & sTargetDB, sfuncName)
                            CheckItemCodeinTargetDB(objForm, sTargetDB, sTrgtDBItemMsg)

                            If p_oTargetCompany.Connected = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sfuncName)
                                If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If p_oTargetCompany.Connected Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sfuncName)
                                If p_oTargetCompany.InTransaction = False Then
                                    If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                Dim oSODraft As SAPbobsCOM.Documents
                                oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                                If sTrgtDraftKey <> "" Then
                                    If oSODraft.GetByKey(sTrgtDraftKey) Then
                                        oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                        oSODraft.DocCurrency = objForm.DataSources.DBDataSources.Item(0).GetValue("DocCur", 0).Trim()
                                        oSODraft.UserFields.Fields.Item("U_ERRMSG").Value = sTrgtDBItemMsg

                                        For i As Integer = 0 To oSODraft.Lines.Count - 1
                                            oSODraft.Lines.Delete()
                                        Next
                                        For i As Integer = 0 To oSODraft.SpecialLines.Count - 1
                                            oSODraft.SpecialLines.Delete()
                                        Next
                                        For i As Integer = 0 To oPOR.Lines.Count - 1
                                            oPOR.Lines.SetCurrentLine(i)
                                            sItemCode = oPOR.Lines.ItemCode.Trim()
                                            If sItemCode <> "" Then
                                                sItemName = oPOR.Lines.ItemDescription.Trim()

                                                sTrgtItemCode = String.Empty
                                                sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                                oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                                If oDs.Tables(0).Rows.Count > 0 Then
                                                    sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                                End If
                                                If sTrgtItemCode = "" Then
                                                    sTrgtItemCode = p_oCompDef.sDummyItemCode
                                                End If

                                                If i > 0 Then
                                                    oSODraft.Lines.Add()
                                                End If
                                                oSODraft.Lines.ItemCode = sTrgtItemCode
                                                oSODraft.Lines.Quantity = CDbl(oPOR.Lines.Quantity)
                                                oSODraft.Lines.UnitPrice = CDbl(oPOR.Lines.UnitPrice)
                                                oSODraft.Lines.DiscountPercent = CDbl(oPOR.Lines.DiscountPercent)
                                            End If
                                        Next
                                        For i As Integer = 0 To oPOR.SpecialLines.Count - 1
                                            oPOR.SpecialLines.SetCurrentLine(i)
                                            If i > 0 Then
                                                oSODraft.SpecialLines.Add()
                                            End If
                                            oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                            oSODraft.SpecialLines.AfterLineNumber = oPOR.SpecialLines.AfterLineNumber
                                            oSODraft.SpecialLines.LineText = oPOR.SpecialLines.LineText.Trim()
                                        Next
                                        If oSODraft.Update() <> 0 Then
                                            sErrDesc = "Error while updating SO Draft in target company while adding PO\ " & p_oTargetCompany.GetLastErrorDescription
                                            BubbleEvent = False
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            Dim RetCode As Integer
                                            RetCode = p_oSBOApplication.MessageBox("Draft Added successfully in the " & sTargetDB & " database", 1, "Ok")
                                        End If
                                    End If
                                Else
                                    oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                    oSODraft.CardCode = sCustCode
                                    oSODraft.DocDate = dPostDate
                                    oSODraft.DocDueDate = dDelivDate
                                    oSODraft.TaxDate = dDocDate
                                    oSODraft.DocCurrency = objForm.DataSources.DBDataSources.Item(0).GetValue("DocCur", 0).Trim()
                                    oSODraft.UserFields.Fields.Item("U_ERRMSG").Value = sTrgtDBItemMsg

                                    For i As Integer = 0 To oPOR.Lines.Count - 1
                                        oPOR.Lines.SetCurrentLine(i)
                                        sItemCode = oPOR.Lines.ItemCode.Trim()
                                        If sItemCode <> "" Then
                                            sItemName = oPOR.Lines.ItemDescription.Trim()

                                            sTrgtItemCode = String.Empty
                                            sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                            oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                            If oDs.Tables(0).Rows.Count > 0 Then
                                                sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                            End If
                                            If sTrgtItemCode = "" Then
                                                sTrgtItemCode = p_oCompDef.sDummyItemCode
                                            End If

                                            If i > 0 Then
                                                oSODraft.Lines.Add()
                                            End If
                                            oSODraft.Lines.ItemCode = sTrgtItemCode
                                            oSODraft.Lines.Quantity = CDbl(oPOR.Lines.Quantity)
                                            oSODraft.Lines.UnitPrice = CDbl(oPOR.Lines.UnitPrice)
                                            oSODraft.Lines.DiscountPercent = CDbl(oPOR.Lines.DiscountPercent)
                                        End If
                                    Next
                                    For i As Integer = 0 To oPOR.SpecialLines.Count - 1
                                        oPOR.SpecialLines.SetCurrentLine(i)
                                        If i > 0 Then
                                            oSODraft.SpecialLines.Add()
                                        End If
                                        oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                        oSODraft.SpecialLines.AfterLineNumber = oPOR.SpecialLines.AfterLineNumber
                                        oSODraft.SpecialLines.LineText = oPOR.SpecialLines.LineText.Trim()
                                    Next
                                    If oSODraft.Add() <> 0 Then
                                        sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                        BubbleEvent = False
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        Dim RetCode As Integer
                                        RetCode = p_oSBOApplication.MessageBox("Draft Added successfully in the " & sTargetDB & " database", 1, "Ok")

                                        Dim sSoDraftKey As String = String.Empty
                                        sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                        objForm.Items.Item("edEntity").Enabled = True
                                        objForm.Items.Item("edDraft").Enabled = True

                                        objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                        objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                        oPOR.UserFields.Fields.Item("U_ENTITY").Value = sTargetDB
                                        oPOR.UserFields.Fields.Item("U_DRAFTKEY").Value = sSoDraftKey
                                        If oPOR.Update() <> 0 Then
                                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                                            Call WriteToLogFile("Error while updating Purchase order " & sErrDesc, sfuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        End If

                                        objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        objForm.Items.Item("edEntity").Enabled = False
                                        objForm.Items.Item("edDraft").Enabled = False
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
            ElseIf (busObj.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And busObj.ActionSuccess = True) Then
                Dim sPODocentry As Integer
                Dim sTrgtDraftKey As String = String.Empty
                Dim sTargetDB As String = String.Empty
                Dim sSuppCode As String = String.Empty
                Dim sCustCode As String = String.Empty
                Dim oDs As DataSet
                Dim sDraftStatus As String = String.Empty
                Dim sItemName As String = String.Empty
                Dim sTrgtItemCode As String = String.Empty
                Dim dPostDate, dDelivDate, dDocDate As Date
                Dim sPOType As String = String.Empty
                Dim sItemCode As String = String.Empty
                Dim sTrgtDBItemMsg As String = String.Empty

                sPODocentry = objForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).Trim()
                sSuppCode = objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim()
                sPOType = objForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).Trim()
                sTargetDB = objForm.DataSources.DBDataSources.Item(0).GetValue("U_ENTITY", 0).Trim()
                sTrgtDraftKey = objForm.DataSources.DBDataSources.Item(0).GetValue("U_DRAFTKEY", 0).Trim()

                If sPOType = "I" Then
                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dPostDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDueDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDelivDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("TaxDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

                    Dim oPOR As SAPbobsCOM.Documents = Nothing
                    oPOR = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

                    If oPOR.GetByKey(sPODocentry) Then
                        sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value
                            If sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking itemcode in target database " & sTargetDB, sfuncName)
                                CheckItemCodeinTargetDB(objForm, sTargetDB, sTrgtDBItemMsg)

                                If p_oTargetCompany.Connected = False Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sfuncName)
                                    If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If p_oTargetCompany.Connected Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sfuncName)
                                    If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                    Dim oSODraft As SAPbobsCOM.Documents
                                    oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                                    If sTrgtDraftKey <> "" Then
                                        sSQL = "SELECT ""WddStatus"" FROM " & sTargetDB & ".""ODRF"" WHERE ""DocEntry"" = '" & sTrgtDraftKey & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                        oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                        If oDs.Tables(0).Rows.Count > 0 Then
                                            sDraftStatus = oDs.Tables(0).Rows(0).Item(0).ToString
                                        End If
                                        If sDraftStatus = "-" Then
                                            If oSODraft.GetByKey(sTrgtDraftKey) Then
                                                If oSODraft.Remove() <> 0 Then
                                                    sErrDesc = p_oTargetCompany.GetLastErrorDescription
                                                    Call WriteToLogFile("Error while removing Draft", sfuncName)
                                                    Throw New ArgumentException(sErrDesc)
                                                End If
                                            End If
                                        Else
                                            sErrDesc = "Cannot Update Purchase Order. Sales order Draft Status in Target company is not Draft"
                                            BubbleEvent = False
                                            Throw New ArgumentException(sErrDesc)
                                        End If
                                    End If

                                    oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                    oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                    oSODraft.CardCode = sCustCode
                                    oSODraft.DocDate = dPostDate
                                    oSODraft.DocDueDate = dDelivDate
                                    oSODraft.TaxDate = dDocDate
                                    oSODraft.DocCurrency = objForm.DataSources.DBDataSources.Item(0).GetValue("DocCur", 0).Trim()
                                    oSODraft.UserFields.Fields.Item("U_ERRMSG").Value = sTrgtDBItemMsg

                                    For i As Integer = 0 To oPOR.Lines.Count - 1
                                        oPOR.Lines.SetCurrentLine(i)
                                        sItemCode = oPOR.Lines.ItemCode.Trim()
                                        If sItemCode <> "" Then
                                            sItemName = oPOR.Lines.ItemDescription.Trim()

                                            sTrgtItemCode = String.Empty
                                            sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                            oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                            If oDs.Tables(0).Rows.Count > 0 Then
                                                sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                            End If
                                            If sTrgtItemCode = "" Then
                                                sTrgtItemCode = p_oCompDef.sDummyItemCode
                                            End If

                                            If i > 0 Then
                                                oSODraft.Lines.Add()
                                            End If
                                            oSODraft.Lines.ItemCode = sTrgtItemCode
                                            oSODraft.Lines.Quantity = CDbl(oPOR.Lines.Quantity)
                                            oSODraft.Lines.UnitPrice = CDbl(oPOR.Lines.UnitPrice)
                                            oSODraft.Lines.DiscountPercent = CDbl(oPOR.Lines.DiscountPercent)
                                        End If
                                    Next
                                    For i As Integer = 0 To oPOR.SpecialLines.Count - 1
                                        oPOR.SpecialLines.SetCurrentLine(i)
                                        If i > 0 Then
                                            oSODraft.SpecialLines.Add()
                                        End If
                                        oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                        oSODraft.SpecialLines.AfterLineNumber = oPOR.SpecialLines.AfterLineNumber
                                        oSODraft.SpecialLines.LineText = oPOR.SpecialLines.LineText.Trim()
                                    Next

                                    If oSODraft.Add() <> 0 Then
                                        sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                        BubbleEvent = False
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        Dim RetCode As Integer
                                        RetCode = p_oSBOApplication.MessageBox("Draft updated successfully in the " & sTargetDB & " database", 1, "Ok")

                                        Dim sSoDraftKey As String = String.Empty
                                        sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                        objForm.Items.Item("edEntity").Enabled = True
                                        objForm.Items.Item("edDraft").Enabled = True

                                        objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                        objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                        oPOR.UserFields.Fields.Item("U_ENTITY").Value = sTargetDB
                                        oPOR.UserFields.Fields.Item("U_DRAFTKEY").Value = sSoDraftKey
                                        If oPOR.Update() <> 0 Then
                                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                                            Call WriteToLogFile("Error while updating Purchase order " & sErrDesc, sfuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        End If

                                        objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        objForm.Items.Item("edEntity").Enabled = False
                                        objForm.Items.Item("edDraft").Enabled = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sfuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sfuncName)
            BubbleEvent = False
            Exit Sub
        End Try

    End Sub
    Public Sub PO_SBO_Application_FormDataEvent_Backup(ByRef busObj As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sfuncName As String = "PO_SBO_Application_FormDataEvent"
        Dim sErrDesc As String = String.Empty

        Try
            If (busObj.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And busObj.ActionSuccess = True) Then
                Dim sPODocentry As Integer
                Dim sTargetDB As String = String.Empty
                Dim sSuppCode As String = String.Empty
                Dim sCustCode As String = String.Empty
                Dim oDs As DataSet
                Dim sDraftStatus As String = String.Empty
                Dim sItemName As String = String.Empty
                Dim sTrgtItemCode As String = String.Empty
                Dim dPostDate, dDelivDate, dDocDate As Date
                Dim sPOType As String = String.Empty
                Dim sItemCode As String = String.Empty

                sPODocentry = objForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).Trim()
                sSuppCode = objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim()
                sPOType = objForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).Trim()

                If sPOType = "I" Then

                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dPostDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDueDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDelivDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("TaxDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

                    Dim oPOR As SAPbobsCOM.Documents = Nothing
                    oPOR = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    If oPOR.GetByKey(sPODocentry) Then
                        sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value
                            sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value

                            If p_oTargetCompany.Connected = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sfuncName)
                                If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If p_oTargetCompany.Connected Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sfuncName)
                                If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                Dim oSODraft As SAPbobsCOM.Documents
                                oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                                oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                oSODraft.CardCode = sCustCode
                                oSODraft.DocDate = dPostDate
                                oSODraft.DocDueDate = dDelivDate
                                oSODraft.TaxDate = dDocDate
                                oSODraft.DocCurrency = objForm.DataSources.DBDataSources.Item(0).GetValue("DocCur", 0).Trim()

                                For i As Integer = 0 To oPOR.Lines.Count - 1
                                    oPOR.Lines.SetCurrentLine(i)
                                    sItemCode = oPOR.Lines.ItemCode.Trim()
                                    If sItemCode <> "" Then
                                        sItemName = oPOR.Lines.ItemDescription.Trim()

                                        sTrgtItemCode = String.Empty
                                        sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                        oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                        If oDs.Tables(0).Rows.Count > 0 Then
                                            sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                        End If
                                        If sTrgtItemCode = "" Then
                                            sTrgtItemCode = p_oCompDef.sDummyItemCode
                                        End If

                                        If i > 0 Then
                                            oSODraft.Lines.Add()
                                        End If
                                        oSODraft.Lines.ItemCode = sTrgtItemCode
                                        oSODraft.Lines.Quantity = CDbl(oPOR.Lines.Quantity)
                                        oSODraft.Lines.UnitPrice = CDbl(oPOR.Lines.UnitPrice)
                                        oSODraft.Lines.DiscountPercent = CDbl(oPOR.Lines.DiscountPercent)
                                    End If
                                Next
                                For i As Integer = 0 To oPOR.SpecialLines.Count - 1
                                    oPOR.SpecialLines.SetCurrentLine(i)
                                    If i > 0 Then
                                        oSODraft.SpecialLines.Add()
                                    End If
                                    oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                    oSODraft.SpecialLines.AfterLineNumber = oPOR.SpecialLines.AfterLineNumber
                                    oSODraft.SpecialLines.LineText = oPOR.SpecialLines.LineText.Trim()
                                Next

                                If oSODraft.Add() <> 0 Then
                                    sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                    BubbleEvent = False
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    Dim RetCode As Integer
                                    RetCode = p_oSBOApplication.MessageBox("Draft Added successfully in the " & sTargetDB & " database", 1, "Ok")

                                    Dim sSoDraftKey As String = String.Empty
                                    sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                    objForm.Items.Item("edEntity").Enabled = True
                                    objForm.Items.Item("edDraft").Enabled = True

                                    objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                    objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                    oPOR.UserFields.Fields.Item("U_ENTITY").Value = sTargetDB
                                    oPOR.UserFields.Fields.Item("U_DRAFTKEY").Value = sSoDraftKey
                                    If oPOR.Update() <> 0 Then
                                        sErrDesc = p_oDICompany.GetLastErrorDescription()
                                        Call WriteToLogFile("Error while updating Purchase order " & sErrDesc, sfuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    End If

                                    objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objForm.Items.Item("edEntity").Enabled = False
                                    objForm.Items.Item("edDraft").Enabled = False
                                End If

                            End If
                        End If
                    End If
                End If
            ElseIf (busObj.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And busObj.ActionSuccess = True) Then
                Dim sPODocentry As Integer
                Dim sTrgtDraftKey As String = String.Empty
                Dim sTargetDB As String = String.Empty
                Dim sSuppCode As String = String.Empty
                Dim sCustCode As String = String.Empty
                Dim oDs As DataSet
                Dim sDraftStatus As String = String.Empty
                Dim sItemName As String = String.Empty
                Dim sTrgtItemCode As String = String.Empty
                Dim dPostDate, dDelivDate, dDocDate As Date
                Dim sPOType As String = String.Empty
                Dim sItemCode As String = String.Empty

                sPODocentry = objForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).Trim()
                sSuppCode = objForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim()
                sPOType = objForm.DataSources.DBDataSources.Item(0).GetValue("DocType", 0).Trim()
                sTargetDB = objForm.DataSources.DBDataSources.Item(0).GetValue("U_ENTITY", 0).Trim()
                sTrgtDraftKey = objForm.DataSources.DBDataSources.Item(0).GetValue("U_DRAFTKEY", 0).Trim()

                If sPOType = "I" Then
                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dPostDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("DocDueDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDelivDate)
                    Date.TryParseExact((objForm.DataSources.DBDataSources.Item(0).GetValue("TaxDate", 0)), format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

                    Dim oPOR As SAPbobsCOM.Documents = Nothing
                    oPOR = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

                    If oPOR.GetByKey(sPODocentry) Then
                        sSQL = "SELECT * FROM ""@AE_CARDCODESETUP"" WHERE ""U_SOURCEDB"" = '" & p_oDICompany.CompanyDB & "' AND ""U_SUPPCODE"" = '" & sSuppCode & "' "
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sCustCode = oRecordSet.Fields.Item("U_CUSTCODE").Value
                            If sTargetDB = oRecordSet.Fields.Item("U_TARGETDB").Value Then

                                If p_oTargetCompany.Connected = False Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany() ", sfuncName)
                                    If ConnectToTargetCompany(p_oTargetCompany, sTargetDB, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If p_oTargetCompany.Connected Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target company connected", sfuncName)
                                    If StartTransaction(p_oTargetCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                    Dim oSODraft As SAPbobsCOM.Documents
                                    oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

                                    If sTrgtDraftKey <> "" Then
                                        sSQL = "SELECT ""WddStatus"" FROM " & sTargetDB & ".""ODRF"" WHERE ""DocEntry"" = '" & sTrgtDraftKey & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                        oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                        If oDs.Tables(0).Rows.Count > 0 Then
                                            sDraftStatus = oDs.Tables(0).Rows(0).Item(0).ToString
                                        End If
                                        If sDraftStatus = "-" Then
                                            If oSODraft.GetByKey(sTrgtDraftKey) Then
                                                If oSODraft.Remove() <> 0 Then
                                                    sErrDesc = p_oTargetCompany.GetLastErrorDescription
                                                    Call WriteToLogFile("Error while removing Draft", sfuncName)
                                                    Throw New ArgumentException(sErrDesc)
                                                End If
                                            End If
                                        Else
                                            sErrDesc = "Cannot Update Purchase Order. Sales order Draft Status in Target company is not Draft"
                                            BubbleEvent = False
                                            Throw New ArgumentException(sErrDesc)
                                        End If
                                    End If

                                    oSODraft = p_oTargetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                                    oSODraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oOrders
                                    oSODraft.CardCode = sCustCode
                                    oSODraft.DocDate = dPostDate
                                    oSODraft.DocDueDate = dDelivDate
                                    oSODraft.TaxDate = dDocDate
                                    oSODraft.DocCurrency = objForm.DataSources.DBDataSources.Item(0).GetValue("DocCur", 0).Trim()

                                    For i As Integer = 0 To oPOR.Lines.Count - 1
                                        oPOR.Lines.SetCurrentLine(i)
                                        sItemCode = oPOR.Lines.ItemCode.Trim()
                                        If sItemCode <> "" Then
                                            sItemName = oPOR.Lines.ItemDescription.Trim()

                                            sTrgtItemCode = String.Empty
                                            sSQL = "SELECT ""ItemCode"" FROM " & sTargetDB & ".""OITM"" WHERE ""ItemName"" = '" & sItemName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sfuncName)
                                            oDs = ExecuteSQLQuery_Hana(sSQL, sTargetDB)
                                            If oDs.Tables(0).Rows.Count > 0 Then
                                                sTrgtItemCode = oDs.Tables(0).Rows(0).Item(0).ToString
                                            End If
                                            If sTrgtItemCode = "" Then
                                                sTrgtItemCode = p_oCompDef.sDummyItemCode
                                            End If

                                            If i > 0 Then
                                                oSODraft.Lines.Add()
                                            End If
                                            oSODraft.Lines.ItemCode = sTrgtItemCode
                                            oSODraft.Lines.Quantity = CDbl(oPOR.Lines.Quantity)
                                            oSODraft.Lines.UnitPrice = CDbl(oPOR.Lines.UnitPrice)
                                            oSODraft.Lines.DiscountPercent = CDbl(oPOR.Lines.DiscountPercent)
                                        End If
                                    Next
                                    For i As Integer = 0 To oPOR.SpecialLines.Count - 1
                                        oPOR.SpecialLines.SetCurrentLine(i)
                                        If i > 0 Then
                                            oSODraft.SpecialLines.Add()
                                        End If
                                        oSODraft.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                        oSODraft.SpecialLines.AfterLineNumber = oPOR.SpecialLines.AfterLineNumber
                                        oSODraft.SpecialLines.LineText = oPOR.SpecialLines.LineText.Trim()
                                    Next

                                    If oSODraft.Add() <> 0 Then
                                        sErrDesc = "Error while adding SO Draft in target company\ " & p_oTargetCompany.GetLastErrorDescription
                                        BubbleEvent = False
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        Dim RetCode As Integer
                                        RetCode = p_oSBOApplication.MessageBox("Draft updated successfully in the " & sTargetDB & " database", 1, "Ok")

                                        Dim sSoDraftKey As String = String.Empty
                                        sSoDraftKey = p_oTargetCompany.GetNewObjectKey()

                                        objForm.Items.Item("edEntity").Enabled = True
                                        objForm.Items.Item("edDraft").Enabled = True

                                        objForm.Items.Item("edEntity").Specific.value = sTargetDB
                                        objForm.Items.Item("edDraft").Specific.value = sSoDraftKey

                                        oPOR.UserFields.Fields.Item("U_ENTITY").Value = sTargetDB
                                        oPOR.UserFields.Fields.Item("U_DRAFTKEY").Value = sSoDraftKey
                                        If oPOR.Update() <> 0 Then
                                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                                            Call WriteToLogFile("Error while updating Purchase order " & sErrDesc, sfuncName)
                                            Throw New ArgumentException(sErrDesc)
                                        End If

                                        'Dim oHeader As SAPbouiCOM.DBDataSource
                                        'oHeader = objForm.DataSources.DBDataSources.Item(0)
                                        'oHeader.SetValue("U_ENTITY", 0, sTargetDB)
                                        'oHeader.SetValue("U_DRAFTKEY", 0, sSoDraftKey)

                                        objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        objForm.Items.Item("edEntity").Enabled = False
                                        objForm.Items.Item("edDraft").Enabled = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sfuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sfuncName)
            BubbleEvent = False
            Exit Sub
        End Try

    End Sub
#End Region
#Region "Menu Events"
    Public Sub PurchaseOrder_SBO_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                If pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    objForm.Items.Item("edEntity").Enabled = False
                    objForm.Items.Item("edDraft").Enabled = False

                ElseIf pVal.MenuUID = "1282" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    objForm.Items.Item("edEntity").Enabled = False
                    objForm.Items.Item("edDraft").Enabled = False

                ElseIf pVal.MenuUID = "1281" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    objForm.Items.Item("edEntity").Enabled = False
                    objForm.Items.Item("edDraft").Enabled = False

                ElseIf pVal.MenuUID = "1287" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    objForm.Items.Item("edEntity").Specific.value = ""
                    objForm.Items.Item("edDraft").Specific.value = ""
                    objForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objForm.Items.Item("edEntity").Enabled = False
                    objForm.Items.Item("edDraft").Enabled = False

                End If
            End If
        Catch ex As Exception
            GC.Collect()
            objForm.Freeze(False)
            objForm.Update()
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        Finally
            GC.Collect()
        End Try
    End Sub
#End Region

End Module


Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports SAPbouiCOM.Framework
Imports System.Globalization

Namespace AE_ISDN_A06
    <FormAttribute("AE_ISDN_A06.Master_Data_Replication_b1f", "Master Data Replication.b1f")>
    Public Class Master_Data_Replication_b1f
        Inherits UserFormBase

        Private WithEvents oForm As SAPbouiCOM.Form
        Private _company As SAPbobsCOM.Company
        Public WithEvents SBO_Application As SAPbouiCOM.Application


        Sub New(company As SAPbobsCOM.Company, application As SAPbouiCOM.Application, ByRef sErrDesc As String)
            Try
                _company = company
                SBO_Application = application
                EditText1.Item.Enabled = False
                EditText2.Item.Enabled = False
                'SBO_Application.ItemEvent += New SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent)
                'SBO_Application.FormDataEvent += New SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent)
                'SBO_Application.MenuEvent += New SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent)

                'oForm = application.Forms.ActiveForm
                If EntityLoad(_company, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Exit Try


            Catch ex As Exception
                sErrDesc = ex.Message
            End Try
        End Sub
        Public Function EntityLoad(company As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long


            Dim sSQL As String = String.Empty
            Dim oform As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm

            'Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_10").Specific
            'Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_0").Specific
            Try
                sFuncName = "EntityLoad()"

                oform.Freeze(True)

                Matrix0.Columns.Item("Col_4").Visible = False
                Matrix0.Columns.Item("Col_5").Visible = False


                ComboBox1.ValidValues.Add("--Select--", "0")
                ComboBox1.ValidValues.Add("Exchange Rate", "EXR")
                ComboBox1.ValidValues.Add("Chart of Accounts", "COA")
                ComboBox1.ValidValues.Add("Business Partner", "BP")
                ComboBox1.ValidValues.Add("Item Master", "ITM")
                ComboBox1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                '' sSQL = "SELECT T0.[U_AB_COMCODE], T0.[U_AB_COMPANYNAME], T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD]  FROM [dbo].[@AB_COMPANYDATA]  T0"

                sSQL = "SELECT ""U_DBNAME"",T0.""Name"" , T0.""U_SAPUSER"", ""U_SAPPWD"" FROM  ""@AE_ENTITYNAME""  T0 "

                Dim sRs As SAPbobsCOM.Recordset = Nothing
                sRs = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                sRs.DoQuery(sSQL)
                Dim dbName As String = String.Empty
                Dim Name As String = String.Empty
                Dim SAPUser As String = String.Empty
                Dim SAPPWD As String = String.Empty

                If sRs.RecordCount > 0 Then


                    'For i As Integer = 1 To sRs.RecordCount
                    '    Matrix0.Columns.Item("Col_1").Cells.Item(Matrix0.RowCount).Specific.string = dbName
                    '    Matrix0.Columns.Item("Col_2").Cells.Item(Matrix0.RowCount).Specific.string = Name
                    '    Matrix0.Columns.Item("Col_4").Cells.Item(Matrix0.RowCount).Specific.string = SAPUser
                    '    Matrix0.Columns.Item("Col_5").Cells.Item(Matrix0.RowCount).Specific.string = SAPUser
                    '    sRs.MoveNext()
                    'Next
                    While Not sRs.EoF
                        dbName = sRs.Fields.Item(0).Value
                        Name = sRs.Fields.Item(1).Value
                        SAPUser = sRs.Fields.Item(2).Value
                        SAPPWD = sRs.Fields.Item(3).Value
                        Matrix0.AddRow(1)
                        Matrix0.Columns.Item("Col_1").Cells.Item(Matrix0.RowCount).Specific.string = dbName
                        Matrix0.Columns.Item("Col_2").Cells.Item(Matrix0.RowCount).Specific.string = Name
                        Matrix0.Columns.Item("Col_4").Cells.Item(Matrix0.RowCount).Specific.string = SAPUser
                        Matrix0.Columns.Item("Col_5").Cells.Item(Matrix0.RowCount).Specific.string = SAPPWD
                        sRs.MoveNext()

                    End While
                End If

                oform.Freeze(False)
                EntityLoad = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                oform.Freeze(False)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                EntityLoad = RTN_ERROR
            End Try
        End Function
        Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            'BubbleEvent = True
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID
                        Case "MSD"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "Item_3" Then
                                'Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                If oForm.Items.Item("Item_3").Specific.selected.description = "EXR" Then
                                    oForm.Items.Item("Item_5").Enabled = True
                                    oForm.Items.Item("Item_6").Enabled = True
                                    oForm.Items.Item("Item_1").Enabled = False
                                    oForm.Items.Item("Item_2").Enabled = False
                                    'oForm.Items.Item("Item_5").Click()

                                    'If p_iExchangeCount = 0 Then
                                    '    p_iCOACount = 0
                                    '    p_iBPCount = 0
                                    '    p_iItemCount = 0
                                    'End If
                                ElseIf oForm.Items.Item("Item_3").Specific.selected.description = "COA" Then                                                                    

                                    oForm.Items.Item("Item_1").Enabled = True
                                    oForm.Items.Item("Item_2").Enabled = True
                                    oForm.Items.Item("Item_5").Enabled = False
                                    oForm.Items.Item("Item_6").Enabled = False
                                   
                                    If p_iCOACount = 0 Then
                                        If AddChooseFromList_COA(oForm, "1", SBO_Application, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If CFL_DataBindingCOA(oForm, "AcctCode", "CFL1", "CFL2", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                ElseIf oForm.Items.Item("Item_3").Specific.selected.description = "BP" Then
                                    oForm.Items.Item("Item_5").Enabled = False
                                    oForm.Items.Item("Item_6").Enabled = False
                                    oForm.Items.Item("Item_1").Enabled = True
                                    oForm.Items.Item("Item_2").Enabled = True
                                    If p_iBPCount = 0 Then
                                        If AddChooseFromList_BP(oForm, "2", SBO_Application, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If CFL_DataBindingBP(oForm, "CardCode", "CFL3", "CFL4", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                ElseIf oForm.Items.Item("Item_3").Specific.selected.description = "ITM" Then
                                    oForm.Items.Item("Item_5").Enabled = False
                                    oForm.Items.Item("Item_6").Enabled = False
                                    oForm.Items.Item("Item_1").Enabled = True
                                    oForm.Items.Item("Item_2").Enabled = True
                                    If p_iItemCount = 0 Then
                                        If AddChooseFromList_Item(oForm, "4", SBO_Application, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If CFL_DataBindingItem(oForm, "ItemCode", "CFL5", "CFL6", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                'Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                Try
                                    If oCFLEvento.BeforeAction = False Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects

                                        If oForm.Items.Item("Item_3").Specific.selected.description = "COA" Then
                                            If pVal.ItemUID = "Item_1" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                            ElseIf pVal.ItemUID = "Item_2" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                            End If
                                        ElseIf oForm.Items.Item("Item_3").Specific.selected.description = "BP" Then
                                            If pVal.ItemUID = "Item_1" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("CardCode", 0)
                                            ElseIf pVal.ItemUID = "Item_2" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("CardCode", 0)
                                            End If
                                        ElseIf oForm.Items.Item("Item_3").Specific.selected.description = "ITM" Then
                                            If pVal.ItemUID = "Item_1" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                            ElseIf pVal.ItemUID = "Item_2" Then
                                                p_iReplication = 0
                                                oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                            End If
                                        End If
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                                If pVal.ItemUID = "Item_3" Then
                                    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing
                                    oForm.Items.Item("Item_1").Specific.string = String.Empty
                                    oForm.Items.Item("Item_2").Specific.string = String.Empty
                                    oForm.Items.Item("Item_5").Specific.String = String.Empty
                                    oForm.Items.Item("Item_6").Specific.String = String.Empty
                                    p_iCOACount = 0
                                    p_iBPCount = 0
                                    p_iItemCount = 0
                                    p_iExchangeCount = 0
                                    p_iReplication = 0

                                    For imjs As Integer = 1 To Matrix0.RowCount
                                        oCheck = Matrix0.Columns.Item("Col_0").Cells.Item(imjs).Specific
                                        oCheck.Checked = False
                                    Next
                                End If
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                Dim oCheckCol As SAPbouiCOM.CheckBox
                                Try
                                    Matrix0 = oForm.Items.Item("Item_10").Specific
                                    If pVal.ItemUID = "Item_10" And pVal.ColUID = "Col_0" Then
                                        oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(1).Specific

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                        Call DisplayStatus(oForm, "Please wait... while selcting/Deselcting the Entities", sErrDesc, p_oSBOApplication)

                                        oForm.Freeze(True)

                                        If oCheckCol.Checked = True Then

                                            For iRow As Integer = 1 To Matrix0.RowCount
                                                oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(iRow).Specific
                                                oCheckCol.Checked = False
                                            Next
                                        Else
                                            For iRow As Integer = 1 To Matrix0.RowCount
                                                oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(iRow).Specific
                                                oCheckCol.Checked = True
                                            Next
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Clear_Matrix()", sFuncName)
                                        If Clear_Matrix(oForm, "Item_10", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                        oForm.Freeze(False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                        Call EndStatus(sErrDesc, SBO_Application)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                    Call EndStatus(sErrDesc, SBO_Application)
                                End Try

                            End If
                    End Select
                Else
                    Select Case pVal.FormUID
                        Case "MSD"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                p_iCOACount = 0
                                p_iBPCount = 0
                                p_iItemCount = 0
                                p_iExchangeCount = 0
                                p_iReplication = 0

                               
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                'Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sSQL As String = String.Empty
                                dtTable = New DataTable
                                Dim sCheck As String = String.Empty
                                Dim oDICompany() As SAPbobsCOM.Company = Nothing
                                Dim sMasterDataType As String = String.Empty
                                Dim sMasterDataCodeF As String = String.Empty
                                Dim sMasterDataCodeT As String = String.Empty
                                Dim dDateFrom As Date
                                Dim dDateTo As Date
                                'Dim dDateFrom As String = String.Empty
                                'Dim dDateTo As String = String.Empty
                                Dim dDtFrom As String = String.Empty
                                Dim dDtTo As String = String.Empty

                                Matrix0 = oForm.Items.Item("Item_10").Specific
                                Dim sCompanyDB As String = String.Empty
                                'If pVal.ItemUID = "Item_11" Then
                                '    SBO_Application.SetStatusBarMessage("button pressed ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                'End If

                                ''Replicate button pressed
                                If pVal.ItemUID = "Item_11" And p_iReplication = 0 Then
                                    Try
                                        SBO_Application.SetStatusBarMessage("Clearing the status messages...", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        p_iReplication += 1

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Clear_Matrix()", sFuncName)
                                        If Clear_Matrix(oForm, "Item_10", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        oForm.Items.Item("Item_11").Enabled = False
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        oDT_ErrorMsg = New DataTable
                                        oDT_ErrorMsg.Columns.Add("ErrorMsg", GetType(String))

                                        SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If HeaderValidation(oForm, sErrDesc, SBO_Application) = 0 Then
                                            oForm.Items.Item("Item_11").Enabled = True
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        SBO_Application.SetStatusBarMessage("Validation Completed ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        sMasterDataType = oForm.Items.Item("Item_3").Specific.selected.description.trim()
                                        sMasterDataCodeF = oForm.Items.Item("Item_1").Specific.String
                                        sMasterDataCodeT = oForm.Items.Item("Item_2").Specific.String


                                        ReDim oDICompany(oDT_Entities.Rows.Count)
                                        If oDT_Entities.Rows.Count > 0 Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                            Call DisplayStatus(oForm, "Master data synching.  Please wait....", sErrDesc, SBO_Application)

                                            For imjs As Integer = 0 To oDT_Entities.Rows.Count - 1
                                                Dim irow As Integer = oDT_Entities.Rows(imjs).Item(0).ToString

                                                oDICompany(imjs) = New SAPbobsCOM.Company

                                                sCompanyDB = oDT_Entities.Rows(imjs).Item("Entity").ToString.Trim()
                                                Dim sCurrentComp As String = p_oDICompany.CompanyDB.ToString().ToUpper()
                                                If sCompanyDB.ToString().ToUpper() = sCurrentComp Then Continue For

                                                Matrix0.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "Processing..."
                                                Matrix0.Columns.Item("Col_6").Cells.Item(irow).Specific.String = ""

                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                                                SBO_Application.SetStatusBarMessage("Connecting to the Target Company " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                If ConnectTargetDB(oDICompany(imjs), sCompanyDB, oDT_Entities.Rows(imjs).Item("UserName").ToString.Trim(), oDT_Entities.Rows(imjs).Item("Password").ToString.Trim(), sErrDesc) <> RTN_SUCCESS Then
                                                    Throw New ArgumentException(sErrDesc)
                                                End If

                                                SBO_Application.SetStatusBarMessage("Connecting to the target company Successfull " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)

                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting transaction on company database " & oDICompany(imjs).CompanyDB, sFuncName)

                                                SBO_Application.SetStatusBarMessage("Started Master Data Synchronization " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                                If MasterDataSync(oForm, irow, p_oDICompany, oDICompany(imjs), sMasterDataType, sMasterDataCodeF, sMasterDataCodeT, sErrDesc, SBO_Application) <> RTN_SUCCESS Then
                                                    SBO_Application.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                                    'Matrix0.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "FAIL"
                                                    'Matrix0.Columns.Item("Col_6").Cells.Item(irow).Specific.String = "Error occurred.. please refer to error log for more information."

                                                    SBO_Application.SetStatusBarMessage("Completed with ERROR " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                Else
                                                    'Matrix0.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "SUCCESS"
                                                    'Matrix0.Columns.Item("Col_6").Cells.Item(irow).Specific.String = ""
                                                    SBO_Application.SetStatusBarMessage("Completed with SUCCESS " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If

                                                If sMasterDataType = "EXR" Then
                                                    'dDateFrom = oForm.Items.Item("Item_5").Specific.String
                                                    'Dim dDtFrom As String = dDateFrom.ToString("yyyy-dd-MM")
                                                    'dDateTo = oForm.Items.Item("Item_6").Specific.String
                                                    'Dim dDtTo As String = dDateTo.ToString("yyyy-dd-MM")

                                                    'dDateFrom = oForm.Items.Item("Item_5").Specific.String
                                                    dDateFrom = DateTime.ParseExact(oForm.Items.Item("Item_5").Specific.String, "dd.MM.yyyy", CultureInfo.InvariantCulture)
                                                    dDtFrom = dDateFrom.ToString("yyyy-MM-dd")
                                                    'dDateTo = oForm.Items.Item("Item_6").Specific.String
                                                    dDateTo = DateTime.ParseExact(oForm.Items.Item("Item_6").Specific.String, "dd.MM.yyyy", CultureInfo.InvariantCulture)
                                                    dDtTo = dDateTo.ToString("yyyy-MM-dd")
                                                    If MasterDataSync_ExRate(oForm, irow, p_oDICompany, oDICompany(imjs), sMasterDataType, dDtFrom, dDtTo, sErrDesc, SBO_Application) <> RTN_SUCCESS Then
                                                        SBO_Application.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                                        Matrix0.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "FAIL"
                                                        Matrix0.Columns.Item("Col_6").Cells.Item(irow).Specific.String = "Error occurred.. please refer to error log for more information."

                                                        SBO_Application.SetStatusBarMessage("Completed with ERROR " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    Else
                                                        Matrix0.Columns.Item("Col_3").Cells.Item(irow).Specific.String = "SUCCESS"
                                                        Matrix0.Columns.Item("Col_6").Cells.Item(irow).Specific.String = ""
                                                        SBO_Application.SetStatusBarMessage("Completed with SUCCESS " & oDT_Entities.Rows(imjs).Item("Entity").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    End If
                                                End If


                                            Next imjs
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling WriteLogFiles()", sFuncName)
                                            'Call WriteLogFiles(sErrDesc)
                                            Call WriteToLogFile(sErrDesc, sFuncName)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                            Call EndStatus(sErrDesc, SBO_Application)
                                        End If

                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                            If Not oDICompany(lCounter) Is Nothing Then
                                                If oDICompany(lCounter).Connected = True Then
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                    oDICompany(lCounter).Disconnect()
                                                    oDICompany(lCounter) = Nothing
                                                End If
                                            End If
                                        Next
                                        '' oMatrix.AutoResizeColumns()
                                        Matrix0.AutoResizeColumns()

                                        If oDT_ErrorMsg.Rows.Count > 0 Then
                                            If Write_TextFile_Error(oDT_ErrorMsg, System.Windows.Forms.Application.StartupPath.ToString, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS .......", sFuncName)
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    Finally
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                                        Call EndStatus(sErrDesc, SBO_Application)
                                        oForm.Items.Item("Item_11").Enabled = True
                                    End Try

                                End If
                            End If
                    End Select
                End If
            Catch ex As Exception
                BubbleEvent = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling WriteLogFiles()", sFuncName)
                Call WriteToLogFile(sErrDesc, sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                ''  Call EndStatus(sErrDesc, SBO_Application)

                sErrDesc = ex.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try
        End Sub
        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            BubbleEvent = True
        End Sub
        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            BubbleEvent = True
        End Sub
        Public Overrides Sub OnInitializeComponent()
            Me.Matrix0 = CType(Me.GetItem("Item_10").Specific, SAPbouiCOM.Matrix)
            Me.ComboBox1 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.ComboBox)
            Me.EditText1 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.EditText)
            Me.EditText2 = CType(Me.GetItem("Item_6").Specific, SAPbouiCOM.EditText)
            Me.EditText3 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.EditText)
            Me.EditText4 = CType(Me.GetItem("Item_2").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub
        Public Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            ' throw new System.NotImplementedException();
        End Sub
        Public Overrides Sub OnInitializeFormEvents()

        End Sub


        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox

        Private Sub OnCustomInitialize()

        End Sub
        Private Sub ComboBox1_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox1.ComboSelectAfter
            'Throw New System.NotImplementedException()
            'Dim sComboSelect As String = String.Empty
            ''sComboSelect = oForm.Items.Item("Item_3").Specific.selected.description.trim()
            'sComboSelect = ComboBox1.Selected.Description.Trim()
            'If sComboSelect = "EXR" Then
            '    'If EditText1.Value <> "" And EditText2.Value <> "" Then
            '    '    EditText1.Value = ""
            '    '    EditText2.Value = ""
            '    'End If
            '    EditText1.Item.Enabled = True
            '    EditText2.Item.Enabled = True
            '    EditText1.Item.Click()
            '    EditText3.Item.Enabled = False
            '    EditText4.Item.Enabled = False
            'Else
            '    If EditText1.Value <> "" And EditText2.Value <> "" Then
            '        EditText1.Value = ""
            '        EditText2.Value = ""
            '    End If
            '    EditText3.Item.Enabled = True
            '    EditText4.Item.Enabled = True
            '    EditText3.Item.Click()
            '    EditText1.Item.Enabled = False
            '    EditText2.Item.Enabled = False

            'End If
        End Sub
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Button0 As SAPbouiCOM.Button


        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            'Throw New System.NotImplementedException()

        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub Button1_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button1.ClickBefore
            'Throw New System.NotImplementedException()
            'Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
            'oForm.Close()
        End Sub

        Private Sub EditText1_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub

        Private Sub EditText3_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub

        Private Sub EditText2_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub
    End Class
End Namespace

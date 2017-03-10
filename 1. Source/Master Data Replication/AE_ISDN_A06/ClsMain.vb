Option Explicit On
Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A06
    Public Class ClsMain
        Public WithEvents SBO_Application As SAPbouiCOM.Application

        Public Sub New()
            SBO_Application = Application.SBO_Application
            Single_Signon(SBO_Application)
        End Sub

        Public Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus = Nothing
            Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
            oMenus = Application.SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = Application.SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "DRP"
            oCreationPackage.String = "Customization"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\Logo.bmp"
            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage)
            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("DRP")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "UDO_SYS"
                oCreationPackage.String = "Update System Rate"
                oCreationPackage.Position = 0
                oMenus.AddEx(oCreationPackage)

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "MSD"
                oCreationPackage.String = "Master Data Replication"
                oCreationPackage.Position = 1
                oMenus.AddEx(oCreationPackage)
            Catch
                'Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub


        Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent


            Try
                If (pVal.BeforeAction And pVal.MenuUID = "MSD") Then


                    If p_iDRP = 0 Then

                        Dim activeForm As Master_Data_Replication_b1f
                        activeForm = New Master_Data_Replication_b1f(ClsFunction.Company, Application.SBO_Application, sErrDesc)
                        If String.IsNullOrEmpty(sErrDesc) Then
                            activeForm.Show()
                        Else
                            'Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.Item("MSD")

                        End If

                    Else
                        BubbleEvent = False
                        Exit Sub
                    End If


                End If


                If (pVal.BeforeAction And pVal.MenuUID = "UDO_SYS") Then


                    'Application.SBO_Application.Menus.Item("SYSRATE").Activate()
                    'Dim oForm As SAPbouiCOM.Form = Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "SYSRATE", "")

                    'Dim activeForm1 As SysRate
                    'activeForm1 = New SysRate(ClsFunction.Company, Application.SBO_Application)
                    'activeForm1.Show()

                    Dim activeForm1 As UDOSYS_b1f
                    activeForm1 = New UDOSYS_b1f(ClsFunction.Company, Application.SBO_Application, sErrDesc)
                    If String.IsNullOrEmpty(sErrDesc) Then
                        activeForm1.Show()

                    End If

                 
                End If
            Catch ex As System.Exception
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "")
            End Try

        End Sub

        Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_AppEvent()
            '   Purpose    :    This function will be handling the SAP Application Event
            '               
            '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
            '                       EventType = set the SAP UI Application Eveny Object        
            ' **********************************************************************************
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty
            Dim sMessage As String = String.Empty

            Try
                sFuncName = "SBO_Application_AppEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Select Case EventType
                    Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                        sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                        p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End
                End Select

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ShowErr(sErrDesc)
            Finally
                GC.Collect()  'Forces garbage collection of all generations.
            End Try
        End Sub

        'Public Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
        '        ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        '    ' **********************************************************************************
        '    '   Function   :    SBO_Application_ItemEvent()
        '    '   Purpose    :    This function will be handling the SAP Menu Event
        '    '               
        '    '   Parameters :    ByVal FormUID As String
        '    '                       FormUID = set the FormUID
        '    '                   ByRef pVal As SAPbouiCOM.ItemEvent
        '    '                       pVal = set the SAP UI ItemEvent Object
        '    '                   ByRef BubbleEvent As Boolean
        '    '                       BubbleEvent = set the True/False        
        '    ' **********************************************************************************

        '    Dim sErrDesc As String = String.Empty
        '    Dim sFuncName As String = String.Empty
        '    Dim p_oDVJE As DataView = Nothing
        '    Dim oDTDistinct As DataTable = Nothing
        '    Dim oDTRowFilter As DataTable = Nothing

        '    Try
        '        sFuncName = "SBO_Application_ItemEvent()"
        '        ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

        '        If Not IsNothing(p_oDICompany) Then
        '            If Not p_oDICompany.Connected Then
        '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
        '                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '            End If
        '        End If

        '        If pVal.BeforeAction = False Then

        '            Select Case pVal.FormUID
        '                Case "SOA"
        '                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
        '                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        '                        oCFLEvento = pVal
        '                        Dim sCFL_ID As String
        '                        sCFL_ID = oCFLEvento.ChooseFromListUID
        '                        Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.Item(FormUID)
        '                        Dim oCFL As SAPbouiCOM.ChooseFromList
        '                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
        '                        Try
        '                            If oCFLEvento.BeforeAction = False Then
        '                                Dim oDataTable As SAPbouiCOM.DataTable
        '                                oDataTable = oCFLEvento.SelectedObjects
        '                                If pVal.ItemUID = "BPFrom" Then 'BP From
        '                                    oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("CardName", 0)
        '                                    oForm.Items.Item("BPFrom").Specific.string = oDataTable.GetValue("CardCode", 0)
        '                                End If
        '                                If pVal.ItemUID = "BPTo" Then 'BP To
        '                                    oForm.Items.Item("Item_3").Specific.string = oDataTable.GetValue("CardName", 0)
        '                                    oForm.Items.Item("BPTo").Specific.string = oDataTable.GetValue("CardCode", 0)
        '                                End If
        '                            End If
        '                        Catch ex As Exception
        '                        End Try
        '                    End If

        '                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
        '                        If pVal.ItemUID = "Item_9" Then
        '                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
        '                            Try
        '                                Dim oMAtrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_8").Specific
        '                                Dim oCheck As SAPbouiCOM.CheckBox = oForm.Items.Item("Item_9").Specific
        '                                Dim ocheckColumn As SAPbouiCOM.CheckBox

        '                                If oCheck.Checked = True Then
        '                                    For mjs As Integer = 1 To oMAtrix.RowCount
        '                                        ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
        '                                        ocheckColumn.Checked = True
        '                                    Next mjs
        '                                Else
        '                                    For mjs As Integer = 1 To oMAtrix.RowCount
        '                                        ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
        '                                        ocheckColumn.Checked = False
        '                                    Next mjs
        '                                End If

        '                            Catch ex As Exception
        '                                BubbleEvent = False
        '                                sErrDesc = ex.Message
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '                                WriteToLogFile(Err.Description, sFuncName)
        '                                ShowErr(sErrDesc)
        '                            End Try
        '                            Exit Sub
        '                        End If

        '                    End If

        '            End Select
        '        Else
        '            Select Case pVal.FormUID
        '                Case "SOA"
        '                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
        '                        If pVal.ItemUID = "Item_10" Then
        '                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
        '                            Try
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)

        '                                SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                If HeaderValidation(oForm, sErrDesc) = 0 Then
        '                                    BubbleEvent = False
        '                                    Exit Sub
        '                                End If
        '                                SBO_Application.SetStatusBarMessage("Loading Data ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Loading_AgingDetails()", sFuncName)
        '                                If Loading_AgingDetails(oForm, SBO_Application, p_oDICompany, sErrDesc) = 0 Then
        '                                    BubbleEvent = False
        '                                    Exit Sub
        '                                End If
        '                                oForm.Items.Item("Item_12").Specific.String = ""
        '                                SBO_Application.SetStatusBarMessage("Loading Data Completed Successfully ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        '                                Exit Sub
        '                            Catch ex As Exception
        '                                BubbleEvent = False
        '                                sErrDesc = ex.Message
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '                                WriteToLogFile(Err.Description, sFuncName)
        '                                ShowErr(sErrDesc)
        '                            End Try
        '                        End If

        '                        If pVal.ItemUID = "Item_11" Then
        '                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
        '                            Dim sTargetFileName As String = String.Empty
        '                            Dim sRptFileName As String = String.Empty
        '                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        '                            Dim oCheck As SAPbouiCOM.CheckBox = Nothing

        '                            Try
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
        '                                SBO_Application.SetStatusBarMessage("Validating the Records .... !", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

        '                                If RowValidation(oForm, SBO_Application, sErrDesc) = 0 Then
        '                                    BubbleEvent = False
        '                                    Exit Sub
        '                                End If
        '                                oMatrix = oForm.Items.Item("Item_8").Specific
        '                                sTargetFileName = "Statement of Account_" & Format(Now.Date, "dd-MM-yyyy") & ".pdf"
        '                                sTargetFileName = System.Windows.Forms.Application.StartupPath & "\" & sTargetFileName
        '                                sRptFileName = System.Windows.Forms.Application.StartupPath & "\Statement_of_Account_V2.rpt"

        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF() ", sFuncName)

        '                                For mjs As Integer = 1 To oMatrix.RowCount
        '                                    oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
        '                                    If oCheck.Checked And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Successfully Sent Email" Then
        '                                        oForm.Items.Item("Item_12").Specific.String = "Processing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
        '                                        SBO_Application.SetStatusBarMessage("Exporting SOA to PDF .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                        If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
        '                                                      System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
        '                                            oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
        '                                            Throw New ArgumentException(sErrDesc)
        '                                        End If
        '                                        SBO_Application.SetStatusBarMessage("Sending Email .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                        If SendEmailNotification(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
        '                                            oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
        '                                            Dim sErrMsg As String = sErrDesc
        '                                            sErrDesc = ""
        '                                            '' Throw New ArgumentException(sErrDesc)
        '                                            SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                            If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
        '                                                              System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
        '                                                              oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Fail To Send", sErrMsg, p_oDICompany, sErrDesc) = RTN_SUCCESS Then
        '                                            End If
        '                                        Else
        '                                            oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Successfully Sent Email"
        '                                            SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                            If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
        '                                                             System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
        '                                                             oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Successfully Sent Email", "", p_oDICompany, sErrDesc) = RTN_SUCCESS Then
        '                                            End If

        '                                        End If
        '                                    End If
        '                                Next mjs
        '                                oForm.Items.Item("Item_12").Specific.String = "Email Processing is Completed ......... "
        '                                SBO_Application.SetStatusBarMessage("Email Processing is Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        '                                Exit Sub
        '                            Catch ex As Exception
        '                                BubbleEvent = False
        '                                sErrDesc = ex.Message
        '                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '                                WriteToLogFile(Err.Description, sFuncName)
        '                                ShowErr(sErrDesc)
        '                            End Try
        '                        End If

        '                    End If

        '            End Select
        '        End If


        '        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        '    Catch exc As Exception
        '        BubbleEvent = False
        '        sErrDesc = exc.Message
        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '        WriteToLogFile(Err.Description, sFuncName)
        '        ShowErr(sErrDesc)
        '    End Try

        'End Sub
    End Class
End Namespace
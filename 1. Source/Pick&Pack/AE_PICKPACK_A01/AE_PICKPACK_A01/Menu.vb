Imports SAPbouiCOM.Framework

Namespace AE_PICKPACK_A01
    Public Class Menu

        Private WithEvents SBO_Application As SAPbouiCOM.Application

        Sub New()
            SBO_Application = Application.SBO_Application

            Single_Signon(SBO_Application)

        End Sub

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = Application.SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = Application.SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "AE_PICKPACK_A01"
            oCreationPackage.String = "AE_PICKPACK_A01"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage)
            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("AE_PICKPACK_A01")
                oMenus = oMenuItem.SubMenus

                ''Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING

                ''Please replace following 2 "Form1" with real form class in current project
                'oCreationPackage.UniqueID = "AE_PICKPACK_A01.Form1"
                'oCreationPackage.String = "Form1"
                'oMenus.AddEx(oCreationPackage)
            Catch
                'Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty
            Dim sMessage As String = String.Empty

            Try
                sFuncName = "SBO_Application_AppEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Select Case EventType
                    Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                        sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                        Application.SBO_Application.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        End
                End Select

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                Application.SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ShowErr(sErrDesc, Application.SBO_Application)
            Finally
                GC.Collect()  'Forces garbage collection of all generations.
            End Try
        End Sub

        Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty

            Try
                If pVal.FormTypeEx = "81" Then
                    If pVal.Before_Action = True Then

                        ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ALL_EVENTS Then
                        ''    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetFormByTypeAndCount("81", pVal.FormTypeCount)
                        ''    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("10").Specific
                        ''    Dim oColumns As SAPbouiCOM.Columns = Nothing
                        ''    Dim oColumn As SAPbouiCOM.Column = Nothing
                        ''    '------------ Creating a New column dynamically
                        ''    oMatrix.Clear()
                        ''    oColumns = oMatrix.Columns
                        ''    oColumn = oColumns.Add("UnConQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                        ''    oColumn.Width = 100
                        ''    oColumn.TitleObject.Caption = "UnConfirm Qty"
                        ''    oColumn.Editable = False
                        ''    oColumn.DataBind.SetBound(True, "RDR1", "U_UnConfirmQty")
                        ''    '------------
                        ''End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DRAW Then
                            Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetFormByTypeAndCount("81", pVal.FormTypeCount)
                            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("10").Specific
                            Dim oDTOSLD As DataTable = Nothing
                            Dim oDVOSLD As DataView = Nothing
                            Dim oDTORDR As DataTable = Nothing
                            Dim oDVORDR As DataView = Nothing
                            Dim sSQL As String = String.Empty
                            Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim sUpdateSQL As String = String.Empty
                            ''Dim oColumns As SAPbouiCOM.Columns = Nothing
                            ''Dim oColumn As SAPbouiCOM.Column = Nothing

                            Try
                                sFuncName = "Form_Draw()"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                                p_oDTPick.Rows.Clear()
                                ' ''------------ Creating a New column dynamically
                                ''oColumns = oMatrix.Columns
                                ''oColumn = oColumns.Add("UnConQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                ''oColumn.Width = 100
                                ''oColumn.TitleObject.Caption = "UnConfirm Qty"
                                ''oColumn.Editable = False
                                ''oColumn.DataBind.SetBound(True, "RDR1", "U_UnConfirmQty")
                                ' ''------------
                                SBO_Application.SetStatusBarMessage("Pls. wait ...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                ''  sSQL = "SELECT T0.""CfmQty"", T0.""DocEntry"", T0.""ItemCode"" , T0.""DocLineNum"" , T0.""ObjType"" FROM OSLD T0"
                                sSQL = "SELECT T0.""CfmQty"", T0.""DocEntry"", T0.""ItemCode"" , T0.""DocLineNum"" , T0.""ObjType"" , T1.""DocNum""  FROM OSLD T0 join ORDR T1 on T0.""DocEntry"" = T1.""DocEntry"" and T0.""ObjType"" = T1.""ObjType"""
                                oDTOSLD = New DataTable
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting from OSDL " & sSQL, sFuncName)
                                oRset.DoQuery(sSQL)
                                oDTOSLD = ConvertRecordset(oRset, sErrDesc)

                                sUpdateSQL = " update RDR1 set ""U_UnConfirmQty"" = 0 "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cleaning the UnConfirmQty " & sUpdateSQL, sFuncName)
                                oRset.DoQuery(sUpdateSQL)


                                ''sSQL = "SELECT T0.""DocNum"", T0.""DocEntry"" FROM ORDR T0"
                                ''oDTORDR = New DataTable
                                ''oRset.DoQuery(sSQL)
                                ''oDTORDR = ConvertRecordset(oRset, sErrDesc)
                                ''oDVORDR = New DataView(oDTORDR)

                                If String.IsNullOrEmpty(sErrDesc) Then
                                    If oDTOSLD.Rows.Count > 0 Then
                                        oDVOSLD = New DataView(oDTOSLD)
                                        For imjs As Integer = 1 To oMatrix.RowCount
                                            '   SBO_Application.MessageBox("DocEntry='" & oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String & "' and ItemCode='" & oMatrix.Columns.Item("8").Cells.Item(imjs).Specific.String & "'  and DocLineNum= " & CInt(oMatrix.Columns.Item("12").Cells.Item(imjs).Specific.String) - 1 & " and Confirm Qty ='" & oMatrix.Columns.Item("1810000128").Cells.Item(imjs).Specific.String & "' ")
                                            If String.IsNullOrEmpty(oMatrix.Columns.Item("1810000128").Cells.Item(imjs).Specific.String) Or oMatrix.Columns.Item("1810000128").Cells.Item(imjs).Specific.String = "0.00" Then
                                                ' If CDbl(oMatrix.Columns.Item("1810000128").Cells.Item(imjs).Specific.String) > 0 Then

                                                Dim ss = "DocEntry='" & oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String & "' and ItemCode='" & oMatrix.Columns.Item("8").Cells.Item(imjs).Specific.String & "'  and DocLineNum= " & CInt(oMatrix.Columns.Item("12").Cells.Item(imjs).Specific.String) - 1 & ""
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(ss, sFuncName)
                                                ''  sUpdateSQL += " update RDR1 set ""U_UnConfirmQty"" = 8 where ""DocEntry"" = '" & oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String & "' and " & _
                                                ''       """ItemCode"" = '" & oMatrix.Columns.Item("8").Cells.Item(imjs).Specific.String & "' and ""LineNum"" = " & CInt(oMatrix.Columns.Item("12").Cells.Item(imjs).Specific.String) - 1 & ""

                                                oDVOSLD.RowFilter = "DocNum='" & oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String & "' and ItemCode='" & oMatrix.Columns.Item("8").Cells.Item(imjs).Specific.String & "'  and DocLineNum= " & CInt(oMatrix.Columns.Item("12").Cells.Item(imjs).Specific.String) - 1 & ""
                                                If oDVOSLD.Count Then
                                                    'SELECT T0."LineNum", T0."ItemCode", T0."U_UnConfirmQty" FROM RDR1 T0 WHERE T0."DocEntry"  = 350
                                                    sUpdateSQL = " update RDR1 set ""U_UnConfirmQty"" = " & CDbl(oDVOSLD.Item(0)("CfmQty").ToString()) & " where ""DocEntry"" = '" & CDbl(oDVOSLD.Item(0)("DocEntry").ToString()) & "' and " & _
                                                        """ItemCode"" = '" & oMatrix.Columns.Item("8").Cells.Item(imjs).Specific.String & "' and ""LineNum"" = " & CInt(oMatrix.Columns.Item("12").Cells.Item(imjs).Specific.String) - 1 & ""
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Sales Order" & sUpdateSQL, sFuncName)
                                                    oRset.DoQuery(sUpdateSQL)
                                                End If
                                                'End If
                                            End If
                                        Next
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        SBO_Application.SetStatusBarMessage("Completed Successfully ...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    End If
                                End If
                            Catch ex As Exception
                                sErrDesc = ex.Message
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "Item_0" Then
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetFormByTypeAndCount("81", pVal.FormTypeCount)
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("10").Specific
                                Dim oChoose As SAPbouiCOM.CheckBox = Nothing
                                Try
                                    sFuncName = "UnConfirm Qty()"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                                    oForm.Freeze(True)
                                    If p_oDTPick.Rows.Count = 0 Then
                                        oForm.Freeze(False)
                                        SBO_Application.SetStatusBarMessage("Choose to Copy Un-Confirmed Qty  ...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    oMatrix.Columns.Item("3").Editable = True
                                    For Each odr As DataRow In p_oDTPick.Rows
                                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_UnConfirmQty").Cells.Item(CInt(odr(0))).Specific.String) And oMatrix.Columns.Item("U_UnConfirmQty").Cells.Item(CInt(odr(0))).Specific.String <> "0.00" And oMatrix.Columns.Item("2").Cells.Item(CInt(odr(0))).Specific.String <> "0.00" Then
                                            oMatrix.Columns.Item("3").Cells.Item(CInt(odr(0))).Specific.String = oMatrix.Columns.Item("U_UnConfirmQty").Cells.Item(CInt(odr(0))).Specific.String
                                        End If
                                    Next

                                    ''For imjs As Integer = 1 To oMatrix.RowCount
                                    ''    oChoose = oMatrix.Columns.Item("1").Cells.Item(imjs).Specific
                                    ''    If oChoose.Checked = True Then
                                    ''        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_UnConfirmQty").Cells.Item(imjs).Specific.String) Then
                                    ''            oMatrix.Columns.Item("3").Editable = True
                                    ''            oMatrix.Columns.Item("3").Cells.Item(imjs).Specific.String = oMatrix.Columns.Item("U_UnConfirmQty").Cells.Item(imjs).Specific.String
                                    ''        End If
                                    ''    End If
                                    ''Next
                                    oForm.Freeze(False)

                                    SBO_Application.SetStatusBarMessage("Completed Successfully ...!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    sErrDesc = ex.Message
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                    Exit Sub

                                End Try
                            End If
                        End If
                    Else
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                            If pVal.ItemUID = "10" And pVal.ColUID = "1" And pVal.Row = 0 Then
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetFormByTypeAndCount("81", pVal.FormTypeCount)
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("10").Specific
                                Dim oChoose As SAPbouiCOM.CheckBox = Nothing
                                Dim sCheck As String = String.Empty
                                Try
                                    oChoose = oMatrix.Columns.Item("1").Cells.Item(1).Specific
                                    If oChoose.Checked = True Then
                                        sCheck = "Y"
                                    Else
                                        sCheck = "N"
                                    End If

                                    p_oDTPick.Rows.Clear()

                                    For imjs As Integer = 1 To oMatrix.RowCount
                                        '' oChoose = oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific
                                        If sCheck = "Y" Then
                                            ' p_oDTPick.Rows.Add(oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String)
                                            p_oDTPick.Rows.Add(imjs)
                                        End If
                                    Next

                                Catch ex As Exception
                                    sErrDesc = ex.Message
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                    Exit Sub
                                End Try
                            End If

                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "10" And pVal.ColUID = "1" Then
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.GetFormByTypeAndCount("81", pVal.FormTypeCount)
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("10").Specific
                                Dim oChoose As SAPbouiCOM.CheckBox = Nothing
                                Try
                                    If pVal.Row = 0 Then Exit Sub

                                    oChoose = oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific
                                    If oChoose.Checked = True Then
                                        p_oDTPick.Rows.Add(pVal.Row)
                                    Else
                                        Dim foundRow() As DataRow = p_oDTPick.Select("row='" & pVal.Row & "'")
                                        If foundRow.Count > 0 Then
                                            For Each row As DataRow In foundRow
                                                row.Delete()
                                            Next
                                        End If
                                    End If


                                Catch ex As Exception
                                    sErrDesc = ex.Message
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                    Exit Sub
                                End Try
                            End If

                        End If
                    End If
                End If

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                BubbleEvent = False
                Exit Sub
            End Try


        End Sub


        Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            BubbleEvent = True

            Try
                If (pVal.BeforeAction And pVal.MenuUID = "AE_PICKPACK_A01.Form1") Then
                    ''Please replace following 3 "Form1" with real form class in current project
                    'Dim activeForm As Form1
                    'activeForm = New Form1
                    'activeForm.Show()
                End If
            Catch ex As System.Exception
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "")
            End Try

        End Sub

    End Class
End Namespace

Module EventHandler

#Region " ... Common Variables For SAP ..."
    Public WithEvents oApplication As SAPbouiCOM.Application
#End Region


    Private Sub oApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApplication.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select pVal.MenuUID

                End Select

                oForm = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286", "1290", "5890"
                        Select Case oForm.UniqueID
                        End Select
                        Select Case oForm.TypeEx
                            Case GRPOFormID
                                oGRPO.MenuEvent(pVal, BubbleEvent)
                        End Select
                End Select

                If pVal.MenuUID = "526" Then
                    oCompany.Disconnect()
                    oApplication.StatusBar.SetText("Addon is DisConnected . . .", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End
                End If
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Purchase Menu Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub


    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    System.Windows.Forms.Application.Exit()
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText("Application Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try

    End Sub


    Private Sub oApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles oApplication.ItemEvent
        Try
            Select Case pVal.FormType
            End Select
            Select Case pVal.FormUID
            End Select
            Select Case pVal.FormTypeEx
                Case GRPOFormID
                    oGRPO.ItemEvent(GRPOFormID, pVal, BubbleEvent)
                Case "-143"
                    If pVal.FormTypeEx = "-143" And pVal.Before_Action = False Then
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                Try
                                    Select Case pVal.ItemUID
                                        Case "U_LCMethod"
                                            If pVal.BeforeAction = False And pVal.ItemChanged = True Then
                                                Dim oform As SAPbouiCOM.Form = oApplication.Forms.GetFormByTypeAndCount(143, pVal.FormTypeCount)
                                                oGRPO.ChargeableWeightCalcFromMethod(oform)
                                            End If
                                    End Select
                                Catch ex As Exception
                                    oApplication.StatusBar.SetText("Combo Select Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Finally
                                End Try
                        End Select

                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(" ItemEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub



    Private Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles oApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormUID
            End Select
            Select Case BusinessObjectInfo.FormTypeEx
                Case GRPOFormID
                    oGRPO.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
            'End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("student FormDataEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub



    Public Sub SetEventFilter()
        Try

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub



End Module
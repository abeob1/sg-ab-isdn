Imports SAPbouiCOM

Public Class GRPO

#Region "VariableDeclaration"
    Public frmGRPO As SAPbouiCOM.Form
    Public frmGRPOUDF As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSHeader1, oDBDSDetail As SAPbouiCOM.DBDataSource
    Dim boolFormLoaded As Boolean = False
    Dim sFormUID As String
    Dim RowID As Integer
    Dim dtMainTable As New System.Data.DataTable
    Dim FormDataLoad As Boolean = False
    Dim oFolder As SAPbouiCOM.Folder
    Dim oItem As SAPbouiCOM.Item
    Dim oGrid As SAPbouiCOM.Grid
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim oMatrix1 As SAPbouiCOM.Matrix
    Dim oFormNew As SAPbouiCOM.Form = Nothing
#End Region

    Sub LoadXML(ByVal FormUID As String)
        Try
            frmGRPO = oApplication.Forms.Item(FormUID)

            oMatrix1 = frmGRPO.Items.Item("38").Specific
            oDBDSHeader = frmGRPO.DataSources.DBDataSources.Item("OPDN")
            oDBDSHeader1 = frmGRPO.DataSources.DBDataSources.Add("@GRPK_OPDN")
            oDBDSDetail = frmGRPO.DataSources.DBDataSources.Add("@GRPK_PDN1")
            boolFormLoaded = False

            FormDataLoad = True
            Me.Load_Form_Controls(FormUID)
            Me.LoadGridDetails()

            boolFormLoaded = True
        Catch ex As Exception
            boolFormLoaded = False
            oApplication.StatusBar.SetText("Load GRPO Form Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub Load_Form_Controls(ByVal FormUID As String)
        Try
            Dim oItem As SAPbouiCOM.Item
            Dim oItemRef As SAPbouiCOM.Item = Nothing
            'Dim oLabel As SAPbouiCOM.StaticText
            'Dim oComboBox As SAPbouiCOM.ComboBox
            'Dim oEditText As SAPbouiCOM.EditText

            frmGRPO.DataSources.DataTables.Add("DataTable")

            ' 1. Folder - Activity 
            oItemRef = frmGRPO.Items.Item("1320002137")
            oItem = frmGRPO.Items.Add("Packing", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = oItemRef.Top
            oItem.Height = oItemRef.Height
            oItem.Left = oItemRef.Left + oItemRef.Width
            'oItem.Width = oItemRef.Width
            oItem.Width = 10
            oItem.Visible = True
            oItem.AffectsFormMode = True

            oFolder = oItem.Specific
            oFolder.Caption = "Packing Details"
            oFolder.AutoPaneSelection = True
            oFolder.GroupWith(oItemRef.UniqueID)
            oFolder.Pane = 0

            oItem = frmGRPO.Items.Add("Matrix", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.FromPane = 23
            oItem.ToPane = 23
            oItem.Top = 170
            oItem.Left = 10
            oItem.Width = 550
            oItem.Height = 180
            oGrid = oItem.Specific
            'oGrid.

            ' User Data Source ...
            frmGRPO.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            oFolder.DataBind.SetBound(True, "", "FolderDS")
            oFolder.ValOn = True
            oFolder.ValOff = False

            'oItem = frmGRPO.Items.Add("t_LCTCWght", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Visible = False
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OPDN", "U_LCToCWeight")

            frmGRPO.PaneLevel = 1

        Catch ex As Exception
            oApplication.StatusBar.SetText("Load Matrix Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function ValidateAll() As Boolean
        Try
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Validation Function Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

          

            If pVal.FormTypeEx = GRPOFormID Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        Try
                            Select Case pVal.ItemUID
                                Case "Matrix"
                                    Select Case pVal.ColUID
                                        Case "Packing Type"
                                    End Select
                            End Select
                        Catch ex As Exception
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Try
                            Select Case pVal.ItemUID
                                Case "Matrix"
                                    Select Case pVal.ColUID
                                        Case "PACKINGTYPE"
                                            If pVal.BeforeAction = False Then
                                                If oGrid.DataTable.GetValue("PACKINGTYPE", pVal.Row) <> "" Then
                                                    oForm = oApplication.Forms.ActiveForm
                                                    Dim Currentrow As Integer = pVal.Row + 1
                                                    Dim LastRow As Integer = oGrid.DataTable.Rows.Count
                                                    If Currentrow = LastRow Then
                                                        oGrid.DataTable.Rows.Add()
                                                    End If
                                                End If
                                            End If
                                    End Select
                            End Select
                        Catch ex As Exception
                            'StatusBarErrorMsg("Lost Focus Function Failed : " & ex.Message)
                        Finally
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Try
                            frmGRPO = oApplication.Forms.Item(pVal.FormUID)
                            oDBDSHeader = frmGRPO.DataSources.DBDataSources.Item("OPDN")  'OWOR
                            If pVal.BeforeAction = False And boolFormLoaded = False Then Me.LoadXML(pVal.FormUID)

                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Form Load Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Try
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                            oDataTable = oCFLE.SelectedObjects
                            If Not oDataTable Is Nothing And pVal.BeforeAction = False Then
                                Select Case pVal.ItemUID
                                    Case "Matrix"

                                End Select
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Try
                            Select Case pVal.ItemUID
                                Case "1"
                                    If pVal.BeforeAction = True And (frmGRPO.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmGRPO.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                        If Me.ValidateAll() = False Then
                                            System.Media.SystemSounds.Asterisk.Play()
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                            End Select
                        Catch ex As Exception
                            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            oApplication.StatusBar.SetText("Click Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        Try
                            If boolFormLoaded And pVal.BeforeAction = False Then
                                boolFormLoaded = False
                                FormDataLoad = False
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Form Close Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                 
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        Try
                            Select Case pVal.ItemUID
                                Case "Matrix"
                                    Select Case pVal.ColUID
                                        Case "LENGTH", "HEIGHT", "WIDTH"
                                            'If pVal.BeforeAction = True Then
                                            '    Dim oformUDF As SAPbouiCOM.Form = oApplication.Forms.GetFormByTypeAndCount(-143, pVal.FormTypeCount)
                                            '    Dim oCombo As SAPbouiCOM.ComboBox = oformUDF.Items.Item("U_LCMethod").Specific
                                            '    If oCombo.Value.ToString.Trim() = "None" Then
                                            '        oApplication.SetStatusBarMessage("Pls. Choose Method ....!", BoMessageTime.bmt_Medium, False)
                                            '        BubbleEvent = False
                                            '        Exit Sub
                                            '    End If

                                            'End If
                                            If pVal.BeforeAction = False And pVal.ItemChanged Then
                                                HighDimensionCalculation(pVal.Row)
                                                ChargeableWeightCalc(pVal.Row)
                                                DimensionSurChargesCalc(pVal.Row)
                                                TotalChargeableValue(pVal.FormTypeCount)
                                            End If
                                        Case "ACTUALWEIGHT"
                                            'If pVal.BeforeAction = True Then
                                            '    Dim oformUDF As SAPbouiCOM.Form = oApplication.Forms.GetFormByTypeAndCount(-143, pVal.FormTypeCount)
                                            '    Dim oCombo As SAPbouiCOM.ComboBox = oformUDF.Items.Item("U_LCMethod").Specific
                                            '    If oCombo.Value.ToString.Trim() = "None" Then
                                            '        oApplication.SetStatusBarMessage("Pls. Choose Method ....!", BoMessageTime.bmt_Medium, False)
                                            '        BubbleEvent = False
                                            '        Exit Sub
                                            '    End If

                                            'End If

                                            If pVal.BeforeAction = False And pVal.ItemChanged Then
                                                ChargeableWeightCalc(pVal.Row)
                                                DimensionSurChargesCalc(pVal.Row)
                                                TotalChargeableValue(pVal.FormTypeCount)
                                            End If
                                    End Select
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Vaidate Event Failed :" & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Try
                            Select Case pVal.ItemUID
                                Case "1"
                                    If pVal.ActionSuccess And frmGRPO.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        Me.LoadGridDetails()
                                    End If
                                Case "Packing"
                                    If pVal.BeforeAction = True Then
                                        frmGRPO.PaneLevel = 23
                                        frmGRPO.Items.Item("Packing").AffectsFormMode = True
                                        'oFolder = oForm.Items.Item("Packing").Specific

                                        'Dim oform As SAPbouiCOM.Form
                                        'oForm = oApplication.Forms.GetFormByTypeAndCount(GRPOFormID, 1)
                                        'oFolder.ValOn = True
                                    End If
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Item Pressed Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Event Handler Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    If pVal.BeforeAction = False Then
                        Me.LoadGridDetails()
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub TotalChargeableValue(ByVal iFormtypecount As Integer)
        Try
            oForm.Freeze(True)
            frmGRPOUDF = oApplication.Forms.GetFormByTypeAndCount(-143, iFormtypecount)
            Dim Row1 As Integer = oGrid.DataTable.Rows.Count
            If Row1 > 0 Then
                Dim TotalValue As Double = 0
                For Row As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    Dim ChargeWeight As Double = oGrid.DataTable.GetValue("CHARGEABLEWEIGHT", Row)
                    If ChargeWeight > 0 Then
                        TotalValue = TotalValue + ChargeWeight
                    End If
                Next
                'frmGRPO.Items.Item("t_LCTCWght").Specific.string = TotalValue
                'frmGRPO.Items.Item("t_LCTCWght").Click(BoCellClickType.ct_Regular)
                frmGRPOUDF.Items.Item("U_LCToCWeight").Specific.string = TotalValue
                frmGRPOUDF.Items.Item("U_LCToCWeight").Click(BoCellClickType.ct_Regular)

            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Total Chargeable Weight Calculation Failed.." & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub ChargeableWeightCalc(ByVal Row As Integer)
        Try
            oForm.Freeze(True)
            Dim method As String = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
            Dim ActualWeight As Double = oGrid.DataTable.GetValue("ACTUALWEIGHT", Row)
            Dim LengthVal As Double = oGrid.DataTable.GetValue("LENGTH", Row)
            Dim HeightVal As Double = oGrid.DataTable.GetValue("HEIGHT", Row)
            Dim WidthVal As Double = oGrid.DataTable.GetValue("WIDTH", Row)
            If method = "Courier" Then
                If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                    Dim TotalWeight As Double = ((LengthVal * HeightVal * WidthVal) / 5000)
                    If ActualWeight < TotalWeight Then
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, TotalWeight)
                    Else
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                    End If
                Else
                    If ActualWeight > 0 Then
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                    Else
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, 0.0)
                    End If

                End If
            ElseIf method = "Forwarder" Then
                If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                    Dim BPCode As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
                    Dim ShipToCntry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString)
                    Dim TotalWeight As Double = 0.0
                    If BPCode = "Logwin" And ShipToCntry = "Switzerland" Then
                        TotalWeight = ((LengthVal * HeightVal * WidthVal) / 3000)
                    Else
                        TotalWeight = ((LengthVal * HeightVal * WidthVal) / 6000)
                    End If

                    If ActualWeight < TotalWeight Then
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, TotalWeight)
                    Else
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                    End If
                Else
                    If ActualWeight > 0 Then
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                    Else
                        oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, 0.0)
                    End If
                End If
            ElseIf ActualWeight > 0 Then
                oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub ChargeableWeightCalcFromMethod(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            'Dim oMatrix As SAPbouiCOM.Grid
            'oMatrix = oForm.Items.Item("Matrix").Specific
            Dim Row1 As Integer = oGrid.DataTable.Rows.Count - 1
            If Row1 > 0 Then
                For Row As Integer = 0 To Row1
                    Dim method As String = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
                    Dim ActualWeight As Double = oGrid.DataTable.GetValue("ACTUALWEIGHT", Row)
                    Dim LengthVal As Double = oGrid.DataTable.GetValue("LENGTH", Row)
                    Dim HeightVal As Double = oGrid.DataTable.GetValue("HEIGHT", Row)
                    Dim WidthVal As Double = oGrid.DataTable.GetValue("WIDTH", Row)
                    If method = "Courier" Then
                        If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                            Dim TotalWeight As Double = ((LengthVal * HeightVal * WidthVal) / 5000)
                            If ActualWeight < TotalWeight Then
                                oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, TotalWeight)
                            Else
                                oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                            End If
                        End If
                    ElseIf method = "Forwarder" Then
                        If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                            Dim BPCode As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
                            Dim ShipToCntry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString)
                            Dim TotalWeight As Double = 0.0
                            If BPCode = "Logwin" And ShipToCntry = "Switzerland" Then
                                TotalWeight = ((LengthVal * HeightVal * WidthVal) / 3000)
                            Else
                                TotalWeight = ((LengthVal * HeightVal * WidthVal) / 6000)
                            End If

                            If ActualWeight < TotalWeight Then
                                oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, TotalWeight)
                            Else
                                oGrid.DataTable.SetValue("CHARGEABLEWEIGHT", Row, ActualWeight)
                            End If
                        End If
                    End If
                Next
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub HighDimensionCalculation(ByVal Row As Integer)
        Try
            oForm.Freeze(True)
            Dim LengthVal As Double = oGrid.DataTable.GetValue("LENGTH", Row)
            Dim HeightVal As Double = oGrid.DataTable.GetValue("HEIGHT", Row)
            Dim WidthVal As Double = oGrid.DataTable.GetValue("WIDTH", Row)
            If (LengthVal > HeightVal) And (LengthVal > WidthVal) Then
                oGrid.DataTable.SetValue("HIGHDIMENSION", Row, LengthVal)
            ElseIf (HeightVal > LengthVal) And (HeightVal > WidthVal) Then
                oGrid.DataTable.SetValue("HIGHDIMENSION", Row, HeightVal)
            ElseIf (WidthVal > LengthVal) And (WidthVal > HeightVal) Then
                oGrid.DataTable.SetValue("HIGHDIMENSION", Row, WidthVal)
            Else
                oGrid.DataTable.SetValue("HIGHDIMENSION", Row, 0.0)
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub DimensionSurChargesCalc(ByVal Row As Integer)
        Try
            oForm.Freeze(True)
            Dim DimSurCharge As Double = 0.0
            Dim BPForwarder As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
            Dim HighDimension As Double = oGrid.DataTable.GetValue("HIGHDIMENSION", Row)
            Dim ChargeWeight As Double = oGrid.DataTable.GetValue("CHARGEABLEWEIGHT", Row)
            If (HighDimension > 0 And ChargeWeight > 0 And BPForwarder <> "") Then
                Dim Sqr As String = "Select ""U_Weight"", ""U_WeightSurcharge"", ""U_Dimension"", ""U_DimensionSurcharge"", ""U_ExtraCharge"" from ""@DIMENSIONSURG"" Where ""U_ForwarderOrCourier"" = '" & BPForwarder & "'"
                Dim Rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Sqr)
                If Rset.RecordCount > 0 Then
                    Rset.MoveFirst()
                    Dim Weight As Double = Rset.Fields.Item("U_Weight").Value
                    Dim WeightSurCharge As Double = Rset.Fields.Item("U_WeightSurcharge").Value
                    Dim Dimension As Double = Rset.Fields.Item("U_Dimension").Value
                    Dim DimensionSurCharge As Double = Rset.Fields.Item("U_DimensionSurcharge").Value
                    Dim ExtraSurCharge As String = Rset.Fields.Item("U_ExtraCharge").Value
                    If HighDimension > Dimension And ChargeWeight > Weight Then
                        If ExtraSurCharge = "YES" Then
                            DimSurCharge = WeightSurCharge + DimensionSurCharge
                            oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, DimSurCharge)
                        Else
                            If WeightSurCharge > DimensionSurCharge Then
                                oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, WeightSurCharge)
                            Else
                                oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, DimensionSurCharge)
                            End If
                        End If
                    ElseIf HighDimension > Dimension And ChargeWeight <= Weight Then
                        oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, DimensionSurCharge)
                    ElseIf HighDimension <= Dimension And ChargeWeight > Weight Then
                        oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, WeightSurCharge)
                    Else
                        oGrid.DataTable.SetValue("DIMENSIONSURCHARGES", Row, 0.0)
                    End If
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Dimension SurCharges Calculation Failed.." & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    Try
                        If BusinessObjectInfo.BeforeAction Then
                            If Me.ValidateAll() = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            Else
                                If oCompany.InTransaction = False Then oCompany.StartTransaction()
                                If Not TransactionManagement() Then
                                    If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    BubbleEvent = False
                                    Return
                                Else
                                    If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data Add ,Update Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try
                    Try
                        If BusinessObjectInfo.ActionSuccess Then
                            Try
                                Dim method As String = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
                                If method <> "None" Then
                                    Dim ErrMsg As String = ""
                                    Dim iRowIndex As Integer = 0

                                    oForm = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                    Dim sVendorCode As String = oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim

                                    Dim oGRPO As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

                                    If oGRPO.GetByKey(oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0)) Then

                                    End If
                                    Dim sPostingDate As String = oGRPO.DocDate

                                    Dim strQuery As String = String.Empty

                                    Dim svrLandedCost As SAPbobsCOM.LandedCostsService = oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService)
                                    Dim oLandedCost As SAPbobsCOM.LandedCost = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost)
                                    oLandedCost.DocumentCurrency = oGRPO.DocCurrency
                                    oLandedCost.VendorCode = oGRPO.CardCode
                                    oLandedCost.DueDate = oGRPO.DocDueDate
                                    oLandedCost.PostingDate = oGRPO.TaxDate
                                    Dim oLandedCostEntry As Long = 0
                                    Dim boolGRN As Boolean = False
                                    Dim Success As Boolean = False
                                    Dim oLandedCost_ItemLine As SAPbobsCOM.LandedCost_ItemLine
                                    Dim sTotalAmount As Double = 0
                                    If oGRPO.Lines.Count > 0 Then
                                        For Rows As Integer = 0 To oGRPO.Lines.Count - 1
                                            oLandedCost_ItemLine = oLandedCost.LandedCost_ItemLines.Add
                                            Dim LIN As String = oGRPO.Lines.LineNum
                                            oLandedCost_ItemLine.BaseDocumentType = SAPbobsCOM.LandedCostBaseDocumentTypeEnum.asGoodsReceiptPO
                                            oLandedCost_ItemLine.BaseEntry = oGRPO.Lines.DocEntry
                                            oLandedCost_ItemLine.BaseLine = Rows
                                        Next
                                    End If

                                    strQuery = "select ""AlcCode"", ""AlcName"", ""OhType"", ""LaCAllcAcc"" from ""OALC"""
                                    Dim rsetCostType As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rsetCostType.DoQuery(strQuery)
                                    Dim oLandedCost_CostLine As SAPbobsCOM.LandedCost_CostLine

                                    If rsetCostType.RecordCount > 0 Then
                                        rsetCostType.MoveFirst()
                                    End If
                                    For JKS As Integer = 0 To rsetCostType.RecordCount - 1
                                        If rsetCostType.Fields.Item("AlcCode").Value <> String.Empty Then
                                            Dim Prce1 As Double = 0
                                            oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add
                                            oLandedCost_CostLine.CostType = SAPbobsCOM.LCCostTypeEnum.asFixedCosts
                                            oLandedCost_CostLine.LandedCostCode = rsetCostType.Fields.Item("AlcCode").Value
                                            oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asQuantity
                                            If rsetCostType.Fields.Item("AlcCode").Value = "FR" Then
                                                Dim Amount As Double = 0
                                                If method = "Courier" Then
                                                    oLandedCost_CostLine.amount = CourierCost(Amount)
                                                End If
                                            ElseIf rsetCostType.Fields.Item("AlcCode").Value = "IC" Then
                                                Dim IncoTerms As String = oDBDSHeader.GetValue("U_IncoTerm", 0).Trim
                                                If IncoTerms <> "" Then
                                                    Dim DocTotal As Double = oGRPO.DocTotal
                                                    Dim Sqr As String = "select ""U_IncoTermCode"", ""U_CFRvalue"", ""U_InsuranceRate"" from ""@INCOTERM"" where ""U_IncoTermCode"" = '" & IncoTerms & "';"
                                                    Dim Rse As SAPbobsCOM.Recordset = oGFun.DoQuery(Sqr)
                                                    If Rse.RecordCount > 0 Then
                                                        Dim CFRVal As Double = Rse.Fields.Item("U_CFRvalue").Value
                                                        Dim InsVal As Double = Rse.Fields.Item("U_InsuranceRate").Value
                                                        Dim IncCharges As Double = DocTotal * CFRVal * InsVal
                                                        oLandedCost_CostLine.amount = Math.Round(IncCharges)
                                                    End If
                                                End If
                                            Else
                                                oLandedCost_CostLine.amount = 12
                                            End If
                                        End If
                                        rsetCostType.MoveNext()
                                    Next

                                    Dim oLandedCostParams As SAPbobsCOM.LandedCostParams = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams)

                                    Try
                                        oLandedCostEntry = oLandedCostParams.LandedCostNumber
                                        oLandedCostParams = svrLandedCost.AddLandedCost(oLandedCost)
                                        Success = True
                                    Catch ex As Exception
                                        oApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Success = False
                                    End Try

                                    If Success = True Then
                                        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    Else
                                        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If
                                End If
                            Catch ex As Exception
                                oApplication.StatusBar.SetText("Landed Costs Posting Method Faild " & ex.Message)
                            Finally
                            End Try
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data Add ,Update Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    Try
                        If BusinessObjectInfo.ActionSuccess Then
                            frmGRPO = oApplication.Forms.ActiveForm
                            frmGRPO.Items.Item("Packing").Visible = True
                            frmGRPO.PaneLevel = 1
                            Me.LoadGridDetails()
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data Load Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function CourierCost(ByVal CourierLandedCost As Double)
        Try
            Dim Flg As Boolean = False
            Dim Forwarder As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
            Dim FromCountry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString)
            Dim TotalWeight As Double = oDBDSHeader.GetValue("U_LCTocWeight", 0)
            Dim ZOne As Double = oDBDSHeader.GetValue("U_LCZone", 0)
            Dim ZOneWIseValue As Double = 0
            Dim FuelSurCharge As Double = 0
            Dim CouriertableCost As Double = 0
            Dim DimensionTableSurCharge As Double = 0
            Dim TotalDimSurCharge As Double = 0
            If TotalWeight > 0 Then
                ''---------------------------------------------------------------
                ''''' Courier Table Vlaues...................... Calculation....
                ''----------------------------------------------------------------
                Dim Query As String = "select ""U_ForwarderCode"", ""U_Weight"", ""U_Zone1"", ""U_Zone2"", ""U_Zone3"", ""U_Zone7"" from ""@COURIER"" where ""U_ForwarderCode"" = '" & Forwarder & "'"
                Dim Rset As SAPbobsCOM.Recordset = oGFun.DoQuery(Query)
                Dim Rset1 As SAPbobsCOM.Recordset = oGFun.DoQuery(Query)
                If Rset.RecordCount > 0 Then
                    Rset.MoveFirst()
                    For i As Integer = 0 To Rset.RecordCount - 1
                        Dim Weight As Double = Rset.Fields.Item("U_Weight").Value
                        Rset1.MoveFirst()
                        For J As Integer = 0 To Rset1.RecordCount - 1
                            If J = i + 1 Then
                                Dim nextWeight As Double = Rset1.Fields.Item("U_Weight").Value
                                If TotalWeight >= Weight And TotalWeight <= nextWeight Then
                                    If ZOne = "1" Then
                                        ZOneWIseValue = Rset1.Fields.Item("U_Zone1").Value
                                    ElseIf ZOne = "2" Then
                                        ZOneWIseValue = Rset1.Fields.Item("U_Zone2").Value
                                    ElseIf ZOne = "3" Then
                                        ZOneWIseValue = Rset1.Fields.Item("U_Zone3").Value
                                    ElseIf ZOne = "7" Then
                                        ZOneWIseValue = Rset1.Fields.Item("U_Zone7").Value
                                    End If
                                    Flg = True
                                    GoTo 89
                                End If
                            End If
                            Rset1.MoveNext()
                        Next
                        Rset.MoveNext()
                    Next
                End If
89:
                If Flg = True Then
                    CouriertableCost = TotalWeight * ZOneWIseValue
                End If
                ''---------------------------------------------------------------
                '''''Fuel Sur Charges Table Vlaues...................... Calculation....
                ''----------------------------------------------------------------
                Dim Query1 As String = "select ""U_Courier"", ""U_Country"", ""U_ByAmtOrPrcnt"", ""U_ByKgOrNet"", ""U_Rate"", ""U_Month"" from ""@FUELSURCHARGE"" Where ""U_Courier"" = '" & Forwarder & "' and ""U_Country"" = '" & FromCountry & "'"
                Dim Rset2 As SAPbobsCOM.Recordset = oGFun.DoQuery(Query1)
                If Rset2.RecordCount > 0 Then
                    Rset2.MoveFirst()
                    Dim ByAmntOrPercnt As String = Rset2.Fields.Item("U_ByAmtOrPrcnt").Value
                    Dim ByKgOrNet As String = Rset2.Fields.Item("U_ByKgOrNet").Value
                    Dim Rate As Double = Rset2.Fields.Item("U_Rate").Value
                    If ByAmntOrPercnt = "Percent" And ByKgOrNet = "KG" Then
                        FuelSurCharge = (Rate / 100) * TotalWeight
                    ElseIf ByAmntOrPercnt = "Amount" And ByKgOrNet = "KG" Then
                        FuelSurCharge = Rate * TotalWeight
                    ElseIf ByAmntOrPercnt = "Amount" And ByKgOrNet = "NET" Then
                        FuelSurCharge = Rate
                    ElseIf ByAmntOrPercnt = "Percent" And ByKgOrNet = "NET" Then
                        FuelSurCharge = (Rate / 100) * TotalWeight
                    End If
                End If

                ''---------------------------------------------------------------
                '''''Dimension Surcharges Table Vlaues...................... Calculation....
                ''----------------------------------------------------------------
                Dim Sqr1 As String = "Select ""U_Weight"", ""U_WeightSurcharge"", ""U_Dimension"", ""U_DimensionSurcharge"", ""U_ExtraCharge"" from ""@DIMENSIONSURG"" Where ""U_ForwarderOrCourier"" = '" & Forwarder & "'"
                Dim Rset3 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset3.DoQuery(Sqr1)
                If Rset3.RecordCount > 0 Then
                    Rset3.MoveFirst()
                    Dim Weight As Double = Rset3.Fields.Item("U_Weight").Value
                    Dim WeightSurCharge As Double = Rset3.Fields.Item("U_WeightSurcharge").Value
                    Dim Dimension As Double = Rset3.Fields.Item("U_Dimension").Value
                    Dim DimensionSurCharge As Double = Rset3.Fields.Item("U_DimensionSurcharge").Value
                    Dim ExtraSurCharge As String = Rset3.Fields.Item("U_ExtraCharge").Value
                    If Weight > DimensionSurCharge Then
                        DimensionTableSurCharge = Weight
                    Else
                        DimensionTableSurCharge = DimensionSurCharge
                    End If
                    If ExtraSurCharge = "YES" Then DimensionTableSurCharge = DimensionTableSurCharge * 2
                End If
                ''---------------------------------------------------------------
                '''''Total Dimension Surcharges ...................... Calculation....
                ''----------------------------------------------------------------

                Dim strSQL1 As String = String.Empty
                strSQL1 = "Select ""U_DimSurChrg"" from ""@GRPK_PDN1"" Where ""U_BaseNum""='" & frmGRPO.Items.Item("8").Specific.value & "'"
                Dim Rset4 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset4.DoQuery(strSQL1)
                If Rset4.RecordCount > 0 Then
                    Rset4.MoveFirst()
                    For i As Integer = 1 To Rset4.RecordCount
                        TotalDimSurCharge = TotalDimSurCharge + Rset4.Fields.Item("U_DimSurChrg").Value
                        Rset4.MoveNext()
                    Next
                End If
            End If
            CourierLandedCost = CouriertableCost + FuelSurCharge + DimensionTableSurCharge + (FuelSurCharge * TotalDimSurCharge)
            Return CourierLandedCost
        Catch ex As Exception
            Return CourierLandedCost
        End Try
    End Function

    Sub RightClickEvent(ByRef EventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            '
            frmGRPO = oApplication.Forms.ActiveForm
            If frmGRPO.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Right Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function TransactionManagement() As Boolean

        TransactionManagement = InsertTermsDetails()
        Return TransactionManagement

    End Function

    Function InsertTermsDetails() As Boolean
        Try

            Dim SaveDataTableRecords As SAPbouiCOM.DataTable = Nothing
            Dim dtProdutionActivity As SAPbouiCOM.DataTable = frmGRPO.DataSources.DataTables.Item("DataTable")

            For i As Integer = 0 To dtProdutionActivity.Rows.Count - 1
                If Trim(dtProdutionActivity.Columns.Item("PACKINGTYPE").Cells.Item(i).Value).Equals("") = False Then
                    SaveDataTableRecords = frmGRPO.DataSources.DataTables.Item("DataTable")
                    Exit For
                End If
            Next

            If Not SaveDataTableRecords Is Nothing Then
                Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strSQL As String = String.Empty
                strSQL = "Delete from ""@GRPK_PDN1"" Where ""U_BaseNum""='" & frmGRPO.Items.Item("8").Specific.value & "'"
                rsetCode.DoQuery(strSQL)

                Dim strSQL1 As String = String.Empty
                strSQL1 = String.Empty
                Dim sDocEntry As String = Trim(frmGRPO.Items.Item("8").Specific.value)

                For i As Integer = 0 To SaveDataTableRecords.Rows.Count - 1
                    If SaveDataTableRecords.GetValue("PACKINGTYPE", i).ToString <> String.Empty Then
                        strSQL1 = "Insert into ""@GRPK_PDN1"" (""DocEntry"",""LineId"",""U_PackType"", ""U_length"", ""U_Width"", ""U_Height"", ""U_ActWeight"", ""U_ChrWeight"", ""U_DimHeight"", ""U_DimSurChrg"",""U_BaseEntry"",""U_BaseNum"") Values  ('" & _
                                frmGRPO.Items.Item("8").Specific.value & "','" & i & "','" & SaveDataTableRecords.GetValue("PACKINGTYPE", i).ToString & "','" & _
                                SaveDataTableRecords.GetValue("LENGTH", i).ToString & "','" & SaveDataTableRecords.GetValue("WIDTH", i).ToString & "','" & SaveDataTableRecords.GetValue("HEIGHT", i).ToString & "','" & SaveDataTableRecords.GetValue("ACTUALWEIGHT", i).ToString & "','" & SaveDataTableRecords.GetValue("CHARGEABLEWEIGHT", i).ToString & "','" & SaveDataTableRecords.GetValue("HIGHDIMENSION", i).ToString & "','" & SaveDataTableRecords.GetValue("DIMENSIONSURCHARGES", i).ToString & "','" & sDocEntry & "','" & frmGRPO.Items.Item("8").Specific.value & "');"
                        rsetCode.DoQuery(strSQL1)
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Insert Terms Details Method Failed : " & ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Sub LoadGridDetails()
        Try
            oGrid.DataTable = Nothing
            sQuery = "select ""U_PackType"" As PackingType, ""U_ActWeight"" As ActualWeight,""U_length"" As Length , ""U_Width"" As Width, ""U_Height"" As Height ,  ""U_ChrWeight"" As ChargeableWeight, ""U_DimHeight"" As HighDimension, ""U_DimSurChrg"" As DimensionSurcharges  from ""@GRPK_PDN1"" where  ""U_BaseNum"" ='" & frmGRPO.Items.Item("8").Specific.value & "'"

            frmGRPO.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQuery)
            oGrid.DataTable = frmGRPO.DataSources.DataTables.Item("DataTable")
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oGrid.Columns.Item(0).Editable = True
            oGrid.Columns.Item(1).Editable = True
            oGrid.Columns.Item(2).Editable = True
            oGrid.Columns.Item(3).Editable = True
            oGrid.Columns.Item(4).Editable = True
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(7).Editable = False

            frmGRPO.Items.Item("35").Click(BoCellClickType.ct_Regular)
            oGrid.AutoResizeColumns()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Load Matrix Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

End Class

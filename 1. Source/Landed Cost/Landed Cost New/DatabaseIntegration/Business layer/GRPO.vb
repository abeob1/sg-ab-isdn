''Imports SAPbouiCOM
Imports System.Reflection
Imports System.IO
Imports System.Xml
''Imports SAPbouiCOM
Imports System.Text
Imports System.Windows.Forms

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
    Dim oGrid As SAPbouiCOM.Matrix
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
            Me.Load_Form_Controls(frmGRPO)
            Me.LoadGridDetails(frmGRPO)

            boolFormLoaded = True
        Catch ex As Exception
            boolFormLoaded = False
            oApplication.StatusBar.SetText("Load GRPO Form Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub Load_Form_Controls(ByVal frmGRPO As SAPbouiCOM.Form)
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
            '' oItem.AffectsFormMode = True

            oFolder = oItem.Specific
            oFolder.Caption = "Packing Details"
            '' oFolder.AutoPaneSelection = True
            oFolder.GroupWith(oItemRef.UniqueID)
            '' oFolder.GroupWith(112)
            '' oFolder.Pane = 1

            oItem = frmGRPO.Items.Add("Matrix", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oItem.FromPane = 23
            oItem.ToPane = 23
            oItem.Top = 170
            oItem.Left = 10
            oItem.Width = 550
            oItem.Height = 180
            oGrid = oItem.Specific

            'oItem = frmGRPO.Items.Add("Matrix", SAPbouiCOM.BoFormItemTypes.it_GRID)
            'oItem.FromPane = 23
            'oItem.ToPane = 23
            'oItem.Top = 170
            'oItem.Left = 10
            'oItem.Width = 550
            'oItem.Height = 180
            'oGrid = oItem.Specific
            'oGrid.

            ' User Data Source ...
            ''frmGRPO.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            ''oFolder.DataBind.SetBound(True, "", "FolderDS")
            ''oFolder.ValOn = True
            ''oFolder.ValOff = False

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

    Function ValidateAll(ByVal oform As SAPbouiCOM.Form, ByVal oform_UDF As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean

        Dim Rs As SAPbobsCOM.Recordset = Nothing
        Dim odtCourier As DataTable
        Dim odtForwarded As DataTable
        Dim odtDimensionSurcharge As DataTable
        Dim odtFuelSurcharge As DataTable
        Dim odtIncoTable As DataTable

        Dim oDvCourier As DataView
        Dim oDvForwarded As DataView
        Dim oDvSpotRate As DataView
        Dim oDvLumpSum As DataView

        Dim dInsurance As Double = 0
        Dim dFrieght As Double = 0

        Try
            Dim sSplt() As String
            Dim smethod As String = Trim(oform_UDF.Items.Item("U_LCMethod").Specific.value)   ''              oform.DataSources.DBDataSources.Item(0).GetValue("U_LCMethod", 0))
            Dim sBPCode As String = Trim(oform_UDF.Items.Item("U_Forwarder").Specific.value)   ''  Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_Forwarder", 0))
            Dim sZone As String = Trim(oform_UDF.Items.Item("U_LCZone").Specific.value)   ''  Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_LCZone", 0))
            Dim sCountry As String = Trim(oform_UDF.Items.Item("U_LCShipFromCountry").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_LCShipFromCountry", 0))
            Dim dCWeight As Double = Trim(oform_UDF.Items.Item("U_LCToCWeight").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_LCToCWeight", 0))
            Dim dSpotrate As Double = Trim(oform_UDF.Items.Item("U_LCSpotandLS").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_LCSpotandLS", 0))
            Dim sIncoTerms As String = Trim(oform_UDF.Items.Item("U_IncoTerm").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_IncoTerm", 0))
            Dim sPort As String = Trim(oform_UDF.Items.Item("U_Port").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_Port", 0))
            Dim sDocDate As String = Trim(oform.Items.Item("10").Specific.value)   '' Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0))

            sSplt = oform.Items.Item("22").Specific.String.ToString.Split(" ")
            Dim dDoctotal As Double = CDbl(Trim(sSplt(1))) ''CDbl(Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocTotal", 0))) - CDbl(Trim(oform.DataSources.DBDataSources.Item(0).GetValue("VatSum", 0)))
            Dim sConfig As String = sBPCode & "," & sZone & "," & sCountry & "," & CStr(dCWeight) & "," & CStr(dSpotrate) & "," & sIncoTerms & "," & sPort
            If smethod.ToUpper = "NONE" Then
                Exit Try
            End If

            If String.IsNullOrEmpty(sBPCode) Then
                Throw New ArgumentException("Forwarder Should not be Blank ...!")
            End If

            If String.IsNullOrEmpty(sZone) Then
                Throw New ArgumentException("Zone Should not be Blank ...!")
            End If

            If String.IsNullOrEmpty(sCountry) Then
                Throw New ArgumentException("Country Should not be Blank ...!")
            End If

            If String.IsNullOrEmpty(dCWeight) Then
                Throw New ArgumentException("Weight Should not be Blank ...!")
            Else
                If dCWeight = 0 Then
                    Throw New ArgumentException("Weight Should not be Blank ...!")
                End If
            End If

            If String.IsNullOrEmpty(sIncoTerms) Then
                Throw New ArgumentException("Inco Terms Should not be Blank ...!")
            End If

            Dim dDocDate As Date

            If oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                Dim ssql As String = "SELECT ""DocEntry"" FROM OPDN T0 WHERE T0.""DocNum"" = '" & oform.Items.Item("8").Specific.String & "' and  T0.""Series"" = " & oform.Items.Item("88").Specific.value & ""
                Rs = oGFun.DoQuery(ssql)

                Dim SS As String = "select ""DocEntry"" ,""BaseEntry"" from ""IPF1"" where ""BaseEntry"" = '" & Rs.Fields.Item("DocEntry").Value & "' and ""BaseType"" = '20';"
                Rs = oGFun.DoQuery(SS)
                If Rs.RecordCount > 0 Then
                    ''Throw New ArgumentException("Landed Cost Already Exists - No. " & Rs.Fields.Item("DocEntry").Value)
                    oApplication.StatusBar.SetText("Landed Cost Already Exists - No. " & Rs.Fields.Item("DocEntry").Value, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return True
                End If
            End If

            odtCourier = New DataTable
            odtCourier = p_odtCourier

            odtForwarded = New DataTable
            odtForwarded = p_odtForwarded

            odtDimensionSurcharge = New DataTable
            odtDimensionSurcharge = p_odtDimensionSurcharge

            odtFuelSurcharge = New DataTable
            odtFuelSurcharge = p_odtFuelSurcharge

            odtIncoTable = New DataTable
            odtIncoTable = p_odtIncoTable

            Dim GlobalClass As New GlobalFunctions

            ''---------------------------  Landed Cost Portion

            Select Case smethod.ToUpper

                Case "COURIER"
                    Dim dw As Double = 0
                    Dim dz As Double = 0
                    Dim dCourierTmp As Double = 0
                    Dim dDimensionSC As Double = 0
                    Dim dFuelSC As Double = 0
                    Dim dIncoTerms As Double = 0
                    Dim dWeightPlus As Double = 0

                    dDocDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", Nothing)

                    '' ------------------------- Courier
                    oDvCourier = New DataView(odtCourier)
                    oDvCourier.RowFilter = "FromDate <= '" & dDocDate & "' and ToDate >= '" & dDocDate & "'"
                    oDvCourier.Sort = "U_Weight DESC"
                    dWeightPlus = CDbl(oDvCourier.Item(0)("U_Weight").ToString)
                    If dCWeight >= dWeightPlus Then
                    Else
                        dWeightPlus = dCWeight
                    End If

                    oDvCourier.RowFilter = "U_ForwarderCode = '" & sBPCode & "' and U_Weight >= " & dWeightPlus & " "
                    oDvCourier.Sort = "U_Weight ASC"

                    If oDvCourier.Count > 0 Then
                        Select Case sZone
                            Case "1"
                                ''   dw = oDvCourier.Item(0)("U_Weight").ToString
                                dz = oDvCourier.Item(0)("U_Zone1").ToString
                            Case "2"
                                ''  dw = oDvCourier.Item(0)("U_Weight").ToString
                                dz = oDvCourier.Item(0)("U_Zone2").ToString
                            Case "3"
                                ''  dw = oDvCourier.Item(0)("U_Weight").ToString
                                dz = oDvCourier.Item(0)("U_Zone3").ToString
                            Case "4"
                                ''  dw = oDvCourier.Item(0)("U_Weight").ToString
                                dz = oDvCourier.Item(0)("U_Zone4").ToString
                            Case "7"
                                ''  dw = oDvCourier.Item(0)("U_Weight").ToString
                                dz = oDvCourier.Item(0)("U_Zone5").ToString
                        End Select

                        dCourierTmp = dCWeight * dz
                    Else
                        dCourierTmp = 0
                    End If

                    ''-------------------- Dimension SurCharge

                    oDvCourier = New DataView(odtDimensionSurcharge)
                    oDvCourier.RowFilter = "U_ForwarderOrCourier = '" & sBPCode & "' and U_Weight < " & dCWeight & ""
                    oDvCourier.Sort = "U_Weight ASC"
                    If oDvCourier.Count > 0 Then
                        If oDvCourier.Item(0)("U_ExtraCharge").ToString.ToUpper = "YES" Then
                            dDimensionSC = CDbl(oDvCourier.Item(0)("U_WeightSurcharge").ToString) + CDbl(oDvCourier.Item(0)("U_DimensionSurcharge").ToString)
                        Else
                            dDimensionSC = CDbl(oDvCourier.Item(0)("U_WeightSurcharge").ToString)
                        End If
                        If oDvCourier.Item(0)("U_Currency").ToString.ToUpper <> "SGD" Then
                            Dim dexchangerate As Double = 0
                            dexchangerate = GlobalClass.GetExchangeRate(oCompany, oDvCourier.Item(0)("U_Currency").ToString.ToUpper.Trim, dDocDate)
                            dDimensionSC = dDimensionSC * dexchangerate
                        End If
                    Else
                        dDimensionSC = 0

                    End If

                    ''-------------------- Fuel SurCharge

                    oDvCourier = New DataView(odtFuelSurcharge)
                    oDvCourier.RowFilter = "U_Courier = '" & sBPCode & "' "
                    If oDvCourier.Count > 0 Then
                        If oDvCourier.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "PERCENT" And oDvCourier.Item(0)("U_ByKgOrNet").ToString.ToUpper = "KG" Then
                            dFuelSC = dCourierTmp * (CDbl(oDvCourier.Item(0)("U_Rate").ToString) / 100)
                        ElseIf oDvCourier.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "AMOUNT" And oDvCourier.Item(0)("U_ByKgOrNet").ToString.ToUpper = "KG" Then
                            dFuelSC = dCourierTmp * CDbl(oDvCourier.Item(0)("U_Rate").ToString)
                        ElseIf oDvCourier.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "AMOUNT" And oDvCourier.Item(0)("U_ByKgOrNet").ToString.ToUpper = "NET" Then
                            dFuelSC = CDbl(oDvCourier.Item(0)("U_Rate").ToString)
                        Else
                            dFuelSC = 0
                        End If
                    Else
                        dFuelSC = 0
                    End If

                    ''-------------------- Inco Terms
                    '' sQuery = "SELECT T0.""U_IncoTermCode"", T0.""U_POPrintOutDscrip"", T0.""U_CFRvalue"", T0.""U_InsuranceRate"" FROM ""@INCOTERM""  T0"
                    oDvCourier = New DataView(odtIncoTable)
                    oDvCourier.RowFilter = "U_IncoTermCode = '" & sIncoTerms & "' "
                    If oDvCourier.Count > 0 Then
                        dIncoTerms = dDoctotal * CDbl(oDvCourier.Item(0)("U_CFRvalue").ToString) * CDbl(oDvCourier.Item(0)("U_InsuranceRate").ToString)
                    Else
                        dIncoTerms = 0
                    End If


                    p_dFrieghtAmount = 0
                    p_dInsuranceAmount = 0

                    p_dFrieghtAmount = dCourierTmp + dDimensionSC + dFuelSC
                    p_dInsuranceAmount = dIncoTerms



                Case "FORWARDER"

                    Dim oDT As New DataTable
                    Dim iMin As Double = 0
                    Dim iFlat As Double = 0
                    Dim iBasic As Double = 0
                    Dim sCurr As String = String.Empty
                    Dim odvtmp As DataView = Nothing
                    Dim dDimensionSC As Double = 0
                    Dim dIncoTerms As Double = 0
                    Dim sFilter As String = String.Empty

                    dDocDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", Nothing)

                    '' ------------------------- Forwarded

                    oDT.Columns.Add("Component", GetType(String))
                    oDT.Columns.Add("Amont", GetType(Double))


                    oDvForwarded = New DataView(odtForwarded)

                    sFilter = "U_Forwarder = '" & sBPCode & "' and U_FromDate <= '" & dDocDate & "' and U_ToDate >= '" & dDocDate & "'"
                    If Not String.IsNullOrEmpty(sCountry) Then
                        sFilter += " and U_Country= '" & sCountry.ToUpper.Trim() & "'"
                    End If

                    If Not String.IsNullOrEmpty(sPort) Then
                        sFilter += " and (U_Port= '" & sPort.ToUpper.Trim() & "' or U_Port= 'ALL') "
                    End If
                    If Not String.IsNullOrEmpty(sIncoTerms) Then
                        sFilter += "and (U_Incoterms= '" & sIncoTerms.ToUpper.Trim() & "' or U_Incoterms= 'ALL' ) "
                    End If

                    oDvForwarded.RowFilter = sFilter
                    Dim odttmp As DataTable = New DataTable
                    odttmp = oDvForwarded.ToTable
                    odvtmp = New DataView(odttmp)

                    If CDbl(oDvForwarded.Item(0)("U_Weight").ToString) > 0 Then
                        odvtmp.RowFilter = "U_Weight >= '" & dCWeight & "'"
                        odvtmp.Sort = "U_Weight ASC"
                        If odvtmp.Count > 0 Then
                            If CDbl(odvtmp.Item(0)("U_Flat").ToString) = 0 Then
                                dFrieght = CDbl(odvtmp.Item(0)("U_Basic").ToString)
                                If odvtmp.Item(0)("U_Cur").ToString() <> "SGD" Then
                                    Dim dexchangerate As Double = 0
                                    dexchangerate = GlobalClass.GetExchangeRate(oCompany, odvtmp.Item(0)("U_Cur").ToString().ToUpper.Trim, dDocDate)
                                    dFrieght = dFrieght * dexchangerate
                                End If
                            Else
                                dFrieght = CDbl(odvtmp.Item(0)("U_Flat").ToString) * dCWeight
                                If odvtmp.Item(0)("U_Cur").ToString() <> "SGD" Then
                                    Dim dexchangerate As Double = 0
                                    dexchangerate = GlobalClass.GetExchangeRate(oCompany, odvtmp.Item(0)("U_Cur").ToString().ToUpper.Trim, dDocDate)
                                    dFrieght = dFrieght * dexchangerate
                                End If
                            End If

                        Else
                            dFrieght = 0
                        End If
                    Else
                        For Each odr As DataRowView In oDvForwarded

                            iMin = odr("U_Min")
                            iFlat = (CDbl(odr("U_Flat").ToString) * Math.Ceiling(dCWeight)) + CDbl(odr("U_Basic").ToString)
                            If iMin > iFlat Then
                                iFlat = iMin
                            End If
                            If odr("U_Cur").ToString() <> "SGD" Then
                                Dim dexchangerate As Double = 0
                                dexchangerate = GlobalClass.GetExchangeRate(oCompany, odr("U_Cur").ToString().ToUpper.Trim, dDocDate)
                                iFlat = iFlat * dexchangerate
                            End If
                            dFrieght += iFlat
                        Next
                    End If


                    ''-------------------- Dimension SurCharge

                    oDvForwarded = New DataView(odtDimensionSurcharge)
                    oDvForwarded.RowFilter = "U_ForwarderOrCourier = '" & sBPCode & "' and U_Weight < " & dCWeight & ""
                    oDvForwarded.Sort = "U_Weight ASC"
                    If oDvForwarded.Count > 0 Then
                        If oDvForwarded.Item(0)("U_ExtraCharge").ToString.ToUpper = "YES" Then
                            dDimensionSC = CDbl(oDvForwarded.Item(0)("U_WeightSurcharge").ToString) + CDbl(oDvForwarded.Item(0)("U_DimensionSurcharge").ToString)
                        Else
                            dDimensionSC = CDbl(oDvForwarded.Item(0)("U_WeightSurcharge").ToString)
                        End If
                        If oDvForwarded.Item(0)("U_Currency").ToString.ToUpper <> "SGD" Then
                            Dim dexchangerate As Double = 0
                            dexchangerate = GlobalClass.GetExchangeRate(oCompany, oDvForwarded.Item(0)("U_Currency").ToString.ToUpper.Trim, dDocDate)
                            dDimensionSC = dDimensionSC * dexchangerate
                        End If
                    Else
                        dDimensionSC = 0

                    End If

                    ''-------------------- Inco Terms
                    oDvForwarded = New DataView(odtIncoTable)
                    oDvForwarded.RowFilter = "U_IncoTermCode = '" & sIncoTerms & "' "
                    If oDvForwarded.Count > 0 Then
                        dIncoTerms = dDoctotal * CDbl(oDvForwarded.Item(0)("U_CFRvalue").ToString) * CDbl(oDvForwarded.Item(0)("U_InsuranceRate").ToString)
                    Else
                        dIncoTerms = 0
                    End If


                    dFrieght += dDimensionSC
                    dInsurance = dIncoTerms

                    p_dFrieghtAmount = 0
                    p_dInsuranceAmount = 0

                    p_dFrieghtAmount = dFrieght
                    p_dInsuranceAmount = dInsurance

                Case "SPOTRATE"

                    Dim dFuelSC As Double = 0
                    Dim dCourierTmp As Double = 0
                    Dim dDimensionSC As Double = 0
                    Dim dIncoTerms As Double = 0

                    ''-------------------- Fuel SurCharge

                    If dSpotrate = 0 Then
                        Throw New ArgumentException("Spot Rate Should not be blank .......!")
                    End If


                    oDvSpotRate = New DataView(odtFuelSurcharge)
                    oDvSpotRate.RowFilter = "U_Courier = '" & sBPCode & "' "
                    If oDvSpotRate.Count > 0 Then
                        If oDvSpotRate.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "PERCENT" And oDvSpotRate.Item(0)("U_ByKgOrNet").ToString.ToUpper = "KG" Then
                            dFuelSC = dCourierTmp * (CDbl(oDvSpotRate.Item(0)("U_Rate").ToString) / 100)
                        ElseIf oDvSpotRate.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "AMOUNT" And oDvSpotRate.Item(0)("U_ByKgOrNet").ToString.ToUpper = "KG" Then
                            dFuelSC = dCourierTmp * CDbl(oDvSpotRate.Item(0)("U_Rate").ToString)
                        ElseIf oDvSpotRate.Item(0)("U_ByAmtOrPrcnt").ToString.ToUpper = "AMOUNT" And oDvSpotRate.Item(0)("U_ByKgOrNet").ToString.ToUpper = "NET" Then
                            dFuelSC = CDbl(oDvSpotRate.Item(0)("U_Rate").ToString)
                        Else
                            dFuelSC = 0
                        End If
                    Else
                        dFuelSC = 0
                    End If


                    ''-------------------- Dimension SurCharge

                    oDvSpotRate = New DataView(odtDimensionSurcharge)
                    oDvSpotRate.RowFilter = "U_ForwarderOrCourier = '" & sBPCode & "' and U_Weight < " & dCWeight & ""
                    oDvSpotRate.Sort = "U_Weight ASC"
                    If oDvSpotRate.Count > 0 Then
                        If oDvSpotRate.Item(0)("U_ExtraCharge").ToString.ToUpper = "YES" Then
                            dDimensionSC = CDbl(oDvSpotRate.Item(0)("U_WeightSurcharge").ToString) + CDbl(oDvSpotRate.Item(0)("U_DimensionSurcharge").ToString)
                        Else
                            dDimensionSC = CDbl(oDvSpotRate.Item(0)("U_WeightSurcharge").ToString)
                        End If
                        If oDvSpotRate.Item(0)("U_Currency").ToString.ToUpper <> "SGD" Then
                            Dim dexchangerate As Double = 0
                            dexchangerate = GlobalClass.GetExchangeRate(oCompany, oDvSpotRate.Item(0)("U_Currency").ToString.ToUpper.Trim, dDocDate)
                            dDimensionSC = dDimensionSC * dexchangerate
                        End If
                    Else
                        dDimensionSC = 0
                    End If

                    ''-------------------- Inco Terms
                    oDvSpotRate = New DataView(odtIncoTable)
                    oDvSpotRate.RowFilter = "U_IncoTermCode = '" & sIncoTerms & "' "
                    If oDvSpotRate.Count > 0 Then
                        dIncoTerms = dDoctotal * CDbl(oDvSpotRate.Item(0)("U_CFRvalue").ToString) * CDbl(oDvSpotRate.Item(0)("U_InsuranceRate").ToString)
                    Else
                        dIncoTerms = 0
                    End If


                    dFrieght = (dSpotrate * dCWeight) + dDimensionSC + dFuelSC
                    dInsurance = dIncoTerms


                    p_dFrieghtAmount = 0
                    p_dInsuranceAmount = 0

                    p_dFrieghtAmount = dFrieght
                    p_dInsuranceAmount = dInsurance


                Case "LUMP SUM"

                    If dSpotrate = 0 Then
                        Throw New ArgumentException("Spot Rate Should not be blank .......!")
                    End If

                    p_dFrieghtAmount = 0
                    p_dInsuranceAmount = 0

                    p_dFrieghtAmount = dSpotrate
                    p_dInsuranceAmount = 0

            End Select

            Return True

        Catch ex As Exception
            If Left(ex.Message, 19) = "Landed Cost Already" Then
                oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Else
                oApplication.StatusBar.SetText("Validation Function Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If


            Return False
        Finally
            ''  System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Rs)
            Rs = Nothing
        End Try
    End Function



#Region "SAP Events"

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
                                        Case "PackingTyp"
                                            If pVal.BeforeAction = False Then
                                                ''oGrid.DataTable.GetValue("PACKINGTYP", pVal.Row)
                                                If oGrid.Columns.Item("PackingTyp").Cells.Item(pVal.Row).Specific.String <> "" Then
                                                    oForm = oApplication.Forms.ActiveForm
                                                    Dim Currentrow As Integer = pVal.Row
                                                    Dim LastRow As Integer = oGrid.RowCount
                                                    If Currentrow = LastRow Then
                                                        oGrid.AddRow(1)
                                                        oGrid.Columns.Item("#").Cells.Item(oGrid.RowCount).Specific.String = oGrid.RowCount
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
                            frmGRPO = oApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            oDBDSHeader = frmGRPO.DataSources.DBDataSources.Item("OPDN")  'OWOR
                            If pVal.BeforeAction = True And boolFormLoaded = False Then Me.LoadXML(pVal.FormUID)

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
                                    Dim oform As SAPbouiCOM.Form = Nothing
                                    oform = oApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                                    If pVal.BeforeAction = True And (oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                                        Dim sErrDesc As String = String.Empty
                                        oform_UDF = oApplication.Forms.GetFormByTypeAndCount(-pVal.FormTypeEx, pVal.FormTypeCount)
                                        If Me.ValidateAll(oform, oform_UDF, sErrDesc) = False Then
                                            '' System.Media.SystemSounds.Asterisk.Play()
                                            BubbleEvent = False
                                            Exit Sub
                                        ElseIf oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If oCompany.InTransaction = False Then oCompany.StartTransaction()
                                            If Not TransactionManagement(oform) Then
                                                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                                Return
                                            Else
                                                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            End If
                                        End If
                                    End If
                            End Select
                        Catch ex As Exception
                            BubbleEvent = False
                            Exit Sub
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
                                        Case "Length", "Height", "Width"
                                            Dim oform As SAPbouiCOM.Form = Nothing
                                            Dim oGrid As SAPbouiCOM.Matrix = Nothing
                                            Dim method As String = String.Empty
                                            oform = oApplication.Forms.GetFormByTypeAndCount(143, pVal.FormTypeCount)
                                            oDBDSHeader = oform.DataSources.DBDataSources.Item("OPDN")
                                            oGrid = oform.Items.Item("Matrix").Specific
                                            If pVal.BeforeAction = True And pVal.ItemChanged Then
                                                method = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
                                                ''oGrid.Columns.Item("PackingTyp").Cells.Item(pVal.Row).Specific.String
                                                '' If method.ToUpper = "NONE" And oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row) > 0 Then
                                                If method.ToUpper = "NONE" And CDbl(oGrid.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String) > 0 Then
                                                    oApplication.StatusBar.SetText("Method should not be Blank ...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If

                                                HighDimensionCalculation(pVal.Row, oform, oGrid)
                                                ChargeableWeightCalc(pVal.Row, oform, oGrid)
                                                DimensionSurChargesCalc(pVal.Row, oform, oGrid)
                                                TotalChargeableValue(pVal.FormTypeCount, oform, oGrid)
                                                oGrid.Columns.Item(pVal.ColUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        Case "ActualWeig"

                                            If pVal.BeforeAction = True And pVal.ItemChanged Then
                                                Dim oform As SAPbouiCOM.Form = Nothing
                                                Dim oGrid As SAPbouiCOM.Matrix = Nothing
                                                Dim method As String = String.Empty
                                                oform = oApplication.Forms.GetFormByTypeAndCount(143, pVal.FormTypeCount)
                                                oDBDSHeader = oform.DataSources.DBDataSources.Item("OPDN")
                                                oGrid = oform.Items.Item("Matrix").Specific
                                                method = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
                                                If method.ToUpper = "NONE" And CDbl(oGrid.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String) > 0 Then
                                                    oApplication.StatusBar.SetText("Method should not be Blank ...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                                ChargeableWeightCalc(pVal.Row, oform, oGrid)
                                                DimensionSurChargesCalc(pVal.Row, oform, oGrid)
                                                TotalChargeableValue(pVal.FormTypeCount, oform, oGrid)

                                                '   oGrid.Columns.Item("ACTUALWEIGHT").Click(BoCellClickType.ct_Regular)
                                                ''oGrid.GetCellFocus()

                                            End If
                                    End Select
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Vaidate Event Failed :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Try
                            Dim oform As SAPbouiCOM.Form = Nothing
                            oform = oApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            Select Case pVal.ItemUID
                                Case "1"
                                    If pVal.ActionSuccess And oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        Me.LoadGridDetails(oform)
                                    End If
                                Case "Packing"
                                    If pVal.BeforeAction = False And pVal.InnerEvent = False Then
                                        Dim ofds As SAPbouiCOM.Folder = Nothing
                                        ofds = oform.Items.Item("Packing").Specific
                                        ofds.Select()
                                        oform.PaneLevel = 23
                                        oform.Refresh()

                                        'BubbleEvent = False
                                        'Exit Sub
                                        ' '' oform.PaneLevel = 23
                                        '' oform.Items.Item("Packing").AffectsFormMode = True
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
                        Dim oform As SAPbouiCOM.Form = Nothing
                        oform = oApplication.Forms.ActiveForm
                        Me.LoadGridDetails(oform)
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    Try
                        If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                            Dim Rs As SAPbobsCOM.Recordset = Nothing
                            Try
                                oForm = oApplication.Forms.ActiveForm
                                Dim smethod As String = Trim(oForm.DataSources.DBDataSources.Item(0).GetValue("U_LCMethod", 0))
                                Dim sDocEntry As String = Trim(oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0))
                                Dim sErrDesc As String = String.Empty
                                Dim sLDocEntry As String = String.Empty

                                If smethod.ToUpper <> "NONE" Then
                                    Dim SS As String = "select ""BaseEntry"" from ""IPF1"" where ""BaseEntry"" = '" & sDocEntry & "' and ""BaseType"" = '20';"
                                    Rs = oGFun.DoQuery(SS)
                                    If Rs.RecordCount = 0 Then

                                        If LandedCostPosting(sDocEntry, p_dFrieghtAmount, p_dInsuranceAmount, sLDocEntry, sErrDesc) <> True Then
                                            Throw New ArgumentException(sErrDesc)
                                        Else
                                            oApplication.MessageBox("Landed Cost No " & sLDocEntry & " has been posted for this GRPO ", 1, "Ok")
                                        End If
                                    End If
                                End If

                            Catch ex As Exception
                                oApplication.StatusBar.SetText("Landed Costs Posting Method Faild " & ex.Message)
                            Finally
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Rs)
                                Rs = Nothing
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
                            Me.LoadGridDetails(frmGRPO)
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

#End Region


#Region "Packing Type Calculation"



    Sub TotalChargeableValue(ByVal iFormtypecount As Integer, ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Matrix)
        Try
            oForm.Freeze(True)
            frmGRPOUDF = oApplication.Forms.GetFormByTypeAndCount(-143, iFormtypecount)
            Dim Row1 As Integer = oGrid.DataTable.Rows.Count
            If Row1 > 0 Then
                Dim TotalValue As Double = 0
                For Row As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    ''CDbl(oGrid.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)
                    ''Dim ChargeWeight As Double = oGrid.DataTable.GetValue("ChaWeight", Row)
                    Dim ChargeWeight As Double = CDbl(oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String)
                    If ChargeWeight > 0 Then
                        TotalValue = TotalValue + ChargeWeight
                    End If
                Next
                'frmGRPO.Items.Item("t_LCTCWght").Specific.string = TotalValue
                'frmGRPO.Items.Item("t_LCTCWght").Click(BoCellClickType.ct_Regular)
                frmGRPOUDF.Items.Item("U_LCToCWeight").Specific.string = TotalValue
                frmGRPOUDF.Items.Item("U_LCToCWeight").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Total Chargeable Weight Calculation Failed.." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub ChargeableWeightCalc(ByVal Row As Integer, ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Matrix)
        Try
            oForm.Freeze(True)
            Dim method As String = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
            Dim sBPCode As String = Trim(oDBDSHeader.GetValue("CardCode", 0).ToString.ToUpper)
            Dim sCountry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString.ToUpper)
            Dim sZone As String = Trim(oDBDSHeader.GetValue("U_LCZone", 0).ToString)

            Dim ActualWeight As Double = CDbl(oGrid.Columns.Item("ActualWeig").Cells.Item(Row).Specific.String)
            Dim LengthVal As Double = CDbl(oGrid.Columns.Item("Length").Cells.Item(Row).Specific.String)
            Dim HeightVal As Double = CDbl(oGrid.Columns.Item("Height").Cells.Item(Row).Specific.String)
            Dim WidthVal As Double = CDbl(oGrid.Columns.Item("Width").Cells.Item(Row).Specific.String)
            If method = "Courier" Then
                If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                    Dim TotalWeight As Double = ((LengthVal * HeightVal * WidthVal) / 5000)
                    If ActualWeight < TotalWeight Then
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = TotalWeight
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, TotalWeight)
                    Else
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                    End If
                Else
                    If ActualWeight > 0 Then
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                        ''  oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                    Else
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = 0.0
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, 0.0)
                    End If

                End If
            ElseIf method = "Forwarder" Then
                If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                    Dim oDV As DataView
                    oDV = New DataView(p_odtZoneDetails)
                    '' sQuery = "SELECT UPPER( T0.""U_BPCOde"") ""U_BPCOde"", upper(T0.""U_Country"") ""U_Country"", T0.""U_Zone"", T0.""U_Volumetric"" FROM ""@ZONEDETAILS""  T0"
                    oDV.RowFilter = "U_BPCOde='" & sBPCode & "' and U_Country='" & sCountry & "' and U_Zone='" & sZone & "' "
                    Dim BPCode As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
                    Dim ShipToCntry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString)
                    Dim TotalWeight As Double = 0.0

                    If oDV.Count > 0 Then
                        TotalWeight = ((LengthVal * HeightVal * WidthVal) / CDbl(oDV.Item(0)("U_Volumetric").ToString))
                    End If

                    If ActualWeight < TotalWeight Then
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = TotalWeight
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, TotalWeight)
                    Else
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                        ''oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                    End If
                Else
                    If ActualWeight > 0 Then
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                    Else
                        oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = 0.0
                        '' oGrid.DataTable.SetValue("ChaWeight", Row, 0.0)
                    End If
                End If
            ElseIf ActualWeight > 0 Then
                '' oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub ChargeableWeightCalcFromMethod(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            'Dim oMatrix As SAPbouiCOM.Grid
            Dim method As String = Trim(oDBDSHeader.GetValue("U_LCMethod", 0).ToString)
            Dim sBPCode As String = Trim(oDBDSHeader.GetValue("CardCode", 0).ToString.ToUpper)
            Dim sCountry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString.ToUpper)
            Dim sZone As String = Trim(oDBDSHeader.GetValue("U_LCZone", 0).ToString)

            oGrid = oForm.Items.Item("Matrix").Specific
            Dim Row1 As Integer = oGrid.RowCount
            If Row1 > 0 Then
                For Row As Integer = 1 To Row1
                    Dim ActualWeight As Double = CDbl(oGrid.Columns.Item("ActualWeig").Cells.Item(Row).Specific.String)
                    Dim LengthVal As Double = CDbl(oGrid.Columns.Item("Length").Cells.Item(Row).Specific.String)
                    Dim HeightVal As Double = CDbl(oGrid.Columns.Item("Height").Cells.Item(Row).Specific.String)
                    Dim WidthVal As Double = CDbl(oGrid.Columns.Item("Width").Cells.Item(Row).Specific.String)

                    If method = "Courier" Then
                        If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then
                            Dim TotalWeight As Double = ((LengthVal * HeightVal * WidthVal) / 5000)
                            If ActualWeight < TotalWeight Then
                                oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = TotalWeight '
                                ' oGrid.DataTable.SetValue("ChaWeight", Row, TotalWeight)
                            Else
                                oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                                ''oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                            End If
                        End If
                    ElseIf method = "Forwarder" Then
                        If LengthVal > 0 And HeightVal > 0 And WidthVal > 0 Then

                            Dim oDV As DataView
                            oDV = New DataView(p_odtZoneDetails)
                            '' sQuery = "SELECT UPPER( T0.""U_BPCOde"") ""U_BPCOde"", upper(T0.""U_Country"") ""U_Country"", T0.""U_Zone"", T0.""U_Volumetric"" FROM ""@ZONEDETAILS""  T0"
                            oDV.RowFilter = "U_BPCOde='" & sBPCode & "' and U_Country='" & sCountry & "' and U_Zone='" & sZone & "' "
                            Dim BPCode As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
                            Dim ShipToCntry As String = Trim(oDBDSHeader.GetValue("U_LCShipFromCountry", 0).ToString)
                            Dim TotalWeight As Double = 0.0

                            If oDV.Count > 0 Then
                                TotalWeight = ((LengthVal * HeightVal * WidthVal) / CDbl(oDV.Item(0)("U_Volumetric").ToString))
                            End If

                            If ActualWeight < TotalWeight Then
                                oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = TotalWeight
                                '' oGrid.DataTable.SetValue("ChaWeight", Row, TotalWeight)
                            Else
                                oGrid.Columns.Item("ChaWeight").Cells.Item(Row).Specific.String = ActualWeight
                                '' oGrid.DataTable.SetValue("ChaWeight", Row, ActualWeight)
                            End If
                        End If
                    End If
                Next
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub HighDimensionCalculation(ByVal Row As Integer, ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Matrix)
        Try
            oForm.Freeze(True)
            Dim LengthVal As Double = CDbl(oGrid.Columns.Item("Length").Cells.Item(Row).Specific.String)
            Dim HeightVal As Double = CDbl(oGrid.Columns.Item("Height").Cells.Item(Row).Specific.String)
            Dim WidthVal As Double = CDbl(oGrid.Columns.Item("Width").Cells.Item(Row).Specific.String)
            If (LengthVal > HeightVal) And (LengthVal > WidthVal) Then
                oGrid.Columns.Item("HighDimen").Cells.Item(Row).Specific.String = LengthVal
                '' oGrid.DataTable.SetValue("HighDimen", Row, LengthVal)
            ElseIf (HeightVal > LengthVal) And (HeightVal > WidthVal) Then
                oGrid.Columns.Item("HighDimen").Cells.Item(Row).Specific.String = HeightVal
                ''  oGrid.DataTable.SetValue("HighDimen", Row, HeightVal)
            ElseIf (WidthVal > LengthVal) And (WidthVal > HeightVal) Then
                oGrid.Columns.Item("HighDimen").Cells.Item(Row).Specific.String = WidthVal
                ''oGrid.DataTable.SetValue("HighDimen", Row, WidthVal)
            Else
                oGrid.Columns.Item("HighDimen").Cells.Item(Row).Specific.String = 0.0
                '' oGrid.DataTable.SetValue("HighDimen", Row, 0.0)
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("High Dimension Calculation Failed.." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Sub DimensionSurChargesCalc(ByVal Row As Integer, ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Matrix)
        Try
            oForm.Freeze(True)
            Dim DimSurCharge As Double = 0.0
            Dim BPForwarder As String = Trim(oDBDSHeader.GetValue("U_Forwarder", 0).ToString)
            Dim HighDimension As Double = oGrid.DataTable.GetValue("HighDimen", Row)
            Dim ChargeWeight As Double = oGrid.DataTable.GetValue("ChaWeight", Row)
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
            oApplication.StatusBar.SetText("Dimension SurCharges Calculation Failed.." & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region


#Region "Landed Cost Posting"

    Function LandedCostPosting(ByVal sGRPODocEntry As String, ByVal sFreight As Double, ByVal sInsurance As Double, ByRef sLDocentry As String, ByRef sErrDesc As String) As Boolean

        Dim sFuncName As String = String.Empty
        Dim oGRPO As SAPbobsCOM.Documents = Nothing
        Dim svrLandedCost As SAPbobsCOM.LandedCostsService = Nothing
        Dim oLandedCost As SAPbobsCOM.LandedCost = Nothing
        Dim strQuery As String = String.Empty
        Dim oLandedCostParams As SAPbobsCOM.LandedCostParams = Nothing

        Try
            sFuncName = "Landed Cost Posting"
            If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("Landed Cost Posting Started", sFuncName)
            oGRPO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
            If oGRPO.GetByKey(sGRPODocEntry) Then
                svrLandedCost = oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService)
                oLandedCost = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost)

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
                        oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add
                        oLandedCost_CostLine.CostType = SAPbobsCOM.LCCostTypeEnum.asFixedCosts
                        oLandedCost_CostLine.LandedCostCode = rsetCostType.Fields.Item("AlcCode").Value
                        oLandedCost_CostLine.AllocationBy = SAPbobsCOM.LandedCostAllocationByEnum.asQuantity
                        If rsetCostType.Fields.Item("AlcCode").Value = "FR" Then
                            If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("p_dFrieghtAmount " & Math.Round(p_dFrieghtAmount, 2), sFuncName)
                            oLandedCost_CostLine.amount = Math.Round(p_dFrieghtAmount, 2)

                        ElseIf rsetCostType.Fields.Item("AlcCode").Value = "IC" Then

                            If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("p_dInsuranceAmount" & Math.Round(p_dInsuranceAmount, 2), sFuncName)
                            oLandedCost_CostLine.amount = Math.Round(p_dInsuranceAmount, 2)
                        End If
                    End If
                    rsetCostType.MoveNext()
                Next

                oLandedCostParams = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams)

                If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("Landed Cost Creation", sFuncName)

                oLandedCostEntry = oLandedCostParams.LandedCostNumber
                oLandedCostParams = svrLandedCost.AddLandedCost(oLandedCost)
                sLDocentry = oLandedCostParams.LandedCostNumber
                Success = True
                If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            End If

            sErrDesc = String.Empty
            LandedCostPosting = True

        Catch ex As Exception
            sErrDesc = ex.Message
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            If p_iDebugMode = DEBUG_ON Then Call oGFun.WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
            LandedCostPosting = False

        Finally
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oGRPO)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(svrLandedCost)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLandedCost)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLandedCostParams)

            oGRPO = Nothing
            svrLandedCost = Nothing
            oLandedCost = Nothing
            oLandedCostParams = Nothing

        End Try

    End Function

#End Region






    Function TransactionManagement(ByVal oform As SAPbouiCOM.Form) As Boolean

        TransactionManagement = InsertTermsDetails(oform)
        Return TransactionManagement

    End Function

    Function InsertTermsDetails(ByVal oform As SAPbouiCOM.Form) As Boolean
        Try

            Dim SaveDataTableRecords As SAPbouiCOM.DataTable = Nothing
            frmGRPO = oform
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
                                SaveDataTableRecords.GetValue("Length", i).ToString & "','" & SaveDataTableRecords.GetValue("Width", i).ToString & "','" & SaveDataTableRecords.GetValue("Height", i).ToString & "','" & SaveDataTableRecords.GetValue("ActualWeig", i).ToString & "','" & SaveDataTableRecords.GetValue("ChaWeight", i).ToString & "','" & SaveDataTableRecords.GetValue("HighDimen", i).ToString & "','" & SaveDataTableRecords.GetValue("DIMENSIONSURCHARGES", i).ToString & "','" & sDocEntry & "','" & frmGRPO.Items.Item("8").Specific.value & "');"
                        rsetCode.DoQuery(strSQL1)
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Insert Terms Details Method Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Sub LoadGridDetails(ByVal oform As SAPbouiCOM.Form)
        Try
            oGrid = oform.Items.Item("Matrix").Specific

            ''  oGrid.DataTable = Nothing
            ''Try
            ''    oform.DataSources.DataTables.Add("oMatrixDT")
            ''Catch ex As Exception
            ''End Try
            ''oform.DataSources.DataTables.Item("oMatrixDT").Clear()

            sQuery = "select  ""LineId"" , ""U_PackType"" As PackingType, ""U_ActWeight"" As ActualWeight,""U_length"" As Length , ""U_Width"" As Width, ""U_Height"" As Height ,  ""U_ChrWeight"" As ChargeableWeight, ""U_DimHeight"" As HighDimension, ""U_DimSurChrg"" As DimensionSurcharges  from ""@GRPK_PDN1"" where  ""U_BaseNum"" ='" & oform.Items.Item("8").Specific.value & "'"
            '' oform.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(sQuery)
            oform.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQuery)

            Dim oColumn As SAPbouiCOM.Column = oGrid.Columns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.Editable = False
            oColumn.DataBind.Bind("DataTable", "LineId")
            oColumn.TitleObject.Caption = "#"
            oColumn = oGrid.Columns.Add("PackingTyp", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.Bind("DataTable", "PackingType")
            oColumn.TitleObject.Caption = "Packing Type"
            oColumn = oGrid.Columns.Add("ActualWeig", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.Bind("DataTable", "ActualWeight")
            oColumn.TitleObject.Caption = "Actual Weight"
            oColumn = oGrid.Columns.Add("Length", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.Bind("DataTable", "Length")
            oColumn.TitleObject.Caption = "Length"
            oColumn = oGrid.Columns.Add("Width", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.Bind("DataTable", "Width")
            oColumn.TitleObject.Caption = "Width"

            ' sQuery = "select ""U_PackType"" As PackingType, ""U_ActWeight"" As ActualWeight,""U_length"" As Length , ""U_Width"" As Width, ""U_Height"" As Height ,
            '""U_ChrWeight"" As ChargeableWeight, ""U_DimHeight"" As HighDimension, ""U_DimSurChrg"" As DimensionSurcharges  from ""@GRPK_PDN1"" where  ""U_BaseNum"" ='" & oform.Items.Item("8").Specific.value & "'"

            oColumn = oGrid.Columns.Add("Height", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.DataBind.Bind("DataTable", "Height")
            oColumn.TitleObject.Caption = "Height"
            oColumn = oGrid.Columns.Add("ChaWeight", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.Editable = False
            oColumn.DataBind.Bind("DataTable", "ChargeableWeight")
            oColumn.TitleObject.Caption = "Chargeable Weight"
            oColumn = oGrid.Columns.Add("HighDimen", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.Editable = False
            oColumn.DataBind.Bind("DataTable", "HighDimension")
            oColumn.TitleObject.Caption = "High Dimension"
            oColumn = oGrid.Columns.Add("DimSurch", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.Editable = False
            oColumn.DataBind.Bind("DataTable", "DimensionSurcharges")
            oColumn.TitleObject.Caption = "Dimension Surcharges"
            oColumn.DisplayDesc = False
            oGrid.LoadFromDataSource()
            oGrid.AddRow(1)
            oGrid.Columns.Item("#").Cells.Item(1).Specific.String = oGrid.RowCount
           
            ''oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            ''oGrid.Columns.Item(0).Editable = True
            ''oGrid.Columns.Item(1).Editable = True
            ''oGrid.Columns.Item(2).Editable = True
            ''oGrid.Columns.Item(3).Editable = True
            ''oGrid.Columns.Item(4).Editable = True
            ''oGrid.Columns.Item(5).Editable = False
            ''oGrid.Columns.Item(6).Editable = False
            ''oGrid.Columns.Item(7).Editable = False

            frmGRPO.Items.Item("35").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oGrid.AutoResizeColumns()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Load Matrix Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

End Class

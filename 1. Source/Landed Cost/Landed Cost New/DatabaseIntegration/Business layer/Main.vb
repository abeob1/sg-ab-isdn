Module Main
#Region "... Main ..."
    Sub Main()
        Dim sErrDesc As String = String.Empty

        Try
            oGFun.SetApplication()
            If Not oGFun.CookieConnect() = 0 Then
                oApplication.MessageBox("DIAPI Connection Failed")
                End
            End If
            If Not oGFun.ConnectionContext() = 0 Then
                oApplication.MessageBox("Company Connection Failed")
                End
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Application Not Found")
            System.Windows.Forms.Application.ExitThread()
        Finally
        End Try


        Dim orsLanded As SAPbobsCOM.Recordset = Nothing

        Dim GlobalClass As New GlobalFunctions


        ''-------------Courier
        Dim sQuery = "SELECT  T0.""Code"" ,upper(T0.""U_ForwarderCode"") ""U_ForwarderCode"", T0.""U_Weight"", T0.""U_Zone1"", T0.""U_Zone2"", T0.""U_Zone3"", T0.""U_Zone4"", T0.""U_Zone7"", to_dats(T0.""U_FDate"") ""U_FDate"" , to_dats(T0.""U_TDate"")  ""U_TDate""  , T0.""U_ApplyFlat""  FROM  ""@COURIER""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtCourier = New DataTable
        sErrDesc = String.Empty
        p_odtCourier = GlobalClass.Courier_ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If
        ''-------------- Forwarded
        sQuery = "SELECT upper(T0.""U_Forwarder"") ""U_Forwarder"", upper(T0.""U_Country"") ""U_Country"", upper(T0.""U_Port"") ""U_Port"", upper(T0.""U_Incoterms"") ""U_Incoterms"", upper( T0.""U_Cur"") ""U_Cur"", upper(T0.""U_Component"") ""U_Component"", upper(T0.""U_ByAmtOrPrcnt"") ""U_ByAmtOrPrcnt"", upper(T0.""U_ByKgOrNet"") ""U_ByKgOrNet"", T0.""U_Weight"", T0.""U_Min"", T0.""U_Flat"", T0.""U_Basic"", to_dats(T0.""U_FromDate"") ""U_FromDate"", to_dats(T0.""U_ToDate"") ""U_ToDate"", T0.""Code"" FROM  ""@FORWARDER""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtForwarded = New DataTable
        sErrDesc = String.Empty
        p_odtForwarded = GlobalClass.Forwarded_ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If
        ''-------------- Dimension Surcharge
        sQuery = "SELECT upper(T0.""U_ForwarderOrCourier"") ""U_ForwarderOrCourier"", upper(T0.""U_Currency"") ""U_Currency"", upper(T0.""U_ExtraCharge"") ""U_ExtraCharge"", T0.""U_Weight"", T0.""U_WeightSurcharge"", T0.""U_Dimension"", T0.""U_DimensionSurcharge"", T0.""Code"" FROM  ""@DIMENSIONSURG""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtDimensionSurcharge = New DataTable
        sErrDesc = String.Empty
        p_odtDimensionSurcharge = GlobalClass.ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If
        ''-------------  Fuel SurCharge
        sQuery = "SELECT upper(T0.""U_Courier"") ""U_Courier"", T0.""U_ByAmtOrPrcnt"", upper(T0.""U_ByKgOrNet"") ""U_ByKgOrNet"", T0.""U_Rate"",  to_dats(T0.""U_FromDate"") ""U_FromDate"", to_dats(T0.""U_ToDate"") ""U_ToDate"" FROM ""@FUELSURCHARGE""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtFuelSurcharge = New DataTable
        sErrDesc = String.Empty
        p_odtFuelSurcharge = GlobalClass.ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If
        '' ------------- Inco Table
        sQuery = "SELECT upper(T0.""U_IncoTermCode"") ""U_IncoTermCode"", upper(T0.""U_POPrintOutDscrip"") ""U_POPrintOutDscrip"", T0.""U_CFRvalue"", T0.""U_InsuranceRate"" FROM ""@INCOTERM""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtIncoTable = New DataTable
        sErrDesc = String.Empty
        p_odtIncoTable = GlobalClass.ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If

        '' ------------- Zone Table

        sQuery = "SELECT UPPER( T0.""U_BPCOde"") ""U_BPCOde"", upper(T0.""U_Country"") ""U_Country"", T0.""U_Zone"", T0.""U_Volumetric"" FROM ""@ZONEDETAILS""  T0"
        orsLanded = oGFun.DoQuery(sQuery)
        p_odtZoneDetails = New DataTable
        sErrDesc = String.Empty
        p_odtZoneDetails = GlobalClass.ConvertRecordset(orsLanded, sErrDesc)
        If Not String.IsNullOrEmpty(sErrDesc) Then
            Throw New ArgumentException(sErrDesc)
        End If


        Try
            Try
                Dim oTableCreation As New TableCreation
                EventHandler.SetEventFilter()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                System.Windows.Forms.Application.ExitThread()
            Finally
            End Try
            oApplication.StatusBar.SetText("Connected Successfully.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Addon Connection Method Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region
End Module


Public Class TableCreation
    Dim DocType As String(,) = New String(,) {{"YES", "YES"}, {"NO", "NO"}}
    Dim DType As String(,) = New String(,) {{"Amount", "Amount"}, {"Percent", "Percent"}}
    Dim DType1 As String(,) = New String(,) {{"KG", "KG"}, {"NET", "NET"}}
    Dim MethodType As String(,) = New String(,) {{"None", "None"}, {"Courier", "Courier"}, {"Forwarder", "Forwarder"}, {"SpotRate", "SpotRate"}, {"Lump Sum", "Lump Sum"}}
    Sub New()

        Me.TableCreation()
        Me.UDFCreation()
        Me.GRPOpacking()

    End Sub

#Region " Table Creation "

    Sub TableCreation()
        Try

            Me.Zone()
            Me.Courier()
            Me.DimensionSurCharges()
            Me.Forwarder()
            Me.FuelSurCharge()
            Me.IncoTerms()

        Catch ex As Exception
            oApplication.StatusBar.SetText("Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

#End Region

    Sub Zone()
        Try
            CreateTable("ZONEDETAILS", "Zone", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@ZONEDETAILS", "BPCOde", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@ZONEDETAILS", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@ZONEDETAILS", "Zone", "Zone", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            CreateUserFields("@ZONEDETAILS", "Volumetric", "Volumetric", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Rate)

        Catch ex As Exception
            oApplication.StatusBar.SetText("Zone Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub Courier()
        Try
            CreateTable("COURIER", "Courier", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@COURIER", "ForwarderCode", "Forwarder Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@COURIER", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "ApplyFlat", "ApplyFlat", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            CreateUserFields("@COURIER", "Zone1", "Zone1", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "Zone2", "Zone2", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "Zone3", "Zone3", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "Zone4", "Zone4", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "Zone7", "Zone7", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@COURIER", "FDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, SAPbobsCOM.BoFldSubTypes.st_None)
            CreateUserFields("@COURIER", "TDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, SAPbobsCOM.BoFldSubTypes.st_None)

        Catch ex As Exception
            oApplication.StatusBar.SetText("Courier Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub DimensionSurCharges()
        Try
            CreateTable("DIMENSIONSURG", "Dimension Sur Charge", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@DIMENSIONSURG", "ForwarderOrCourier", "Forwarder Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@DIMENSIONSURG", "Currency", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            CreateUserFieldsComboBox("@DIMENSIONSURG", "ExtraCharge", "Extra Charges", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , DocType, "")
            CreateUserFields("@DIMENSIONSURG", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@DIMENSIONSURG", "WeightSurcharge", "Weight Surcharge", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            CreateUserFields("@DIMENSIONSURG", "Dimension", "Dimension", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@DIMENSIONSURG", "DimensionSurcharge", "Dimension Surcharge", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)

        Catch ex As Exception
            oApplication.StatusBar.SetText("Dimension SurCharges Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub Forwarder()
        Try
            CreateTable("FORWARDER", "Forwarder", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@FORWARDER", "Forwarder", "Forwarder Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@FORWARDER", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            CreateUserFields("@FORWARDER", "Port", "Port", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            CreateUserFields("@FORWARDER", "Incoterms", "Inco Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            CreateUserFields("@FORWARDER", "Cur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            CreateUserFields("@FORWARDER", "Component", "Component", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFieldsComboBox("@FORWARDER", "ByAmtOrPrcnt", "By Amount/Percent", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, , , DType, "")
            CreateUserFieldsComboBox("@FORWARDER", "ByKgOrNet", "By KG/Net", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, , , DType1, "")
            CreateUserFields("@FORWARDER", "Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@FORWARDER", "Min", "Minimum", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@FORWARDER", "Max", "Minimum", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@FORWARDER", "Flat", "Flat", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@FORWARDER", "Basic", "Basic", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@FORWARDER", "FromDate", "FromDate", SAPbobsCOM.BoFieldTypes.db_Date)
            CreateUserFields("@FORWARDER", "ToDate", "ToDate", SAPbobsCOM.BoFieldTypes.db_Date)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Forwarder Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub FuelSurCharge()
        Try
            CreateTable("FUELSURCHARGE", "Fuel SurCharge", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@FUELSURCHARGE", "Courier", "Courier Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFieldsComboBox("@FUELSURCHARGE", "ByAmtOrPrcnt", "By Amount/Percent", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, , , DType, "")
            CreateUserFieldsComboBox("@FUELSURCHARGE", "ByKgOrNet", "By KG/Net", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, , , DType1, "")
            CreateUserFields("@FUELSURCHARGE", "Rate", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Rate)
            CreateUserFields("@FUELSURCHARGE", "FromDate", "FromDate", SAPbobsCOM.BoFieldTypes.db_Date)
            CreateUserFields("@FUELSURCHARGE", "ToDate", "ToDate", SAPbobsCOM.BoFieldTypes.db_Date)

        Catch ex As Exception
            oApplication.StatusBar.SetText("Fuel SurCharge Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub IncoTerms()
        Try
            CreateTable("INCOTERM", "Inco Terms", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            CreateUserFields("@INCOTERM", "IncoTermCode", "Inco Term Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            CreateUserFields("@INCOTERM", "POPrintOutDscrip", "Document Print Out Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 249)
            CreateUserFields("@INCOTERM", "CFRvalue", "CFR Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)
            CreateUserFields("@INCOTERM", "InsuranceRate", "Insurance Rate", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Rate)

        Catch ex As Exception
            oApplication.StatusBar.SetText("Inco Terms Table Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

#Region "       ...GRPO Packing ...        "

    Sub GRPOpacking()
        Try
            Me.GRPOpackingHeader()
            Me.GRPOpackingDetails()
            If Not UDOExists("GRPK") Then
                Dim findAliasNDescription As String(,) = New String(,) {{"DocNum", "DocNum"}, {"DocEntry", "DocEntry"}}
                RegisterUDO("GRPK", "GRPO Packing Details", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "GRPK_OPDN", "GRPK_PDN1")
                findAliasNDescription = Nothing
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText("Table Creation: ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Sub GRPOpackingHeader()
        CreateTable("GRPK_OPDN", "GRPO Header", SAPbobsCOM.BoUTBTableType.bott_Document)

    End Sub

    Sub GRPOpackingDetails()
        Try
            CreateTable("GRPK_PDN1", "GRPO Packing Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            Dim sTableName As String = "@GRPK_PDN1"

            CreateUserFields("@GRPK_PDN1", "BaseEntry", "Base Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 11)
            CreateUserFields("@GRPK_PDN1", "BaseNum", "Base Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 11)

            CreateUserFields("@GRPK_PDN1", "PackType", "Packing Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("@GRPK_PDN1", "ActWeight", "Actual Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@GRPK_PDN1", "length", "Length", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)

            CreateUserFields("@GRPK_PDN1", "Width", "Width", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@GRPK_PDN1", "Height", "Height", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@GRPK_PDN1", "ChrWeight", "Chargeable Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)
            CreateUserFields("@GRPK_PDN1", "DimHeight", "Height Dimension", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Measurement)

            CreateUserFields("@GRPK_PDN1", "DimSurChrg", "Dimension SurCharges", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Price)

        Catch ex As Exception
            oApplication.StatusBar.SetText("TableCreation GRPO Packing Details Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub


#End Region

#Region "UDF Creation "
    Sub UDFCreation()
        Try
            CreateUserFields("OCRD", "IncoTerm", "Inco Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("OPDN", "IncoTerm", "Inco Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFieldsComboBox("OPDN", "LCMethod", "Method", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, , , MethodType, "")
            CreateUserFields("OPDN", "Forwarder", "Forwarder", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("OPDN", "LCZone", "LC Zone", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("OPDN", "LCShipFromCountry", "Ship From Country", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            CreateUserFields("OPDN", "LCToCWeight", "Total Charge Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
            CreateUserFields("OPDN", "LCSpotandLS", "Spot Rate/ Lump Sum", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Sum)
        Catch ex As Exception
            oApplication.StatusBar.SetText("UDF Creation Failed: " & ex.Message)
        Finally
        End Try
    End Sub
#End Region

    Function CreateUserFields(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal DefaultValue As String = "") As Boolean
        Try
            If TableName.StartsWith("@") = True Then
                If Not Me.ColumnExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    If DefaultValue <> "" Then v_UserField.DefaultValue = DefaultValue
                    If FieldName = "ApplyFlat" Then
                        v_UserField.ValidValues.SetCurrentLine(0)
                        v_UserField.ValidValues.Value = "Y"
                        v_UserField.ValidValues.Description = "Y"
                        v_UserField.ValidValues.Add()
                        v_UserField.ValidValues.SetCurrentLine(1)
                        v_UserField.ValidValues.Value = "N"
                        v_UserField.ValidValues.Description = "N"
                        v_UserField.ValidValues.Add()
                    End If

                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField masterid" & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText("[" & TableName & "] - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
            End If

            If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If

                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function

    Function UDFExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            Dim aa = "Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'"
            rs.DoQuery("Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function TableExists(ByVal TableName As String) As Boolean
        Dim oTables As SAPbobsCOM.UserTablesMD
        Dim oFlag As Boolean
        oTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        oFlag = oTables.GetByKey(TableName)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
        Return oFlag
    End Function

    Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            Dim SS As String = "Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'"
            rs.DoQuery("Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function CreateTable(ByVal TableName As String, ByVal TableDesc As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        CreateTable = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.TableExists(TableName) Then
                Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                oApplication.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                v_UserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                v_UserTableMD.TableName = TableName
                v_UserTableMD.TableDescription = TableDesc
                v_UserTableMD.TableType = TableType
                v_RetVal = v_UserTableMD.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Table " & TableDesc & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText("[" & TableName & "] - " & TableDesc & " Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function

    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function

    Function RegisterUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", Optional ByVal ChildTable1 As String = "", _
     Optional ByVal ChildTable2 As String = "", Optional ByVal ChildTable3 As String = "", Optional ByVal ChildTable4 As String = "", Optional ByVal ChildTable5 As String = "", _
         Optional ByVal ChildTable6 As String = "", Optional ByVal ChildTable7 As String = "", Optional ByVal ChildTable8 As String = "", Optional ByVal ChildTable9 As String = "", _
         Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try
            RegisterUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName
            v_udoMD.CanCreateDefaultForm() = SAPbobsCOM.BoYesNoEnum.tYES
            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable1 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable1
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable2 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable2
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable3 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable3
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable4 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable4
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable5 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable5
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable6 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable6
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable7 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable7
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable8 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable8
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable9 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable9
                v_udoMD.ChildTables.Add()
            End If

            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDO = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
    End Function

    Function CreateUserFieldsComboBox(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean
        Try
            'If TableName.StartsWith("@") = False Then
            If Not Me.UDFExists(TableName, FieldName) Then
                Dim v_UserField As SAPbobsCOM.UserFieldsMD
                v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                v_UserField.TableName = TableName
                v_UserField.Name = FieldName
                v_UserField.Description = FieldDescription
                v_UserField.Type = type
                If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    If size <> 0 Then
                        v_UserField.Size = size
                    End If
                End If
                If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    v_UserField.SubType = subType
                End If

                For i As Int16 = 0 To ComboValidValues.GetLength(0) - 1
                    If i > 0 Then v_UserField.ValidValues.Add()
                    v_UserField.ValidValues.Value = ComboValidValues(i, 0)
                    v_UserField.ValidValues.Description = ComboValidValues(i, 1)
                Next
                If DefaultValidValues <> "" Then v_UserField.DefaultValue = DefaultValidValues

                If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                v_RetVal = v_UserField.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return True
                End If
            Else
                Return False
            End If
            ' End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function
End Class

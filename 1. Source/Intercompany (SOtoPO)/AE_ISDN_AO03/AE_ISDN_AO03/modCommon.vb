Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration
Imports System.Data
Imports System.Data.Common

Module modCommon

    Public Function GetCompanyInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long
        Dim sFunctName As String = String.Empty
        Dim sConnection As String = String.Empty

        Try
            sFunctName = "Get Company Initialization"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company Initialization", sFunctName)

            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty
            oCompDef.sDummyItemCode = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPassword")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DummyItem")) Then
                oCompDef.sDummyItemCode = ConfigurationManager.AppSettings("DummyItem")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFunctName)
            GetCompanyInfo = RTN_SUCCESS

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFunctName)
            GetCompanyInfo = RTN_ERROR
        End Try

    End Function

    Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    ConnectDICompSSO()
        '   Purpose    :    Connect To DI Company Object
        '
        '   Parameters :    ByRef objCompany As SAPbobsCOM.Company
        '                       objCompany = set the SAP Company Object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sCookie As String = String.Empty
        Dim sConnStr As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim lRetval As Long
        Dim iErrCode As Int32
        Try
            sFuncName = "ConnectDICompSSO()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            objCompany = New SAPbobsCOM.Company

            sCookie = objCompany.GetContextCookie
            sConnStr = p_oUICompany.GetConnectionContext(sCookie)
            'sConnStr = p_oSBOApplication.Company.GetConnectionContext(sCookie)
            lRetval = objCompany.SetSboLoginContext(sConnStr)

            If Not lRetval = 0 Then
                Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
            End If
            p_oSBOApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objCompany.LicenseServer = ConfigurationManager.AppSettings("LicenseServer")
            lRetval = objCompany.Connect
            If lRetval <> 0 Then
                objCompany.GetLastError(iErrCode, sErrDesc)
                Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
            Else
                p_oSBOApplication.StatusBar.SetText("Company Connection Has Established with the " & objCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If
            ConnectDICompSSO = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectDICompSSO = RTN_ERROR
        End Try
    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, ByVal sDBName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim sSAPUser As String = String.Empty
        Dim sSAPPWd As String = String.Empty
        Dim sTrgtDBName As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            sFuncName = "ConnectToTargetCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Dim sConnectedCompany As String = p_oDICompany.CompanyDB

            sSQL = "SELECT * FROM ""@AE_COMPANYDATA"" WHERE ""U_DATABASE"" = '" & sDBName & "' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sSQL)
            If oRecordSet.RecordCount > 0 Then
                sTrgtDBName = sDBName
                sSAPUser = oRecordSet.Fields.Item("U_SAPUSER").Value
                sSAPPWd = oRecordSet.Fields.Item("U_SAPPASSWORD").Value

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                oCompany = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name : " & sTrgtDBName, sFuncName)
                oCompany.Server = p_oDICompany.Server

                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

                oCompany.LicenseServer = p_oDICompany.LicenseServer
                oCompany.CompanyDB = sTrgtDBName
                oCompany.UserName = sSAPUser
                oCompany.Password = sSAPPWd

                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                oCompany.UseTrusted = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                iRetValue = oCompany.Connect()

                If iRetValue <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)

                    sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                        oCompany.CompanyDB, System.Environment.NewLine, _
                                    vbTab, sErrDesc)

                    Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection established with " & oCompany.CompanyName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

    Public Sub ShowErr(ByVal sErrMsg As String)
        ' ***********************************************************************************
        '   Function   :    ShowErr()
        '   Purpose    :    Show Error Message
        '   Parameters :  
        '                   ByVal sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Dev
        '   Date       :    23 Jan 2007
        '   Change     :
        ' ***********************************************************************************
        Try
            If sErrMsg <> "" Then
                If Not p_oSBOApplication Is Nothing Then
                    If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                        p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                    ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                        p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                    End If
                End If
            End If
        Catch exc As Exception
            WriteToLogFile(exc.Message, "ShowErr()")
        End Try
    End Sub

    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Try
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            sPath = Application.StartupPath.ToString
            'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
            oXmlDoc.Load(sPath & "\" & FileName)
            ' MsgBox(Application.StartupPath)

            Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Public Function Udoform(ByVal n As String) As Integer 'function to load Udt forms
        Dim oMenu As SAPbouiCOM.MenuItem
        Dim i As Integer
        oMenu = p_oSBOApplication.Menus.Item("43535")
        For i = 0 To oMenu.SubMenus.Count - 1
            If oMenu.SubMenus.Item(i).String = n Then
                Return oMenu.SubMenus.Item(i).UID
                Exit For
            End If
        Next
    End Function

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Function CreateUDOTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not isColumnExist(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If ValidValues <> "" Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    sErrDesc = p_oDICompany.GetLastErrorCode() & ":" & p_oDICompany.GetLastErrorDescription()
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) AS ""MNO"" FROM ""CUFD"" WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName & "'")
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try

    End Function

    Public Function CreateUDODocumentChild(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                ByVal strChild As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean

        Try
            Dim oUserObjects As SAPbobsCOM.UserObjectsMD
            oUserObjects = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If (oUserObjects.GetByKey(strUDO) <> True) Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.ChildTables.TableName = strChild
                '//'oUserObjects.ChildTables.SetCurrentLine(1)
                oUserObjects.ChildTables.Add()

                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
                oUserObjects.TableName = strTable

                If (oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES) Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    oUserObjects.FindColumns.Add()
                    oUserObjects.FindColumns.SetCurrentLine(1)
                    oUserObjects.FindColumns.ColumnAlias = strName
                    oUserObjects.FindColumns.Add()
                End If
                If (oUserObjects.Add() <> 0) Then
                    MsgBox(p_oDICompany.GetLastErrorCode & " " & p_oDICompany.GetLastErrorDescription)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    Return False
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
            Return True

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
            Throw ex
        Finally
            GC.Collect()
        End Try
    End Function

    Public Function ExecuteSQLQuery_Hana(ByVal sSql As String, ByVal sDbName As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery_Hana"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oDICompany.Server & ";CS=" & sDbName
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods

    End Function

    Public Function StartTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "StartTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Transaction", sFuncName)

            If oCompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback hanging transactions", sFuncName)
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            oCompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Trancation Started Successfully", sFuncName)
            StartTransaction = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while starting Trancation", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function

    Public Function CommitTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CommitTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If oCompany.InTransaction Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Transaction is Active", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit Transaction Complete", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while committing Transaciton", sFuncName)
            CommitTransaction = RTN_ERROR
        End Try
    End Function

    Public Function RollbackTransaction(ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollbackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oCompany.InTransaction Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No transaction is active", sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
            RollbackTransaction = RTN_SUCCESS
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            RollbackTransaction = RTN_ERROR
        End Try

    End Function

End Module




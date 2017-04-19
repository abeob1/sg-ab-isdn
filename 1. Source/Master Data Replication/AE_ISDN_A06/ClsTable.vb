Option Explicit On
Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A06
    Public Class ClsTable
#Region "Variable Declaration"
        'Public Shared Company As SAPbobsCOM.Company
        Private sFuncName As String = String.Empty
        Public Shared sErrMsg As String = Nothing
        Public Shared lErrCode As Integer = 0
        Public Shared lRetCode As Integer
#End Region
        Public Sub CreateFields(ByVal Company As SAPbobsCOM.Company)
            '  ************************************************************************************************
            '    Function    :   CreateFields()
            '    Purpose     :   This function to to call add fields/UDT/FMS in table 
            '                
            '    Parameters  :  
            '                                

            '    Author      :   SHIBIN
            '    Date        :   Mar 2017
            '  ************************************************************************************************
            'UDT
            AddUDT("AE_ENTITYNAME", "ENTITY NAME", Company)
            Try
                AddTable("AE_SYSRATE", "Update System Rate", SAPbobsCOM.BoUTBTableType.bott_MasterData, Company)
            Catch ex As Exception

            End Try
            Try
                AddTable("AE_SYSRATE1", "Update System Rate Details1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, Company)
            Catch ex As Exception

            End Try


            Try
                'UDT-Fields

                AddFields(Company, "@AE_ENTITYNAME", "DBNAME", "Database Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, 50, SAPbobsCOM.BoFldSubTypes.st_None)
                AddFields(Company, "@AE_ENTITYNAME", "SAPUSER", "SAP User", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, 50, SAPbobsCOM.BoFldSubTypes.st_None)
                AddFields(Company, "@AE_ENTITYNAME", "SAPPWD", "SAP Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, 50, SAPbobsCOM.BoFldSubTypes.st_None)
                'AddFields(Company, "@AE_ENTITYNAME", "DBCODE", "DB Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, 20, SAPbobsCOM.BoFldSubTypes.st_None)

            Catch ex As Exception

            End Try

            Try
                'UDO-Fields
                AddFields(Company, "@AE_SYSRATE", "DATEFROM", "Date From", SAPbobsCOM.BoFieldTypes.db_Date, 1, 50, SAPbobsCOM.BoFldSubTypes.st_None)
                AddFields(Company, "@AE_SYSRATE", "DATETO", "Date To", SAPbobsCOM.BoFieldTypes.db_Date, 2, 50, SAPbobsCOM.BoFldSubTypes.st_None)
                AddFields(Company, "@AE_SYSRATE", "TOLERANCE", "Tolerance", SAPbobsCOM.BoFieldTypes.db_Float, 3, , SAPbobsCOM.BoFldSubTypes.st_Rate)

                Try
                    AddFields(Company, "@AE_SYSRATE1", "CHKBOX", "Checkbox", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, 5, SAPbobsCOM.BoFldSubTypes.st_None)
                    AddFields(Company, "@AE_SYSRATE1", "CURRENCY", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, 20, SAPbobsCOM.BoFldSubTypes.st_None)
                    AddFields(Company, "@AE_SYSRATE1", "BANKRATE", "Bank Rate", SAPbobsCOM.BoFieldTypes.db_Float, 3, , SAPbobsCOM.BoFldSubTypes.st_Rate)
                    AddFields(Company, "@AE_SYSRATE1", "SYSRATE", "System Rate", SAPbobsCOM.BoFieldTypes.db_Float, 4, , SAPbobsCOM.BoFldSubTypes.st_Rate)

                Catch ex As Exception

                End Try
            Catch ex As Exception

            End Try

            Try
                AddUDO(Company)
            Catch ex As Exception

            End Try

            'Try
            '    AddFields(Company, "OITM", "GLOBAL", "GLOBAL ITEM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20, SAPbobsCOM.BoFldSubTypes.st_None)
            'Catch ex As Exception

            'End Try
        End Sub

        Private Sub AddTable(strTab As String, strDesc As String, nType As SAPbobsCOM.BoUTBTableType, ByVal Company As SAPbobsCOM.Company)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD = Nothing
            Try
                sFuncName = "AddTable() - clsTable class"

                oUserTablesMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating  Table....", sFuncName)

                If Not oUserTablesMD.GetByKey(strTab) Then

                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType
                    lRetCode = oUserTablesMD.Add()

                    ' check for errors in the process
                    If lRetCode <> 0 Then
                        Dim transTemp5 As Integer = System.Convert.ToInt32(lErrCode)
                        Company.GetLastError(transTemp5, sErrMsg)
                        'Application.SBO_Application.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        WriteToLogFile(sErrMsg, sFuncName)
                    Else


                        Application.SBO_Application.StatusBar.SetText("Table: " + oUserTablesMD.TableName + " was added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                End If
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                ' GC.WaitForPendingFinalizers();
                GC.Collect()
            End Try

        End Sub
        Private Sub AddUDT(ByVal Table As String, ByVal Description As String, ByVal Company As SAPbobsCOM.Company)
            '  ************************************************************************************************
            '    Function    :   AddUDT()
            '    Purpose     :   This function to add UDT in table 
            '                
            '    Parameters  :  
            '                     
            '    Author      :   SHIBIN
            '    Date        :   Mar 2017
            '  ************************************************************************************************
            sFuncName = "AddUDT - clsTable"
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD = Nothing
            Try

                'oUserTablesMD = DirectCast(ClsFunction.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables), SAPbobsCOM.UserTablesMD)
                'oUserTablesMD = ClsFunction.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                oUserTablesMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating UDT Table....", sFuncName)

                '  set the 2 mandatory fields
                oUserTablesMD.TableName = Table '"AE_CRYSTALDETAILS"
                oUserTablesMD.TableDescription = Description '"AE_CRYSTALDETAILS"
                lRetCode = oUserTablesMD.Add()

                ' check for errors in the process
                If lRetCode <> 0 Then
                    Dim transTemp5 As Integer = System.Convert.ToInt32(lErrCode)
                    Company.GetLastError(transTemp5, sErrMsg)
                    'Application.SBO_Application.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrMsg, sFuncName)
                Else


                    Application.SBO_Application.StatusBar.SetText("Table: " + oUserTablesMD.TableName + " was added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                ' GC.Collect();
                ' GC.WaitForPendingFinalizers();
                oUserTablesMD = Nothing
            End Try

        End Sub


        Private Sub AddFields(ByVal Company As SAPbobsCOM.Company, ByVal strTab As String, _
                                ByVal strCol As String, _
                                    ByVal strDesc As String, _
                                        ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                            Optional ByVal i As Integer = 0, _
                                                Optional ByVal nEditSize As Integer = 10, _
                                                    Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                        Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
            sFuncName = "AddFields - clsTable"
            Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD '= Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating UDF Fileds....", sFuncName)
            Try
                If strTab = "OITM" Then
                    Try
                        oUserFieldMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                        'oUserFieldMD.Remove()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message)
                    End Try
                End If
                If strTab = "OADM" Or strTab = "OITM" Or strTab = "POR1" Or strTab = "PDN1" Or strTab = "ODSC" _
                       Or strTab = "OWOR" Or strTab = "OPOR" Or strTab = "OPDN" Or strTab = "OINS" Or strTab = "OIBT" Or strTab = "OPOR" Or strTab = "OWHS" Or strTab = "OADM" Or strTab = "ORDR" Then

                    oUserFieldMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


                    If Not IsColumnExists(strTab, strCol, Company) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating field if not exist ....", sFuncName)
                        'oUserFieldMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


                        oUserFieldMD.Description = strDesc
                        oUserFieldMD.Name = strCol
                        oUserFieldMD.Type = nType
                        oUserFieldMD.SubType = nSubType
                        oUserFieldMD.TableName = strTab
                        oUserFieldMD.EditSize = nEditSize
                        oUserFieldMD.Mandatory = Mandatory

                        'If strCol = "PFLim" And strTab = "OADM" Then
                        '    oUserFieldMD.LinkedTable = "Test"
                        'End If
                        'If strTab = "ORDR" Then
                        '    If strCol = "FollowUp" Then
                        '        oUserFieldMD.ValidValues.Value = "1"
                        '        oUserFieldMD.ValidValues.Description = "Yes"
                        '        oUserFieldMD.ValidValues.Add()
                        '        oUserFieldMD.ValidValues.Value = "2"
                        '        oUserFieldMD.ValidValues.Description = "No"
                        '        oUserFieldMD.ValidValues.Add()
                        '    End If
                        'End If
                        If oUserFieldMD.Add <> 0 Then
                            Throw New Exception(Company.GetLastErrorDescription)
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

                    End If
                Else
                    oUserFieldMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    If Not oUserFieldMD.GetByKey("@" & strTab, i) Then
                        oUserFieldMD.Description = strDesc
                        oUserFieldMD.Name = strCol
                        oUserFieldMD.Type = nType
                        oUserFieldMD.SubType = nSubType
                        oUserFieldMD.TableName = strTab
                        oUserFieldMD.EditSize = nEditSize
                        oUserFieldMD.Mandatory = Mandatory
                        'If strTab = "FIN2" Then
                        '    If strCol = "FollowUp" Then
                        '        oUserFieldMD.ValidValues.Value = "1"
                        '        oUserFieldMD.ValidValues.Description = "Yes"
                        '        oUserFieldMD.ValidValues.Add()
                        '        oUserFieldMD.ValidValues.Value = "2"
                        '        oUserFieldMD.ValidValues.Description = "No"
                        '        oUserFieldMD.ValidValues.Add()
                        '    End If
                        'End If

                        'If strTab = "OIOGE" Then
                        '    If strCol = "Status" Then

                        '        oUserFieldMD.ValidValues.Value = "1"
                        '        oUserFieldMD.ValidValues.Description = "Received"
                        '        oUserFieldMD.ValidValues.Add()
                        '        oUserFieldMD.ValidValues.Value = "2"
                        '        oUserFieldMD.ValidValues.Description = "PMT Processed"
                        '        oUserFieldMD.ValidValues.Add()
                        '        oUserFieldMD.ValidValues.Value = "3"
                        '        oUserFieldMD.ValidValues.Description = "Despatched"
                        '        oUserFieldMD.ValidValues.Add()
                        '        oUserFieldMD.DefaultValue = "1"
                        '    End If
                        'End If


                        If oUserFieldMD.Add <> 0 Then
                            Throw New Exception(Company.GetLastErrorDescription)
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)
                oUserFieldMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

#Region "Check Whether Column Exsists or not"

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String, ByVal Company As SAPbobsCOM.Company) As Boolean

            Dim strSQL As String
            Dim oRecordSet As SAPbobsCOM.Recordset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try

                strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                oRecordSet = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oRecordSet = Nothing
                GC.Collect()
            End Try
        End Function

#End Region
        'Private Function AddFields(strTab As String, strCol As String, strDesc As String, iFieldSize As Integer, oType As SAPbobsCOM.BoFieldTypes, ByVal Company As SAPbobsCOM.Company) As Boolean
        '    '  ************************************************************************************************
        '    '    Function    :   AddFields()
        '    '    Purpose     :   This function to add fields in table
        '    '                
        '    '    Parameters  :  
        '    '                     string strTab           
        '    '                          string strDesc
        '    '                              int iFieldSize
        '    '                                  SAPbobsCOM.BoFieldTypes oType)
        '    '    Author      :   SHIBIN
        '    '    Date        :   Mar 2017
        '    '  ************************************************************************************************
        '    Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        '    'oUserFieldsMD = DirectCast(ClsFunction.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields), SAPbobsCOM.UserFieldsMD)
        '    oUserFieldsMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        '    sFuncName = "AddFields - clsTable class"

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding fields to the database table.....", sFuncName)
        '    Try
        '        oUserFieldsMD.TableName = strTab
        '        oUserFieldsMD.Name = strCol
        '        oUserFieldsMD.Description = strDesc
        '        oUserFieldsMD.Type = oType
        '        If iFieldSize > 0 Then
        '            oUserFieldsMD.EditSize = iFieldSize
        '        End If


        '        lRetCode = oUserFieldsMD.Add()

        '        If (lRetCode <> 0) Then
        '            If (lRetCode = -2035 Or lRetCode = -1120 Or lRetCode = -5002) Then
        '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Failed or fields already exist in table.....", sFuncName)
        '                Return False
        '            End If

        '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Failed or fields already exist in table.....", sFuncName)

        '            Return False

        '            Throw New Exception(Company.GetLastErrorDescription())
        '        End If

        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fields added to the database table successfully.....", sFuncName)


        '        Return True
        '    Catch ex As Exception
        '        WriteToLogFile(ex.Message, sFuncName)
        '        Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '        Throw ex
        '    Finally
        '        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
        '        ' GC.Collect();
        '        oUserFieldsMD = Nothing
        '    End Try


        'End Function


        Private Sub AddUDO(ByVal Company As SAPbobsCOM.Company)
            sFuncName = "AddUDO() - clsTable class"
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD = Nothing
            Dim oUDOFind As SAPbobsCOM.UserObjectMD_FindColumns = Nothing
            Dim oUDOForm As SAPbobsCOM.UserObjectMD_FormColumns = Nothing
            Dim oUDOEnhancedForm As SAPbobsCOM.UserObjectMD_EnhancedFormColumns = Nothing
            Try
                oUserObjectMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                oUDOFind = oUserObjectMD.FindColumns
                oUDOForm = oUserObjectMD.FormColumns
                oUDOEnhancedForm = oUserObjectMD.EnhancedFormColumns
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating UDO....", sFuncName)
                If Not oUserObjectMD.GetByKey("SYSRATE") Then
                    '    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                    '    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                    '    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                    '    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                    '    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    '    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                    '    'oUserObjectMD.LogTableName = "AE_SYSRATE1"
                    '    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES

                    '    oUserObjectMD.Code = "SYSRATE"
                    '    oUserObjectMD.Name = "System Rate"
                    '    oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                    '    oUserObjectMD.TableName = "AE_SYSRATE"
                    '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Child Table....", sFuncName)
                    '    oUserObjectMD.ChildTables.TableName = "AE_SYSRATE1"

                    '    lRetCode = oUserObjectMD.Add()



                    'oUserObjectMD = DirectCast(Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD), SAPbobsCOM.UserObjectsMD)
                    'oUserObjectMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                    oUserObjectMD.Code = "SYSRATE"
                    oUserObjectMD.Name = "System Rate"
                    oUserObjectMD.TableName = "AE_SYSRATE"
                    oUserObjectMD.ChildTables.TableName = "AE_SYSRATE1"
                    oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO
                    'oUserObjectMD.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES

                    'oUserObjectMD.FatherMenuID = 43520
                    'oUserObjectMD.MenuCaption = "UDO_SYS"
                    'oUserObjectMD.MenuUID = "UDO_SYS"
                    'oUserObjectMD.Position = 0
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO


                    oUserObjectMD.FindColumns.ColumnAlias = "DocEntry"
                    oUserObjectMD.FindColumns.ColumnDescription = "DocEntry"
                    oUserObjectMD.FindColumns.Add()

                    oUserObjectMD.FindColumns.ColumnAlias = "U_DATEFROM"
                    oUserObjectMD.FindColumns.ColumnDescription = "Date From"
                    oUserObjectMD.FindColumns.Add()

                    oUserObjectMD.FindColumns.ColumnAlias = "U_DATETO"
                    oUserObjectMD.FindColumns.ColumnDescription = "Date To"
                    oUserObjectMD.FindColumns.Add()

                    oUserObjectMD.FindColumns.ColumnAlias = "U_TOLERANCE"
                    oUserObjectMD.FindColumns.ColumnDescription = "Tolerance"
                    oUserObjectMD.FindColumns.Add()


                    oUDOEnhancedForm.ColumnAlias = "U_CHKBOX"
                    oUDOEnhancedForm.ColumnDescription = "Checkbox"
                    oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES

                    oUDOEnhancedForm.ColumnNumber = 1
                    oUDOEnhancedForm.ChildNumber = 1
                    oUDOEnhancedForm.Add()

                    oUDOEnhancedForm.ColumnAlias = "U_CURRENCY"
                    oUDOEnhancedForm.ColumnDescription = "Currency"
                    oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.ColumnNumber = 2
                    oUDOEnhancedForm.ChildNumber = 1
                    oUDOEnhancedForm.Add()

                    oUDOEnhancedForm.ColumnAlias = "U_BANKRATE"
                    oUDOEnhancedForm.ColumnDescription = "Bank Rate"
                    oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.ColumnNumber = 3
                    oUDOEnhancedForm.ChildNumber = 1
                    oUDOEnhancedForm.Add()

                    oUDOEnhancedForm.ColumnAlias = "U_SYSRATE"
                    oUDOEnhancedForm.ColumnDescription = "Bank Rate"
                    oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                    oUDOEnhancedForm.ColumnNumber = 4
                    oUDOEnhancedForm.ChildNumber = 1
                    oUDOEnhancedForm.Add()

                    'oUserObjectMD.ChildTables.TableName = "LINES"

                    'oUserObjectMD.FormColumns.SetCurrentLine(0)
                    'oUserObjectMD.FormColumns.FormColumnAlias = "U_CHKBOX"
                    ''ChildTable ColumnAlias
                    'oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES
                    'oUserObjectMD.FormColumns.FormColumnDescription = "Checkbox"
                    ''ChildTable ColumnDescription
                    'oUserObjectMD.FormColumns.SonNumber = 1

                    '' ChildTable SonNumber
                    'oUserObjectMD.FormColumns.Add()

                    ' check for errors in the process
                    If oUserObjectMD.Add() <> 0 Then
                        Company.GetLastErrorDescription()
                        Dim transTemp5 As Integer = System.Convert.ToInt32(lErrCode)
                        Company.GetLastError(transTemp5, sErrMsg)

                        'Application.SBO_Application.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        WriteToLogFile(sErrMsg, sFuncName)
                    Else


                        Application.SBO_Application.StatusBar.SetText("UDO: " + oUserObjectMD.Name + " was added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If




                    ' check for errors in the process
                    'If lRetCode <> 0 Then
                    '    Dim transTemp5 As Integer = System.Convert.ToInt32(lErrCode)
                    '    Company.GetLastError(transTemp5, sErrMsg)

                    '    'Application.SBO_Application.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    '    WriteToLogFile(sErrMsg, sFuncName)
                    'Else


                    '    Application.SBO_Application.StatusBar.SetText("UDO: " + oUserObjectMD.Name + " was added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'End If
                End If
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                ' GC.Collect();
                ' GC.WaitForPendingFinalizers();
                oUserObjectMD = Nothing
            End Try
        End Sub
    End Class
End Namespace

Option Explicit On
Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A02
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
            '    Date        :   DEC 2016
            '  ************************************************************************************************

            AddUDT("AE_CRYSTALDETAILS", "CRYSTAL DETAILS", Company)
            AddUDT("AE_COMPANYDATA", "COMPANY DATA", Company)
            AddUDT("AE_ESOALOG", "LOG TRANSACTION", Company)

            'T0.[U_SMTP], T0.[U_PORT], T0.[U_Email], T0.[U_Username], T0.[U_Password]
            Try

                AddFields("@AE_COMPANYDATA", "SMTP", "SMTP", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_COMPANYDATA", "PORT", "PORT", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_COMPANYDATA", "Email", "Email", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_COMPANYDATA", "Username", "Username", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_COMPANYDATA", "Password", "Password", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
            Catch ex As Exception

            End Try
            Try
                AddFields("@AE_ESOALOG", "BPCode", "BPCode", 40, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "BPName", "BPName", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "Soa_Date", "Soa_Date", 0, SAPbobsCOM.BoFieldTypes.db_Date, Company)
                AddFields("@AE_ESOALOG", "Balance", "Balance", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "EmailID", "EmailID", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "Status", "Status", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "ErrMsg", "ErrMsg", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
                AddFields("@AE_ESOALOG", "EDate", "EDate", 0, SAPbobsCOM.BoFieldTypes.db_Date, Company)
                AddFields("@AE_ESOALOG", "user", "user", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, Company)
            Catch ex As Exception

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
            '    Date        :   DEC 2016
            '  ************************************************************************************************
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

        


        Private Function AddFields(strTab As String, strCol As String, strDesc As String, iFieldSize As Integer, oType As SAPbobsCOM.BoFieldTypes, ByVal Company As SAPbobsCOM.Company) As Boolean
            '  ************************************************************************************************
            '    Function    :   AddFields()
            '    Purpose     :   This function to add fields in table
            '                
            '    Parameters  :  
            '                     string strTab           
            '                          string strDesc
            '                              int iFieldSize
            '                                  SAPbobsCOM.BoFieldTypes oType)
            '    Author      :   SHIBIN
            '    Date        :   April 2016
            '  ************************************************************************************************
            Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
            'oUserFieldsMD = DirectCast(ClsFunction.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields), SAPbobsCOM.UserFieldsMD)
            oUserFieldsMD = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            sFuncName = "AddFields - clsTable class"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding fields to the database table.....", sFuncName)
            Try
                oUserFieldsMD.TableName = strTab
                oUserFieldsMD.Name = strCol
                oUserFieldsMD.Description = strDesc
                oUserFieldsMD.Type = oType
                If iFieldSize > 0 Then
                    oUserFieldsMD.EditSize = iFieldSize
                End If


                lRetCode = oUserFieldsMD.Add()

                If (lRetCode <> 0) Then
                    If (lRetCode = -2035 Or lRetCode = -1120 Or lRetCode = -5002) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Failed or fields already exist in table.....", sFuncName)
                        Return False
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Failed or fields already exist in table.....", sFuncName)

                    Return False

                    Throw New Exception(Company.GetLastErrorDescription())
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fields added to the database table successfully.....", sFuncName)


                Return True
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                ' GC.Collect();
                oUserFieldsMD = Nothing
            End Try


        End Function


    End Class
End Namespace

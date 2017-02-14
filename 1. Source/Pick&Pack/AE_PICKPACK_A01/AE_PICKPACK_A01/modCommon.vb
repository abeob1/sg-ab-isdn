Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration


Module modCommon

    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oDTPick As DataTable = Nothing
    Public p_iDebugMode As Int16 = 1
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2
    Public DateConversion As New System.Globalization.CultureInfo("fr-FR", True)
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String


    Public Function AddUserDataSrc(ByRef oForm As SAPbouiCOM.Form, ByVal sDSUID As String, _
                                   ByRef sErrDesc As String, ByVal oDataType As SAPbouiCOM.BoDataType, _
                                   Optional ByVal lLen As Long = 0) As Long
        ' ***********************************************************************************
        '   Function   :    AddUserDataSrc()
        '   Purpose    :    Add User Data Source
        '
        '   Parameters :    ByVal oForm As SAPbouiCOM.Form
        '                       oForm = set the SAP UI Form Object
        '                   ByVal sDSUID As String
        '                       sDSUID = Data Set UID
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '                   ByVal oDataType As SAPbouiCOM.BoDataType
        '                       oDataType = set the SAP UI BoDataType Object
        '                   Optional ByVal lLen As Long = 0
        '                       lLen= Length
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Dev
        '   Date       :    23 Jan 2007
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "AddUserDataSrc()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If lLen = 0 Then
                oForm.DataSources.UserDataSources.Add(sDSUID, oDataType)
            Else
                oForm.DataSources.UserDataSources.Add(sDSUID, oDataType, lLen)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddUserDataSrc = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AddUserDataSrc = RTN_ERROR
        End Try
    End Function

    Public Sub Single_Signon(ByRef oApplication As SAPbouiCOM.Application)
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty


        Try
            sFuncName = "Single_Signon()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sconn = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056" 'Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            oApplication = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = p_oDICompany.GetContextCookie
            str = oApplication.Company.GetConnectionContext(scook)
            ret = p_oDICompany.SetSboLoginContext(str)
            oApplication.StatusBar.SetText("Initializing License Server... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            ''  p_oDICompany.LicenseServer = ConfigurationManager.AppSettings("LicenseServer")
            oApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            p_oDICompany.Connect()
            p_oDICompany.GetLastError(ret, str)
            If ret <> 0 Then

                Throw New ArgumentException("Connect of Single SignOn failed : " & str)
            Else
                oApplication.StatusBar.SetText("Connected to the Company ........ " & p_oDICompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Public Function StartTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    StartTransaction()
        '   Purpose    :    Start DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Sri
        '   Date       :   29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "StartTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oDICompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Found hanging transaction.Rolling it back.", sFuncName)
                p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            p_oDICompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            StartTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function

    Public Function RollBackTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    RollBackTransaction()
        '   Purpose    :    Roll Back DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollBackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oDICompany.InTransaction Then
                p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for rollback", sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            RollBackTransaction = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            RollBackTransaction = RTN_ERROR
        Finally
            GC.Collect()
        End Try

    End Function

    Public Function CommitTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    CommitTransaction()
        '   Purpose    :    Commit DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "CommitTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oDICompany.InTransaction Then
                p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for commit", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CommitTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function


    Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.

        Dim dtTable As New DataTable
        Dim NewCol As DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer

        Try
            For ColCount = 0 To SAPRecordset.Fields.Count - 1
                NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                dtTable.Columns.Add(NewCol)
            Next

            Do Until SAPRecordset.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To SAPRecordset.Fields.Count - 1

                    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)


                SAPRecordset.MoveNext()
            Loop

            sErrDesc = String.Empty
            Return dtTable

        Catch ex As Exception
            sErrDesc = ex.Message
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Exit Function
        End Try


    End Function

    Public Function GetExchangeRate(oholdingCompany As SAPbobsCOM.Company, sCurrency As String, dCurDate As Date) As Double
        Dim ovObj As SAPbobsCOM.SBObob = Nothing
        Dim ors As SAPbobsCOM.Recordset = Nothing
        Try
            Dim Result As Double = 1
            ovObj = oholdingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            ors = oholdingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ors = ovObj.GetCurrencyRate(sCurrency, dCurDate)
            Result = ors.Fields.Item(0).Value
            Return Result
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            ovObj = Nothing
            ors = Nothing
        End Try
    End Function

    Public Sub ShowErr(ByVal sErrMsg As String, ByRef oApp As SAPbouiCOM.Application)
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
                If Not oApp Is Nothing Then
                    If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                        oApp.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                    ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                        oApp.MessageBox("Error : " & sErrMsg)
                    End If
                End If
            End If
        Catch exc As Exception
            WriteToLogFile(exc.Message, "ShowErr()")
        End Try
    End Sub

End Module

Imports System.Reflection
Imports System.IO
Imports System.Xml
''Imports SAPbouiCOM
Imports System.Text
Imports System.Windows.Forms

''' <summary>
''' Globally whatever Function and method do you want define here 
''' We can use any class and module from here  
''' </summary>
''' <remarks></remarks>
Public Class GlobalFunctions


#Region " ...  Common For Company ..."

    Public Function FormExist(ByVal FormID As String) As Boolean
        FormExist = False
        For Each uid As SAPbouiCOM.Form In oApplication.Forms
            If uid.UniqueID = FormID Then
                FormExist = True
                Exit Function
            End If
        Next
        If FormExist Then
            oApplication.Forms.Item(FormID).Visible = True
            oApplication.Forms.Item(FormID).Select()
        End If
    End Function

    Public Function ConnectionContext() As Integer
        Try
            Dim strErrorCode As String
            If oCompany.Connected = True Then oCompany.Disconnect()

            oApplication.StatusBar.SetText("Please Wait ..........Connecting Landed Cost Addon With The Company....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strErrorCode = oCompany.Connect
            ConnectionContext = strErrorCode
            If strErrorCode = 0 Then
                oApplication.StatusBar.SetText("Landed Cost Addon Connection Established  !!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Media.SystemSounds.Asterisk.Play()
            Else
                oApplication.StatusBar.SetText("Failed To Connect, Please Check The License Configuration....." & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function

    Public Function CookieConnect() As Integer
        Try
            Dim strCkie, strContext As String
            oCompany = New SAPbobsCOM.Company
            Debug.Print(oCompany.CompanyDB)
            strCkie = oCompany.GetContextCookie()
            strContext = oApplication.Company.GetConnectionContext(strCkie)
            CookieConnect = oCompany.SetSboLoginContext(strContext)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function

    Public Sub SetApplication()
        Try
            Dim oGUI As New SAPbouiCOM.SboGuiApi
            oGUI.AddonIdentifier = ""
            oGUI.Connect(Environment.GetCommandLineArgs.GetValue(1).ToString())
            oApplication = oGUI.GetApplication()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub
#End Region

    Private Const MAXFILESIZE_IN_MB As Int16 = 5 '(2 MB)
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0


    Public P_odtCourier As DataTable()
    Public P_odtForwarded As DataTable()



    Function DoQuery(ByVal strSql As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCode.DoQuery(strSql)
            Return rsetCode
        Catch ex As Exception
            oApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try
    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long

        ' **********************************************************************************
        '   Function   :    WriteToLogFile_Debug()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    SRI
        '   Date       :    29 April 2013
        '   Changes    : 
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double
        Dim iFileCount As Integer = 0

        Try
            strTempString = Space(40 - Len(strSourceName))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText
            strFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.Windows.Forms.Application.StartupPath, LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
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

    Public Function Courier_ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.

        Dim dtTable As New DataTable
        Try
            dtTable.Columns.Add("U_ForwarderCode", GetType(String))
            dtTable.Columns.Add("U_Weight", GetType(Double))
            dtTable.Columns.Add("U_Zone1", GetType(Double))
            dtTable.Columns.Add("U_Zone2", GetType(Double))
            dtTable.Columns.Add("U_Zone3", GetType(Double))
            dtTable.Columns.Add("U_Zone4", GetType(Double))
            dtTable.Columns.Add("U_Zone7", GetType(Double))
            dtTable.Columns.Add("FromDate", GetType(Date))
            dtTable.Columns.Add("ToDate", GetType(Date))

            Do Until SAPRecordset.EoF

                'Add the row to the datatable
                dtTable.Rows.Add(SAPRecordset.Fields.Item("U_ForwarderCode").Value, SAPRecordset.Fields.Item("U_Weight").Value, SAPRecordset.Fields.Item("U_Zone1").Value, _
                                 SAPRecordset.Fields.Item("U_Zone2").Value, SAPRecordset.Fields.Item("U_Zone3").Value, SAPRecordset.Fields.Item("U_Zone4").Value, _
                                 SAPRecordset.Fields.Item("U_Zone7").Value, DateTime.ParseExact(SAPRecordset.Fields.Item("U_FDate").Value, "yyyyMMdd", Nothing), DateTime.ParseExact(SAPRecordset.Fields.Item("U_TDate").Value, "yyyyMMdd", Nothing))
                SAPRecordset.MoveNext()
            Loop
            '' DateTime.ParseExact(Right(sString(1), 8), "yyyyMMdd", Nothing)
            sErrDesc = String.Empty
            Return dtTable

        Catch ex As Exception
            sErrDesc = ex.Message
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Exit Function
        End Try


    End Function

    Public Function Forwarded_ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.
        'SELECT T0.""U_Forwarder"", T0.""U_Country"", T0.""U_Port"", T0.""U_Incoterms"", T0.""U_Cur"", T0.""U_Component"", T0.""U_ByAmtOrPrcnt"", 
        ''T0.""U_ByKgOrNet"", T0.""U_Weight"", T0.""U_Min"", T0.""U_Flat"", 
        ''          T0.""U_Basic"", T0.""U_FromDate"", T0.""U_ToDate"", T0.""Code"" FROM  ""@FORWARDER""  T0"
        Dim dtTable As New DataTable
        Try
            dtTable.Columns.Add("U_Forwarder", GetType(String))
            dtTable.Columns.Add("U_Country", GetType(String))
            dtTable.Columns.Add("U_Port", GetType(String))
            dtTable.Columns.Add("U_Incoterms", GetType(String))
            dtTable.Columns.Add("U_Cur", GetType(String))
            dtTable.Columns.Add("U_Component", GetType(String))
            dtTable.Columns.Add("U_ByAmtOrPrcnt", GetType(String))
            dtTable.Columns.Add("U_ByKgOrNet", GetType(String))
            dtTable.Columns.Add("U_Weight", GetType(Double))
            dtTable.Columns.Add("U_Min", GetType(Double))
            dtTable.Columns.Add("U_Flat", GetType(Double))
            dtTable.Columns.Add("U_Basic", GetType(Double))
            dtTable.Columns.Add("U_FromDate", GetType(Date))
            dtTable.Columns.Add("U_ToDate", GetType(Date))

            Do Until SAPRecordset.EoF

                'Add the row to the datatable
                dtTable.Rows.Add(SAPRecordset.Fields.Item("U_Forwarder").Value, SAPRecordset.Fields.Item("U_Country").Value, SAPRecordset.Fields.Item("U_Port").Value, _
                                 SAPRecordset.Fields.Item("U_Incoterms").Value, SAPRecordset.Fields.Item("U_Cur").Value, SAPRecordset.Fields.Item("U_Component").Value, _
                                 SAPRecordset.Fields.Item("U_ByAmtOrPrcnt").Value, SAPRecordset.Fields.Item("U_ByKgOrNet").Value, SAPRecordset.Fields.Item("U_Weight").Value, _
                                 SAPRecordset.Fields.Item("U_Min").Value, SAPRecordset.Fields.Item("U_Flat").Value, SAPRecordset.Fields.Item("U_Basic").Value, _
                                 DateTime.ParseExact(SAPRecordset.Fields.Item("U_FromDate").Value, "yyyyMMdd", Nothing), DateTime.ParseExact(SAPRecordset.Fields.Item("U_ToDate").Value, "yyyyMMdd", Nothing))
                SAPRecordset.MoveNext()
            Loop
            '' DateTime.ParseExact(Right(sString(1), 8), "yyyyMMdd", Nothing)
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
            Throw New Exception(ex.Message & " Curr : " & sCurrency)
        Finally
            ovObj = Nothing
            ors = Nothing
        End Try
    End Function

End Class

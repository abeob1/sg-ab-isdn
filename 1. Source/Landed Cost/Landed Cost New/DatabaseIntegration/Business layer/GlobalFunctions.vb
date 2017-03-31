Imports System.Reflection
Imports System.IO
Imports SAPbouiCOM
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

End Class

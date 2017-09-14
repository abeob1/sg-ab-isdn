Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A06
    Module modMain

        Public Structure CompanyDefault

            Public sSQL_Name As String
            Public sSQL_password As String
            Public sSMTPServer As String
            Public sSMTPPort As String
            Public sEmailFrom As String
            Public sSMTPUser As String
            Public sSMTPPassword As String

        End Structure

        Public p_iDRP As Integer = 0
        Public p_iUDO As Integer = 0
        Public p_iEXR As Integer = 0
        Public p_iSearch As Integer = 0
        Public p_iCOACount As Integer = 0
        Public p_iBPCount As Integer = 0
        Public p_iItemCount As Integer = 0
        Public p_iExchangeCount As Integer = 0
        Public p_iReplication As Integer = 0
        Public p_iAddMode As Integer = 0
        Public p_oApps As SAPbouiCOM.SboGuiApi
        '  Public p_oEventHandler As clsEventHandler
        Public WithEvents p_oSBOApplication As SAPbouiCOM.Application

        Public p_oDICompany As SAPbobsCOM.Company
        Public p_oUICompany As SAPbouiCOM.Company
        Public Company As SAPbobsCOM.Company
        Public sFuncName As String
        Public sErrDesc As String
        Public p_stype As String
        Public dtTable As DataTable = Nothing
        Public oDT_ErrorMsg As DataTable = Nothing

        Public oDT_Entities As DataTable = Nothing
        Public p_oDtSyncLogCOA As DataTable = Nothing
        Public p_oDtSyncLogBP As DataTable = Nothing
        Public p_oDtSyncLogItem As DataTable = Nothing
        Public p_oDtSyncLogER As DataTable = Nothing

        Public p_iDebugMode As Int16
        Public p_iErrDispMethod As Int16
        Public p_iDeleteDebugLog As Int16
        Public oDT_ExchRates As DataTable = Nothing

        Public p_sSQLName As String = String.Empty
        Public p_sSQLPass As String = String.Empty

        Public Const RTN_SUCCESS As Int16 = 1
        Public Const RTN_ERROR As Int16 = 0

        Public Const DEBUG_ON As Int16 = 1
        Public Const DEBUG_OFF As Int16 = 0

        Public Const ERR_DISPLAY_STATUS As Int16 = 1
        Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2
        Public format1 As New System.Globalization.CultureInfo("fr-FR", True)
        Public p_oCompDef As CompanyDefault
        Public p_sEmailID As String = String.Empty

        Public p_sSelectedFilepath As String = String.Empty




        <STAThread()>
        Sub Main(ByVal args() As String)

            ''Dim oApp As Application
            Dim sconn As String = String.Empty
            ''If (args.Length < 1) Then
            ''    oApp = New Application
            ''Else
            ''    oApp = New Application(args(0))
            ''End If

            sFuncName = "Main()"
            Try
                p_iDebugMode = DEBUG_ON
                p_iErrDispMethod = ERR_DISPLAY_STATUS

                Dim oApp As Application
                If (args.Length < 1) Then
                    oApp = New Application
                Else
                    oApp = New Application(args(0))
                End If
                Application.SBO_Application.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                p_oDICompany = New SAPbobsCOM.Company

                'p_oDICompany = Application.SBO_Application.Company.GetDICompany()

                'If p_oDICompany.Connected Then
                '    Application.SBO_Application.StatusBar.SetText("Company Connection Has Established with the " & p_oDICompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    '' Application.SBO_Application.MessageBox("Connected Successfully ... !" & "   " & p_oDICompany.CompanyName)
                'Else
                '    Application.SBO_Application.StatusBar.SetText("Connect of Single SignOn failed : ... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                'End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
                'p_oEventHandler.AddMenuItems()
                Dim MyMenu As ClsMain
                MyMenu = New ClsMain(Application.SBO_Application)
                MyMenu.AddMenuItems()


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Add tables", sFuncName)
                Dim oCreateFields As ClsTable
                oCreateFields = New ClsTable()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling function CreateFields() for creating UDF.....", sFuncName)
                Application.SBO_Application.SetStatusBarMessage("Creating UDT`s & UDF`s ....!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                oCreateFields.CreateFields(p_oDICompany)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateDataTable() for COA", sFuncName)
                p_oDtSyncLogCOA = CreateDataTable_COA("Entity", "GL Code", "GL Name", "Status", "Error Message")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateDataTable_BP() for BP", sFuncName)
                p_oDtSyncLogBP = CreateDataTable_BP("Entity", "CardCode", "CardName", "Status", "Error Message")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateDataTable_Item() for ItemMaster", sFuncName)
                p_oDtSyncLogItem = CreateDataTable_Item("Entity", "ItemCode", "ItemName", "Status", "Error Message")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateDataTable_Item() for ItemMaster", sFuncName)
                p_oDtSyncLogItem = CreateDataTable_ExRate("Entity", "FromDate", "ToDate", "Status", "Error Message")

                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication Function", sFuncName)
                ''    Call p_oEventHandler.SetApplication(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

                'Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                ' Call EndStatus(sErrDesc)
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing Recordset ", "Main()")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
               
                'If GetSystemIntializeInfo(p_oCompDef, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


                Application.SBO_Application.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
             
                oApp.Run()
                p_oDICompany.Disconnect()
                System.Windows.Forms.Application.Exit()

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try


        End Sub

        Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes)
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    System.Windows.Forms.Application.Exit()
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            End Select
        End Sub


        'Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes)
        '    ' **********************************************************************************
        '    '   Function   :    SBO_Application_AppEvent()
        '    '   Purpose    :    This function will be handling the SAP Application Event
        '    '               
        '    '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
        '    '                       EventType = set the SAP UI Application Eveny Object        
        '    ' **********************************************************************************
        '    Dim sFuncName As String = String.Empty
        '    Dim sErrDesc As String = String.Empty
        '    Dim sMessage As String = String.Empty

        '    Try
        '        sFuncName = "SBO_Application_AppEvent()"
        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

        '        Select Case EventType
        '            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
        '                sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
        '                p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        '                End
        '        End Select

        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        '    Catch ex As Exception
        '        sErrDesc = ex.Message
        '        WriteToLogFile(sErrDesc, sFuncName)
        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '        ShowErr(sErrDesc)
        '    Finally
        '        GC.Collect()  'Forces garbage collection of all generations.
        '    End Try
        'End Sub

    End Module
End Namespace
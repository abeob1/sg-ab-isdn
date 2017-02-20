Module modMain

    Public Structure CompanyDefault
        Public sDBUser As String
        Public sDBPwd As String
        Public sDummyItemCode As String
    End Structure

    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public p_oEventHandler As clsEventHandler
    Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oUICompany As SAPbouiCOM.Company
    Public p_oTargetCompany As SAPbobsCOM.Company
    Public p_oCompDef As CompanyDefault
    Public sFuncName As String
    Public sErrDesc As String

    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

    <STAThread()>
    Sub Main(ByVal args() As String)
        sFuncName = "Main()"
        Try
            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            p_oApps = New SAPbouiCOM.SboGuiApi
            p_oApps.Connect(args(0))

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
            p_oSBOApplication = p_oApps.GetApplication

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = p_oSBOApplication.Company

            p_oDICompany = New SAPbobsCOM.Company
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting Company", sFuncName)
            If Not p_oDICompany.Connect() Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
            p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ()", sFuncName)
            If GetCompanyInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
            'p_oEventHandler.AddMenuItems()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateTable()", sFuncName)
            CreateTable()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

            p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Private Sub CreateTable()
        Try
            'CREATE UDF COLUMNS FOR MARKETING DOCUMENTS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("OPOR", "ENTITY", "TARGET DATABASE", SAPbobsCOM.BoFieldTypes.db_Alpha, 130, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("OPOR", "DRAFTKEY", "DRAFT KEY", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            'CREATE UDO TABLE FOR COMPANY DATA TABLE
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            CreateUDOTable("AE_COMPANYDATA", "COMPANY DATA", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_COMPANYDATA", "DATABASE", "DATABASE", SAPbobsCOM.BoFieldTypes.db_Alpha, 130, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_COMPANYDATA", "SAPUSER", "SAP USER", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_COMPANYDATA", "SAPPASSWORD", "PASSWORD", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            'CREATE UDO TABLE FOR CARDCODE SETUP
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            CreateUDOTable("AE_CARDCODESETUP", "CARDCODE SETUP", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "SOURCEDB", "SOURCE DATABASE", SAPbobsCOM.BoFieldTypes.db_Alpha, 130, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "SUPPCODE", "SUPPLIER CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "SUPPNAME", "SUPPLIER NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "TARGETDB", "TARGET DATABASE", SAPbobsCOM.BoFieldTypes.db_Alpha, 130, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "CUSTCODE", "CUSTOMER CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")
            addField("@AE_CARDCODESETUP", "CUSTNAME", "CUSTOMER NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub

End Module

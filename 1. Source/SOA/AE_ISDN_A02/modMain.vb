﻿Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A02
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

        Public p_iSOA As Integer = 0
        Public p_iSearch As Integer = 0
        Public p_iEmail As Integer = 0
        Public p_iPrint As Integer = 0

        Public p_oApps As SAPbouiCOM.SboGuiApi
        '  Public p_oEventHandler As clsEventHandler
        Public WithEvents p_oSBOApplication As SAPbouiCOM.Application

        Public p_oDICompany As SAPbobsCOM.Company
        Public p_oUICompany As SAPbouiCOM.Company
        Public sFuncName As String
        Public sErrDesc As String


        Public p_iDebugMode As Int16
        Public p_iErrDispMethod As Int16
        Public p_iDeleteDebugLog As Int16

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
                MyMenu = New ClsMain()
                MyMenu.AddMenuItems()


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Add tables", sFuncName)
                Dim oCreateFields As ClsTable
                oCreateFields = New ClsTable()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling function CreateFields() for creating UDF.....", sFuncName)
                Application.SBO_Application.SetStatusBarMessage("Creating UDT`s & UDF`s ....!", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                oCreateFields.CreateFields(p_oDICompany)


                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication Function", sFuncName)
                ''    Call p_oEventHandler.SetApplication(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

                'Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                ' Call EndStatus(sErrDesc)
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing Recordset ", "Main()")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
                ' '' '' '' '' '' '' ''Commented on 08/12/2016 for Checking
                ' '' '' '' '' '' '' '''Need to uncomment
                If GetSystemIntializeInfo(p_oCompDef, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


                Application.SBO_Application.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'System.Windows.Forms.Application.Run()
                'AddHandler Application.SBO_Application.AppEvent, AddressOf SBO_Application_AppEvent
                oApp.Run()

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



    End Module
End Namespace
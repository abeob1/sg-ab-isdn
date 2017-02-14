Imports SAPbouiCOM.Framework

Namespace AE_PICKPACK_A01
    Module Module1

        <STAThread()>
        Sub Main(ByVal args() As String)

            Try

                Dim oApp As Application
                If (args.Length < 1) Then
                    oApp = New Application
                Else
                    oApp = New Application(args(0))
                End If
                Application.SBO_Application.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                p_oDICompany = New SAPbobsCOM.Company
                ''p_oDICompany = Application.SBO_Application.Company.GetDICompany()
                ''If p_oDICompany.Connected Then
                ''    Application.SBO_Application.StatusBar.SetText("Company Connection Has Established with the " & p_oDICompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                ''    '' Application.SBO_Application.MessageBox("Connected Successfully ... !" & "   " & p_oDICompany.CompanyName)
                ''Else
                ''    Application.SBO_Application.StatusBar.SetText("Connect of Single SignOn failed : ... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                ''End If

                Dim MyMenu As Menu
                MyMenu = New Menu()
                p_oDTPick = New DataTable()
                p_oDTPick.Columns.Add("row", GetType(String))
                ' ''MyMenu.AddMenuItems()
                ' ''AddHandler Application.SBO_Application.AppEvent, AddressOf SBO_Application_AppEvent
                oApp.Run()

            Catch ex As Exception
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
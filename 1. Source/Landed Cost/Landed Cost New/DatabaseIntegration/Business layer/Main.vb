Module Main
#Region "... Main ..."
    Sub Main()
        Try
            oGFun.SetApplication()
            If Not oGFun.CookieConnect() = 0 Then
                oApplication.MessageBox("DIAPI Connection Failed")
                End
            End If
            If Not oGFun.ConnectionContext() = 0 Then
                oApplication.MessageBox("Company Connection Failed")
                End
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Application Not Found")
            System.Windows.Forms.Application.ExitThread()
        Finally
        End Try
        Try
            Try
                Dim oTableCreation As New TableCreation
                EventHandler.SetEventFilter()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                System.Windows.Forms.Application.ExitThread()
            Finally
            End Try
            oApplication.StatusBar.SetText("Connected Successfully.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Addon Connection Method Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region
End Module

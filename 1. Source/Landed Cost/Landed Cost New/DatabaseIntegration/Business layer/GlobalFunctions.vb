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

End Class

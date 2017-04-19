
Imports System.Xml
Imports System.IO

Namespace AE_ISDN_A06

    Module modExchangeRate

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim igval As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function ExchangeRate(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, _
                                        ByVal sCurrency As String, ByVal dSysRate As Double, ByVal dDate As Date, ByRef sErrDesc As String) As Long

            'Function   :   ExchangeRate()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   Shibin
            'Date       :   08/03/2017
            'Change     :

            Dim sFuncName As String = String.Empty

            Dim sSQL As String = String.Empty


            Try
                sFuncName = "ItemMaster()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & dDate, sFuncName)


                ExchangeRate = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                ExchangeRate = RTN_ERROR

                sErrDesc = Err.Description


                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try
        End Function



    End Module
End Namespace





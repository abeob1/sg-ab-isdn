
Imports System.Xml
Imports System.IO

Namespace AE_ISDN_A06

    Module modItemMaster

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim igval As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function ItemMaster(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, _
                                        ByVal sMasterCode As String, ByRef sErrDesc As String) As Long

            'Function   :   ItemMaster()
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
            Dim iHashCode As Integer = 0
            Dim flg1 As Boolean = False

            Dim oItemMaster As SAPbobsCOM.Items = Nothing
            Dim oTItemMaster As SAPbobsCOM.Items = Nothing

            Dim orsGroup As SAPbobsCOM.Recordset = Nothing

            Dim ItemName As String = ""
            Dim sSQL As String = String.Empty

            oItemMaster = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            Try
                sFuncName = "ItemMaster()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterCode, sFuncName)

                '' initialize the objects 
                oItemMaster = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                oTItemMaster = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting ItemMaster Sync Function" & oTragetCompany.CompanyDB, sFuncName)

                If oItemMaster.GetByKey(sMasterCode) Then

                    orsGroup = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    sSQL = "Select ""ItemName"" from ""OITM"" where ""ItemCode"" = '" & sMasterCode & "'"
                    'orsGroup.DoQuery(String.Format("Select ""ItemName"" from ""OITM"" where ""ItemCode"" = {0}", sMasterCode))
                    orsGroup.DoQuery(sSQL)
                    ItemName = orsGroup.Fields.Item(0).Value
                    sSQL = ""
                    orsGroup = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'orsGroup.DoQuery(String.Format("Select ""ItemCode"" from ""OITM"" where ""ItemName"" = '{0}'", ItemName))
                    sSQL = "Select ""ItemName"" from ""OITM"" where ""ItemName"" = '" & ItemName & "'"
                    orsGroup.DoQuery(sSQL)
                    If orsGroup.RecordCount = 1 Then
                        'oTItemMaster.ItemsGroupCode = orsGroup.Fields.Item(0).Value
                        oItemMaster.UserFields.Fields.Item("U_Global").Value = "Global Item"
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Item Master", sFuncName)
                    '
                    ival = oItemMaster.Update() ''Update the Item Master UDF in Source company
                    If ival <> 0 Then
                        IsError = True
                        oTragetCompany.GetLastError(iErr, sErr)
                        Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                        sErrDesc = sErr
                        ItemMaster = RTN_ERROR
                        Exit Function
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterCode
                    ItemMaster = RTN_ERROR
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                ItemMaster = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                ItemMaster = RTN_ERROR
                iHashCode = Err.Number
                sErrDesc = Err.Description


                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try
        End Function

        

    End Module
End Namespace





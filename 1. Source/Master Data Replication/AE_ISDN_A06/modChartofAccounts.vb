
Imports System.Xml
Imports System.IO

Namespace AE_ISDN_A06

    Module modChartofAccounts

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function ChartofAccounts(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, _
                                        ByVal sMasterdatacode As String, ByRef sErrDesc As String) As Long

            'Function   :   ChartofAccounts()
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
            'Date       :   07/03/2017
            'Change     :

            Dim sFuncName As String = String.Empty
            Dim iHashCode As Integer = 0
            Dim flg1 As Boolean = False
            Dim oCOA As SAPbobsCOM.ChartOfAccounts = Nothing
            Dim oTargetCOA As SAPbobsCOM.ChartOfAccounts = Nothing

            Try
                sFuncName = "ChartofAccounts()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

                'oHoldingCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                'oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                '' initialize the objects 
                oCOA = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                oTargetCOA = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting COA Sync Function" & oTragetCompany.CompanyDB, sFuncName)
               
                If oCOA.GetByKey(sMasterdatacode) Then
                    If oTargetCOA.GetByKey(sMasterdatacode) = True Then
                        flg1 = True
                        oTargetCOA.Name = oCOA.Name
                        oTargetCOA.FatherAccountKey = oCOA.FatherAccountKey

                        oTargetCOA.ActiveAccount = oCOA.ActiveAccount
                        'oTargetCOA.Category = oCOA.Category
                        oTargetCOA.BudgetAccount = oCOA.BudgetAccount
                        oTargetCOA.CashAccount = oCOA.CashAccount
                        oTargetCOA.AcctCurrency = oCOA.AcctCurrency
                        oTargetCOA.ExternalCode = oCOA.ExternalCode
                        oTargetCOA.Protected = oCOA.Protected
                        oTargetCOA.LockManualTransaction = oCOA.LockManualTransaction
                        oTargetCOA.Details = oCOA.Details
                        oTargetCOA.ValidFor = oCOA.ValidFor
                        oTargetCOA.FrozenFor = oCOA.FrozenFor
                        oTargetCOA.ValidFrom = oCOA.ValidFrom
                        oTargetCOA.ValidTo = oCOA.ValidTo
                        oTargetCOA.ValidRemarks = oCOA.ValidRemarks
                        oTargetCOA.FrozenFrom = oCOA.FrozenFrom
                        oTargetCOA.FrozenTo = oCOA.FrozenTo
                        oTargetCOA.FrozenRemarks = oCOA.FrozenRemarks
                        oTargetCOA.LoadingType = oCOA.LoadingType
                        oTargetCOA.PlanningLevel = oCOA.PlanningLevel
                        oTargetCOA.ProjectRelevant = oCOA.ProjectRelevant
                        oTargetCOA.ProjectCode = oCOA.ProjectCode

                    Else

                        oTargetCOA.Code = oCOA.Code
                        oTargetCOA.Name = oCOA.Name

                        oTargetCOA.AccountType = oCOA.AccountType
                        oTargetCOA.FatherAccountKey = oCOA.FatherAccountKey
                        oTargetCOA.ActiveAccount = oCOA.ActiveAccount
                        oTargetCOA.Category = oCOA.Category
                        oTargetCOA.BudgetAccount = oCOA.BudgetAccount
                        oTargetCOA.CashAccount = oCOA.CashAccount
                        oTargetCOA.AcctCurrency = oCOA.AcctCurrency
                        oTargetCOA.ExternalCode = oCOA.ExternalCode
                        oTargetCOA.Protected = oCOA.Protected
                        oTargetCOA.LockManualTransaction = oCOA.LockManualTransaction
                        oTargetCOA.Details = oCOA.Details
                        oTargetCOA.ValidFor = oCOA.ValidFor
                        oTargetCOA.FrozenFor = oCOA.FrozenFor
                        oTargetCOA.ValidFrom = oCOA.ValidFrom
                        oTargetCOA.ValidTo = oCOA.ValidTo
                        oTargetCOA.ValidRemarks = oCOA.ValidRemarks
                        oTargetCOA.FrozenFrom = oCOA.FrozenFrom
                        oTargetCOA.FrozenTo = oCOA.FrozenTo
                        oTargetCOA.FrozenRemarks = oCOA.FrozenRemarks
                        oTargetCOA.LoadingType = oCOA.LoadingType
                        oTargetCOA.PlanningLevel = oCOA.PlanningLevel
                        oTargetCOA.ProjectRelevant = oCOA.ProjectRelevant
                        oTargetCOA.ProjectCode = oCOA.ProjectCode

                    End If
                    If flg1 = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating COA", sFuncName)
                        '
                        ival = oTargetCOA.Update() ''Update the Account in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ChartofAccounts = RTN_ERROR
                            Exit Function
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding COA", sFuncName)
                        ival = oTargetCOA.Add() ''Add the Account in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ChartofAccounts = RTN_ERROR
                            Exit Function
                        End If

                    End If


               
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    ChartofAccounts = RTN_ERROR
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                    Exit Function
                End If

                ChartofAccounts = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                ChartofAccounts = RTN_ERROR
                iHashCode = Err.Number
                sErrDesc = Err.Description
                If iHashCode = "-1114" Then
                    sErrDesc = "UDFs' missing in the target company"
                Else
                    sErrDesc = exc.Message.ToString()
                End If

                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCOA)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetCOA)
                oCOA = Nothing
                oTargetCOA = Nothing

            End Try
        End Function

        Public Sub COA_Assignment(ByRef oChartofAccount As SAPbobsCOM.ChartOfAccounts, ByRef oChartofAccounts As SAPbobsCOM.ChartOfAccounts)

            sFuncName = "COA_Assignment()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oChartofAccount.Name = oChartofAccounts.Name
            oChartofAccount.AcctCurrency = oChartofAccounts.AcctCurrency
            'oChartofAccount.TaxExemptAccount = SAPbobsCOM.BoYesNoEnum.tYES
            'oChartofAccount.TaxLiableAccount = SAPbobsCOM.BoYesNoEnum.tYES
            'oChartofAccount.RevaluationCoordinated = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.ReconciledAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.LoadingType = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.LiableForAdvances = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.DatevFirstDataEntry = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.DatevAutoAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES


            If oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Expenses Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Expenses
            ElseIf oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Other Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Other
            ElseIf oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Revenues Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Revenues
            End If

            If oChartofAccounts.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tYES

            Else
                oChartofAccount.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            oChartofAccount.DataExportCode = oChartofAccounts.DataExportCode

            If oChartofAccounts.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tNO

            ''If oChartofAccounts.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES Then
            ''    oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES
            ''Else
            ''    oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tNO
            ''End If
            oChartofAccount.FrozenFrom = oChartofAccounts.FrozenFrom
            oChartofAccount.FrozenTo = oChartofAccounts.FrozenTo

            If oChartofAccounts.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            oChartofAccount.FatherAccountKey = oChartofAccounts.FatherAccountKey

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

        End Sub

    End Module
End Namespace





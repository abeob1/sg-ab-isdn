Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics


Namespace AE_ISDN_A02
    <FormAttribute("AE_ISDN_A02.SOA", "SOA.b1f")>
    Public Class SOA
        Inherits UserFormBase

        Private _company As SAPbobsCOM.Company
        Public WithEvents SBO_Application As SAPbouiCOM.Application

        Sub New(company As SAPbobsCOM.Company, application As SAPbouiCOM.Application, ByRef sErrDesc As String)
            ' TODO: Complete member initialization 
            Try
                _company = company
                SBO_Application = application
            Catch ex As Exception
                sErrDesc = ex.Message
            End Try

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("BPFrom").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("Item_2").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("BPTo").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_6").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("Item_7").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.Matrix)
            Me.CheckBox0 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.CheckBox)
            Me.Button0 = CType(Me.GetItem("Item_10").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.EditText5 = CType(Me.GetItem("Item_12").Specific, SAPbouiCOM.EditText)
            Me.Button4 = CType(Me.GetItem("Item_14").Specific, SAPbouiCOM.Button)
            Me.ComboBox0 = CType(Me.GetItem("Item_15").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText4 = CType(Me.GetItem("Item_16").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox1 = CType(Me.GetItem("Item_18").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText5 = CType(Me.GetItem("Item_19").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents EditText5 As SAPbouiCOM.EditText


        Public Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_ItemEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByVal FormUID As String
            '                       FormUID = set the FormUID
            '                   ByRef pVal As SAPbouiCOM.ItemEvent
            '                       pVal = set the SAP UI ItemEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************

            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim p_oDVJE As DataView = Nothing
            Dim oDTDistinct As DataTable = Nothing
            Dim oDTRowFilter As DataTable = Nothing

            'Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm

            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                ''If Not IsNothing(p_oDICompany) Then
                ''    If Not p_oDICompany.Connected Then
                ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                ''        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ''    End If
                ''End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID
                        Case "SOA"

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.Item(FormUID)
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                Try
                                    If oCFLEvento.BeforeAction = False Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        If pVal.ItemUID = "BPFrom" Then 'BP From
                                            p_iSearch = 0
                                            p_iEmail = 0
                                            p_iPrint = 0

                                            oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("CardName", 0)
                                            oForm.Items.Item("BPFrom").Specific.string = oDataTable.GetValue("CardCode", 0)
                                        End If
                                        If pVal.ItemUID = "BPTo" Then 'BP To
                                            Try
                                                p_iSearch = 0
                                                p_iEmail = 0
                                                p_iPrint = 0

                                                oForm.Items.Item("Item_3").Specific.string = oDataTable.GetValue("CardName", 0)
                                                oForm.Items.Item("BPTo").Specific.string = oDataTable.GetValue("CardCode", 0)

                                            Catch ex As Exception

                                            End Try
                                        End If

                                    End If
                                Catch ex As Exception
                                End Try

                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_9" Then
                                    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                    Try
                                        Dim oMAtrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_8").Specific
                                        Dim oCheck As SAPbouiCOM.CheckBox = oForm.Items.Item("Item_9").Specific
                                        Dim ocheckColumn As SAPbouiCOM.CheckBox

                                        If oCheck.Checked = True Then
                                            For mjs As Integer = 1 To oMAtrix.RowCount
                                                ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                ocheckColumn.Checked = True
                                            Next mjs
                                        Else
                                            For mjs As Integer = 1 To oMAtrix.RowCount
                                                ocheckColumn = oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                ocheckColumn.Checked = False
                                            Next mjs
                                        End If

                                        p_iSearch = 0
                                        p_iEmail = 0
                                        p_iPrint = 0

                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                    Exit Sub
                                End If

                            End If
                            'Or pVal.ItemUID = "BPFrom" Or pVal.ItemUID = "BPTo" 
                            If pVal.ItemChanged = True And (pVal.ItemUID = "BPFrom" Or pVal.ItemUID = "BPTo" Or pVal.ItemUID = "Item_7" Or pVal.ItemUID = "Item_15" Or pVal.ItemUID = "Item_18" Or pVal.ItemUID = "Item_8") Then

                                p_iSearch = 0
                                p_iEmail = 0
                                p_iPrint = 0
                            End If

                    End Select
                Else
                    Select Case pVal.FormUID
                        Case "SOA"

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                                'Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                ' ''  Set_Conditions(oForm, "")
                                'Dim oItem1 As SAPbouiCOM.Item = oForm.Items.Item("Item_18")
                                'Dim oCombo1 As SAPbouiCOM.ComboBox = oItem1.Specific
                                'oCombo1.Select(0)

                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                p_iSOA = 0
                                p_iSearch = 0
                                p_iEmail = 0
                                p_iPrint = 0

                            End If
                            'And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE 
                            If pVal.ItemUID = "Item_8" And pVal.BeforeAction = True And pVal.ColUID = "Col_0" And pVal.ItemChanged = False Then
                                p_iPrint = 0
                                p_iEmail = 0
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_10" And p_iSearch = 0 Then
                                    'SBO_Application.MessageBox(" Inner = " & pVal.InnerEvent)
                                    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm

                                    Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        p_iSearch += 1
                                        SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If HeaderValidation(oForm, sErrDesc, SBO_Application) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        'SBO_Application.SetStatusBarMessage("SOAValidation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        'If SOAValidation(oForm, sErrDesc, SBO_Application) = 0 Then
                                        '    BubbleEvent = False
                                        '    Exit Sub
                                        'End If
                                        SBO_Application.SetStatusBarMessage("Loading Data ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Loading_AgingDetails()", sFuncName)
                                        If Loading_AgingDetails(oForm, SBO_Application, p_oDICompany, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oForm.Items.Item("Item_12").Specific.String = ""
                                        SBO_Application.SetStatusBarMessage("Loading Data Completed Successfully ........!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        'p_iSearch = 0
                                        Exit Sub


                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try

                                End If


                                If pVal.ItemUID = "Item_11" And p_iEmail = 0 Then
                                    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                    Dim sTargetFileName As String = String.Empty
                                    Dim sRptFileName As String = String.Empty

                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing

                                    Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Item_18")
                                    Dim oCombo As SAPbouiCOM.ComboBox = oItem.Specific
                                    Dim oItem1 As SAPbouiCOM.Item = oForm.Items.Item("Item_15")
                                    Dim oCombo1 As SAPbouiCOM.ComboBox = oItem1.Specific
                                    Try
                                        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SOAValidation()", sFuncName)
                                        p_iEmail += 1

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RowValidation()", sFuncName)
                                        SBO_Application.SetStatusBarMessage("Validating the Records .... !", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                        If RowValidation(oForm, SBO_Application, sErrDesc, SBO_Application) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        Try
                                            If oCombo.Selected.Value = "-" Then
                                                SBO_Application.StatusBar.SetText("Validation Error!  SOA format couldn`t be empty ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Catch ex As Exception
                                            SBO_Application.StatusBar.SetText("Validation Error!  SOA format couldn`t be empty ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        End Try


                                        oMatrix = oForm.Items.Item("Item_8").Specific
                                      
                                        sRptFileName = System.Windows.Forms.Application.StartupPath & "\Statement_of_Account_V2.rpt"
                                        'If oCombo.Selected.Value <> "" Then
                                        If oCombo.Selected.Value = "PDF" Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening PDF.......", sFuncName)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToPDF() ", sFuncName)
                                            sTargetFileName = "Statement of Account_" & Format(Now.Date, "dd-MM-yyyy") & ".PDF"
                                            sTargetFileName = System.Windows.Forms.Application.StartupPath & "\" & sTargetFileName
                                            For mjs As Integer = 1 To oMatrix.RowCount
                                                oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                If oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String = "EMAIL" And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Successfully Sent Email" Then
                                                    oForm.Items.Item("Item_12").Specific.String = "Processing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                                    SBO_Application.SetStatusBarMessage("Exporting SOA to PDF .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    'If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                    '              System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
                                                    '    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"

                                                    If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                                 System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then

                                                        Throw New ArgumentException(sErrDesc)
                                                    End If

                                                    SBO_Application.SetStatusBarMessage("Sending Email .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    ''   If SendEmailNotification(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_6").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
                                                    If SendEmailNotification(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_6").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
                                                        oForm.Items.Item("Item_12").Specific.String = "Email Processing is Failed ......... "
                                                        Dim sErrMsg As String = sErrDesc
                                                        sErrDesc = ""
                                                        '' Throw New ArgumentException(sErrDesc)
                                                        SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                          System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                          oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Fail To Send", sErrMsg, p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                        End If
                                                    Else
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Successfully Sent Email"
                                                        oForm.Items.Item("Item_12").Specific.String = "Email Processing is Completed ......... "
                                                        SBO_Application.SetStatusBarMessage("Email Processing is Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                         System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                         oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Successfully Sent Email", "", p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                        End If

                                                    End If
                                                ElseIf oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String <> "EMAIL" Then
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "EMail is not send has SOA Method is 'MAIL' "
                                                    oForm.Items.Item("Item_12").Specific.String = "Email Processing is Not Completed -  SOA Method is 'MAIL' ......... "
                                                    SBO_Application.SetStatusBarMessage("Email Processing is Not Completed -  SOA Method is 'MAIL' ..........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                            Next mjs
                                            'oForm.Items.Item("Item_12").Specific.String = "Email Processing is Completed ......... "
                                            'SBO_Application.SetStatusBarMessage("Email Processing is Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        Else
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Opening HTML.......", sFuncName)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExportToHTML() ", sFuncName)
                                            ''  sTargetFileName = "Statement of Account_" & Format(Now.Date, "dd-MM-yyyy") & ".HTML"
                                            sTargetFileName = System.Windows.Forms.Application.StartupPath '' & "\" & sTargetFileName ''
                                            For mjs As Integer = 1 To oMatrix.RowCount
                                                oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                If oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String = "EMAIL" And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Successfully Sent Email" Then
                                                    oForm.Items.Item("Item_12").Specific.String = "Processing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                                    SBO_Application.SetStatusBarMessage("Exporting SOA to HTML .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    'If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                    '              System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
                                                    '    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"

                                                    If ExportToHTML(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                                 System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then

                                                        Throw New ArgumentException(sErrDesc)
                                                    End If

                                                    SBO_Application.SetStatusBarMessage("Sending Email .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    ''   If SendEmailNotification(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_6").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
                                                    If SendEmailNotificationHTML(sTargetFileName, oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_6").Cells.Item(mjs).Specific.String, sErrDesc) <> RTN_SUCCESS Then
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Fail To Send"
                                                        oForm.Items.Item("Item_12").Specific.String = "Email Processing is Failed ......... "
                                                        Dim sErrMsg As String = sErrDesc
                                                        sErrDesc = ""
                                                        '' Throw New ArgumentException(sErrDesc)
                                                        SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                          System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                          oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Fail To Send", sErrMsg, p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                        End If
                                                    Else
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Successfully Sent Email"
                                                        oForm.Items.Item("Item_12").Specific.String = "Email Processing is Completed ......... "
                                                        SBO_Application.SetStatusBarMessage("Email Processing is Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        SBO_Application.SetStatusBarMessage("Updating the Transaction in Log Table .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        If TransactionLog(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, _
                                                                         System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String, _
                                                                         oMatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String, "Successfully Sent Email", "", p_oDICompany, sErrDesc) = RTN_SUCCESS Then
                                                        End If

                                                    End If
                                                ElseIf oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String <> "EMAIL" Then
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "EMail is not send has SOA Method is 'MAIL' "
                                                    oForm.Items.Item("Item_12").Specific.String = "Email Processing is Not Completed -  SOA Method is 'MAIL' ......... "
                                                    SBO_Application.SetStatusBarMessage("Email Processing is Not Completed -  SOA Method is 'MAIL' ..........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                End If
                                            Next mjs
                                        End If
                                        '  p_iEmail = 0
                                        Exit Sub
                                        'Else
                                        'SBO_Application.SetStatusBarMessage("Please select a SOA Format.........", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                        'End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        SBO_Application.SetStatusBarMessage(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                                If pVal.ItemUID = "Item_14" And p_iPrint = 0 Then
                                    Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                                    Dim sRptFileName As String = String.Empty
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing
                                    Dim sTargetFileName As String = String.Empty

                                    Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Item_18")
                                    Dim oCombo As SAPbouiCOM.ComboBox = oItem.Specific
                                    Try
                                        p_iPrint += 1
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        SBO_Application.SetStatusBarMessage("Validating the Records .... !", SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                                        If RowValidation(oForm, SBO_Application, sErrDesc, SBO_Application) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oMatrix = oForm.Items.Item("Item_8").Specific
                                        sRptFileName = System.Windows.Forms.Application.StartupPath & "\Statement_of_Account_V2.rpt"

                                        Dim dir As New DirectoryInfo(System.Windows.Forms.Application.StartupPath)
                                        Dim hdDirectoryInWhichToSearch As New DirectoryInfo(System.Windows.Forms.Application.StartupPath)

                                        Dim filesInDirDelete As FileInfo() = hdDirectoryInWhichToSearch.GetFiles("*.pdf")
                                        Dim FilesDelete As FileInfo() = dir.GetFiles("*.pdf")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Delete already exist PDF files.......", sFuncName)
                                        For Each file As FileInfo In filesInDirDelete
                                            Dim pdfFileDelete As String = file.FullName
                                            System.IO.File.Delete(pdfFileDelete)
                                        Next
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Print_Report() ", sFuncName)
                                        If (SBO_Application.MessageBox("Do you need to Print or Preview in PDF?", 1, "Print", "PDF")) <> 1 Then
                                            BubbleEvent = False
                                            For mjs As Integer = 1 To oMatrix.RowCount
                                                sTargetFileName = "Statement of Account_" & mjs & "_" & Format(Now.Date, "dd-MM-yyyy") & ".pdf"
                                                sTargetFileName = System.Windows.Forms.Application.StartupPath & "\" & sTargetFileName
                                                oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific

                                                'If oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String = "MAIL" And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Printed Successfully " Then
                                                If oCheck.Checked And oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String <> "Printed Successfully " Then
                                                    oForm.Items.Item("Item_12").Specific.String = "Processing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                                    SBO_Application.SetStatusBarMessage("Exporting SOA to PDF .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    If ExportToPDF(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                                  System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), "2", sTargetFileName, sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "PDF preview been Failed"
                                                        oForm.Items.Item("Item_12").Specific.String = "PDF Preview been Failed ......... "
                                                        SBO_Application.SetStatusBarMessage("PDF Preview been Failed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        Throw New ArgumentException(sErrDesc)
                                                    Else
                                                        oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "PDF Preview been Success"
                                                        oForm.Items.Item("Item_12").Specific.String = "PDF Preview been Completed ......... "
                                                        SBO_Application.SetStatusBarMessage("PDF Preview been Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    End If
                                                    'ElseIf oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String <> "MAIL" Then
                                                    '    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Export to PDF is not done has SOA Method is 'MAIL' "
                                                End If

                                            Next mjs
                                            'oForm.Items.Item("Item_12").Specific.String = "PDF Preview been Completed ......... "
                                            'SBO_Application.SetStatusBarMessage("PDF Preview been Completed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PDF Files are opened to view.....", sFuncName)
                                            Dim filesInDir As FileInfo() = hdDirectoryInWhichToSearch.GetFiles("*.pdf")
                                            Dim Files As FileInfo() = dir.GetFiles("*.pdf")
                                            For Each file As FileInfo In filesInDir
                                                Dim pdfFile As String = file.FullName
                                                System.Diagnostics.Process.Start(pdfFile)
                                            Next
                                        Else
                                            For mjs As Integer = 1 To oMatrix.RowCount
                                                oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific
                                                'If oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String = "MAIL" Then
                                                If oCheck.Checked Then
                                                    oForm.Items.Item("Item_12").Specific.String = "Printing the BP -  " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                                    SBO_Application.SetStatusBarMessage("Printing SOA .... ! " & oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    If Print_Report(oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String, _
                                                                  System.DateTime.Parse(oForm.Items.Item("Item_7").Specific.String, format1, Globalization.DateTimeStyles.None), sRptFileName, SBO_Application, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then
                                                        oForm.Items.Item("Item_12").Specific.String = "Printing Failed ......... "
                                                        SBO_Application.SetStatusBarMessage("Printed Failed .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                        Throw New ArgumentException(sErrDesc)
                                                    End If
                                                    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Printed Successfully "
                                                    oForm.Items.Item("Item_12").Specific.String = "Printed Successfully ......... "
                                                    SBO_Application.SetStatusBarMessage("Printed Successfully .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                    'ElseIf oCheck.Checked And oMatrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.String <> "MAIL" Then
                                                    '    oMatrix.Columns.Item("Col_5").Cells.Item(mjs).Specific.String = "Cannot Print has SOA Method is 'EMAIL' "
                                                End If
                                            Next mjs
                                            'System.Threading.Thread.Sleep(100)
                                            'oForm.Items.Item("Item_12").Specific.String = "Printed Successfully ......... "
                                            'SBO_Application.SetStatusBarMessage("Printed Successfully .........", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Printed Successfully ", sFuncName)
                                            'oForm.Items.Item("Item_12").Specific.String = "Printing is Completed ......... "
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Print Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If
                            End If
                    End Select
                End If


                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                sErrDesc = exc.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try

        End Sub

        Private WithEvents Button4 As SAPbouiCOM.Button
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText


        Private Sub ComboBox0_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            'Throw New System.NotImplementedException()
            Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
            Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Item_15")
            Dim oItem1 As SAPbouiCOM.Item = oForm.Items.Item("Item_18")
            Dim oCombo As SAPbouiCOM.ComboBox = oItem.Specific
            Dim oCombo1 As SAPbouiCOM.ComboBox = oItem1.Specific
            Try
                If oCombo.Selected.Value = "MAIL" Then
                    '    Button4.Item.Enabled = False
                    '    Button1.Item.Enabled = True
                    'Else
                    oItem1.Enabled = False
                    oCombo1.Select(0)
                    Button1.Item.Enabled = False
                    Button4.Item.Enabled = True
                Else
                    oItem1.Enabled = True
                    Button1.Item.Enabled = True
                    Button4.Item.Enabled = False
                End If

            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText

        Private Sub ComboBox1_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox1.ComboSelectAfter
            Dim oForm As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
            Dim oItem As SAPbouiCOM.Item = oForm.Items.Item("Item_18")
            Dim oCombo As SAPbouiCOM.ComboBox = oItem.Specific
            Try
                If oCombo.Selected.Value = "PDF" Or oCombo.Selected.Value = "InLine" Then
                    Button4.Item.Enabled = False
                    Button1.Item.Enabled = True
                Else
                    Button1.Item.Enabled = False
                    Button4.Item.Enabled = True
                End If

            Catch ex As Exception

            End Try
        End Sub

 
    End Class
End Namespace

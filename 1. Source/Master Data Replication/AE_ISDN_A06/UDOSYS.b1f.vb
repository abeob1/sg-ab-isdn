
Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A06
    <FormAttribute("AE_ISDN_A06.UDOSYS_b1f", "UDOSYS.b1f")>
    Public Class UDOSYS_b1f
        Inherits UserFormBase
        Private WithEvents oForm As SAPbouiCOM.Form
        Private _company As SAPbobsCOM.Company
        Public WithEvents SBO_Application As SAPbouiCOM.Application
        Public variableAdd As Boolean = False

        Public Sub New(company As SAPbobsCOM.Company, application As SAPbouiCOM.Application, ByRef sErrDesc As String)
            Try
                _company = company
                SBO_Application = application
                oForm = SBO_Application.Forms.ActiveForm
              
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                EditText0.Item.Click()
                If CurrencyLoad(_company, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Exit Try

            Catch ex As Exception

            End Try
        End Sub
        Public Function CurrencyLoad(company As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

            Dim sSQLCode As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oform As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm

            Try
                sFuncName = "CurrencyLoad()"

                oform.Freeze(True)

                'Matrix0.Columns.Item("Col_4").Visible = False
                'Matrix0.Columns.Item("Col_5").Visible = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                sSQLCode = "SELECT Count(T0.""Code"")+1 FROM ""@AE_SYSRATE""  T0"
                Dim sRSCode As SAPbobsCOM.Recordset = Nothing
                sRSCode = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sRSCode.DoQuery(sSQLCode)
                If sRSCode.RecordCount > 0 Then
                    EditText3.Item.Specific.value = sRSCode.Fields.Item(0).Value
                End If

                sSQL = "SELECT ""CurrCode"" FROM  ""OCRN""  T0 "
                Dim sRs As SAPbobsCOM.Recordset = Nothing
                sRs = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sRs.DoQuery(sSQL)
                Dim curr As String = String.Empty

                If sRs.RecordCount > 0 Then
                    While Not sRs.EoF
                        curr = sRs.Fields.Item(0).Value
                        Matrix0.AddRow(1)
                        Matrix0.Columns.Item("Col_3").Cells.Item(Matrix0.RowCount).Specific.string = curr
                        sRs.MoveNext()
                    End While
                End If
                oform.Freeze(False)
                CurrencyLoad = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                oform.Freeze(False)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                CurrencyLoad = RTN_ERROR
            End Try
        End Function
        Public Overrides Sub OnInitializeComponent()
            Me.EditText0 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.EditText)
            Me.EditText2 = CType(Me.GetItem("Item_7").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.Matrix)
            Me.EditText3 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub OnCustomInitialize()

        End Sub

        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
            'BubbleEvent = True
            If BusinessObjectInfo.BeforeAction = False Then

                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                    If BusinessObjectInfo.ActionSuccess Then
                        variableAdd = True
                    End If
                End If

            End If
        End Sub
        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            'BubbleEvent = True
            Select Case pVal.MenuUID
                Case "1282"
                    Matrix0.Clear()
                    EditText0.Value = ""
                    EditText1.Value = ""
                    EditText2.Value = ""
                    EditText3.Value = ""

                    Try
                        If CurrencyLoad(_company, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Exit Try
                    Catch ex As Exception

                    End Try


                    'End If

            End Select
        End Sub
        Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
            'BubbleEvent = True
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Try

                If pVal.BeforeAction = False Then
                    Select Case pVal.FormUID
                        Case "SYSRATE"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                                If pVal.ItemUID = "1" Then
                                    If variableAdd = True Then
                                        variableAdd = False
                                        Try
                                            If CurrencyLoad(_company, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            Exit Try
                                        Catch ex As Exception

                                        End Try
                                    End If
                                End If
                            End If


                            'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_GOT_FOCUS Then
                            '    If pVal.ColUID = "Col_2" Then
                            '        Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()
                            '    End If
                            'End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then

                                If pVal.ColUID = "Col_1" Then
                                    Dim oCheckCol As SAPbouiCOM.CheckBox

                                    'Dim stest As String = "11"
                                    Dim dTolerance As Double = EditText2.Value
                                    Dim dDatefrom As String = EditText0.Value
                                    Dim dDateTo As String = EditText1.Value
                                    Dim oEditText As SAPbouiCOM.EditText
                                    Dim dValue As Double = 0.0
                                    oEditText = DirectCast(Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)
                                    oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                                    If oCheckCol.Checked = False Then
                                        'Matrix0.Columns.Item("Col_0").Editable = True
                                        oCheckCol.Checked = True
                                        'Matrix0.Columns.Item("Col_0").Editable = False
                                    End If

                                    If dTolerance = 0.0 Or Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.value = 0.0 Then
                                        SBO_Application.SetStatusBarMessage("Make sure Tolerance and Bank Rate value is entered", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()
                                        BubbleEvent = False
                                        Exit Try
                                    Else
                                        Matrix0.Columns.Item("Col_2").Editable = True
                                        Matrix0.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String = dTolerance * oEditText.Value
                                        Matrix0.Columns.Item("Col_2").Editable = False
                                        Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()

                                    End If
                                End If
                            End If

                    End Select
                Else

                    Select Case pVal.FormUID
                        Case "SYSRATE"
                            Dim sSQL As String = String.Empty
                            Dim oform As SAPbouiCOM.Form = SBO_Application.Forms.ActiveForm
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                'If p_iAddMode = 0 Then
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ItemUID = "1" Then
                                    Dim dTolerance As Double = EditText2.Value
                                    Dim dDatefrom As String = EditText0.Value
                                    Dim dDateTo As String = EditText1.Value

                                    If dDatefrom = "" And dDateTo = "" Then
                                        SBO_Application.SetStatusBarMessage("Please enter From and To Date befor Adding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        EditText0.Item.Click()
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    If dTolerance = 0.0 Then
                                        SBO_Application.SetStatusBarMessage("Please enter Tolerance Value", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        EditText2.Item.Click()
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    'If Matrix0.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.value = 0.0 Then
                                    '    SBO_Application.SetStatusBarMessage("Please enter  Bank Rate value", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()
                                    '    BubbleEvent = False
                                    '    Exit Try
                                    'End If

                                    sSQL = "select T1.""U_DATEFROM"",T1.""U_DATETO"" from ""@AE_SYSRATE"" T1 WHERE T1.""U_DATEFROM"" BETWEEN TO_DATE('" & dDatefrom & "','YYYYMMDD') AND TO_DATE('" & dDateTo & "','YYYYMMDD') OR T1.""U_DATETO"" BETWEEN TO_DATE('" & dDatefrom & "','YYYYMMDD') AND TO_DATE('" & dDateTo & "','YYYYMMDD')"
                                    Dim sRSCode As SAPbobsCOM.Recordset = Nothing
                                    sRSCode = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    sRSCode.DoQuery(sSQL)
                                    If sRSCode.RecordCount > 0 Then
                                        SBO_Application.StatusBar.SetText("Entry Already Exists for selected dates,Please check", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Return
                                    End If
                                    'p_iAddMode = 1

                                End If
                                'End If


                            End If

                    End Select
                End If

            Catch ex As Exception

            End Try
        End Sub
        Private Sub EditText0_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub


        Private Sub EditText1_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub


        Private Sub EditText2_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub
        Private WithEvents EditText3 As SAPbouiCOM.EditText

        Private Sub EditText3_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.KeyDownAfter
            'Throw New System.NotImplementedException()

        End Sub
    End Class
End Namespace

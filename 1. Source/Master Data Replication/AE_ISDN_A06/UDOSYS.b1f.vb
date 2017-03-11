
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

        Public Sub New(company As SAPbobsCOM.Company, application As SAPbouiCOM.Application, ByRef sErrDesc As String)
            Try
                _company = company
                SBO_Application = application

                'SBO_Application.ItemEvent += New SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent)
                'SBO_Application.FormDataEvent += New SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent)
                'SBO_Application.MenuEvent += New SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent)
                oForm = SBO_Application.Forms.ActiveForm
                'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                'Dim sboBrowser As SAPbouiCOM.DataBrowser = oForm.DataBrowser
                'sboBrowser.BrowseBy = "Item_1"
                'Dim FrmID As String = SBO_Application.Forms.ActiveForm.Type.ToString()
                EditText0.Item.Click()
                'EditText3.Item.Enabled = False
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
        
        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            BubbleEvent = True
        End Sub
        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            'BubbleEvent = True
            Select Case pVal.MenuUID


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
                            'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            '    If pVal.ColUID = "Col_0" Then
                            '        Dim oCheckCol As SAPbouiCOM.CheckBox
                            '        Dim oEditText As SAPbouiCOM.EditText
                            '        Dim dValue As Double = 0.0
                            '        oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                            '        oEditText = DirectCast(Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)
                            '        If oCheckCol.Checked = True Then
                            '            'oEditText.Item.Enabled = True
                            '            'oEditText.Value = dValue.ToString
                            '            Matrix0.Columns.Item("Col_1").Editable = True

                            '        End If
                            '    End If
                            'End If
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
                                    Dim dDatefrom As String = EditText1.Value
                                    Dim dDateTo As String = EditText2.Value
                                    Dim oEditText As SAPbouiCOM.EditText
                                    Dim dValue As Double = 0.0
                                    oEditText = DirectCast(Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText)
                                    oCheckCol = Matrix0.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                                    If oCheckCol.Checked = False Then
                                        'Matrix0.Columns.Item("Col_0").Editable = True
                                        oCheckCol.Checked = True
                                        'Matrix0.Columns.Item("Col_0").Editable = False
                                    End If
                                    If dDatefrom = "" And dDateTo = "" Then
                                        SBO_Application.SetStatusBarMessage("Please enter From and To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    If dTolerance = 0.0 And Matrix0.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.value = 0.0 Then
                                        SBO_Application.SetStatusBarMessage("Please enter Tolerance and Bank Rate value", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()
                                        BubbleEvent = False
                                        Exit Try
                                    Else
                                        Matrix0.Columns.Item("Col_2").Editable = True
                                        Matrix0.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String = dTolerance * oEditText.Value
                                        Matrix0.Columns.Item("Col_2").Editable = False
                                        Matrix0.Columns.Item("Col_1").Cells.Item(pVal.Row).Click()
                                      
                                    End If
                                  
                                    oForm.Freeze(False)
                                   
                                End If
                            End If

                    End Select
                Else
                   
                    Select Case pVal.FormUID
                        Case "SYSRATE"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "1" Then
                                    'Dim ocheck As CheckBox

                                End If

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

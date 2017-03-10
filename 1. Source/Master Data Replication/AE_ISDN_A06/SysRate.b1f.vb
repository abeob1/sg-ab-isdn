
Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AE_ISDN_A06
    <FormAttribute("UDO_FT_SYSRATE", "SysRate.b1f")>
    Public Class SysRate
        Inherits UserFormBase
        'Inherits UDOFormBase



        Private _company As SAPbobsCOM.Company
        Private WithEvents SBO_Application As SAPbouiCOM.Application
        Private WithEvents oForm As SAPbouiCOM.Form
        'Dim _company As SAPbobsCOM.Company
        'Dim SBO_Application As SAPbouiCOM.Application
        'Dim oForm As SAPbouiCOM.Form
        Public Sub New(company As SAPbobsCOM.Company, application As SAPbouiCOM.Application)
            Try
                _company = company
                SBO_Application = application
                'oForm = application.SBO_Application.Forms.ActiveForm
                'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            Catch ex As Exception
                'sErrDesc = ex.Message
            End Try
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Matrix0 = CType(Me.GetItem("0_U_G").Specific, SAPbouiCOM.Matrix)
            Me.EditText0 = CType(Me.GetItem("0_U_E").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("2_U_E").Specific, SAPbouiCOM.EditText)
            Me.EditText2 = CType(Me.GetItem("13_U_E").Specific, SAPbouiCOM.EditText)
            Me.EditText3 = CType(Me.GetItem("14_U_E").Specific, SAPbouiCOM.EditText)
            Me.EditText4 = CType(Me.GetItem("15_U_E").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            BubbleEvent = True
        End Sub
        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            BubbleEvent = True
        End Sub
        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            'Throw New System.NotImplementedException()

        End Sub
    End Class
End Namespace

''' <summary>
''' GGlobally whatever variable do you want declare here 
''' We can use any class and module from here  
''' </summary>
''' <remarks></remarks>
Module GolabalVariables
    Public oCompany As SAPbobsCOM.Company
    Public oGFun As New GlobalFunctions
    Public oForm As SAPbouiCOM.Form
    Public oFilters As SAPbouiCOM.EventFilters

    Public oFilter As SAPbouiCOM.EventFilter
    'GRPO 
    Public GRPOFormID As String = "143"
    Public oGRPO As New GRPO

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0
    Public sFuncName As String
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public p_iDebugMode As Int16 = 1
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public v_RetVal, v_ErrCode As Long
    Public v_ErrMsg As String = ""
    Public addonName As String = "Landed Cost"
    Public ShowFolderBrowserThread As Threading.Thread
    Public BankFileName As String
    Public sQuery As String

    Public p_dFrieghtAmount As Double = 0
    Public p_dInsuranceAmount As Double = 0


    Public p_odtCourier As DataTable
    Public p_odtForwarded As DataTable
    Public p_odtDimensionSurcharge As DataTable
    Public p_odtFuelSurcharge As DataTable
    Public p_odtIncoTable As DataTable
    Public p_odtZoneDetails As DataTable



End Module

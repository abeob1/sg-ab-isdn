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


    Public v_RetVal, v_ErrCode As Long
    Public v_ErrMsg As String = ""
    Public addonName As String = "Landed Cost"
    Public ShowFolderBrowserThread As Threading.Thread
    Public BankFileName As String
    Public sQuery As String


End Module

Namespace AE_ISDN_A02
    Public Class ClsFunction

        Private Shared m_Company As SAPbobsCOM.Company
        Public Shared Property Company() As SAPbobsCOM.Company
            Get
                Return m_Company
            End Get
            Set(value As SAPbobsCOM.Company)
                m_Company = value
            End Set
        End Property
    End Class
End Namespace


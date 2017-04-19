
Imports System.Xml
Imports System.IO

Namespace AE_ISDN_A06

    Module modBusinessPartner

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function BusinessPartner(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, _
                                        ByVal sMasterdatacode As String, ByRef sErrDesc As String) As Long

            'Function   :   BusinessPartner()
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

            Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
            Dim orsB As SAPbobsCOM.Recordset = Nothing
            Dim orsTarget As SAPbobsCOM.Recordset = Nothing
            Dim orsGroup As SAPbobsCOM.Recordset = Nothing
            Dim oTargetBP As SAPbobsCOM.BusinessPartners = Nothing
            Dim GroupName As String = ""

            Dim sSQL As String = String.Empty
            Dim oDVContact As DataView = Nothing
            Dim sSrcShipTypeName As String = String.Empty
            Dim sSrcPymntTrmsCod As String = String.Empty
            Try
                sFuncName = "BusinessPartner()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                '' initialize the objects 
                oBP = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                oTargetBP = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting BP Sync Function" & oTragetCompany.CompanyDB, sFuncName)

                If oBP.GetByKey(sMasterdatacode) Then
                    If oTargetBP.GetByKey(sMasterdatacode) = True Then
                        ' ''Update BP MAster
                        flg1 = True
                        orsTarget = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update BP " & sMasterdatacode, sFuncName)

                        oTargetBP.CardName = oBP.CardName
                        oTargetBP.CardType = oBP.CardType
                        oTargetBP.CardForeignName = oBP.CardForeignName
                        oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                        oTargetBP.DiscountPercent = oBP.DiscountPercent
                        oTargetBP.Address = oBP.Address
                        oTargetBP.EmailAddress = oBP.EmailAddress
                        oTargetBP.Phone1 = oBP.Phone1
                        oTargetBP.Phone2 = oBP.Phone2
                        oTargetBP.Cellular = oBP.Cellular
                        oTargetBP.Fax = oBP.Fax
                        oTargetBP.Password = oBP.Password
                        oTargetBP.BusinessType = oBP.BusinessType
                        oTargetBP.AdditionalID = oBP.AdditionalID
                        oTargetBP.VatIDNum = oBP.VatIDNum
                        oTargetBP.FederalTaxID = oBP.FederalTaxID
                        oTargetBP.Notes = oBP.Notes
                        oTargetBP.FreeText = oBP.FreeText
                        oTargetBP.AliasName = oBP.AliasName
                        oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                        oTargetBP.Valid = oBP.Valid
                        oTargetBP.Frozen = oBP.Frozen

                        oTargetBP.Website = oBP.Website
                        oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                        orsGroup = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                        GroupName = orsGroup.Fields.Item(0).Value

                        orsGroup = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))
                        If orsGroup.RecordCount = 1 Then
                            oTargetBP.GroupCode = orsGroup.Fields.Item(0).Value
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                        orsGroup = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        sSQL = "SELECT ""TrnspName"" FROM ""OSHP"" WHERE ""TrnspCode"" = (SELECT ""ShipType"" FROM ""OCRD"" WHERE ""CardCode"" = '" & oBP.CardCode & "')"
                        orsGroup.DoQuery(sSQL)
                        If orsGroup.RecordCount = 1 Then
                            sSrcShipTypeName = orsGroup.Fields.Item(0).Value
                        End If
                        If sSrcShipTypeName <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsGroup = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orsGroup.DoQuery(String.Format("SELECT ""TrnspCode"" FROM ""OSHP"" WHERE ""TrnspName"" = '{0}'", sSrcShipTypeName))
                            If orsGroup.RecordCount = 1 Then
                                oTargetBP.ShippingType = orsGroup.Fields.Item(0).Value
                            End If
                        End If

                        If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.ValidFrom = oBP.ValidFrom
                            oTargetBP.ValidTo = oBP.ValidTo
                            oTargetBP.ValidRemarks = oBP.ValidRemarks
                        End If
                        If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.FrozenFrom = oBP.FrozenFrom
                            oTargetBP.FrozenTo = oBP.FrozenTo
                            oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                        End If
                        If oTargetBP.Addresses.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.Addresses.Count - 1
                                oTargetBP.Addresses.SetCurrentLine(oTargetBP.Addresses.Count - 1)
                                oTargetBP.Addresses.Delete()
                                If oTargetBP.Addresses.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        'Handle add update/add new Address
                        If oBP.Addresses.Count > 0 And oBP.Addresses.AddressName <> "" Then
                            For i As Integer = 0 To oBP.Addresses.Count - 1
                                oBP.Addresses.SetCurrentLine(i)
                                oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                                oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                                oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                                oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                                oTargetBP.Addresses.Block = oBP.Addresses.Block
                                oTargetBP.Addresses.City = oBP.Addresses.City
                                oTargetBP.Addresses.County = oBP.Addresses.County
                                oTargetBP.Addresses.Country = oBP.Addresses.Country
                                oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                                oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                                oTargetBP.Addresses.State = oBP.Addresses.State
                                oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                                oTargetBP.Addresses.Street = oBP.Addresses.Street
                                oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                                oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber

                                'oTargetBP.Addresses.UserFields.Fields.Item("U_FAX").Value = oBP.Addresses.UserFields.Fields.Item("U_FAX").Value
                                oTargetBP.Addresses.Add()
                            Next
                            oTargetBP.BilltoDefault = oBP.BilltoDefault
                            oTargetBP.ShipToDefault = oBP.ShipToDefault

                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Payment Code", sFuncName)
                        orsGroup = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsGroup.DoQuery(String.Format("SELECT B.""PymntGroup"" FROM ""OCRD"" A INNER JOIN ""OCTG"" B ON B.""GroupNum"" = A.""GroupNum"" WHERE A.""CardCode"" = '{0}'", oBP.CardCode))
                        If orsGroup.RecordCount = 1 Then
                            sSrcPymntTrmsCod = orsGroup.Fields.Item(0).Value
                        End If
                        If sSrcPymntTrmsCod <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsGroup = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orsGroup.DoQuery(String.Format("SELECT ""GroupNum"" FROM ""OCTG"" WHERE ""PymntGroup"" = '{0}'", sSrcPymntTrmsCod))
                            If orsGroup.RecordCount = 1 Then
                                oTargetBP.PayTermsGrpCode = orsGroup.Fields.Item(0).Value
                            End If
                        End If

                        'oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                        oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                        oTargetBP.PriceListNum = oBP.PriceListNum
                        oTargetBP.DiscountPercent = oBP.DiscountPercent

                        oTargetBP.CreditLimit = oBP.CreditLimit
                        oTargetBP.MaxCommitment = oBP.MaxCommitment
                        oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                        oTargetBP.HouseBank = oBP.HouseBank
                        oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                        oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                        oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                        oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                        For imjs As Integer = 1 To oBP.BPPaymentMethods.PaymentMethodCode.Count
                            oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode
                            oTargetBP.BPPaymentMethods.Add()
                            'oRset_Tar.MoveNext()
                        Next imjs

                        ' ''BP Bank Details 
                        If oTargetBP.BPBankAccounts.Count > 0 Then

                            Dim delete As Boolean = False
                            For i As Integer = 0 To oTargetBP.BPBankAccounts.Count - 1
                                oTargetBP.BPBankAccounts.SetCurrentLine(oTargetBP.BPBankAccounts.Count - 1)
                                oTargetBP.BPBankAccounts.Delete()
                                If oTargetBP.BPBankAccounts.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                            oBP.BPBankAccounts.SetCurrentLine(i)
                            'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                            'If orsTarget.RecordCount = 1 Then
                            oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                            oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                            oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                            oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                            oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                            oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                            oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                            oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                            oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                            oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                            oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.State
                            oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Block
                            oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                            oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                            oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                            oTargetBP.BPBankAccounts.Add()
                            'Else
                            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Bank code is not available in target bank setup " & oBP.BPBankAccounts.BankCode, sFuncName)
                            'Console.WriteLine("In Target Company Bank Setup the Bank Code- " & oBP.BPBankAccounts.BankCode, sFuncName & "doesn't exist ")
                            'End If

                        Next
                        orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                        If orsTarget.RecordCount = 1 Then
                            Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                            Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                            oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                            oTargetBP.DefaultAccount = oBP.DefaultAccount

                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Contact Employees", sFuncName)

                        If oTargetBP.ContactEmployees.Count = 1 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                'oTargetBP.ContactEmployees.Add()
                                oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                oTargetBP.ContactEmployees.Add()
                            Next
                        ElseIf oTargetBP.ContactEmployees.Count > 0 Then
                            For imjs As Integer = 0 To oBP.ContactEmployees.Count - 1
                                oBP.ContactEmployees.SetCurrentLine(imjs)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oBP.ContactEmployees.Name " & oBP.ContactEmployees.Name, sFuncName)
                                If oBP.ContactEmployees.Name = "" Then Continue For
                                oDVContact.RowFilter = "Name='" & oBP.ContactEmployees.Name & "'"
                                If oDVContact.Count > 0 Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Index " & oDVContact(0)("No").ToString(), sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Name " & oDVContact(0)("Name").ToString(), sFuncName)
                                    oTargetBP.ContactEmployees.SetCurrentLine(Convert.ToInt32(oDVContact(0)("No").ToString()))
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigned", sFuncName)
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                Else

                                    oTargetBP.ContactEmployees.Add()
                                    oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                                    oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                                    oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                                    oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                                    oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2

                                    oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                                    oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                                    oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                                    oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                                    oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                                    oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                                    oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                                    oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                                    oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                                    oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.Remarks2
                                    oTargetBP.ContactEmployees.Password = oBP.ContactEmployees.Password
                                    oTargetBP.ContactEmployees.PlaceOfBirth = oBP.ContactEmployees.PlaceOfBirth
                                    oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                                    oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                                    oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                                    oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                                End If
                                ''oTargetBP.ContactEmployees.Add()
                            Next
                        End If
                        oTargetBP.ContactPerson = oBP.ContactPerson

                    Else
                        ' ''Create BP MAster
                        Dim ors As SAPbobsCOM.Recordset = Nothing
                        ors = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsTarget = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If oTargetBP.GetByKey(oBP.CardCode) Then
                            Throw New Exception(String.Format("BP : {0} aldready existed in Branch : {1}", oBP.CardCode, oTragetCompany.CompanyName))
                        End If
                        oTargetBP.CardCode = oBP.CardCode
                        oTargetBP.CardName = oBP.CardName
                        oTargetBP.CardType = oBP.CardType
                        oTargetBP.CardForeignName = oBP.CardForeignName

                        oTargetBP.CompanyPrivate = oBP.CompanyPrivate
                        If oBP.CardType = SAPbobsCOM.BoCardTypes.cCustomer Then
                            oTargetBP.Currency = "##"
                        End If
                        oTargetBP.DiscountPercent = oBP.DiscountPercent
                        oTargetBP.Address = oBP.Address
                        oTargetBP.EmailAddress = oBP.EmailAddress
                        oTargetBP.Phone1 = oBP.Phone1
                        oTargetBP.Phone2 = oBP.Phone2
                        oTargetBP.Cellular = oBP.Cellular
                        oTargetBP.Fax = oBP.Fax
                        oTargetBP.Password = oBP.Password
                        oTargetBP.BusinessType = oBP.BusinessType
                        oTargetBP.AdditionalID = oBP.AdditionalID
                        oTargetBP.VatIDNum = oBP.VatIDNum
                        oTargetBP.FederalTaxID = oBP.FederalTaxID
                        oTargetBP.Notes = oBP.Notes
                        oTargetBP.FreeText = oBP.FreeText
                        oTargetBP.AliasName = oBP.AliasName
                        oTargetBP.GlobalLocationNumber = oBP.GlobalLocationNumber
                        oTargetBP.Valid = oBP.Valid
                        oTargetBP.Frozen = oBP.Frozen
                        oTargetBP.Website = oBP.Website
                        oTargetBP.UnifiedFederalTaxID = oBP.UnifiedFederalTaxID

                        ors.DoQuery(String.Format("Select ""GroupName"" from ""OCRG"" where ""GroupCode"" = {0}", oBP.GroupCode))
                        GroupName = ors.Fields.Item(0).Value

                        ors = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ors.DoQuery(String.Format("Select ""GroupCode"" from ""OCRG"" where ""GroupName"" = '{0}'", GroupName))

                        If ors.RecordCount = 1 Then
                            oTargetBP.GroupCode = ors.Fields.Item(0).Value
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                        ors.DoQuery(String.Format("SELECT ""TrnspName"" FROM ""OSHP"" WHERE ""TrnspCode"" = (SELECT ""ShipType"" FROM ""OCRD"" WHERE ""CardCode"" = '{0}')", oBP.CardCode))
                        If ors.RecordCount = 1 Then
                            sSrcShipTypeName = ors.Fields.Item(0).Value
                        End If
                        If sSrcShipTypeName <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsTarget.DoQuery(String.Format("SELECT ""TrnspCode"" FROM ""OSHP"" WHERE ""TrnspName"" = '{0}'", sSrcShipTypeName))
                            If orsTarget.RecordCount = 1 Then
                                oTargetBP.ShippingType = orsTarget.Fields.Item(0).Value
                            End If
                        End If

                        'oTargetBP.DebitorAccount = oBP.DebitorAccount
                        If oTargetBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.ValidFrom = oBP.ValidFrom
                            oTargetBP.ValidTo = oBP.ValidTo
                            oTargetBP.ValidRemarks = oBP.ValidRemarks
                        End If
                        If oTargetBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
                            oTargetBP.FrozenFrom = oBP.FrozenFrom
                            oTargetBP.FrozenTo = oBP.FrozenTo
                            oTargetBP.FrozenRemarks = oBP.FrozenRemarks
                        End If

                        'oTargetBP.UserFields.Fields.Item("U_PreferredFowarder").Value = oBP.UserFields.Fields.Item("U_PreferredFowarder").Value

                        If oBP.Addresses.Count > 0 And oBP.Addresses.AddressName <> "" Then
                            For i As Integer = 0 To oBP.Addresses.Count - 1

                                oBP.Addresses.SetCurrentLine(i)
                                oTargetBP.Addresses.AddressName = oBP.Addresses.AddressName
                                oTargetBP.Addresses.AddressName2 = oBP.Addresses.AddressName2
                                oTargetBP.Addresses.AddressName3 = oBP.Addresses.AddressName3
                                oTargetBP.Addresses.AddressType = oBP.Addresses.AddressType
                                oTargetBP.Addresses.Block = oBP.Addresses.Block
                                oTargetBP.Addresses.City = oBP.Addresses.City
                                oTargetBP.Addresses.County = oBP.Addresses.County
                                oTargetBP.Addresses.Country = oBP.Addresses.Country
                                oTargetBP.Addresses.StreetNo = oBP.Addresses.StreetNo
                                oTargetBP.Addresses.TypeOfAddress = oBP.Addresses.TypeOfAddress
                                oTargetBP.Addresses.State = oBP.Addresses.State
                                oTargetBP.Addresses.ZipCode = oBP.Addresses.ZipCode
                                oTargetBP.Addresses.Street = oBP.Addresses.Street
                                oTargetBP.Addresses.BuildingFloorRoom = oBP.Addresses.BuildingFloorRoom
                                oTargetBP.Addresses.GlobalLocationNumber = oBP.Addresses.GlobalLocationNumber
                                'oTargetBP.Addresses.UserFields.Fields.Item("U_FAX").Value = oBP.Addresses.UserFields.Fields.Item("U_FAX").Value
                               
                                oTargetBP.Addresses.Add()
                            Next
                            oTargetBP.BilltoDefault = oBP.BilltoDefault
                            oTargetBP.ShipToDefault = oBP.ShipToDefault
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping Code", sFuncName)
                        ors.DoQuery(String.Format("SELECT B.""PymntGroup"" FROM ""OCRD"" A INNER JOIN ""OCTG"" B ON B.""GroupNum"" = A.""GroupNum"" WHERE A.""CardCode"" = '{0}'", oBP.CardCode))
                        If ors.RecordCount = 1 Then
                            sSrcPymntTrmsCod = ors.Fields.Item(0).Value
                        End If
                        If sSrcPymntTrmsCod <> "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Shipping type code in target ", sFuncName)
                            orsTarget.DoQuery(String.Format("SELECT ""GroupNum"" FROM ""OCTG"" WHERE ""PymntGroup"" = '{0}'", sSrcPymntTrmsCod))
                            If orsTarget.RecordCount = 1 Then
                                oTargetBP.PayTermsGrpCode = orsTarget.Fields.Item(0).Value
                            End If
                        End If
                        '''''''''''oTargetBP.PayTermsGrpCode = oBP.PayTermsGrpCode
                        oTargetBP.IntrestRatePercent = oBP.IntrestRatePercent
                        oTargetBP.PriceListNum = oBP.PriceListNum
                        oTargetBP.DiscountPercent = oBP.DiscountPercent

                        oTargetBP.CreditLimit = oBP.CreditLimit
                        oTargetBP.MaxCommitment = oBP.MaxCommitment
                        oTargetBP.EffectiveDiscount = oBP.EffectiveDiscount

                        oTargetBP.HouseBank = oBP.HouseBank
                        oTargetBP.HouseBankAccount = oBP.HouseBankAccount
                        oTargetBP.HouseBankBranch = oBP.HouseBankBranch
                        oTargetBP.HouseBankCountry = oBP.HouseBankCountry
                        oTargetBP.BPPaymentMethods.PaymentMethodCode = oBP.BPPaymentMethods.PaymentMethodCode

                        ''BP Bank Details 
                        For i As Integer = 0 To oBP.BPBankAccounts.Count - 1
                            oBP.BPBankAccounts.SetCurrentLine(i)
                            'orsTarget.DoQuery(String.Format("SELECT ""BankCode"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.BPBankAccounts.BankCode))
                            'If orsTarget.RecordCount = 1 Then
                            oTargetBP.BPBankAccounts.BankCode = oBP.BPBankAccounts.BankCode
                            oTargetBP.BPBankAccounts.Country = oBP.BPBankAccounts.Country
                            oTargetBP.BPBankAccounts.BPCode = oBP.BPBankAccounts.BPCode
                            oTargetBP.BPBankAccounts.AccountNo = oBP.BPBankAccounts.AccountNo
                            oTargetBP.BPBankAccounts.AccountName = oBP.BPBankAccounts.AccountName
                            oTargetBP.BPBankAccounts.Branch = oBP.BPBankAccounts.Branch
                            oTargetBP.BPBankAccounts.BICSwiftCode = oBP.BPBankAccounts.BICSwiftCode
                            oTargetBP.BPBankAccounts.InternalKey = oBP.BPBankAccounts.InternalKey
                            oTargetBP.BPBankAccounts.ControlKey = oBP.BPBankAccounts.ControlKey
                            oTargetBP.BPBankAccounts.IBAN = oBP.BPBankAccounts.IBAN
                            oTargetBP.BPBankAccounts.Street = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.State = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.Block = oBP.BPBankAccounts.Street
                            oTargetBP.BPBankAccounts.BuildingFloorRoom = oBP.BPBankAccounts.BuildingFloorRoom
                            oTargetBP.BPBankAccounts.City = oBP.BPBankAccounts.City
                            oTargetBP.BPBankAccounts.MandateID = oBP.BPBankAccounts.MandateID

                            oTargetBP.BPBankAccounts.Add()
                            'Else
                            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Bank code is not available in target bank setup " & oBP.BPBankAccounts.BankCode, sFuncName)
                            'Console.WriteLine("In Target Company Bank Setup the Bank Code- " & oBP.BPBankAccounts.BankCode, sFuncName & "doesn't exist ")
                            'End If

                        Next

                        orsTarget.DoQuery(String.Format("SELECT ""BankCode"",""CountryCod"" from ""ODSC"" where ""BankCode"" =  '{0}'", oBP.DefaultBankCode))
                        If orsTarget.RecordCount = 1 Then
                            Dim sBankcode As String = orsTarget.Fields.Item(0).Value
                            Dim sCtrycode As String = orsTarget.Fields.Item(1).Value

                            oTargetBP.DefaultBankCode = oBP.DefaultBankCode
                            oTargetBP.DefaultAccount = oBP.DefaultAccount

                        End If
                        For i As Integer = 0 To oBP.ContactEmployees.Count - 1
                            oBP.ContactEmployees.SetCurrentLine(i)
                            oTargetBP.ContactEmployees.Name = oBP.ContactEmployees.Name
                            oTargetBP.ContactEmployees.FirstName = oBP.ContactEmployees.FirstName
                            oTargetBP.ContactEmployees.MiddleName = oBP.ContactEmployees.MiddleName
                            oTargetBP.ContactEmployees.LastName = oBP.ContactEmployees.LastName
                            oTargetBP.ContactEmployees.Title = oBP.ContactEmployees.Title
                            oTargetBP.ContactEmployees.Position = oBP.ContactEmployees.Position
                            oTargetBP.ContactEmployees.Address = oBP.ContactEmployees.Address
                            oTargetBP.ContactEmployees.Phone1 = oBP.ContactEmployees.Phone1
                            oTargetBP.ContactEmployees.Phone2 = oBP.ContactEmployees.Phone2
                            oTargetBP.ContactEmployees.MobilePhone = oBP.ContactEmployees.MobilePhone
                            oTargetBP.ContactEmployees.Fax = oBP.ContactEmployees.Fax
                            oTargetBP.ContactEmployees.E_Mail = oBP.ContactEmployees.E_Mail
                            oTargetBP.ContactEmployees.Pager = oBP.ContactEmployees.Pager
                            oTargetBP.ContactEmployees.Remarks1 = oBP.ContactEmployees.Remarks1
                            oTargetBP.ContactEmployees.Remarks2 = oBP.ContactEmployees.InternalCode
                            oTargetBP.ContactEmployees.CityOfBirth = oBP.ContactEmployees.CityOfBirth
                            oTargetBP.ContactEmployees.DateOfBirth = oBP.ContactEmployees.DateOfBirth
                            oTargetBP.ContactEmployees.Gender = oBP.ContactEmployees.Gender
                            oTargetBP.ContactEmployees.Profession = oBP.ContactEmployees.Profession
                            oTargetBP.ContactEmployees.Active = oBP.ContactEmployees.Active
                            oTargetBP.ContactEmployees.Add()
                        Next

                        oTargetBP.ContactPerson = oBP.ContactPerson


                    End If


                    If flg1 = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP", sFuncName)
                        '
                        ival = oTargetBP.Update() ''Update the Account in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BusinessPartner = RTN_ERROR
                            Exit Function
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP", sFuncName)
                        ival = oTargetBP.Add() ''Add the Account in target company
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BusinessPartner = RTN_ERROR
                            Exit Function
                        End If

                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    BusinessPartner = RTN_ERROR
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                    Exit Function
                End If

                BusinessPartner = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                BusinessPartner = RTN_ERROR
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetBP)
                oBP = Nothing
                oTargetBP = Nothing
            End Try
        End Function

        

    End Module
End Namespace





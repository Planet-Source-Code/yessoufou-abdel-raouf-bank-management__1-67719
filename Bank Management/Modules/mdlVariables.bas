Attribute VB_Name = "mdlVariables"
Public Const title As String = "Online Banking"
Public UserName As String
Public UserPassword As String
Public UserID As String
Public UserRole As String

Public newUserID As String

Public blAddUser As Boolean
Public blUpdateUser As Boolean

Public blAddCustomer  As Boolean
Public blUpdateCustomer As Boolean

Public lst As ListItem


Public blQuery As Boolean
Public blDeposit As Boolean
Public blTransfer As Boolean
Public blWithdrawal As Boolean
Public blStop As Boolean

Public AccountNumber As String
Public FirstName As String
Public LastName As String
Public OfficeAddress As String
Public HomeAddress As String
Public DateOfBirth As Date
Public AccountType As String
Public Gender As String
Public EmailAddress As String
Public OfficePhone As String
Public HomePhone As String
Public Balance As Currency
Public AccessCode As String
Public DateOpened As Date
Public CustomerPicture As String
Public ATMCardNumber As String
Public ExpiryDate As Date

Public Sub Enable_Menu()
    Dim ctl As control
    
    For Each ctl In frmMain
        If TypeOf ctl Is Menu Then
            ctl.Enabled = True
        End If
    Next
    
End Sub

Public Sub Disable_Menu()
    Dim ctl As control
    
    For Each ctl In frmMain
        If TypeOf ctl Is Menu Then
            ctl.Enabled = False
        End If
    Next
    
End Sub


Public Sub fill_Combo(cbo As ComboBox)

    cbo.Clear
    recCustomers.Requery
    If Not recCustomers.EOF Then
        recCustomers.MoveFirst
        Do While Not recCustomers.EOF
            cbo.AddItem recCustomers!AccountNumber
'            cbo.ItemData(cbo.NewIndex) = recCustomers!AccountNumber
            recCustomers.MoveNext
        Loop
    End If

End Sub

Public Function Check_CheckNo(strCheck As String) As Boolean

    recTrans.Requery
    If Not recTrans.EOF Then
        recTrans.MoveFirst
        Do While Not recTrans.EOF
            If Trim(recTrans!checkNumber) = strCheck Then
                Check_CheckNo = True
                Exit Function
            End If
            recTrans.MoveNext
        Loop
    End If

    Check_CheckNo = False

End Function

Public Function Check_StopCheckNo(strCheck As String, strAccountNumber As String) As Boolean

    recStopPayment.Requery
    If Not recStopPayment.EOF Then
        recStopPayment.MoveFirst
        Do While Not recStopPayment.EOF
            If Trim(recStopPayment!checkNumber) = strCheck And Trim(recStopPayment!AccountNumber) = strAccountNumber Then
                Check_StopCheckNo = True
                Exit Function
            End If
            recStopPayment.MoveNext
        Loop
    End If

    Check_StopCheckNo = False

End Function

Public Sub DisplayCustomerDetails(strAccountNumber As String)

    recCustomers.Requery
    If Not recCustomers.EOF Then
        recCustomers.MoveFirst
        Do While Not recCustomers.EOF
            If recCustomers!AccountNumber = strAccountNumber Then
                
                AccountNumber = recCustomers!AccountNumber & ""
                FirstName = recCustomers!FirstName & ""
                LastName = recCustomers!LastName & ""
                OfficeAddress = recCustomers!OfficeAddress & ""
                HomeAddress = recCustomers!HomeAddress & ""
                DateOfBirth = recCustomers!DateOfBirth & ""
                AccountType = recCustomers!AccountType & ""
                Gender = recCustomers!Gender & ""
                EmailAddress = recCustomers!Email & ""
                OfficePhone = recCustomers!OfficePhone & ""
                HomePhone = recCustomers!HomePhone & ""
                Balance = recCustomers!Balance & ""
                AccessCode = recCustomers!AccessCode & ""
                DateOpened = IIf(IsNull(recCustomers!DateOpened), Date, recCustomers!DateOpened)
                CustomerPicture = recCustomers!CustomerPicture & ""
                ATMCardNumber = recCustomers!ATMCardNumber & ""
                ExpiryDate = recCustomers!ExpiryDate & ""
                
            End If
            recCustomers.MoveNext
        Loop
    End If

End Sub

Public Sub onlyNumbers(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub


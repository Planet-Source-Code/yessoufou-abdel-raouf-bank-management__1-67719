Attribute VB_Name = "mdlConnection"
Public recCheckBookDetails As New ADODB.Recordset
Public recCustomerQuery As New ADODB.Recordset
Public recCustomers As New ADODB.Recordset
Public recStopPayment As New ADODB.Recordset
Public recTrans As New ADODB.Recordset
Public recUsers As New ADODB.Recordset

Public con As New ADODB.Connection



Public Sub ConnectMe()
On Error Resume Next

    con.Open "provider = microsoft.jet.oledb.4.0;data source = " & App.Path & "\Database\OnlineBanking.mdb"
    
    recCheckBookDetails.Open "Select * from CheckBookDetails order by CheckBookNumber", con, adOpenDynamic, adLockOptimistic
    recCustomerQuery.Open "Select * from customerQuery order by QueryID", con, adOpenDynamic, adLockOptimistic
    recStopPayment.Open "Select * from StopPayment", con, adOpenDynamic, adLockOptimistic
    recCustomers.Open "Select * from customer order by AccountNumber", con, adOpenDynamic, adLockOptimistic
    recTrans.Open "select * from Trans order by TransactionID", con, adOpenDynamic, adLockOptimistic
    recUsers.Open "select * from users order by loginid", con, adOpenDynamic, adLockOptimistic

End Sub


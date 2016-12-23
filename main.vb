' Sample: Building a SalesOrder Using QBFC '
Dim SessionManager As QBSessionManager
Set SessionManager = New QBSessionManager

' We’re going to use US QuickBooks and the 6.0 spec '
Dim SalesOrderSet As IMsgSetRequest
Set SalesOrderSet = SessionManager.CreateMsgSetRequest("US", 6, 0)

' Need a separate Append* for each request in the MsgSetRequest object '
' First set properties in the main body of the sales order '
Dim salesOrder As ISalesOrderAdd
Set salesOrder = SalesOrderSet.AppendSalesOrderAddRq
salesOrder.CustomerRef.FullName.setValue "John Hamilton"
salesOrder.RefNumber.setValue "121345"

' Now add the line items. Every transaction with line items will look very similar to this '
Dim SOLineItemAdder As ISalesOrderLineAdd
Set SOLineItemAdder = salesOrder.ORSalesOrderLineAd
dList.Append.salesOrderLineAdd
SOLineItemAdder.ItemRef.FullName.setValue "fee"
SOLineItemAdder.Quantity.setValue 3
SOLineItemAdder.Other1.setValue "gold"

' Add another line item. Notice our re-use of the SOLineItemAdder variable '
Set SOLineItemAdder = salesOrder.ORSalesOrderLineAddList.Append.salesOrderLineAdd
SOLineItemAdder.ItemRef.FullName.setValue "fee"
SOLineItemAdder.Quantity.setValue 5
SOLineItemAdder.Other1.setValue "silver"

' OK, we’re done, send this to QB; for grins show results in a message box '
' We close everything down when done for illustration only: you would keep the session and '
' connection open if you were going to send more requests '
SessionManager.OpenConnection2 appID, appName, ctLocalQBD

' Let’s use whatever company file happens to be open now '
SessionManager.BeginSession "", omDontCare
Dim SOAddResp As IMsgSetResponse
Set SOAddResp = SessionManager.DoRequests(SalesOrderSet)
MsgBox SOAddResp.ToXMLString
SessionManager.EndSession
SessionManager.CloseConnection
Set SessionManager = Nothing


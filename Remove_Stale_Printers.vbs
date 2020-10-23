Set wshNetwork = CreateObject("WScript.Network")
on Error Resume Next

'Deletes all network printers
Set clPrinters = WshNetwork.EnumPrinterConnections
On Error Resume Next
For i = 0 to clPrinters.Count - 1 Step 2
wshNetwork.RemovePrinterConnection clPrinters.Item(i+1), true
Next

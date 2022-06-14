Set objWord = CreateObject("Word.Application")
objWord.Caption = "Test Caption"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "18"
objSelection.TypeText "Network Adapter Report"
objSelection.TypeParagraph()

objSelection.Font.Size = "14"
objSelection.TypeText "" & Date()
objSelection.TypeParagraph()
objSelection.TypeParagraph()

objSelection.Font.Size = "10"

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration")

For Each objItem in colItems

    objSelection.Font.Bold = True
    objSelection.TypeText "ARP Always Source Route: " 
    objSelection.Font.Bold = False
    objSelection.TypeText "" & objItem.ArpAlwaysSourceRoute
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "ARP Use EtherSNAP: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""  & objItem.ArpUseEtherSNAP
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Caption: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""  & objItem.Caption
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Database Path: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""   & objItem.DatabasePath
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Dead GW Detection Enabled: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""   & objItem.DeadGWDetectEnabled
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Default IP Gateway: " 
    objSelection.Font.Bold = False
    objSelection.TypeText "" & objItem.DefaultIPGateway
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Default TOS: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""  & objItem.DefaultTOS
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Default TTL: "
    objSelection.Font.Bold = False
    objSelection.TypeText ""  & objItem.DefaultTTL
    objSelection.TypeParagraph()

    objSelection.Font.Bold = True
    objSelection.TypeText "Description: "
    objSelection.Font.Bold = True
    objSelection.Font.Bold = False
    objSelection.TypeText ""  & objItem.Description
    objSelection.TypeParagraph()

    objSelection.TypeParagraph()

Next

objDoc.SaveAs("C:\Scripts\Word\testdoc.doc")
objWord.Quit
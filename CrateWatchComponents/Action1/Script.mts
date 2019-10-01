On Error Resume Next
'Declaration
Data = "C:\Users\rkrishnakumar.cg.tcs\Desktop\watch\Watch_Components_2.xlsx"
'Data = "C:\Raghu\Workspace\UFT\CreateWachComponents.xlsx"
sSOrgCIS = "CA11,CH11,CN12,DK11,GB11,IE11,US11"
sSOrgAL = "CA11,CH11,US11,GB11,DK11,CN12,IE11,FR11"
sDistChnl = "10,20,30"
Server = "Material_Management"
Set oSession = SAPGuiSession("guicomponenttype:=12")
Set oWindow = oSession.SAPGuiWindow("guicomponenttype:=21")
Set oWindow1 = oSession.SAPGuiWindow("guicomponenttype:=22")
'Import Data Sheet
DataTable.AddSheet "Environment"
DataTable.ImportSheet Data,"Environment","Environment"
DataTable.AddSheet "Components"
DataTable.ImportSheet Data,"Components","Components"
iDSRowCnt = DataTable.GetSheet("Components").GetRowCount
sAssortmentCIS = "CISCANADA,CISUS,DCCANADA,DCAGIAG,DCUSA,9001,9002"
sAssortmentAldo = "DCCANADA,DCAGIAG,DCUSA,ALDOCANADA,ALDOUK,ALDOUS,ALDOPANEU,ALDOIE,ALDOFRANCE,ALDO_9001"
iAssortmentAldo = Split(sAssortmentAldo,",")
iAssortmentCIS = Split(sAssortmentCIS,",")
iSOrgCIS = Split(sSOrgCIS,",")
iSOrgAL = Split(sSOrgAL,",")
iDistChnl = Split(sDistChnl,",")
'Login
sEnv = DataTable("Env","Environment")
sClnt = DataTable("Client","Environment")
sUid = DataTable("UserID","Environment")
sPwd = DataTable("Password","Environment")
sLang = DataTable("Language","Environment")
SAPGuiUtil.AutoLogon sEnv,sClnt,sUid,sPwd,sLang'"Aldo ECC Prod","100","pparthasarat","pune123P","EN" 
Create Article
For i = 1 to 4'iDSRowCnt
	DataTable.GetSheet("Components").SetCurrentRow(i)
	oWindow.Activate
	Call SetTcode("/nmm41")
	Call PressEnter()
	Call SelectGuiComboBoxByAttachedText("Material Type",DataTable("Material_Type","Components"))
	Call SetTextFieldByAttachedText("Mdse Catgry",DataTable("Merchandise_Category","Components"))
	Call SelectGuiComboBoxByAttachedText("Artl category",DataTable("Article_Category","Components"))
	Call SetTextFieldByAttachedText("Sales Org\.",DataTable("Sales_Org","Components"))
	Call SetTextFieldByAttachedText("Distr\. Channel",DataTable("Distro_Chnl","Components"))
	If oWindow.SAPGuiTable("guicomponenttype:=80","name:=SAPLMGMWTAB_CONT_0100").GetROProperty("rowcount") = 15 Then
		Call PressEnter()
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",1)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",2)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",3)
		Call PressEnter()
	ElseIf oWindow.SAPGuiTable("guicomponenttype:=80","name:=SAPLMGMWTAB_CONT_0100").GetROProperty("rowcount") = 18 Then
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",1)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",2)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",3)
		Call PressEnter()
'	ElseIf oWindow.SAPGuiTable("guicomponenttype:=80","name:=SAPLMGMWTAB_CONT_0100").GetROProperty("rowcount") = 15 Then
'		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",1)
'		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",2)
'		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",3)
'		Call PressEnter()
	ElseIf oWindow.SAPGuiTable("guicomponenttype:=80","name:=SAPLMGMWTAB_CONT_0100").GetROProperty("rowcount") = 19 Then
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",1)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",2)
		Call SelectRowSAPTable("SAPLMGMWTAB_CONT_0100",3)
		Call PressEnter()
	End If
	Call SelectTabByName("TABSPR1","Basic Data")
	Call SetTextFieldByAttachedText31("Single article",DataTable("Description","Components"))
	Call SetTextFieldByAttachedText31("Char\. value",DataTable("Char_Values","Components"))
	Call SetTextFieldByAttachedText("Brand",DataTable("Brand","Components"))
	Call SetTextFieldByAttachedText("Ctry of origin",DataTable("CountryOfOrigin","Components"))
	Call SetTextFieldByAttachedText("Comm\./imp\. code",DataTable("HTSCode","Components"))
	ArtNo = Mid(oWindow.GetROProperty("text"),17,3)
'	MsgBox ArtNo
	MatType = DataTable("Component_Type","Components")
	If MatType = "WATCH MOVEMENT" Then
		abb = "WM"
	ElseIf MatType = "WATCH BATTERY" Then
		abb = "WB"
	ElseIf MatType = "WATCH CASE" Then
		abb = "WC"
	ElseIf MatType = "WATCH STRAP" Then
		abb = "WS"
	End If
	ColorWay = ArtNo&"-"&abb
	Call SetTextFieldByAttachedText31("Old artl number",ColorWay)
	Call SelectTabByName("TABSPR1","Basic Data: Fashion")
	Call SetTextFieldByAttachedText("Operating Brand",DataTable("Operating_Brand","Components"))
	Call PressEnter()
	Call SelectTabByName("TABSPR1","Listing")
	Call SetTextFieldByAttachedText("Assortment grade",DataTable("Assortment_Grade","Components"))
	Call PressEnter()
	Call SelectTabByName("TABSPR1","Basic Data")
	Call SetTextFieldByAttachedTextAndName("Valid from","MARA-MSTDV",DataTable("Valid_From","Components"))
	Call SetTextFieldByAttachedText("X-DChain status",DataTable("Status","Components"))
	Call PressEnter()
	Call ClickButton("Save")
	DataTable("Old_Article_No","Components") = ColorWay
	DataTable("Article_No","Components") = GetGuiStatusBarValue("item1")
	DataTable.ExportSheet Data,"Components","Components"
Next
'Enrichment
oWindow.Activate
Call SetTcode("/nZMART_ENRICHMENT")
Call PressEnter()
iAldo = 0
iCis = 0
For i = 1 to 4'iDSRowCnt'DataTable.GetSheet("Components").GetRowCount
	If DataTable("BrandName","Components") = "Aldo" Then
		iAldo = iAldo+1
	ElseIf DataTable("BrandName","Components") = "CIS" Then
		iCis = iCis+1
	End If
Next
'Aldo
If iAldo<>0 Then
	Call ClickButtonByIndex("Multiple selection",0)
	For j = Lbound(iAssortmentAldo) To Ubound(iAssortmentAldo)
		Call SetTableData("SAPLALDBSINGLE",j+1,"#2",iAssortmentAldo(j))		
	Next
	Call ClickButton("Copy")
	Call SetTextFieldByAttachedText31("Max\. Number of Processes",20)
	Call SetTextFieldByAttachedText31("Artl per Process",100)
	Call SetTextFieldByAttachedText("Logon/Server Group","Material_Management")
	Call SetTextFieldByAttachedText("Operating brand","001")
	Call SetTextFieldByAttachedText("Target Article X-Dist Status","20")
	Call SetTextFieldByAttachedText("Source Article X-Dist Status","20")
	Call ClickButtonByIndex("Multiple selection",4)
	For k = Lbound(iDistChnl) To Ubound(iDistChnl)
		Call SetTableData("SAPLALDBSINGLE",k+1,"#2",iDistChnl(k))		
	Next
	Call ClickButton("Copy")
	Call ClickButtonByIndex("Multiple selection",3)
	For l = Lbound(iSOrgAL) To Ubound(iSOrgAL)
		Call SetTableData("SAPLALDBSINGLE",l+1,"#2",iSOrgAL(l))		
	Next
	Call ClickButton("Copy")
	Call ClickButtonByIndex("Multiple selection",2)
	For m = 1 to 4'iDSRowCnt'DataTable.GetSheet("Components").GetRowCount
	DataTable.GetSheet("Components").SetCurrentRow(m)
		If DataTable("BrandName","Components") = "Aldo" Then
			Call SetTableData("SAPLALDBSINGLE",m,"#2",DataTable("Article_No","Components"))			
		End If
	Next
	Call ClickButton("Copy")
	Call PressF8()
	Call ClickButton("Continue")
End If
'CIS
If iCis<>0 Then
	Call ClickButtonByIndex("Multiple selection",0)
	For j = Lbound(iAssortmentCIS) To Ubound(iAssortmentCIS)
		Call SetTableData("SAPLALDBSINGLE",j,"#2",iAssortmentCIS(j))		
	Next
	Call ClickButton("Copy")
	Call SetTextFieldByAttachedText31("Max\. Number of Processes",20)
	Call SetTextFieldByAttachedText31("Artl per Process",100)
	Call SetTextFieldByAttachedText("Logon/Server Group","Material_Management")
	Call SetTextFieldByAttachedText("Operating brand","001")
	Call SetTextFieldByAttachedText("Target Article X-Dist Status","20")
	Call SetTextFieldByAttachedText("Source Article X-Dist Status","20")
	Call ClickButtonByIndex("Multiple selection",4)
	For k = Lbound(iDistChnl) To Ubound(iDistChnl)
		Call SetTableData("SAPLALDBSINGLE",k,"#2",iDistChnl(k))		
	Next
	Call ClickButton("Copy")
	Call ClickButtonByIndex("Multiple selection",3)
	For l = Lbound(iSOrgCIS) To Ubound(iSOrgCIS)
		Call SetTableData("SAPLALDBSINGLE",l,"#2",iSOrgCIS(l))		
	Next
	Call ClickButton("Copy")
	Call ClickButtonByIndex("Multiple selection",2)
	For m = 1 to 4'iDSRowCnt'DataTable.GetSheet("Components").GetRowCount
	DataTable.GetSheet("Components").SetCurrentRow(m)
		If DataTable("BrandName","Components") = "CIS" Then
			Call SetTableData("SAPLALDBSINGLE",m,"#2",DataTable("Article_No","Components"))			
		End If
	Next
	Call ClickButton("Copy")
	Call PressF8()
	Call ClickButton("Continue")
End If
'PIR & PB00
For i = 1 to 4'iDSRowCnt'DataTable.GetSheet("Components").GetRowCount
	DataTable.GetSheet("Components").SetCurrentRow(i)
	oWindow.Activate
	Call SetTcode("/nme11")
	Call PressEnter()
	Call SetTextFieldByAttachedText("Vendor",DataTable("Vendor_ID","Components"))
	Call SetTextFieldByAttachedText("Article",DataTable("Article_No","Components"))
	Call SetTextFieldByAttachedText("Purchasing Org\.","1000")
	Call SetTextFieldByAttachedText("Info record","")
	Call PressEnter()
	Call SetTextFieldByAttachedText("Ctry of Origin",DataTable("CountryOfOrigin","Components"))
	oWindow.SAPGuiEdit("guicomponenttype:=32","attachedtext:=Region","index:=0").Highlight
	oWindow.SAPGuiEdit("guicomponenttype:=32","attachedtext:=Region","index:=0").SetFocus
	Set wsh = CreateObject("WScript.Shell")
	wsh.SendKeys "{DEL}"
'	oWindow.SendKey
	Call SetTextFieldByAttachedText("Region","")
	Call SetGuiCheckBoxByAttachedTextAndIndex("Regular Vendor",0,"ON")
	Call SetTextFieldByAttachedText("Var\. Order Unit",1)
	Call ClickButton("Purchasing Organization Data 1")
	If oWindow.SAPGuiStatusBar("guicomponenttype:=103","type:=GuiStatusbar").GetroProperty("messagetype") = "W" Then
		Call PressEnter()
	End If
	Call SetTextFieldByAttachedText31("Pl\. Deliv\. Time","60")
	Call SetTextFieldByAttachedText("Purch\. Group",DataTable("Purch_Grp","Components"))
	Call SetTextFieldByAttachedText31("Standard Qty","1.00")
	Call SetTextFieldByAttachedText31("Net Price",DataTable("First_Cost","Components"))
	Call SetTextFieldByAttachedTextAndName("Net Price","EINE-WAERS","USD")
	Call ClickButton("Conditions")
	Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY",1,"Valid From",DataTable("FC_ValidFrom","Components"))
	Call PressEnter()
	Call ClickButton("Save")
Next
'ZPELC_TAB
For i = 1 to 4'iDSRowCnt'DataTable.GetSheet("Components").GetRowCount
	DataTable.GetSheet("Components").SetCurrentRow(i)
	oWindow.Activate
	Call SetTcode("/nsm30")
	Call PressEnter()
	Call SetTextFieldByAttachedText("Table/View","ZPELC_TAB")
	Call ClickButton("Maintain")
	Call ClickButton("New Entries")
	Call SetTextFieldByAttachedText31("CNTRY O",DataTable("CountryOfOrigin","Components"))
	Call SetTextFieldByAttachedText("Site","8937")
	Call SetTextFieldByAttachedText31("Season",DataTable("Season","Components"))
	Call SetTextFieldByAttachedText("Article",DataTable("Article_No","Components"))
	Call SelectGuiComboBoxByAttachedText("Active","Y")
	Call SetTextFieldByAttachedText31("Created By",DataTable("Created_By","Components"))
	Call SetTextFieldByAttachedText("Created On",DataTable("Valid_From","Components"))
	Call SetTextFieldByAttachedText31("Shipping Instr\.","01")
	Call SetTextFieldByAttachedText31("YY",DataTable("Season_Year","Components"))
	Call SetTextFieldByAttachedText31("FCST",DataTable("First_Cost","Components"))
	Call SetTextFieldByAttachedText31("FC CURR","USD")
	Call SetTextFieldByAttachedText31("AGI MRGN",DataTable("AGI_MRGN","Components"))
	Call SetTextFieldByAttachedText31("M\.P\. FEE",DataTable("MPF","Components"))
	Call SetTextFieldByAttachedText31("H\.M\. FEES",DataTable("HMF","Components"))
	Call SetTextFieldByAttachedText31("Cust Dty",DataTable("Duty","Components"))
	Call SetTextFieldByAttachedText31("Est\. LC",DataTable("ELC","Components"))
	Call SetTextFieldByAttachedText31("DES CN","US")
	Call SetTextFieldByAttachedText31("CLR WAY",DataTable("Old_Article_No","Components"))
	Call PressEnter()
	Call ClickButton("Save")
	Call ClickButton("Back")
	Call ClickButton("Save")
Next
Call SetTcode("/n")
Call PressEnter()
DataTable.ExportSheet Data,"Components","Components"

''Code to uncomment to run individual components
''Declaration`
'Data = "C:\Raghu\Workspace\UFT\CreateWachComponents.xlsx"
'Set oSession = SAPGuiSession("guicomponenttype:=12")
'Set oWindow = oSession.SAPGuiWindow("guicomponenttype:=21")
'Set oWindow1 = oSession.SAPGuiWindow("guicomponenttype:=22")
''Import Data Sheet
'DataTable.AddSheet "Environment"
'DataTable.ImportSheet Data,"Environment","Environment"
'DataTable.AddSheet "Components"
'DataTable.ImportSheet Data,"Components","Components"
'iDSRowCnt = DataTable.GetSheet("Components").GetRowCount

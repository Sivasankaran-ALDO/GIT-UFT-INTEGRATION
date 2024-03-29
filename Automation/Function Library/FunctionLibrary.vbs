Function SetTextFieldByAttachedText(sAttachedtext,sValue)
If sValue<>"" Then
    Set obj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedtext,"index:=0")
    Set obj1=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=32","attachedtext:="&sAttachedtext,"index:=0")
    If obj.Exist Then
    	If obj.GetROProperty("enabled")=True Then
    		obj.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If   
    ElseIf obj1.Exist Then
    	If obj1.GetROProperty("enabled")=True Then
    		obj1.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If
    Else
        reporter.ReportEvent micFail,"Set Field "&sAttachedtext, "Unable to find the field "&sAttachedtext&" on the screen"
    End If
Else
	Reporter.ReportEvent micWarningSet,"Set Field "&sAttachedtext, "Value to be set not provided."
End If 
End Function

Function PressEnter()    
    Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
    If obj.Exist Then
        obj.SendKey(ENTER)
    ElseIf obj1.Exist Then
        obj1.SendKey(ENTER)
    Else
        reporter.ReportEvent micFail,"Press Eenter", "Unable to recognize SAP Window"
    End If
End Function

Function SetTcode(sTcode)
    If instr(1,sTCode,"/n",1) Then
        tcode = sTcode
    ElseIf Left(sTcode,2) = "/n" Then
        tcode = sTcode
    Else
        tcode = "/n"&sTcode
    End If
    Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiOKCode("type:=GuiOkCodeField","name:=okcd","guicomponenttype:=35")
    If obj.exist Then
        obj.Set tcode
    Else
        reporter.ReportEvent micFail,"Set T-Code", "Set T-Code Failed"
    End If
End Function


Function SelectTreeNodeByName(sName,sNode)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","name:="&sName,"treetype:=SapColumnTree")
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTree("guicomponenttype:=200","name:="&sName,"treetype:=SapColumnTree")
    If obj.Exist Then
        obj.SelectNode(sNode)
    ElseIf obj1.Exist Then
        obj1.SelectNode(sNode)
    Else
        reporter.ReportEvent micFail,"Press Eenter", "Unable to find the SAP Tree"
    End If
End Function


Function SetTextAreaByName(sName,sValue)
If sValue<>"" Then
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTextArea("guicomponenttype:=203","name:="&sName)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTextArea("guicomponenttype:=203","name:="&sName)
    If obj.Exist Then
    	obj.Set sValue
    ElseIf obj1.Exist Then
    	obj1.Set sValue
    Else
    	reporter.ReportEvent micFail,"Set Text Area "& sName, "Unable to Text Area with Name "&sName
    End If
Else
	Reporter.ReportEvent micWarningSet,"Set Text Area "& sName, "Value to be set not provided."
End If 
End Function

Function ClickButton(sTooltip)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("guicomponenttype:=40","tooltip:="& sTooltip &".*","index:=0")
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiButton("guicomponenttype:=40","tooltip:="& sTooltip &".*","index:=0")
    If obj.Exist Then
	    If obj.GetROProperty("enabled")=True Then
	    	obj.Click
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If
    ElseIf obj1.Exist Then
	    If obj1.GetROProperty("enabled")=True Then
	    	obj1.Click
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If    	
    Else
    	reporter.ReportEvent micFail,"Click Button", "Unable to find Button "&sTooltip
    End If
End Function


Function SelectTabByName(sName,sTabName)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTabStrip("guicomponenttype:=90","name:="&sName,"type:=GuiTabStrip")
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTabStrip("guicomponenttype:=90","name:="&sName,"type:=GuiTabStrip")
    If obj.Exist Then
        obj.Select(sTabName)
    ElseIf obj1.Exist Then
        obj1.Select(sTabName)
    Else
        reporter.ReportEvent micFail,"Press Eenter", "Unable to find the SAP Tree"
    End If
End Function

Function SetTableData(sName,sRowNo,sColumnName,sValue)
If sValue<>"" Then
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sName,"type:=GuiTableControl")
	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable("guicomponenttype:=80","name:="&sName,"type:=GuiTableControl")
	If obj.Exist Then
		If obj.IsCellEditable(sRowNo,sColumnName) Then
			obj.SetCellData sRowNo,sColumnName,sValue
		Else
			Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& sRowNo, "Set "& sColumnName &" in Row "& sRowNo &" Failed as Cell is not enabled"
		End If
	ElseIf obj1.Exist Then
		If obj1.IsCellEditable(sRowNo,sColumnName) Then
			obj1.SetCellData sRowNo,sColumnName,sValue
		Else
			Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& sRowNo, "Set "& sColumnName &" in Row "& sRowNo &" Failed as Cell is not enabled"
		End If
	Else
		Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& sRowNo, "Unable to find Table "& sName
	End If
Else
	Reporter.ReportEvent micWarningSet,"Set "& sColumnName &" in Row "& sRowNo, "Value to be set not provided."
End If
End Function


Function SetTextFieldByAttachedTextAndName(sAttachedtext,sName,sValue)
If sValue<>"" Then
    Set obj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=32","name:="&sName,"attachedtext:="&sAttachedtext,"index:=0")
    Set obj1=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=32","name:="&sName,"attachedtext:="&sAttachedtext,"index:=0")
    If obj.Exist Then
    	If obj.GetROProperty("enabled")=True Then
    		obj.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If   
    ElseIf obj1.Exist Then
    	If obj1.GetROProperty("enabled")=True Then
    		obj1.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If
    Else
        reporter.ReportEvent micFail,"Set Field "&sAttachedtext, "Unable to find the field "&sAttachedtext&" on the screen"
    End If
Else
	Reporter.ReportEvent micWarningSet,"Set Field "&sAttachedtext, "Value to be set not provided."
End If 
End Function


Function SetTableDataByRefFieldValue(sName,sRefColumn,sRefValue,sColumnName,sValue)
If sValue<>"" Then
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sName,"type:=GuiTableControl")
	If obj.Exist Then
		iRowNo = obj.FindRowByCellContent(sRefColumn,sRefValue)
		If iRowNo<>"" Then
			If obj.IsCellEditable(iRowNo,sColumnName) = True Then
				obj.SetCellData iRowNo,sColumnName,sValue
			Else
				Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& iRowNo, "Set "& sColumnName &" in Row "& sRowNo &" Failed as Cell is not enabled"
			End If
		Else
			Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& iRowNo, "Unable to find Roe for the Size mentioned."
		End If
	Else
		Reporter.ReportEvent micFail,"Set "& sColumnName &" in Row "& iRowNo, "Unable tom find the Table to set value in."
	End If
Else
	Reporter.ReportEvent micWarningSet,"Set "& sColumnName &" in Row "& iRowNo, "Value to be set not provided."
End If	
End Function

Function GetGuiStatusBarValue(sItemno)
	Set obj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiStatusBar("guicomponenttype:=103","type:=GuiStatusbar")
	Set obj1=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiStatusBar("guicomponenttype:=103","type:=GuiStatusbar")
	If obj.Exist Then
    	If obj.GetROProperty(sItemno)<>"" Then
    		sValue = obj.GetROProperty(sItemno)
    	End If   
    ElseIf obj1.Exist Then
    	If obj1.GetROProperty(sItemno)<>"" Then
    		sValue = obj1.GetROProperty(sItemno)
    	End If
    Else
        reporter.ReportEvent micFail,"Get Statusbar Value", "Unable to find Statusbar"
    End If
    GetGuiStatusBarValue = sValue
    sValue= ""
End Function


Function SetTextFieldByAttachedText31(sAttachedtext,sValue)
If sValue<>"" Then
    Set obj=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedtext,"index:=0")
    Set obj1=SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiEdit("guicomponenttype:=31","attachedtext:="&sAttachedtext,"index:=0")
    If obj.Exist Then
    	If obj.GetROProperty("enabled")=True Then
    		obj.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If   
    ElseIf obj1.Exist Then
    	If obj1.GetROProperty("enabled")=True Then
    		obj1.Set sValue
    	Else
    		reporter.ReportEvent micFail,"Set Value to field "&sAttachedtext, "Set Value to field "& sAttachedtext &" Falied as Text field is Disbaled"
    	End If
    Else
        reporter.ReportEvent micFail,"Set Field "&sAttachedtext, "Unable to find the field "&sAttachedtext&" on the screen"
    End If
Else
	Reporter.ReportEvent micWarningSet,"Set Field "&sAttachedtext, "Value to be set not provided."
End If 
End Function


Function GetRowCountSAPGridMinus1(sGridTitle,sGridName)
	If sGridTitle="" Then
		sGridTitle=".*"
	End If
		If sGridName="" Then
		sGridName=".*"
	End If
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("title:="&sGridTitle,"name:="&sGridName)
	If obj.Exist Then
		GetRowCountSAPGridMinus1 = obj.GetROProperty("rowcount")-1
	Else
		reporter.ReportEvent micFail,"Get Gui Grid "& sGridTitle &" Row Count","Unable to find SAP Gui Grid "& sGridTitle
	End If
End Function

Function GetCellDataSAPGrid(sGridTitle,sGridName,sColumnName,iRowNumber)
	If sGridTitle="" Then
		sGridTitle=".*"
	End If
		If sGridName="" Then
		sGridName=".*"
	End If
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiGrid("title:="&sGridTitle,"name:="&sGridName)
	If obj.Exist Then
		GetCellDataSAPGrid = obj.GetCellData(iRowNumber,sColumnName)
	Else
		reporter.ReportEvent micFail,"Get Cell Data Gui Grid "& sGridTitle &" For row "& iRowNumber &" And Column "& sColumnName,"Unable to find SAP Gui Grid "& sGridTitle
	End If
End Function


Function PressF8()    
    Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
    If obj.Exist Then
        obj.SendKey(F8)
    ElseIf obj1.Exist Then
        obj1.SendKey(F8)
    Else
        reporter.ReportEvent micFail,"Press Eenter", "Unable to recognize SAP Window"
    End If
End Function


Function ToolBarClickContextButtonAndSelectItem(sType,sName,sButton,sPath)
	If sPath<>"" Then
		Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:=202","type:="&sType,"name:="&sName,"index:=0")
    	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiToolbar("guicomponenttype:=202","type:="&sType,"name:="&sName,"index:=0")
    	If obj.Exist Then
    		obj.PressContextButton sButton
    		obj.SelectMenuItem sPath
    	Elseif 	obj1.Exist Then
    		obj1.PressContextButton sButton
    		obj1.SelectMenuItem sPath
    	Else
    		reporter.ReportEvent micFail,"Select Menu Path "& sPath &" in Toolbar "&sName, "Unable to find the Tool bar"
    	End If
    End If
End Function


Function ExpandAndSelectTreeNode(sType,sName,sTreeType,sSelectionMode,sNode)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    If obj.Exist Then
    	obj.Expand sNode
    	obj.ActivateNode sNode &";" &sNode
    Elseif 	obj1.Exist Then
    	obj1.Expand sNode
    	obj1.ActivateNode sNode &";" &sNode
    Else
    		reporter.ReportEvent micFail,"Select Menu Path "& sPath &" in Toolbar "&sName, "Unable to find the Tool bar"
    	End If
End Function


Function ToolBarClicktButton(sType,sName,sButton)
	If sPath<>"" Then
		Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiToolbar("guicomponenttype:=202","type:="&sType,"name:="&sName,"index:=0")
    	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiToolbar("guicomponenttype:=202","type:="&sType,"name:="&sName,"index:=0")
    	If obj.Exist Then
    		obj.PressButton sButton
    	Elseif 	obj1.Exist Then
    		obj1.PressButton sButton
    	Else
    		reporter.ReportEvent micFail,"Select Menu Path "& sPath &" in Toolbar "&sName, "Unable to find the Tool bar"
    	End If
    End If
End Function



Function SelectItemTree(sType,sName,sTreeType,sSelectionMode,sNode)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    If obj.Exist Then
    	obj.SelectItem sNode
    Elseif 	obj1.Exist Then
    	obj1.SelectItem sNode
    Else
    		reporter.ReportEvent micFail,"Select Menu Path "& sPath &" in Toolbar "&sName, "Unable to find the Tool bar"
    	End If
End Function

Function ExtendNodeTree(sType,sName,sTreeType,sSelectionMode,sNode)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTree("guicomponenttype:=200","type:="&sType,"name:="&sName,"treetype:="&sTreeType,"selectionmode:="&sSelectionMode)
    If obj.Exist Then
    	obj.ExtendNode sNode
    Elseif 	obj1.Exist Then
    	obj1.ExtendNode sNode
    Else
    		reporter.ReportEvent micFail,"Select Menu Path "& sPath &" in Toolbar "&sName, "Unable to find the Tool bar"
    	End If
End Function



Function GetRowNoByRefValue(sName,sRefColumn,sRefValue)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sName,"type:=GuiTableControl")
	If obj.Exist Then
		iRowNo = obj.FindRowByCellContent(sRefColumn,sRefValue)
		If iRowNo<>"" Then
			GetRowNoByRefValue = iRowNo
		Else
			Reporter.ReportEvent micFail,"Set "& sRefColumn &" in Row "& iRowNo, "Unable to find Row for the Value mentioned."
		End If
	Else
		Reporter.ReportEvent micFail,"Set "& sRefColumn &" in Row "& iRowNo, "Unable tom find the Table to set value in."
	End If
End Function


Function GetCellDataSAPTable(sTableName,sColumnName,iRowNumber)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	If obj.Exist Then
		GetCellDataSAPTable = obj.GetCellData(iRowNumber,sColumnName)
	ElseIf obj1.Exist Then
		GetCellDataSAPTable = obj1.GetCellData(iRowNumber,sColumnName)
	Else
		reporter.ReportEvent micFail,"Get Cell Data Gui Grid "& sGridTitle &" For row "& iRowNumber &" And Column "& sColumnName,"Unable to find SAP Gui Grid "& sGridTitle
	End If
End Function



Function SetCheckboxSAPTable(sTableName,sColumnName,iRowNumber,sValue)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	If obj.Exist Then
		obj.SetCellData iRowNumber,sColumnName,sValue
	ElseIf obj1.Exist Then
		obj1.SetCellData iRowNumber,sColumnName,sValue
	Else
		reporter.ReportEvent micFail,"Get Cell Data Gui Grid "& sGridTitle &" For row "& iRowNumber &" And Column "& sColumnName,"Unable to find SAP Gui Grid "& sGridTitle
	End If
End Function



Function SelectRowSAPTable(sTableName,iRowNumber)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiTable("guicomponenttype:=80","name:="&sTableName)
	If obj.Exist Then
		obj.SelectRow iRowNumber
	ElseIf obj1.Exist Then
		obj1.SelectRow iRowNumber
	Else
		reporter.ReportEvent micFail,"Get Cell Data Gui Grid "& sGridTitle &" For row "& iRowNumber &" And Column "& sColumnName,"Unable to find SAP Gui Grid "& sGridTitle
	End If
End Function



Function SelectGuiComboBoxByAttachedText(sAttachedText,sKey)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiComboBox("guicomponenttype:=34","attachedtext:="&sAttachedText)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiComboBox("guicomponenttype:=34","attachedtext:="&sAttachedText)
    If obj.Exist Then
        obj.SelectKey sKey 
    ElseIf obj1.Exist Then
        obj1.SelectKey sKey 
    Else
        reporter.ReportEvent micFail,"Press Eenter", "Unable to find the SAP Tree"
    End If
End Function



Function ClickButtonByIndex(sTooltip,iIndex)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiButton("guicomponenttype:=40","tooltip:="& sTooltip &".*","index:="& iIndex)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiButton("guicomponenttype:=40","tooltip:="& sTooltip &".*","index:="& iIndex)
    If obj.Exist Then
	    If obj.GetROProperty("enabled")=True Then
	    	obj.Highlight
	    	obj.Click
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If
    ElseIf obj1.Exist Then
	    If obj1.GetROProperty("enabled")=True Then
	    	obj1.Highlight
	    	obj1.Click
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If    	
    Else
    	reporter.ReportEvent micFail,"Click Button", "Unable to find Button "&sTooltip
    End If
End Function



Function SetGuiCheckBoxByAttachedTextAndIndex(sAttachedText,iIndex,sValue)
	Set obj = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiCheckBox("guicomponenttype:=42","attachedtext:="& sAttachedText &".*","index:="& iIndex)
    Set obj1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22").SAPGuiCheckBox("guicomponenttype:=42","attachedtext:="& sAttachedText &".*","index:="& iIndex)
    If obj.Exist Then
	    If obj.GetROProperty("enabled")=True Then
	    	obj.Set sValue
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If
    ElseIf obj1.Exist Then
	    If obj1.GetROProperty("enabled")=True Then
	    	obj1.Set sValue
	    Else
	    	reporter.ReportEvent micFail, "Click Button", "Button "& sTooltip &" not enabled"
	    End If    	
    Else
    	reporter.ReportEvent micFail,"Click Button", "Unable to find Button "&sTooltip
    End If
End Function


Function LaunchBrowserURL(iBroswer, sURL)
On Error Resume Next
	If iBroswer <>"" And sURL<>"" Then
		SystemUtil.Run iBroswer, sURL,,,3
		Set oBrowser = Browser("CreationTime:=0")
		If oBrowser.Exists and oBrowser.GetROProperty("openurl")=sURL Then
			reporter.ReportEvent micDone,"Open Browser & URL", "Open Browser & URL Successful"
		Else
			reporter.ReportEvent micFail,"Open Browser & URL", "Open Browser & URL Failed"
		End If
		reporter.ReportEvent micFail,"Open Browser & URL", "Parameters Not Provided"
	End If 
On Error Goto 0
End Function

Function Web_ClickRadioButton(sName,iIndex)
	On Error Resume Next
		If sName <>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebRadioGroup("name:="&sName, "index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Click
				reporter.ReportEvent micDone,"Select Radio Button", "Select Radio Button"&sName&" Successful"
			Else
				reporter.ReportEvent micDone,"Select Radio Button", "Select Radio Button"&sName&" Failed"
			End If
			reporter.ReportEvent micFail,"Select Radio Button", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_SelectWebList(sName,iIndex,sValue)
	On Error Resume Next
		If sName <>"" and sValue<>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebList("name:="&sName, "index:="&iIndex)
			If obj.Exist and instr(obj.GetROProperty("innertext"),sValue) Then
				obj.Select sValue
				reporter.ReportEvent micDone,"Select Value from Dropdown", "Select Value from Dropdown "&sName&" Successful"
			Else
				reporter.ReportEvent micDone,"Select Value from Dropdown", "Select Value from Dropdown "&sName&" Failed"
			End If
			reporter.ReportEvent micFail,"Select Value from Dropdown", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_ClickButton(sName,iIndex)
	On Error Resume Next
		If sName <>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebButton("name:="&sName, "index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Click
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Successful"
			Else
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Failed"
			End If
			reporter.ReportEvent micFail,"Click Button", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function


Function Web_SetEditBox(sName,iIndex,sValue)
	On Error Resume Next
		If sName <>"" Or sValue<>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebEdit("name:="&sName, "index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Set sValue
				If instr(obj.GetROProperty("value"),sValue) Then
					reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" With Value "&sValue&" Successfully."
				Else 
					reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" With Value "&sValue&" Failed."
				End If
				reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" Does not exist."
			End If
			reporter.ReportEvent micFail,"Set Web Edit", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_ClickEditBox(sName,iIndex)
	On Error Resume Next
		If sName <>"" Or sValue<>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebEdit("name:="&sName, "index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Click
				If instr(obj.GetROProperty("value"),sValue) Then
					reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" With Value "&sValue&" Successfully."
				Else 
					reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" With Value "&sValue&" Failed."
				End If
				reporter.ReportEvent micDone,"Set Web Edit", "Set Web Edit "&sName&" Does not exist."
			End If
			reporter.ReportEvent micFail,"Set Web Edit", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_ClickWebElement(sName,iIndex)
	On Error Resume Next
		If sName <>"" Then
			Set oPage = Browser("CreationTime:=0").Page("CreationTime:=0")
			Set obj = oPage.WebElement("innertext:="&sName, "index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Click
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Successful"
			Else
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Failed"
			End If
			reporter.ReportEvent micFail,"Click Button", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_SetDialogEditBoxByNatvClsIndx(sDiaTitle,sEdtNatvCls,iIndex,sValue)
	On Error Resume Next
		If sDiaTitle <>"" Then
			Set oBrowser = Browser("CreationTime:=0")
			Set obj = oBrowser.Dialog("regexpwndtitle:="&sDiaTitle).WinEdit("micClass:=WinEdit","nativeclass:="&sEdtNatvCls,"index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Set sValue
				If obj.GetROProperty("text") = sValue Then
					reporter.ReportEvent micDone,"Set Web Dialod", "Set Web Dialod"&sEdtTitle&" Successful"
				Else
					reporter.ReportEvent micDone,"Set Web Dialod", "Click Button"&sName&" Failed"
				End If
			Else
				reporter.ReportEvent micDone,"Set Web Dialod", "Object Doesn't exist"
			End If
		Else
			reporter.ReportEvent micFail,"Set Web Dialod", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function Web_ClickDialogButton(sDiaTitle,sNtvCls,sRegexTitl,iIndex)
	On Error Resume Next
		If sDiaTitle <> "" Or sRegexTitl <> "" Or iIndex <> "" Then
			Set oBrowser = Browser("CreationTime:=0")
			Set obj = oBrowser.Dialog("regexpwndtitle:="&sDiaTitle).WinButton("micClass:=WinButton","nativeclass:="&sNtvCls,"regexpwndtitle:="&sRegexTitl,"index:="&iIndex)
			If obj.Exist Then
				obj.Highlight
				obj.Click
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Successful"
			Else
				reporter.ReportEvent micDone,"Click Button", "Click Button"&sName&" Failed"
			End If
			reporter.ReportEvent micFail,"Click Button", "Parameters Not Provided"
		End If 
	On Error Goto 0
End Function

Function SendKeys(sValue)
	If sValue<>"" Then
		Dim mySendKeys
		set mySendKeys = CreateObject("WScript.shell")
		mySendKeys.SendKeys sValue
	End If
End Function

Function MercuryDRSendKeys(sValue)
	If sValue<>"" Then
		Dim myDeviceReplay
		Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
		myDeviceReplay.PressKey sValue
	End If
End Function

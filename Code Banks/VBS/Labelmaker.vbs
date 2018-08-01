	returns = Print_HardwareTags("YMS-","Cart-N",30)

Function Print_HardwareTags(txtPrefix,txtItem,intNnumber)' Vertion 1.0 Aplha
	Dim objWord,objDoc,objSelection ' Word objects"
	Dim txtCart_labble
	Dim intCurrent_Hardware_Item_number 
	
	intCurrent_Hardware_Item_number = 20
		
	Set objWord 		= CreateObject("Word.Application")
	objWord.Caption 	= "Information Technology Cart Tag"
	objWord.Visible 	= True '<<<<<<<<< Debug point <<<<<<<<< 
	Set objDoc			= objWord.Documents.Add()
	Set objSelection 	= objWord.Selection
	
	
	Do while intCurrent_Hardware_Item_number < intNnumber
	
	If intCurrent_Hardware_Item_number < 9 then
			txtCart_labble = txtPrefix & txtItem & "0" & intCurrent_Hardware_Item_number + 1
	else
			txtCart_labble = txtPrefix & txtItem & intCurrent_Hardware_Item_number + 1
	end if
	
		objSelection.ParagraphFormat.Alignment = 1
		objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
		objSelection.Font.Bold = false
		objSelection.Font.Size = "14"
		objSelection.TypeText "*" & txtCart_labble &"*" 
		objSelection.TypeParagraph()
		objSelection.TypeParagraph()	
		intCurrent_Hardware_Item_number = intCurrent_Hardware_Item_number + 1
		
	loop 
	
	'objDoc.PrintOut()
	'objDoc.SaveAs(txtIncident &".doc")
	'objWord.Quit
	
End Function
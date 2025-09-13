
;;[[                       CLIP CODES                          ]]


clipChange(){
  global
  sleep 75
  if SimpleClip
    return
  if Instr(Clipboard, "[P]",true,1,1){
    ProductTab.AddProductFromClipboard()
		clip.codesRegex()
		; clipBar.Flash()
		If winactive("Edit Formulation")
			ProductTab.AddNewFormulation()
		if winactive("Edit Product")
			ProductTab.AddNewProduct()
		return
  }
  else if InStr(Clipboard, "<<QuIT>>",true, 1,1){
    Clipboard:=
    exitsub()
    Return
  }
  else if InStr(Clipboard, ">>|", true,1,1) {
    if (Iteration >=25) || (Iteration < 0) || !(Iteration)
      iteration:=1
			; Clipbar.FlashIteration()
   ;  ClippedData:=Clipboard
	; 	UsedLimits:=

	; 	AllowPrefixes:=
   ;  FileDelete, ClippedExcelData.txt
   ;  sleep 400
   ;  FileAppend, %ClippedData%, ClippedExcelData.txt
    LMS.AddDataFromClipboard()
    return
  }
	else if InStr(Clipboard, "<<CopyLabelCopy>>",true, 1,1){
		simpleclip:=1
		Gui ClipBar:Default
IF (SubStr(Clipboard, 19,4) != "Temp"){
		Product:=TRIM(SubStr(Clipboard, 19,4))
		sleep 200
		GuiControl,ClipBar:Text, Product, %Product%
}
		; clip.codesRegex()
		; Product:=RegexMatch(Clipboard, RegexProduct,r) ? rProduct : Product
		; clipBar.Flash()
		; ControlsetText, Edit1,%Product%,ClipBar  ;clip.codesRegex()
		; SLEEP 100
		Clipboard:=
		copyLabelCopyDoc()
		; GoSub ShowFinalLabelCopy
		; sleep 100
		; Clipwait,5,0

		; tt(clipboard)
		; Clipboard:=
simpleclip:=
		; FileDelete, %A_ScriptDir%  \LabelCopyText.txt
		; Try FileDelete, %A_ScriptDir%  \LabelCopies\%Product%.txt
		; FileAppend,  %labelcopytext%, %A_ScriptDir%  \LabelCopies\%Product%.txt
		return
  }
  else if InStr(Clipboard, "<<CoMPILE>>",true, 1,1){
		Clipboard:=
    sleep 800
    Run, "\\10.1.2.118\users\vitaquest\mmignin\VQ_Helper\RawFiles\COMPILE.exe"
    exitapp
    Return
  }

;   else if (Winactive("NuGenesis LMS") && InStr(Clipboard, "5-HTP=0`r`n5-MTHF=1`r`n",true, 1,1)){
; 		NewIngredientsString:="[Ingredients]`r`n" Clipboard
; 		filedelete, Ingredients.ini
; 		fileappend, %NewIngredientsString%, Ingredients.ini
;     sleep 400
;     Return
		;   }
	else if (Winactive("NuGenesis LMS") && InStr(Clipboard, "Value",true, 1,1)){
			TrimmedList:=Trim(RegExReplace(Clipboard, "^\s*Value[^\S\r\n]*\R", "", "", 1))
		if instr(Clipboard, "111Skin Limited",true,1,1){
			WorkTab.SaveCustomerList()
			clipboard:=TrimmedList
			msgbox,, Cleaned & Copied the Customer List,
		}
		else if instr(Clipboard, "5-HTP",true,1,1){
			ProductTab.SaveIngredientList()
			clipboard:=TrimmedList
			Msgbox,, Cleaned & Copied the Ingredients List,
		}
		else
			clip.codesRegex()
		; sleep 800
		Return
	}
  ; else if Winactive("Test Definition Editior"){
  ;   DESCRIPTION:=Trim(Clipboard,"`r`n")
  ;   ; TT(Description,2000)
	; 	return
  ; }
    ; iniwrite, %Description%, Settings.ini, CopiedSpecs, Description
  else if Winactive("Results Definition") || Winactive("Results ahk_exe eln.exe") {
    clip.ParseSpecsTable(1)
		; return
	}
  else if Winactive("Composition"){
    clip.ParseIngredientsTable()
	}
						; 	ELSE if (RegexMatch(Clipboard, RegexRequirements, subpat)){  ;THis clips requirments
						; 		; if (RegexMatch(Parse, RegexRequirements, r)){
						; 		Prefix:=
						; 		MinLimit:=
						; 		MaxLimit:=
						; 		Units:=
						; 		Prefix:=Subpat["Prefix"]
						; 						MinLimit:=Subpat["LowerLimit"]
						; 					MaxLimit:=Subpat["UpperLimit"]
						; 		Units:=Subpat["Unit"]
						; if (Prefix="<") && !(Minlimit)
						; 				Minlimit:=Prefix
						; 					; LowerLimit:=rLowerLimit
						; 					; UpperLimit:=rUpperLimit
						; 			ClippedRequirements:=Prefix   "  min:" Minlimit "  max:" maxlimit "  unit:" Units
						; 		; Unit:=rUnit
						; 			; Prefix:=
						; 					sleep 300
						; 		; msgbox, prefix: %Prefix% `nlowerlimit: %MinLimit% `nupperlimit: %Maxlimit% `nunit: %Units%
						; 					ControlsetText, Edit6,%ClippedRequirements%,ClipBar ahk_exe VQ_Helper
						; 		; return
						; }
  else
    clip.codesRegex()
		AllProducts:=GetAllProducts(" ")
		AllBatches:=GetAllBatches(" ")
    sleep 50
		return
	}



Clip(input=50,Wait:="0.95"){
  global tab, Batch, Product, lot, coated, sampleGUID, analytical,micro,retain,physical,CTphysical,CTretain,department, regexProduct, regexBatch,regexlot,regexSampleGUID, regexCoated
  clipboard:=
    Send, ^{c}
  sleep %input%
  clipwait,%Wait%
  if errorlevel
  {
    tt("clip error level")
    sleep 100
  }
  clip.CodesRegex()
  return
}





Class Clip {

VariableFileClip(input:=""){
Global
Gui ClipBar:Default
		; FileRead, FileContents, %Input%
		; FileContents := FileOpen(Input, "r")
		;clipBar.Flash()
  if Instr(Input, "[P]",true,1,1){
		Haystack:= input
		SheetInfo:=[]
		Coated:=
		Lot:=
		SheetInfo:=StrSplit(HayStack,"|")
		Product:=SheetInfo[2]
		ProductName:=SheetInfo[3]
		Customer:=SheetInfo[4]
		CustomerPosition:=SheetInfo[5]
		ShapeAndSize:=SheetInfo[6]
		Color:=SheetInfo[7]
		ServingSize:=SheetInfo[8]
		BatchLot:=SheetInfo[9]
		Batch:=RegexMatch(BatchLot, RegexBatch, r) ? rBatch : Batch
		Lot:=RegexMatch(BatchLot, RegexLot, r) ? rLot : Lot
		Coated:=RegExMatch(BatchLot, RegexCoated, r) ? rCoated : Coated
    ; ProductTab.AddProductFromClipboard(Input)
		Iteration:=SheetInfo[5]
		ProductInfo:=ProductName " : " Customer
		ServingInfo:=ServingSize " " Color
			; GUI, ClipBar:submit,NoHide
			stringUpper, Product, Product
      GuiControl,ClipBar:Text, Product, %Product%
      GuiControl,ClipBar:Text, Batch, %Batch%
      GuiControl,ClipBar:Text, lot, %lot%
      GuiControl,ClipBar:Text, Coated, %Coated%
      GuiControl,ClipBar:Text, Iteration, %Iteration%
		; Clip.SetClipBar()
      GuiControl,ClipBar:Text, Generalbox, %ProductInfo%
      GuiControl,ClipBar:Text, Generalbox2, %ServingInfo%
  }
	else if InStr(Input, "<<CopyLabelCopy>>",true, 1,1){
		; clip.codesRegex()
		simpleclip:=1
		; clip.codesRegex()
		Gui ClipBar:Default
		IF (SubStr(Clipboard, 19,4) != "Temp"){
			Product:=TRIM(SubStr(Input, 19,4))
			sleep 200
			GuiControl,ClipBar:Text, Product, %Product%
		}
		; else
			; return
		; Product:=RegexMatch(Clipboard, RegexProduct,r) ? rProduct : Product
		Clipboard:=
		; copyLabelCopyDocRegex(1,1)
		copyLabelCopyDoc(1)
		; tt(clipboard)
		sleep 2800
		; GoSub ShowFinalLabelCopy
		GoSub ShowScanLabelCopy
		; FileDelete, %A_ScriptDir%  \LabelCopyText.txt
		; Try FileDelete, %A_ScriptDir%  \LabelCopies\%Product%.txt
		; FileAppend,  %labelcopytext%, %A_ScriptDir%  \LabelCopies\%Product%.txt
		simpleclip:=
		return
  }
;   else if InStr(input, ">>|", true,1,1) {
;     if (Iteration >=25) || (Iteration < 0) || !(Iteration)
; 		{
;       iteration:=1
; 		}
;     LMS.AddDataFromClipboard(">>|",Input)
;     ; FileDelete, ClippedExcelData.txt
;     ; sleep 400
;     ; FileAppend, %ClippedData%, ClippedExcelData.txt
;     return
;   }
; else
; {
; 			Product:=RegexMatch(Input, RegexProduct,r) ? rProduct : Product
; 			Batch:= RegexMatch(Input, RegexBatch, r) ? rBatch : Batch
; 			Lot:=RegexMatch(Input, RegexLot, r) ? rLot : Lot
; 			Coated:=RegExMatch(Input, RegexCoated, r) ? rCoated : Coated

; 			if RegexMatch(Input, "\[\[(?P<CustomerPosition>-?\d+)\]\]", r){
; 					; Clipbar.FlashIteration()
; 					Iteration:=Floor(rCustomerPosition)
; 					CustomerPosition:=rCustomerPosition
; 			Clip.SetClipBar()
; 			}
; 	}
	; Clip.CodesRegex(FileContents
	return
}

	CodesRegex(input:=""){
		global RegexProduct, RegexBatch, RegexLot, RegexCoated, RegexSampleGUID, Product, Lot, Batch, Coated, sampleGUID, PriorSampleGUID, CodeString, CodeFile, PriorCodeString, CustomerPosition, Iteration,AllBatches, AllProducts
		Gui ClipBar:Default
		codestring:=
			; AllProducts:=
			coated:=
			; AllBatches:=
			sleep 200
			SampleGUID:=
			Parse:= Input ? Input : Clipboard
			rCustomerPosition:=

			Product:=RegexMatch(Parse, RegexProduct,r) ? rProduct : Product
			Batch:=RegexMatch(Parse, RegexBatch, r) ? rBatch : Batch
			Lot:=RegexMatch(Parse, RegexLot, r) ? rLot : Lot
			Coated:=RegExMatch(Parse, RegexCoated, r) ? rCoated : Coated
			SampleGUID:=RegExMatch(Parse, RegexSampleGUID, r) ? rSampleGUID : SampleGUID

				Ct:=Coated ? " ct#" : ""
				; Coated:=RegExMatch(Parse, RegexCoated, r) ? rCoated : Coated
				if (Batch!=PriorBatch) && (!rlot && !rCoated){
					PriorBatch:=Batch
					Lot:=
					Coated:=
				}
				if RegexMatch(Parse, "\[\[(?P<CustomerPosition>-?\d{1,3})\]\]", r){
					; Clipbar.FlashIteration()
					Iteration:=Floor(rCustomerPosition)
					CustomerPosition:=rCustomerPosition

			}
				if RegexMatch(Parse, "\|(?P<CustomerPosition>-?\d{1,3}+)\|", r){
					; Clipbar.FlashIteration()
					Iteration:=Floor(rCustomerPosition)
					CustomerPosition:=rCustomerPosition

			}

				this.SetClipBar()
				if rCustomerPosition
					GuiControl,ClipBar:Text, Iteration, %CustomerPosition%
				else
					GuiControl,ClipBar:Text, Iteration, %Iteration%


				if (Winactive("NuGenesis LMS") && InStr(Parse, "Ship To",true, 1,1) InStr(Parse, "Ship To",true, 1,1) )
					clip.GetSampleInfo()
				GuiControl,ClipBar:Text, Iteration, %Iteration%
					GUI, ClipBar:submit,NoHide
				if Coated
					codeString:=trim(Product " " Batch " " Lot " ct#" Coated)
				else
					codeString:=trim(Product " " Batch " " Lot)

		If (SampleGUID){
			iniwrite, %SampleGUID%, Settings.ini, SavedVariables, SampleGUID
		}

				TT(trim(Product " " Batch " " Lot Ct Coated "`n" SampleGUID),800,10,100,5,180,"M")
			Return
	}


GetSampleInfo(){ ;on the lms main menu
	global Coated, Customer, Iteration
	ParsedSample:=[]
	clipped_coated:=
	clipped_customer:=
	Loop, parse, Clipboard, `t
		ParsedSample.insert(A_LoopField)
	TotalColumns:=Parsedsample.maxindex()//2

	clipped_Customer:=ParsedSample[HasValue(ParsedSample, "Ship To") + TotalColumns]
	; Name:=Trim(ParsedSample[HasValue(ParsedSample, "Product Trade Name") + TotalColumns],"`r`n")
	Clipped_Coated:=ParsedSample[HasValue(ParsedSample, "Coated Lot #") + TotalColumns]

	GuiControl,ClipBar:Text, Iteration, %Iteration%
	if clipped_Customer
	{
		customer:=Clipped_Customer
		DetectedIteration:=GetIniValue("Customers.ini",Customer)
		tt(customer,4000)
		if DetectedIteration {
			ControlsetText, Edit5,%DetectedIteration%,ClipBar
			Iteration:=DetectedIteration
		}
		; GuiControl,ClipBar:Text, Iteration, %Iteration%
		; ControlsetText, Edit6,%Customer%,ClipBar
		; Clipbar.FlashIteration()
	}
	; else
		; GuiControl,ClipBar:Text, GeneralBox,


	if (clipped_Coated) && (strlen(clipped_coated) < 9)
	{

		Coated:=Clipped_Coated
		GuiControl,ClipBar:Text, Coated, %Coated%

	}
	else
		GuiControl,ClipBar:Text, Coated,

	GUI, ClipBar:submit,NoHide
	; if !ShipTo
	; ShipTo:=ShipToIndex
	; return ;ShiptoIndex
}

		ParseIngredientsTable(){
		global
    clipped_Position:=
    clipped_IngredientId:=
    clipped_LabelName:=
    clipped_LabelClaim:=
    clipped_IngredientGroup:=
		ParsedIngredients:=[]
		Loop, parse, Clipboard, `t
			ParsedIngredients.insert(A_LoopField)
			TotalColumns:=ParsedIngredients.maxindex()//2
			clipped_Position:=Trim(ParsedIngredients[HasValue(ParsedIngredients, "Position") + TotalColumns],"`r`n")
			clipped_IngredientId:=Trim(ParsedIngredients[HasValue(ParsedIngredients, "Ingredient Id") + TotalColumns],"`r`n")
			clipped_LabelName:=Trim(ParsedIngredients[HasValue(ParsedIngredients, "Description") + TotalColumns],"`r`n")
			clipped_LabelClaim:=Trim(ParsedIngredients[HasValue(ParsedIngredients, "Generic 1") + TotalColumns],"`r`n")
			clipped_IngredientGroup:=Trim(ParsedIngredients[HasValue(ParsedIngredients, "Generic 2") + TotalColumns],"`r`n")
      sleep 200
      Clipped_Ingredients:= Clipped_position ": " Clipped_IngredientId "`t"  Clipped_LabelClaim "`n" Clipped_IngredientGroup
      ; msgbox, %Clipped_Ingredients%
      ; Tooltip, %Clipped_Ingredients%, 200,0
			ControlsetText, Edit6,%Clipped_Ingredients%,ClipBar
			ControlsetText, Edit7,%Clipped_LabelName%,ClipBar
			gui, Clipbar:Submit,Nohide
      tt(Clipped_ingredients 5000,10,20,2)
      return
		}
  ParseSpecsTable(EnterData:=""){
		global
    MinLimit:=
    MaxLimit:=
    Percision:=
    Clipped_Requirement:=
    Clipped_ParsedSpecs:=
    FullRequirements:=
		AllowPrefixes:=
		UseLimitsBox:=
    Units:=
    Clipped_SeqNo:=
    Clipped_Method:=
    Clipped_ResultID:=
		ParsedSpecs:=[]
    ParseData:=Clipboard
		Clipboard.Flash()
		Loop, parse, ParseData, `t
			ParsedSpecs.insert(A_LoopField)
			TotalColumns:=ParsedSpecs.maxindex()//2
			UseLimitsBox:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Use the limits from the test") + TotalColumns],"`r`n")
			MinLimit_Units:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Lower Limit") + TotalColumns],"`r`n")
			MaxLimit_Units:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Upper Limit") + TotalColumns],"`r`n")
			Percision:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Precision") + TotalColumns],"`r`n")
			Clipped_Requirement:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Requirement") + TotalColumns], "`r`n")
      Clipped_ParsedSpecs:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Requirement") + TotalColumns],"`r`n")
			FullRequirements:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Requirement") + TotalColumns],"`r`n")
			Units:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Unit") + TotalColumns],"`r`n")
			AllowPrefixes:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Allow Prefixes") + TotalColumns],"`r`n")
			Clipped_SeqNo:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Seq No") + TotalColumns],"`r`n")
      Clipped_Method:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Method Id") + TotalColumns],"`r`n")
			Clipped_ResultID:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Result Id") + TotalColumns],"`r`n")
			MinLimit:=SubStr(MinLimit_Units,1, InStr(MinLimit_Units," "))
			MaxLimit:=SubStr(MaxLimit_Units,1, InStr(MaxLimit_Units," "))
      sleep 200
      Clipped_Specs:= Clipped_ResultID "`t" DESCRIPTION "`n MinLimit: " MinLimit "`n MaxLimit: " MaxLimit "`n Requirement: " Clipped_Requirement " `n Percision: " Percision "`n Units: " Units "`n Allow Prefix/limits: " AllowPrefixes "/" UseLimits "`n UseLimitsBox: " UseLimitsBox
			; if !UseLimits
			; 	UsedLimits:=
			; If !AllowPrefixes
			; 	AllowPrefixes:=
      tt(Clipped_Specs "`n`n" FullRequirement,3000,100,300)
      ; tt(,3000,100,300)
			ControlsetText, Edit6,%Clipped_Specs%,ClipBar

			ControlsetText, Edit7,%FullRequirements%,ClipBar
			; msgbox, Pasre spec table
        ; tooltip, %Clipped_specs%, 200,0
      If (EnterData){
        sleep 300

        ; Spectab.Autofill()
      }
      return
		}
  ParseMainSpecsTopTable(){
		global
    SimpleClip:=1
    Clipboard:=
    WinActivate, NuGenesis LMS
    send, ^{c}
    ClipWait, 6
    if ErrorLevel
      {          ; sleep 800
          ; MsgBox, The attempt to copy text onto the clipboard failed.
          SimpleClip:=
          TT("Errorlevel", 100,100)
          return
          ; exit
      }
    Product:=
    VersionType:=
    VersionStatus:=
    SelectionStatus:=
		ParsedSpecs:=[]
    ParseData:=Clipboard
		Loop, parse, ParseData, `t
			ParsedSpecs.insert(A_LoopField)
			TotalColumns:=ParsedSpecs.maxindex()//2
			Product:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Formulation ID") + TotalColumns],"`r`n")
			VersionType:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Specification version type") + TotalColumns],"`r`n")
			SampleguID:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Sample Guid") + TotalColumns],"`r`n")
			Description:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Description") + TotalColumns],"`r`n")
      Clipped_SpecsTop:= Product  ": " VersionType
      ControlsetText, Edit1,%Product%,ClipBar
      sleep 200
      ; iniRead, VersionStatus, FixedSpecs.ini, Status, %Product%
      SimpleClip:=
      sleep 350
      ; if VersionStatus
        ; return VersionStatus

      ; else
        return
		}


  ParseMainSpecTable(Save:=1){
		global
    ; Clipped_Method:=
    ; Clipped_Requirements:=
    ; Clipped_minmax:=
    ; MinLimit:=
    ; MaxLimit:=
    ; Clipped_TestID:=
    ; Clipped_Department:=
    ; Clipped_SeqNo:=
    ; Clipped_SampleTemplate:=
    ; DESCRIPTION:=
		ParsedMainSpecTable:=[]
    ParseData:=Clipboard
		Loop, parse, ParseData, `t  //`t
			ParsedMainSpecTable.insert(A_LoopField)
    TotalColumns:=ParsedMainSpecTable.maxindex()//2
    Clipped_Method:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Method Id") + TotalColumns],"`r`n")
    FullRequirement:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Requirements") + TotalColumns],"`r`n")
    Clipped_minmax:=Strsplit(Requirements," - ","`n")
    MinLimit:=Trim(Minmax[1],"`r`n")
    MaxLimit:=Trim(minmax[2],"`r`n")
    Clipped_TestID:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Test Id") + TotalColumns],"`r`n")
    Clipped_Department:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Department") + TotalColumns],"`r`n")
    Clipped_SeqNo:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Seq No") + TotalColumns],"`r`n")
    Clipped_SampleTemplate:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Sample Template") + TotalColumns],"`r`n")
    DESCRIPTION:=Trim(ParsedMainSpecTable[HasValue(ParsedMainSpecTable, "Description") + TotalColumns],"`r`n")
    ; StringReplace, Description, Description, `r`n, , All
      sleep 100
      Clipped_Specs:= Clipped_TestID "`t" DESCRIPTION "`n MinMax: " MinLimit " - " MaxLimit "`n Sample Template: " Clipped_SampleTemplate "`n Department: " Clipped_Department
      TT(%Clippsed_Specs%,4000)
			; msgbox, parse main spec table
			ControlsetText, Edit6,%Clipped_Specs%,ClipBar
			ControlsetText, Edit7,%FullRequirement%,ClipBar
        ; tooltip, %Clipped_specs%, 200,0
      return
		}




  HeavyMetalSpecs(){
    global
    HM_Specs:=StrSplit(clipboard,"|")
    HM_Unit:=HM_Specs[2]
    HM_Lower_Limit:=HM_Specs[3] " " HM_Unit
    Arsenic_Limit:=HM_Specs[4]
    Lead_Limit:=HM_Specs[5]
    Cadmium_Limit:=HM_Specs[6]
    Mercury_Limit:=HM_Specs[7]
    Arsenic_Requirement:=HM_Specs[8] " " HM_Unit
    Lead_Requirement:=HM_Specs[9] " " HM_Unit
    Cadmium_Requirement:=HM_Specs[10] " " HM_Unit
    Mercury_Requirement:=HM_Specs[11] " " HM_Unit

	}

EditBoxOld(Input:=""){
  Global EditBox
	try GUI, EditBox:Destroy
  if !Input
		Result := Clipboard
  else
    result := Input
    sleep 200
	Gui EditBox: +AlwaysOnTop +Toolwindow +Resize +owner +HwndEditBoxGUIID
		GUI, EditBox:Font, s12 cBlack, Consolas
    gui, EditBox:Margin,1,1
		GUI, EditBox:Add, Edit, x0 y0 +Resize vEditBox , % Result
		; Gui, Add, DropDownList, vDelimeter, {space}|{enter}|``r|,|`;|
    GUI, EditBox:add, button,default gEditBoxButtonOK, OK
     ;GUI, EditBox:add, radio,default gEditBoxButtonOK, {space}|{enter}|RnCr|;|,|
    If !Input
      try GUI, EditBox:Destroy
    else
      GUI, EditBox:Show,Autosize, Result
    ; winSet, Transparent, 100, EditBox
		return
		OldEditBoxGuiClose:
    editbox:=
    try GUI, EditBox:Destroy
    Return
		OldEditBoxGuiEscape:
    editbox:=
    try GUI, EditBox:Destroy
    Return
		OldEditBoxButtonOK:
		GUI, EditBox:submit
    if !Input
			clipboard:=EditBox
    sleep 200
    try GUI, EditBox:Destroy
    ; sleep 50
  TT(Editbox,1000,ClipBar_X,33,,200)
  if Editbox
    ControlsetText, Edit6,%EditBox%,ClipBar
		return

    OldEditBoxGuiSize:
  if (ErrorLevel = 1)  ; The window has been minimized. No action needed.
      return
    NewWidth := A_GuiWidth
    NewHeight := A_GuiHeight
    ; sleep 200
    GuiControl, Move, EditBox, W%NewWidth% H%NewHeight%
    return

	}



EditBox(Input:=""){
  Global EditBox, Delimiter
	try GUI, EditBox:Destroy
  if !Input
		Result := Clipboard
  else
    result := Input
    sleep 200
	Gui EditBox: +AlwaysOnTop +Toolwindow +Resize +owner +HwndEditBoxGUIID
		GUI, EditBox:Font, s12 cBlack, Consolas
    gui, EditBox:Margin,1,1
		GUI, EditBox:Add, Edit, x0 y0 +Resize vEditBox , % Result
    GUI, EditBox:Add, Radio, x5 y+5 vDelimiter gSetDelimiter, ||{Space}|{Tab}|{`,}|{`;}|{Other}
    GUI, EditBox:Add, Edit, x55 y+7 w50 vOtherDelimiter
    GuiControl, Disable, OtherDelimiter
    GUI, EditBox:Add, Button, Default x+30 y+30 gEditBoxButtonOK, OK
    If !Input
      try GUI, EditBox:Destroy
    else
      GUI, EditBox:Show,Autosize, Result
    ; winSet, Transparent, 100, EditBox
		return
		EditBoxGuiClose:
    editbox:=
    try GUI, EditBox:Destroy
    Return
		EditBoxGuiEscape:
    editbox:=
    try GUI, EditBox:Destroy
    Return
		EditBoxButtonOK:
		GUI, EditBox:submit
    if !Input
			clipboard:=EditBox
    sleep 200
    try GUI, EditBox:Destroy
    ; sleep 50
  TT(Editbox,1000,ClipBar_X,33,,200)
  if Editbox
    ControlsetText, Edit6,%EditBox%,ClipBar
		return

    EditBoxGuiSize:
  if (ErrorLevel = 1)  ; The window has been minimized. No action needed.
      return
    NewWidth := A_GuiWidth
    NewHeight := A_GuiHeight
    ; sleep 200
    GuiControl, Move, EditBox, W%NewWidth% H%NewHeight%
    return

    SetDelimiter:
      ; GuiControlGet, Delimiter, Radio, EditBox
      if (Delimiter = "Other")
        GuiControl, Enable, OtherDelimiter
      else
        GuiControl, Disable, OtherDelimiter
      return
	}














SetClipBar(){
  Global Product, Batch, Lot, Coated
    Gui ClipBar:Default
      stringUpper, Product, Product
      GuiControl,ClipBar:Text, Product, %Product%
      GuiControl,ClipBar:Text, Batch, %Batch%
      GuiControl,ClipBar:Text, lot, %lot%
      GuiControl,ClipBar:Text, Coated, %Coated%
      GuiControl,ClipBar:Text, Iteration, %Iteration%
; Clipbar.Flash()
}


Regex(Category:=""){
    global
    if !Category
      Haystack:=Clipboard
		else
      Haystack:=Category
    sleep 100
      Gui ClipBar:Default
      RegExMatch(HayStack, RegexProduct,r)  ;"i)[abdefghijklmn]\d{3}", rProduct)
        If rProduct {
          GuiControl,ClipBar:Text, Product, %rProduct%
          stringUpper, rProduct, rProduct
          Product:=rProduct
        }
      RegExMatch(HayStack, RegexBatch, r)
        If rBatch {
          GuiControl,ClipBar:Text, Batch, %rBatch%
          Batch:=rBatch
        }
      RegExMatch(HayStack, RegexLot, r)
        If rLot {
          GuiControl,ClipBar:Text, lot, %rlot%
          lot:=rLot
        }
      RegExMatch(HayStack, RegexCoated, r)
        If rCoated {
          GuiControl,ClipBar:Text, Coated, %rCoated%
          Coated:=rCoated
					Ct:=" Ct#"
        }
      RegExMatch(HayStack, RegexSampleGUID, r)
        If rSampleGUID {
          SampleGUID:=rSampleGUID
          ; GuiControl,ClipBar:Text, SampleGUID, %rSampleGUID%
        }
      RegExMatch(HayStack, RegexSampleGUID, r)
        If !rSampleGUID {
          SampleGUID:=
          ; GuiControl,ClipBar:Text, SampleGUID, %rSampleGUID%
        }
      If !rLot {
        GuiControl,ClipBar:Text, lot, %rlot%
        lot:=
      }
      If !rCoated {
        GuiControl,ClipBar:Text, Coated, %rCoated%
        Coated:=
				Ct:=
      }
			; if (rProduct) & (rBatch) & (rLot)
			; 	{
			; 		NewString:=Trim(Product " " Batch " " Lot Ct Coated)
			; 		ControlsetText, Edit6,%NewString%,ClipBar ahk_exe VQ_Helper
			; 	}
      ; GuiControl, ClipBar:MoveDraw, Coated
      gui ClipBar:submit, nohide
      ; tt(Product " " Batch " " lot " ct#" Coated)
      sleep 20
			Return ;Trim(Product " " Batch " " Lot Ct Coated)
		}


Department(DepartmentInput:=""){
  global Department, organolepticTemplate
      Department:=
			OrganolepticTemplate:=
     if !DepartmentInput
      DepartmentHaystack:=Clipboard
    else
      DepartmentHaystack:=DepartmentInput
      Regexmatch(DepartmentHaystack, "F,\sMicro",Micro)
      Regexmatch(DepartmentHaystack, "I,\sRetain", Retain)
      Regexmatch(DepartmentHaystack, "I,\sPhysical", Physical)
      Regexmatch(DepartmentHaystack, "CT,\sPhysical", ctPhysical)
      Regexmatch(DepartmentHaystack, "CT,\sRetain", ctRetain)
      Regexmatch(DepartmentHaystack, "I,\sAnalytical", Analytical)
      Sleep      20
      If Micro
        Department:="Micro"
      else If Retain
        Department:="Retain"
      else If ctRetain
        Department:="ctRetain"
      else If Physical
			{
				Regexmatch(DepartmentHaystack, "Organoleptic", OrganolepticTemplate)
        Department:="Physical"
				OrganolepticTemplate:="Organoleptic"
			}
      else If ctPhysical
        Department:="ctPhysical"
      else If Analytical
        Department:="Analytical"
      else
				return
        ; TT("`nDepartment: " Department "`nmicro: " micro "`nretain: " retain "`nctretain: " ctretain "`nphysical: " physical "`nctphysical: " ctphysical "`nanalytical: " analytical,3000,100,100)
  return %Department%
}









}






NumberMenu(n:=0, runCorrectTestResults:=""){
	global
	MouseGetPos, mx, my
	try Menu, NumberMenu, DeleteAll
	Menu, NumberMenu, Add, &0, NumberMenubutton
	loop %n%,
	{
		Menu, NumberMenu, Add, &%A_index%, NumberMenubutton
	}
	sleep 200
		Try Menu,NumberMenu,show

		return
NumberMenubutton:
	if A_ThisMenuItemPos
	{
		Iteration:=A_ThisMenuItemPos - 1
	sleep 100
	ControlsetText, Edit5,%Iteration%,ClipBar
	; Clipbar.FlashIteration()
	MouseMove, %mx%, %my%, 0
	sleep 300
	if runCorrectTestResults = toggle
		WorkTab.CorrectTestResults("toggle", "Loop")
	else if runCorrectTestResults = 0


		WorkTab.CorrectTestResults(0,5)

		}
		blockinput, off
Return
}

ServingSizeMenu(Preselect:=""){
	global
	; MouseGetPos, mx, my
	if Preselect
	{
		Preselect:=Trim(preselect)
		gosub, ServingSizeMenuButton
		return
	}
	try Menu, ServingSizeMenu, DeleteAll
	Menu, ServingSizeMenu, Add, (1) capsule,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (2) capsules,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (3) capsules,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (4) capsules,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add,
	Menu, ServingSizeMenu, Add, (1) tablet,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (2) tablets,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (3) tablets,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (4) tablets,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add,
	Menu, ServingSizeMenu, Add, (1) scoop,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (2) scoops,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add,
	Menu, ServingSizeMenu, Add, (1) stick packet,ServingSizeMenuButton
	Menu, ServingSizeMenu, Add, (1) packet,ServingSizeMenuButton
		Try Menu,ServingSizeMenu,show
	return

	ServingSizeMenuButton:
	Winactivate, Edit Ingredient
	if A_ThisMenuItem
	{
		; send, {Click 325, 337}
		; click 317, 306
		; click 325, 337
		sleep 300
		winactivate, Edit Ingredient
		click 352, 327 ; click description
		if instr(A_ThisMenuItem,"(2)")
			Send, {end}two %A_ThisMenuItem%
		else if instr(A_ThisMenuItem,"(3)")
			Send, {end}three %A_ThisMenuItem%
		else if instr(A_ThisMenuItem,"(4)")
			Send, {end}four %A_ThisMenuItem%
		; else if instr(A_ThisMenuItem,"(5)")
		; 	Send, five %A_ThisMenuItem%
		else
			Send, {end}%A_ThisMenuItem%
		if instr(A_ThisMenuItem,"scoop")
			send, {space}( g){left 3}
		else if instr(A_ThisMenuItem,"packet")
			send, {space}( g){left 3}
		else
			return
	; MouseMove, %mx%, %my%, 0
	}
	if Preselect
	{
		winactivate,
		sleep 300
		if instr(Preselect,"2 ")
			Sendinput % "{end}two (" SubStr(Preselect, 1,1) ")" substr(Preselect,2)
		else if instr(Preselect,"3 ")
			Sendinput % "{end}three (" SubStr(Preselect, 1,1 )")" substr(Preselect,2)
		else if instr(Preselect,"4 ")
			Sendinput % "{end}four (" SubStr(Preselect, 1,1) ")" substr(Preselect,2)
		else if instr(Preselect,"5 ")
			Sendinput % "{end}five (" SubStr(Preselect, 1,1) ")" substr(Preselect,2)
		else
			Sendinput % "{end}(" SubStr(Preselect, 1,1) ")" substr(Preselect,2)
		preselect:=
		return
		}
	sleep 300
Return
		}

MenuCodeSelect(FileName:="AllProducts"){
	global
	try Menu, CodeMenu, DeleteAll
	FileRead, OutputText, %FileName%.txt
	stringUpper, OutputText, OutputText
	loop, parse, OutputText, " "
		Menu, CodeMenu, Add, %a_LoopField%, CodeMenubutton
			Menu, CodeMenu, Add, All, CodeMenubutton
	FileRead, WholeBatches, WholeBatches.txt
		Menu, CodeMenu, Add, , , +Break
	sleep 200
	Loop, parse, WholeBatches, "`n"
		Menu, CodeMenu, Add, %a_LoopField%, WholeBatchMenubutton,
		Menu, CodeMenu, Add, , WholeBatchMenubutton
		Menu, CodeMenu, Add, %WholeBatches%[2], WholeBatchMenubutton
	Try Menu,CodeMenu,show
	return
MenuCodeSelectMenubutton:
	clip.codesRegex(A_ThisMenuItem)
return


AllproductsMenuButton:
		product := A_ThisMenuItem
		Menu, AllProductsMenu, Check, %A_ThisMenuItem%
		ControlsetText, Edit1,%product%,ClipBar
		; Batch:=
		; lot:=
		; coated:=
		; ControlsetText, Edit2,%batch%,ClipBar
		; ControlsetText, Edit3,%lot%,ClipBar
		; ControlsetText, Edit4,%Coated%,ClipBar
		return

AllBatchesMenuButton:
		batch := A_ThisMenuItem
		; ControlsetText, Edit1,%product%,ClipBar
		; lot:=
		; coated:=
		Menu, AllBatchesMenu, Check, %A_ThisMenuItem%
		ControlsetText, Edit2,%Batch%,ClipBar
		; ControlsetText, Edit3,%lot%,ClipBar
		; ControlsetText, Edit4,%Coated%,ClipBar
		return


CodeMenubutton:
If A_ThisMenuItem contains All
{
	if winactive("NuGenesis LMS")
		LMS.SearchBar(A_ThisMenuItem)
	else
		sendinput, %A_ThisMenuItem%
	return
}

If FileName contains Product
	{
		product := A_ThisMenuItem
		ControlsetText, Edit1,%product%,ClipBar
	}
else If FileName contains Batch
	{
		batch := A_ThisMenuItem
		ControlsetText, Edit2,%Batch%,ClipBar
	}
else If FileName contains SampleGUID
	SampleGUID := A_ThisMenuItem

sleep 200
if winactive("NuGenesis LMS")
		LMS.SearchBar(A_ThisMenuItem,"{enter}",0)
if winactive("ahk_exe explorer.exe")
	send, ^e{*}%A_ThisMenuItem%{*}{enter}
else
	sendinput, %A_ThisMenuItem%

Return
}


WholeBatchMenu(){
	global
	try Menu, CodeMenu, DeleteAll
	FileRead, WholeBatches, WholeBatches.txt
	WholeBatches:=Trim(StrReplace(WholeBatches, "`n`n", ""))
	Loop, parse, WholeBatches, "`r"
		Menu, CodeMenu, Add, &%A_Index% `t %a_LoopField%, WholeBatchMenubutton,
	Try Menu,CodeMenu,show
	return

WholeBatchMenubutton:
if (Instr(Trim(A_ThisMenuItem),"|")=1){
		ProductTab.AddProductFromClipboard(A_ThisMenuItem)
		return
	}
		else clip.codesRegex(A_ThisMenuItem)
return
}

PriorCodesMenu(ShowMenu:=""){
	global
	PreviousCodes:=[]
	try Menu, CodeMenu, DeleteAll
	FileRead, PriorCodes, PriorCodes.txt
	PriorCodes:=Trim(StrReplace(PriorCodes, "`n`n", ""))
	Loop, parse, PriorCodes, "`r"
		{
		Menu, CodeMenu, Add, &%A_Index% `t %a_LoopField%, PriorCodesMenubutton,
		PreviousCodes.Push(A_loopfield)
		}
		if ShowMenu
			Try Menu,CodeMenu,show

	return

PriorCodesMenubutton:
clip.codesRegex(A_ThisMenuItem)
return
}




GetAllProducts(Delimiter:=" ",msg:=""){
	global
	regProducts:=[]
	pos=0
	Haystack:=Clipboard
	If Getkeystate("ROption","P")
		Delimiter:="`;"
	sleep 50
	while pos := RegexMatch(Haystack, RegexProduct, aProduct, pos+1) ; {
		regProducts.insert(aProduct)
	; }
	AllProducts:=[], oTemp := {}
	for vKey, vValue in regProducts
	{
		if (ObjGetCapacity([vValue], 1) = "") ;is numeric
		{
			if !ObjHasKey(oTemp, vValue+0)
				AllProducts.Push(vValue+0), oTemp[vValue+0] := ""
		}
		else
		{
			if !ObjHasKey(oTemp, "" vValue)
				AllProducts.Push("" vValue), oTemp["" vValue] := ""
		}
	}
	AllProducts:=Listarray(AllProducts,Delimiter)
	AllProducts:= Trim(StrReplace(AllProducts, Delimiter Delimiter, Delimiter),Delimiter)
	sleep 200
	StringUpper, AllProducts, AllProducts
	if (StrLen(AllProducts) > 5){
	FileDelete, AllProducts.txt
		; ControlsetText, Edit6,,ClipBar
	FileAppend, %AllProducts%, AllProducts.txt
	ControlsetText, Edit6,%AllProducts%,ClipBar
		; Allproducts:=
		; return
	}
	if (msg && AllProducts!="")
		clip.editbox(AllProducts)
	Else
			Return AllProducts
}

GetAllBatches(Delimiter:=" ",msg:=""){
	global
	regBatches:=[]
	pos=0
	Haystack:=Clipboard
	If Getkeystate("LControl","P")
		Delimiter:="`;"
	If Getkeystate("ROption","P")
		Delimiter:="`;"
	sleep 50
	while pos := RegexMatch(Haystack, RegexBatch, aBatch, pos+1) ; {
		regBatches.insert(aBatch)
	; }
	AllBatches:=[], oTemp := {}
	for vKey, vValue in regBatches
	{
		if (ObjGetCapacity([vValue], 1) = "") ;is numeric
		{
			if !ObjHasKey(oTemp, vValue+0)
				AllBatches.Push(vValue+0), oTemp[vValue+0] := ""
		}
		else
		{
			if !ObjHasKey(oTemp, "" vValue)
				AllBatches.Push("" vValue), oTemp["" vValue] := ""
		}
	}
	AllBatches:=Listarray(AllBatches,Delimiter)
	AllBatches:= Trim(StrReplace(AllBatches, Delimiter Delimiter, Delimiter),Delimiter)
	sleep 200
	if (StrLen(AllBatches) > 10){
	FileDelete, AllBatches.txt
		FileAppend, %AllBatches%, AllBatches.txt
		ControlsetText, Edit7,%AllBatches%,ClipBar
		; ControlsetText, Edit7,,ClipBar
		; AllBatches:=
		; return
	}
	if (msg && trim(AllBatches)!="")
		clip.editbox(AllBatches)
	Else
			Return AllBatches
}

WholeBatchesSave(Input,Overwrite:=""){
	Global
	if Overwrite
		FileDelete, WholeBatches.txt
	sleep 200
	FileAppend, %input%, WholeBatches.txt
	sleep 200
	FileRead, WholeBatches, WholeBatches.txt
	sleep 200
	tt(WholeBatches,1000,0,0,2)
		return
}

GetAllWholeBatches(Delimiter:=" ",msg:=""){
	global
	regWholeBatches:=[]
	pos=0
	Haystack:=Clipboard
	If Getkeystate("LControl","P")
		Delimiter:="`;"
	If Getkeystate("ROption","P")
		Delimiter:="`;"
	sleep 50
	while pos := RegexMatch(Haystack, RegexCombined, aWholeBatch, pos+1) ; {
		regWholeBatches.insert(aWholeBatch)
	; }
	AllWholeBatches:=[], oTemp := {}
	for vKey, vValue in regWholeBatches
	{
		if (ObjGetCapacity([vValue], 1) = "") ;is numeric
		{
			if !ObjHasKey(oTemp, vValue+0)
				AllWholeBatches.Push(vValue+0), oTemp[vValue+0] := ""
		}
		else
		{
			if !ObjHasKey(oTemp, "" vValue)
				AllWholeBatches.Push("" vValue), oTemp["" vValue] := ""
		}
	}
	AllWholeBatches:=Listarray(AllWholeBatches,Delimiter)
	AllWholeBatches:= Trim(StrReplace(AllWholeBatches, Delimiter Delimiter, Delimiter),Delimiter)
	FileDelete, WholeBatches.txt
	sleep 200
	if (StrLen(AllWholeBatches) < 10){
		ControlsetText, Edit7,,ClipBar
		AllWholeBatches:=
		return
	}
		FileAppend, %AllWholeBatches%, WholeBatches.txt
		ControlsetText, Edit7,%AllWholeBatches%,ClipBar
	if (msg && trim(AllWholeBatches)!="")
		clip.editbox(AllWholeBatches)
	Else
			Return AllWholeBatches
}

GetAllWholeBatches2(Delimiter:="`n",msg:=""){
	global
	aWholeBatches:=[]
	; pos=0
	Haystack:=Clipboard
	FileRead, WholeBatches, WholeBatches.txt
	sleep 100
	WholeBatchesArray:=[]
	loop, parse, WholeBatches, "`r`n"
		aWholeBatches.insert(a_LoopField)

	AllWholeBatches:=[], oTemp := {} remove duplicates
	for vKey, vValue in aWholeBatches
	{
		if (ObjGetCapacity([vValue], 1) = "") ;is numeric
		{
			if !ObjHasKey(oTemp, vValue+0)
				AllWholeBatches.Push(vValue+0), oTemp[vValue+0] := ""
		}
		else
		{
			if !ObjHasKey(oTemp, "" vValue)
				AllWholeBatches.Push("" vValue), oTemp["" vValue] := ""
		}
	}
	AllWholeBatches:=Listarray(AllWholeBatches,"`r`n")
	AllWholeBatches:=StrReplace(AllWholeBatches, "`r`n`r`n", "`r`n")
	AllWholeBatches:=StrReplace(AllWholeBatches, A_space A_space, A_Space)
	AllWholebatches:=Trim(AllWholeBatches,"`r`n")
	; AllWholeBatches:= StrReplace(AllWholeBatches, "+")
	FileDelete, WholeBatches.txt
	sleep 200
	FileAppend, %AllWholeBatches%, WholeBatches.txt
	if msg
		clip.editbox(AllWholeBatches)
	Else
			Return AllWholeBatches
}






;;------------------------------------------------------------
;;[[                       ClipBar                          ]]
;;------------------------------------------------------------
Class ClipBar{
	Show(X:=1, Y:=1, Destroy:="Reset"){
		Global
		try GUI,ClipBar:Destroy
		TopScreen:=1
		MidScreen:=A_ScreenWidth//2
		wingetpos, Nugenesis_X, Nugenesis_y, Nugenesis_w, Nugenesis_h, NuGenesis LMS
		wingetpos, Win_X, Win_y, Win_w, Win_h, A
		ClipBar_H=33
		ClipBar_H_max=56
		ClipBar_T:=205
		ClipBar_W=440
		IF (ClipbarLocation= "BOTTOMR") || (ClipbarLocation= "RIGHT") || (ClipbarLocation= "TOPR")
			ClipBar_x:=Nugenesis_W-(Nugenesis_W//6)
		else
			ClipBar_x:=Nugenesis_X+(Nugenesis_W//6)
		; ClipBar_x:=1
		IF (ClipbarLocation= "BOTTOM") || (ClipbarLocation= "BOTTOMR") || (ClipbarLocation= "RIGHT")
			ClipBar_Y:=Nugenesis_h + Nugenesis_y - 40
		else
			ClipBar_Y:=Nugenesis_y -0
		Gui ClipBar: +AlwaysOnTop -Caption +Toolwindow +owner +HwndClipBarGUIID
		Gui ClipBar:Default
		; Gui, ClipBar:+Delimiter`n
		GUI, ClipBar:color,ff6c56,ffffff
			GUI,ClipBar:Font,		ce66000	 s18 Bold , consolas
			; GUI,ClipBar:Add,edit,		vProduct +wrap -multi	gclipbarhandler left h33 x7 y-1 w65,	%Product%
		GUI,ClipBar:Add,edit,		vProduct +wrap	gClipBarHandler left h33 x20 y0 w65,	%Product%
		this.AddEdit("Batch",	 "left h33 x+0 y0 w110", 			"16 c0800ff,")
		this.AddEdit("Lot",		 "left h17 x+0 y0 w77", 			"10 c01750d, ")
		this.AddEdit("Coated",	 "left h16 y+0 w77",		"8 c909B00, Arial Narrow")
		GUI, ClipBar:font, cffffff s9 Norm w500 , Consolas
		This.AddEdit("Iteration", "x+0 h33 left y0 w62 Center",			 "17 Bold 107C41, Consolas")	; Text1
		this.AddEdit("GeneralBox",	 "x+0 h17 left y0 w100",		"7, Arial")
		this.AddEdit("GeneralBox2",	 "h16 left y+0 w100",		"7, Arial")
		this.AddBoxes()
		CoordMode, mouse, screen
		try GUI, ClipBar:Show, x%ClipBar_X% y%ClipBar_y% w%ClipBar_w% h%ClipBar_H% Noactivate, ClipBar
		catch
			GUI, ClipBar:Show, x%Win_X% y%Win_Y% w%ClipBar_w% h%ClipBar_H% Noactivate, ClipBar
		CoordMode, mouse, window
		; OnMessage(0x0201, "WM_LBUTTONDOWN")
		this.loadSavedVariables()
			OnMessage(0x0201, "WM_LBUTTONDOWN")
			OnMessage(0x0203, "WM_LBUTTONDBLCLK")
			; OnMessage(  WM_LBUTTONUP := 0x0202, "WM_LBUTTONUP")
			;OnMessage(0x0200, "WM_MOUSEMOVE")
			; OnMessage(0x203,  "VariableBar_Relocate")
		; OnMessage(0x002C, "ODDDL_MeasureItem") ; WM_MEASUREITEM
		; OnMessage(0x002B, "ODDDL_DrawItem") ; WM_DRAWITEM
		winSet, Transparent, %ClipBar_T%, AHK_id %ClipBarGUIID%
		Return

		ClipBarHandler:
			PriorCodestring:=CodeString
			GUI, ClipBar:submit,NoHide
			Ct:=Coated ? " ct#" : ""
			; CodeString:=clip.CodesRegex(clipboard)
			CodeString:=trim(Product " " Batch " " Lot Ct Coated)
				; if (PriorCodestring!=Codestring){
					; FileDelete, %CodeFile%
					; sleep 200
					; FileAppend, %CodeString%, %CodeFile%
				; }
		return
		ClipBarGuiClose:
		ClipBarButtonOK:
		ClipBarGuiEscape:
			; GUI, ClipBar:submit,NoHide
			sleep 50
			ClipBar.SaveVariables()
		return
	}
	Flash(){
		global
		Gui ClipBar:Default
		Loop 1
		{
			GUI, ClipBar:color,ff0000,ff0000
			winSet, Transparent, 255, AHK_id %ClipBarGUIID%
			sleep 100
			GUI, ClipBar:color,DC734F,ffffff
			winSet, Transparent, 255, AHK_id %ClipBarGUIID%
			sleep 50
		}
	}
		FlashIteration(){
			Global
		Gui ClipBar:Default
			Gui, Font, cff0000 s16  ; This just tells what color to change to
			GuiControl, Font, Edit5
			sleep 200
			Gui, Font, cBlack s16  ; This just tells what color to change to
			GuiControl, Font, Edit5
	}
	AddEdit(Variable,Dimensions:="",Font:=""){
		global
		GUI,ClipBar:Font,			 s%Font% , consolas ;cBlack Bold, %Font%
		GUI,ClipBar:Add,edit,		v%Variable% +wrap -multi	gClipBarHandler %Dimensions%,		% %Variable%
	}
	AddText(Variable,Dimensions:="",Font:=""){
		global
		GUI,ClipBar:Font,			 s%Font% ;cBlack Bold, %Font%
		GUI,ClipBar:Add,Text,		v%Variable% 	%Dimensions%,		% %Variable%
	}

	Menu(){
		global
		If Nsub
				exit
			Nsub:=1
		MouseGetPos,,,,winControl
		try Menu, ClipBarmenu, DeleteAll
		Menu, ClipBarMenu, Add, All Products, AllProductsNoMsgbox
		Menu, ClipBarMenu, Add, All Batches, AllBatchesNoMsgbox
		Menu, ClipBarMenu, Add,
		Menu, ClipBarMenu, add, Show Final Label Copy, ShowFinalLabelCopy
		Menu, ClipBarMenu, add, Show Scan Label Copy, ShowScanLabelCopy
		Menu, ClipBarMenu, add, Show Total CoAs, ShowFINAL_C_O_A
		Menu, ClipBarMenu, add, &Total CoAs, ShowFINAL_C_O_A
		Menu, ClipBarMenu, add, C_O_A, ShowC_O_A
		Menu, ClipbarMenu, add, Manual &COAs folder, ShowManualCOA
		Menu, ClipbarMenu, add, &mfg folder, Showmfg
		Menu, ClipbarMenu, add, &GLOBAL VISION folder, ShowGlobalVision

		Menu, ClipBarMenu, add, Update List via Clipboard, SaveClipboardToList
		Menu, ClipBarMenu, add, Add Data From Clipboard, +!F10
		Menu, ClipBarMenu, Add, Stop Timer, StopTimer
		; Menu, ClipbarMenu, add, Add Sample Log, Add15SampleLog
		Menu, ClipBarMenu, Add, Detect Clipboard, clipChange
		if winactive("Results Definition") || winactive("Composition") || winactive("Results"){
			Menu, ClipBarMenu, Add, USP Heavy Metal, Heavy_metals, +Break
			Menu, ClipBarMenu, Add, Canada Heavy Metal, Heavy_metals
			Menu, ClipBarMenu, Add, Prop65 Heavy Metal, Heavy_metals
			Menu, ClipBarMenu, Add, Report Only Heavy Metal, Heavy_metals
			Menu, ClipBarMenu, Add, Bloom Nutrition Heavy Metal, Heavy_metals
			Menu, ClipBarMenu, Add, Custom Heavy Metal, Heavy_metals
		}
		Try Menu,ClipBarmenu,show
		Nsub:=
	}
	Reset(){
		Global
		coordmode, mouse, Screen
		; keywait, Lbutton, U T2
		MouseGetPos,ClipBar_X,ClipBar_Y
		ClipBar.show(0)
		coordmode, mouse, window
		return
	}

	Focus(Control){
		global
		winactivate, ClipBar ahk_exe VQ_Helper
		GuiControl ClipBar:Focus, %Control%
		Sendinput, +{left}
		return
	}

	Search(input){
		global
		Clip()
		LMS.SearchBar(input)
		return
	}

	AddIteration(speed:=300,showtooltip:="",PreventStall:=""){
		global Iteration, NAdd
		Thread, Interrupt, 0
			; If NAdd && !PreventStall
			; 	{
			; 		Sleep 500
			; 		return
			; 	}
			; #maxthreadsperhotkey, 1
			; #MaxHotkeysPerInterval, 500
			; NAdd:=1
			sleep 55
			Iteration+=1
			ControlsetText, Edit5,%Iteration%,ClipBar
			; sleep 100
			; #MaxHotkeysPerInterval, 70
			if showTooltip
				tt(Iteration)
			CustomerPosition:=Iteration
			sleep %Speed%
			; #maxthreadsperhotkey, 2
			; NAdd:=

			; GuiControl, -redraw, ClipBar
		return
	}
	SubIteration(speed:=300){
		global Iteration, NSub
		; If NSub
			; {
				; sleep 500
			; return
			; }
		; #maxthreadsperhotkey, 1
		; #MaxHotkeysPerInterval, 500
		; NSub:=1
		; GuiControl, -redraw, ClipBar
		sleep 55
		Iteration-=1
		ControlsetText, Edit5,%Iteration%,ClipBar
		; sleep 100
		tt(Iteration)
		CustomerPosition:=Iteration
		; Nsub:=
		; #maxthreadsperhotkey, 2
		sleep %speed%
		return
	}

		Relocate(){
		global
				PostMessage, 0xA1, 2
		return
	}
	SaveVariables(){ ;;_________________SAVING VARIABLES_________________________
	global
		GUI, ClipBar:default
		GUI, ClipBar:Submit, Nohide
		ControlGetText, Product, Edit1, ClipBar
		ControlGetText, Batch, Edit2, ClipBar
		ControlGetText, Lot, Edit3, ClipBar
		ControlGetText, Coated, Edit4, ClipBar
		ControlGetText, Iteration, Edit5, ClipBar
		ControlGetText, GeneralBox, Edit6, ClipBar
		ControlGetText, GeneralBox2, Edit7, ClipBar
		Null:=""
		sleep 100
		if Product
			iniwrite, %Product%, Settings.ini, SavedVariables, Product
		else
			iniwrite, %Null%, Settings.ini, SavedVariables, Product
		if Batch
			iniwrite, %Batch%, Settings.ini, SavedVariables, Batch
		else
			iniwrite, %Null%, Settings.ini, SavedVariables, Batch
		if lot
			iniwrite, %Lot%, Settings.ini, SavedVariables, Lot
		else
			iniwrite, %Null%, Settings.ini, SavedVariables, Lot
		if Coated
			iniwrite, %Coated%, Settings.ini, SavedVariables, Coated
		else
			iniwrite, %Null%, Settings.ini, SavedVariables, Coated
		if AllProducts
		{
			AllProductsTrimmed:=Trim(strreplace(strreplace(StrReplace(strreplace(AllProducts,"`n"," ")"`r"," "),"`t"," "),"`;"," "))
			iniwrite, %AllProductsTrimmed%, Settings.ini, SavedVariables, GeneralBox
		}
		if AllBatches
		{
			AllBatchesTrimmed:=Trim(strreplace(strreplace(strreplace(strreplace(AllBatches,"`n"," ")"`r"," "),"`t"," "),"`;"," "))
			iniwrite, %AllBatchesTrimmed%, Settings.ini, SavedVariables, GeneralBox2
		}
		; else
			; iniwrite, %Null%, Settings.ini, SavedVariables, GeneralBox
		if SampleGUID
			iniwrite, %SampleGUID%, Settings.ini, SavedVariables, SampleGUID
		; if Iteration
			iniwrite, %Iteration%, Settings.ini, SavedVariables, Iteration
		; else
			; iniwrite, %Null%, Settings.ini, SavedVariables, SampleGUID
		; if CustomerPosition
	return
	}
	loadSavedVariables(){ ;;___________________________LOADING VARIABLES_________________________
		global
			GUI, ClipBar:default
			Iniread, Iteration, Settings.ini, SavedVariables, Iteration
			iniRead, Product, Settings.ini, SavedVariables, Product
			iniRead, Batch, Settings.ini, SavedVariables, Batch
			iniRead, Lot, Settings.ini, SavedVariables, Lot
			iniRead, Coated, Settings.ini, SavedVariables, Coated
			iniRead, AllProducts, Settings.ini, SavedVariables, ALLProducts
			iniRead, AllBatches, Settings.ini, SavedVariables, AllBatches

			iniRead, SampleGUID, Settings.ini, SavedVariables, SampleGUID
		}
	}
			return

		WM_LBUTTONDOWN(wParam, lParam){
			If !MouseIsOver("ClipBar")
			return
			PostMessage, 0xA1, 2
					X := lParam & 0xFFFF
					Y := lParam >> 16
					if A_GuiControl
						ctrl := "`n(in control " . A_GuiControl . ")"
					PostMessage, 0xA1, 2
					MouseGetPos,,,,winControl
					return
		}
		WM_LBUTTONDBLCLK(wParam, lParam){
			If !MouseIsOver("ClipBar")
			return
			; PostMessage, 0xA1, 2
					X := lParam & 0xFFFF
					Y := lParam >> 16
					if A_GuiControl
						ctrl := "`n(in control " . A_GuiControl . ")"
					; PostMessage, 0xA1, 2
					MouseGetPos,,,,winControl
					; return
					ControlGetFocus,winControl,ClipBar
				if (winControl = "Edit2") {
					ControlGetText, clipboard, Edit2, ClipBar
				} else if (winControl = "Edit1") {
					ControlGetText, clipboard, Edit1, ClipBar
				} else if (winControl = "Edit3") {
					ControlGetText, clipboard, Edit3, ClipBar
				} else if (winControl = "Edit4") {
					ControlGetText, clipboard, Edit4, ClipBar
				} else if (winControl = "Edit6") {
					ControlGetText, clipboard, Edit6, ClipBar
				} else if (winControl = "Edit7") {
					ControlGetText, clipboard, Edit7, ClipBar
				}
				TrayTip, Copied to Clipboard, %clipboard%, 2
				; TT(Clipboard)
				return
		}






	ClipBar_ResetSub:
		ClipBar.Reset()
		return




	BlockTheWheel:
	; Critical
		sleep 1000
		breaking.point(1)
		; BlockingWheel:=
		; #maxthreadsperhotkey, 2
		return




		ArrayContains(array, value)
		{
				for index, element in array
				{
						if (element = value)
								return true
				}
				return false
		}
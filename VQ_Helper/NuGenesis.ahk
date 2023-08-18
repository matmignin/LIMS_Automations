












copyLabelCopyDocRegex(){
	Global
	RegexIngredients:="is)(%.Daily Value|Amount per serving)(\s+\n?)(?P<Ingredients>\s?[\w.].*?)([\*\s]*?\n\s)?(\*? ?)?(Daily Value|Percent|Other ingredients)"
	RegexIngredient:="Usm)(?P<Ingredient>(.*\n)|(.*))\s+?(?P<claim>[0-9.,]*) ?(?P<unit>m?c?g RAE|m?c?g\b|IU\b|Billions)+?.*$"
	listofIngredients:=
	RegexServingSize:= "im)^\s?Serving size:?\s+(?P<ServingSize>\d+)\s?(?P<ServingType>[\w ]+)?\s?(?P<ServingWeigth>\(.*\))?"
	ServingSize:=
	RegexPillSize:="i)(?P<PillSize>(?<=size: )#[0{2})[ \w]*)(?:(?:.*\d, \d{4}\.)"
	PillSize:=
	regingredient:=[]

FilePattern := "\\netapp\Label Copy Final\L000-L999\*" product "*.docx"
Loop, %FilePattern%, 1, 0
		oW:=ComObjGet(A_LoopFileLongPath)
		sleep 2000
		; oW.Visible :=0
		oW.Range.FormattedText.Copy
		; oW.Close()
		LabelCopyText:=Clipboard
	Ingredients:= RegexMatch(LabelCopyText, RegexIngredients,ri)
	RegexMatch(LabelCopyText, RegexServingSize, ss)
	RegexMatch(LabelCopyText, RegexPillSize, ps)
	ServingSize:=Trim(ssServingSize " " ssServingType " " SsservingWeight)
	PillSize:=Trim(psPillSize)
	Loop, Parse, riIngredients,`n
	{
		if RegexMatch(A_LoopField,"i)(Total Carbohydrate|Added Sugar|Total Sugar|Calories|Cholesterol|Sodium| Fat|Dietary Fiber|folic acid)",nogo)
			Continue
		NewString:=RegexReplace(A_LoopField, RegexIngredient, "${Ingredient}`t${claim} ${unit}`n")
	 	 listofIngredientsPreTrim:=trim(NewString) "`r`n||"
	listofIngredientsPretrim2:=Trim(strreplace(strReplace(listofIngredientsPreTrim, "`r`n||",""),"†",""))
	 	 listofIngredients.=Trim(strReplace(listofIngredientsPreTrim2, "`r",""))
		;  regingredient.insert(listofIngredientsPreTrim2)
			}
		FileDelete, U:\VQ_Helper\LabelCopyText.txt
			sleep 400
	FileAppend,  %listofIngredients%, U:\VQ_Helper\LabelCopyText.txt
	sleep 500
		Clipboard:=listofIngredients
		tt(listofIngredients,2000,10,10,3,190)
	Return
}



;;----------------------------------------------------------
;;[[                  General LMS                         ]]
Class LMS {

	Menu(){
		Global
		try Menu,Menu, deleteAll
		MouseGetPos, getX, getY, getWin
		WingetTitle getTitle, A
		If WinExist("Spec Table ahk_exe VQ_Helper.exe") && !GetKeyState("Shift","P"){
			SpecTab.ShowSpecMenu()
			return
		}
		Menu, Menu, add, &Products from clipboard, PasteAllProducts
		Menu, Menu, add, &Batches from clipboard, PasteAllBatches
		Menu, Menu, add,
		Menu, Menu, add, &Final Label Copy`t[%Product%], ShowFinalLabelCopy
		Menu, Menu, add, &Scan Label Copy,  ShowScanLabelCopy
		Menu, Menu, add, &GLOBAL VISION folder, ShowGlobalVision
		Menu, Menu, add, &mfg folder, Showmfg
		Menu, Menu, add, &Total CoAs, ShowFINAL_C_O_A
		; Menu, Menu, add, Manual &COAs folder, ShowManualCOA
		Menu, Menu, add,
		Menu, Menu, add, mmignin, mmigninFolder
		Menu, Menu, add,

		Menu, Menu, add, Window Info `t%getTitle%, get_window_info
		Menu, Menu, add, Mouse Info `t%getX%`, %GetY%, get_Mouse_info
		if GetKeyState("Shift","P"){
			Try Menu,menu,show
			return
		}
		; Try Menu,menu,show
		if winactive("Edit ingredient")
			productTab.IngredientMenu()
		if winactive("Edit sample")
			worktab.CustomerMenu()
		if winactive("Results Definition") || winactive("Results ahk_exe eln.exe"){

			; This.add("&Spec Table","Tests")
			try menu, menu, deleteAll
			Menu, Menu, Add, 100k TPC, Autofill
			Menu, Menu, Add, 3k TPC, Autofill
			Menu, Menu, Add, P. aeruginosa, Autofill
			Menu, Menu, add
			Menu, Menu, Add, USP Heavy Metal,Autofill
			Menu, Menu, Add, Canada Heavy Metal,Autofill
			Menu, Menu, Add, Prop65 Heavy Metal,Autofill
			Menu, Menu, Add, Report Only Heavy Metal,Autofill
			Menu, Menu, Add, Bloom Nutrition Heavy Metal,Autofill
			Menu, Menu, Add, Custom Heavy Metal,Autofill
			; return
		}
		If winactive("Result Editor"){
			try menu, menu, deleteAll
			Menu, Menu, Add, 100k TPC, Autofill
			Menu, Menu, Add, 3k TPC, Autofill
			Menu, Menu, Add, P. aeruginosa, Autofill
		}
		Try Menu,menu,show
		; else
		return
	}
	AddDataFromClipboard(Pointer:=">>|",Source:=""){
		global
		if !Source
			Source:=Clipboard
		sleep 100
		if RegexMatch(Source, Pointer){
			MethodList:=
			Name:=			[]
			IngredientID:=	[]
			Position:=		[]
			LabelClaim:=	[]
			MinLimit:=		[]
			MaxLimit:=		[]
			Units:=			[]
			Percision:=		[]
			LabelName:=		[]
			Description:=	[]
			Assay:=			[]
			Requirement:=	[]
			Method:= 		[]
			Table_height=0
			Skipped_row:=
			loop, parse, Source, "`n"
			{
				Line:=A_index
				Ingredient:=[]
				ingredient:=StrSplit(A_LoopField,"|")
				Name[Line]:=ingredient[2]
				IngredientID[Line]:=ingredient[3]
				Position[Line]:=ingredient[4]
				LabelName[Line]:=ingredient[5]
				labelClaim[Line]:=ingredient[6]
				if !ingredient[7] 					;If there are spec tabs
					continue
				Assay[Line]:=ingredient[7]
				Method[Line]:=ingredient[8]
				If !instr(MethodList, ingredient[8] "`n")
					MethodList.= ingredient[8]  "`n"
				Description[Line]:=ingredient[9]
				MinLimit[Line]:=ingredient[10]
				MaxLimit[Line]:=ingredient[11]
				Units[Line]:=ingredient[12]
				Percision[Line]:=ingredient[13]
				Requirement[Line]:=ingredient[14]
				Table_height+=1
			}
			FileDelete, MethodList.txt
			sleep 200
			FileAppend, %MethodList%, MethodList.txt
			;  Lms.detectTab()
			If winexist("Edit Ingredient"){
				winactivate, Edit Ingredient
				ProductTab.AddNewIngredient(LabelName[1],LabelClaim[1],Position[1],IngredientID[1],1)
				return
			}
			else if Winexist("Composition") || winactive("Edit Ingredient"){
				winactivate, Composition
				loop % Line
				{
					If !Position[A_index] && !(skipped_row)
					{
						Continue
						Skipped_row:=1
					}
					else If !Position[A_index] && (skipped_row = 1)
					{
						Skipped_row:=
						Break
					}
					else
						Skipped_row:=
					ProductTab.AddNewIngredient(LabelName[a_index],LabelClaim[a_index],Position[a_index],IngredientID[a_index])
					sleep 150
					ifwinnotactive, Composition
						winactivate, Composition
					sleep 175
				}
			}
			else if (winactive("NuGenesis LMS") || winactive("Results Definition") || winactive("Select methods tests")) && (Table_height > 1) {
				Try GUI, Spec_Table:destroy
				SpecTab.Table()
				return
			}
			else if winactive("Result Editor") || winactive("Results Definition") || winactive("Test Definition Editor") || (winactive("NuGenesis LMS") && (Table_height = 1)) {
				SpecData:=[]
				SpecData:=StrSplit(Source,"|")
				Name:=SpecData[2]
				IngredientID:=SpecData[3]
				Method:=SpecData[8]
				Description:=SpecData[9]
				MinLimit:=SpecData[10]
				MaxLimit:=SpecData[11]
				Units:=SpecData[12]
				Percision:=SpecData[13]
				Requirement:=SpecData[14]
				SpecTab.AutoFill()
				return
			}
		}
		else
			; msgbox, add data from clipboard failed
			clip.codesRegex()
		sleep 100
		return

	}



	AddSampleLog(count)
	{
		global
		; clk(mX,mY)
		;SetWinDelay, %NormalWinDelay%
		loop, %count%
		{
			click 46, 842
			Breaking.Point()
			if !winactive("Edit test (Field Configuration: ")
				winactivate
			Breaking.Point()
			Sendinput,{Click, 402, 284}{end}{down 2}{shiftdown}{9}{shiftup}on sample log{shiftdown}{0}{shiftup}{click, 334, 618}
			winwaitactive, NuGenesis LMS,,2
			winactivate, NuGenesis LMS
			Breaking.Point()
			Sendinput,{click, 1290, 703}{down %A_index%}
			Breaking.Point()
		}
		return
	}



	SearchBar(Code:="",PostCmd:="",Overwrite:="true"){
		Global
		; sleep 200
		If Nsub
		{
			sleep 400
			; exit
			return
		}
		Nsub:=1
		; #maxthreadsperhotkey, 1
		Tab:=LMS.DetectTab()
		GuiControl,ClipBar:Text, GeneralBox, %Tab%
		SetTimer, Block_Input, -2000

			sleep 200

		ifwinactive, Select methods tests
		{
			clk(246,77, 2)
		}
		Ifwinactive, Register new samples
			{
			Clk(180, 103, 2)
			Send, {click 180, 103,2}%Product%{enter}
			}
		ifwinactive, NuGenesis LMS
		{
			if (Tab="Products") {
				; If (Code=Batch){
					; lmsclick.Edit_Composition()
					; return
				; }
				clk(x%Tab%Search+45 ,yProductsSearch, "Left")
				clk(x%Tab%Search ,yProductsSearch, "Left", 2)
				; sleep 30
				; clk(x%Tab%Search,yProductsSearch, "Left", 2)
				; clk(x%Tab%Search+35,yProductsSearch, "Left")
				; clk(x%Tab%Search,yProductsSearch)
				Send, ^{a}
				If Overwrite=true
					Send, ^{x}
				If Code
					Sendinput, %Product%^{a}
				If Overwrite=true
					send, {right}{space}^{v}^{a}^{c}
				if PostCmd!=""
					{
					sleep 300
					send % PostCmd
				}
					return
			}
			else if (Tab="Specs") {
				; If (Code=Batch) {
				; 	lmsclick.Edit_Test()
				; 	return
				; }
				; xtabsLocation:=x%Tab%Search
				xtabsLocation2:=x%Tab%Search+10
				xtabsLocation3:=x%Tab%Search+25
				xtabsLocation4:=x%Tab%Search+35
				xtabsLocation5:=x%Tab%Search+40
				; xtabsLocation6:=x%Tab%Search+50
				; ytabsLocation:=%yWorkTabSearch%
				mouseclick,, %xtabslocation5%, %yProductsSearch%
				mouseclick,, %xtabslocation4%, %yProductsSearch%
				mouseclick,, %xtabslocation3%, %yProductsSearch%
				mouseclick,, %xtabslocation2%, %yProductsSearch%
				; mouseclick,, %xtabslocation%, %yProductsSearch%
				; mouseclick,, %xtabslocation5%, %yProductsSearch%
				; mouseclick,, %xtabslocation6%, %yProductsSearch%
				; clk(x%Tab%Search+20,yProductsSearch,,2,"NuGenesis LMS",0)
				; clk(x%Tab%Search+5,yProductsSearch,,2)
				Send, {ctrldown}{a}{ctrlup}
				If Overwrite=Add
					Send, ^{x}
				if Code
					Sendinput, %Product%^{a}
				If Overwrite=Add
					send, {right}{space}^{v}^{a}^{c}
				if PostCmd!=""
					{
					sleep 300
					send % PostCmd
				}
				sleep 500
					return
			}
			else If (Tab="Requests") {
				; clk(x%Tab%Search,yWorkTabSearch,,1,"NuGenesis LMS",0)
				; clk(x%Tab%Search-9,yWorkTabSearch)
				xtabsLocation:=x%Tab%Search
				xtabsLocation2:=x%Tab%Search-9
				xtabsLocation3:=x%Tab%Search-20
				; ytabsLocation:=%yWorkTabSearch%
				mouseclick,, %xtabslocation%, %yWorkTabSearch%
				mouseclick,, %xtabslocation2%, %yWorkTabSearch%
				mouseclick,, %xtabslocation3%, %yWorkTabSearch%
				Send, {ctrldown}{a}{ctrlup}
				If Overwrite=Add
					Send, ^{x}
				Send, %Code%{ctrldown}{a}{ctrlup}
				If Overwrite=Add
					send, {right}{space}^{v}^{a}^{c}
				if PostCmd!=""
					{
					sleep 300
					send % PostCmd
				}
				sleep 500
					 return
			}
			else If (Tab="Tests"|| Tab="Samples" || Tab="Results") {
				clk(x%Tab%Search,yWorkTabSearch,,1,,0)
				clk(x%Tab%Search+10,yWorkTabSearch,,1,,0)
				clk(x%Tab%Search+20,yWorkTabSearch,,2)
				Send, {ctrldown}{a}{ctrlup}
				If Overwrite=Add
					Send, ^{x}
				Send, %Code%{ctrldown}{a}{ctrlup}
				If Overwrite=Add
					send, {right}{space}^{v}^{a}^{c}
				if PostCmd!=""
					{
					sleep 300
					send % PostCmd
				}
				sleep 500
					return
			}
			else
				send, %Code%
		}
		; sleep 4000
		; #maxthreadsperhotkey, 2
		Nsub:=
		return
	}

	SaveCode(){
		global
		clipboard:=
		; Simpleclip:=1
		send, ^{a}^{c}
		sleep 50
		clipwait,2
		if errorlevel {
			sendinput, ^{a}^{c}
			sleep 500
		}
		else
		Send, {enter}
		; simpleclip:=
		return
	}
	SearchbarPaste(Delimiter){
		Clipboard := StrReplace(Clipboard, "`r`n", Delimiter)
		Clipboard := StrReplace(Clipboard, "`n", Delimiter)
		Clipboard := StrReplace(Clipboard, ", ", Delimiter)
		Clipboard := StrReplace(Clipboard, ",", Delimiter)
		Clipboard := StrReplace(Clipboard, "|", "")
		Clipboard := StrReplace(Clipboard, "/", "")
		Clipboard := StrReplace(Clipboard, A_tab, Delimiter)
		Clipboard := StrReplace(Clipboard, A_space, Delimiter)
		Clipboard := StrReplace(Clipboard, Delimiter Delimiter, Delimiter)
		sleep 200
		send, ^v
		sleep 100
		return clipboard
	}

	DetectTab(){
		global
		tab:=
		FoundSamples:=
		FoundRequests:=
		FoundDocuments:=
		FoundResults:=
		FoundTests:=
		FoundSpecs:=
		FoundReagents:=
		Tab1:=
		TAB2:=
		TAB3:=
		TAB4:=
		TAB5:=
		TAB6:=
		TAB7:=
sleep 20
		LMS.Orient()
		; CoordMode, pixel, window
		if winactive("NuGenesis LMS") {
			PIXELSEARCH, Tab2, FoundY, XTAB2, YTabS, XTAB2+3, yTabs+5, 0xfff8c3, 10, Fast RGB ;icon on
			if !Tab1 {
				PixelSearch, FoundSamples, FoundY, SamplesTab, yWorkTabs, SamplesTab+2, yWorkTabs+2, 0xfffd353, 10, Fast RGB
				PixelSearch, FoundRequests, FoundY, RequestsTab, yWorkTabs, RequestsTab+2, yWorkTabs+2, 0xffd353, 10, Fast RGB
				if FoundSamples {
					Tab:="Samples"
					return Tab
				}
				else
				{
					If FoundRequests {
						Tab:="Requests"
						return Tab
					}

					else if Errorlevel
					{

						PIXELSEARCH, FoundSpecs, FoundY, 14, 351, 16, 353, 0x0D77AF, 10, Fast RGB ;icon on
						; PIXELSEARCH, FoundSpecs, FoundY, 13, 355, 15, 358, 0xeaeff3, 10, Fast RGB ;icon on
						If FoundSpecs
							Tab:="Products"
						else
							tab:="Specs"
						return Tab
						; }
					}

				}
			}
		}
		return Tab
	}




	Orient(){
		global
		winGetPos,Nux,NuY,NuW,NuH, NuGenesis LMS
		WbX:=WbX+400

		DocumentMenuSection_X:=NuW-200
		DocumentMenuDocument_X:=NuW-140
		DocumentMenu_Y:=110
		yTabSelect:=45
		yTabDropdown:=45
		SamplesTab:=(NuW/2)-80
		RequestsTab:=(NuW/2)+20
		DocumentsTab:=(NuW/3)+(NuW/3)-50
		TestsTab:=(NuW/3)+(NuW/3)-220
		ResultsTab:=(NuW/3)+(NuW/3)-150
		HScrollBarRightX:=NuW-40
		HScrollBarLeftX:=(NuW/5)+35
		HScrollBarRightY:=HScrollBarLeftY:=(Nuh/2)+38
		yWorkTabs:=71

		xDivider:=(NuW/5)
		xTab1=150
		xTab2=350
		MyWorkTab=350
		xTab3=550
		ProductsTab=550
		xTab4=750
		SpecsTab=750
		xTab5=950
		xTab6=1150
		xTab7=1550
		xTab8=358+1000
		xTab9=358+1200
		xTab10=358+1400
		yTabs:=36
		xWorkTab:=334, 47 ;1st
		yWorkTabSearch:=128
		XCoA:=(NuW-131)
		xClearfilter:=xDivider+16
		yClearfilter:=270
		xFilterIcon:=NuW-22
		yFilterIcon:=131

		xProductsSearch:=xDivider+180
		xSpecsSearch:=xDivider+183
		yProductsSearch:=93
		ySpecsSearch:=93
		xRequestsSearch:=xDivider+190 ;175 ;103
		xRequestsSearchDefault:=xDivider+170 ;155 ;103
		xSamplesSearch:=xDivider+145
		xResultsSearch:=xDivider+185
		xResultsSearch:=xDivider+185
		xTestsSearch:=xDivider+125
		xDocumentsSearch:=xDivider+25
		return
	}

}

	Heavy_metals:
		if (A_ThisMenuItem = "USP Heavy Metal")
		SpecTab.HM_USP()
		else if (A_ThisMenuItem = "Canada Heavy Metal")
		SpecTab.HM_Canada()
		else if (A_ThisMenuItem = "Prop65 Heavy Metal")
		SpecTab.HM_Prop65()
		else if (A_ThisMenuItem = "Report Only Heavy Metal")
		SpecTab.HM_ReportOnly()
		else if (A_ThisMenuItem = "Bloom Nutrition Heavy Metal")
		SpecTab.HM_Bloom()
		else if (A_ThisMenuItem = "Custom Heavy Metal")
		SpecTab.HM_Custom()
		else
			Menu,Menu, deleteAll
	return
	PasteAllBatches:
		Clk(mx,mY)
		sendinput % Trim(GetAllBatches())
	return
	PasteAllProducts:
		Clk(mx,mY)
		sendinput % Trim(GetAllProducts())
	return
	Autofill:
		winactivate, ahk_exe eln.exe
		sleep 200
		if A_thismenuitem contains &Analytical
		SpecTab.Edit_Analytical()
		else if A_thismenuitem contains Copy &Specs
		SpecTab.CopySpecs()
		else if A_thismenuitem contains Paste &Specs
		SpecTab.PasteSpecs()
		else if A_thismenuitem contains Copy &Template
		SpecTab.CopySpecTemplate()
		else if A_thismenuitem contains New &Request
		WorkTab.NewRequest()
		else if (A_ThisMenuItem = "100k TPC"){
			if winactive("Results Definition") ;Selection window
			{
				winactivate, Results Definition
				Send,{click 80, 66} ;click edit
				Breaking.Point()
			}
			Breaking.Point()
			winwaitactive, Result Editor,,4
			SpecTab.ChangeMicroLimit("100,000")
			return
		}
		else if (A_ThisMenuItem = "3k TPC"){
			if winactive("Results Definition") ;Selection window
			{
				winactivate, Results Definition
				Send,{click 80, 66} ;click edit
				Breaking.Point()
			}
			Breaking.Point()
			winwaitactive, Result Editor,,4
			SpecTab.ChangeMicroLimit("3,000", 1)
			; Mouseclick, left, 378, 667,1,0  ;click okay in result editor
			winwaitactive, Results Definition,,4
			My+=26
			click, 445, 164
			Send,{click 80, 66} ;click edit
			winwaitactive, Result Editor,,4
			SpecTab.ChangeMicroLimit("300")
			; Mouseclick, left, 378, 667,1,0  ;click okay in result editor
			return
		}
	else if (A_ThisMenuItem = "P. aeruginosa")
	SpecTab.AddPaeruginosa()
	else if (A_ThisMenuItem = "USP Heavy Metal")
	SpecTab.HM_USP()
	else if (A_ThisMenuItem = "Canada Heavy Metal")
	SpecTab.HM_Canada()
	else if (A_ThisMenuItem = "Prop65 Heavy Metal")
	SpecTab.HM_Prop65()
	else if (A_ThisMenuItem = "Report Only Heavy Metal")
	SpecTab.HM_ReportOnly()
	else if (A_ThisMenuItem = "Bloom Nutrition Heavy Metal")
	SpecTab.HM_Bloom()
	else if (A_ThisMenuItem = "Custom Heavy Metal")
	SpecTab.HM_Custom()
	else if (A_ThisMenuItem = "Delete Retain")
	WorkTab.DeleteRetain()
	Menu,Menu, deleteAll
	return




;;----------------------------------------------------------
;;{{                ProductTab Class                         }}
Class ProductTab {


	/*
	[2] Product
	[3] ProductName
	[4] Customer
	[5] CustomerPosition
	[6] ShapeAndSize
	[7] Color
	[8] ServingSize
	[9] clip.codesRegex
	*/
	AddProductFromClipboard(Input:=""){
		Global
		Haystack:= Input ? Input : Clipboard
		; loop, parse, clipboard, "`n"
		; {
		SheetInfo:=[]
		Coated:=
		Lot:=
		SheetInfo:=StrSplit(HayStack,"|")
		Product:=SheetInfo[2]
		ProductName:=SheetInfo[3]
		Customer:=SheetInfo[4]
		sleep 200
		; CustomerPosition:=SheetInfo[5]
		if !GetIniValue("Customer.ini",Customer)
			CustomerPosition:=SheetInfo[5]
		else
			CustomerPosition:=GetIniValue("Customer.ini",Customer)

		ShapeAndSize:=SheetInfo[6]
		Color:=SheetInfo[7]
		ServingSize:=SheetInfo[8]
		if Instr(HayStack, "ScanLabel >>",true,1,1){
			GoSub ShowScanLabelCopy
			return
		}
		else if Instr(HayStack, "FinalLabel >>",true,1,1){
			GoSub ShowFinalLabelCopy
			return
		}
		else if Instr(HayStack, "GlobalVision >>",true,1,1){
			GoSub ShowGlobalVision
			return
		}
		else if Instr(HayStack, "mfg >>",true,1,1){
			GoSub showmfg
			return
		}
		else
			clip.codesRegex(SheetInfo[9])
		ProductInfo:=ProductName " : " Customer
		ServingInfo:=ServingSize " " Color
		ControlsetText, Edit6,%ProductInfo% ,ClipBar
		ControlsetText, Edit7,%ServingInfo% ,ClipBar
		; ControlsetText, Static1,%CustomerPosition%,ClipBar
		; GuiControl,ClipBar:Text, Iteration, %Iteration%
		; }
		if CustomerPosition {
			Iteration:=StrReplace(strReplace(CustomerPosition,"[[",""),"]]","")
			ControlsetText, Edit5,%Iteration%,ClipBar
			IniWrite, %Iteration%, Settings.ini, SavedVariables, Iteration
		}



		return
	}


	SaveIngredientList(){
		Clipboard:=Trim(strReplace(StrReplace(Clipboard, "`r`n","`n"),"`t","")) ; Normalize line endings
		lines := StrSplit(Clipboard, "`n") ; Split the string into lines
		totalLines := lines.Length()-1
		half := FLOOR(totalLines/2)
		halfAmount:=-(Half)
		Try FileDelete, Ingredients.ini
		IngredientsString:=
		Loop, % totalLines
		{
			if (A_Index=1){
				IngredientsString:="[Ingredients]"
				continue
			}
			line := lines[A_Index]
			if (A_Index <= half)
				lineNumber := A_Index - 2
			else
			{
				HalfAmount+=1
				If (HalfAmount=0)
					Linenumber:="-0"
				else
					Linenumber:=halfAmount
			}
			IngredientsString.="`n"Line "=" lineNumber
			; IngredientName:= StrReplace(Line, "Allergen","")
		}
		sleep 300
		FileAppend, % Trim(strReplace(IngredientsString," Allergen","")), Ingredients.ini
		return IngredientsString
	}


	AddNewIngredient(Ingredient_Name,Ingredient_Claim,Ingredient_Position,Ingredient_Id,Dont_Hit_Okay:=""){
		Global
		ifwinnotexist, Edit Ingredient
		{
			if winactive("Composition"){
				winactivate, Composition
				click 57, 65
			}
		}
		if Dont_Hit_Okay
			click, 150, 73
		else
			ProductTab.Dropdown_IngredientSelect(Ingredient_Id)
		Send,{tab 6}^{a}
		sendinput % TRIM(Ingredient_position," `r`n")
		Sendinput,{tab}^a%Ingredient_Name%
		if Strlen(Ingredient_Name) > 250
			sleep 500
		If Ingredient_Claim contains Heavy Metal,Allergens,Other
			Send,{tab}
		Send,{tab 2}^a
		Send,%Ingredient_Claim%
		If InStr(Ingredient_Name, "* Heavy Metals results are based on a daily dose of "){
			 Send, +{tab 3}
			if ServingSize
			ServingSizeMenu(ServingSize)
			else
				ServingSizeMenu()
			Return
		}
		Breaking.Point()
		If !Dont_Hit_Okay
			Sendinput,{enter}
		Breaking.Point()
		ifwinexist, Duplicate ingredient ID
			exit
		winwaitclose, Edit Ingredient,,3
		if errorlevel
			exit
		return
	}

	Dropdown_IngredientSelect(A_DropdownCount){
		global
		; SetKeyDelay,-1,-1
		ifWinNotActive, Edit Ingredient
			winactivate, Edit Ingredient
		click, 150, 73 ;click dropdown box
		; if HitTab
		; 	Send, {tab}««
		sleep 110
		AbsSelection:=Abs(A_DropdownCount)
		if (A_DropdownCount > 0){
			Sendinput, {tab}{home}{right %A_DropdownCount%}
		}
		if (A_DropdownCount < 0){
			Sendinput, {tab}{end}{left %AbsSelection%}
		}
		if (A_DropdownCount = "-0")
			Sendinput, {tab}{end}
		if (A_DropdownCount = "i"){
			if !(IngredientNoteCount) || (IngredientNoteCount >8)
				IngredientNoteCount:=1
			this.Dropdown_GenericIngredient(IngredientNoteCount,1)
			IngredientNoteCount+=1
		}
		if (A_DropdownCount = ""){
			if Iteration > 25
				Iteration:=1
			this.Dropdown_GenericIngredient(Iteration)

			ClipBar.AddIteration(0,0,"True")
		}
		return
	}




	AddNewProduct(){ ;for naming Product code and customer,
		global Product, ProductName, Customer, ShapeAndSize, color, Iteration

		click 120,80 ;click product box
		Sendinput,%Product%`,{space}
		sendraw, %ProductName%
		sleep 20
		sendinput, {tab 2}%Customer%{tab 2}{End}{tab}{right 3}{tab}%Product%{tab 2} ; new vitaquest selection
		sleep 200
		Sendraw,%ProductName%
		sendinput, {tab 8}
		Breaking.Point()
		winwaitactive, NuGenesis LMS,,8
		if errorlevel
			Exit
		; tt(you broke it,2000)
		winactivate, NuGenesis LMS
		Breaking.Point()
		click, 67, 283
		Breaking.Point()
		This.AddNewFormulation()
		; clk(287, 578) ;click save
		Iteration:=1
		return

	}

	AddNewFormulation(){     ;then click on Edit Formulation, puts in code, then tabs to serving size
		global Product, ShapeAndSize, color, ServingSize, ServingType
		if !winactive("Edit Formulation") && winexist("Edit Formulation")
			winactivate, Edit Formulation,
		Sendinput, {tab}%product%
		Sendinput, {Tab 23}^a ;{click 268, 578}
		; sendinput, {tab}^a


		 If InStr(ServingSize,"1 t",False) || InStr(ServingSize,"1t",False) || InStr(ServingSize,"1 Chewable",False)
			Sendinput % "Each (1) tablet contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"2 t",False) || InStr(ServingSize,"2t",False)  InStr(ServingSize,"2 Chewable",False)
			Sendinput % "Each two (2) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"1 c",False) || InStr(ServingSize,"1c ",False)
			Sendinput % "Each (1) capsule contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"2 c",False) || InStr(ServingSize,"2c",False)
			Sendinput % "Each two (2) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"3 c",False) || InStr(ServingSize,"3c",False)
			Sendinput % "Each three (3) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"4 c",False)
			Sendinput % "Each four (4) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"5 c",False)
			Sendinput % "Each five (5) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"6 c",False)
			Sendinput % "Each six (6) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"7 c",False)
			Sendinput % "Each seven (7) capsules contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"3 t",False) || InStr(ServingSize,"3t",False)
			Sendinput % "Each three (3) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"4 t",False)
			Sendinput % "Each four (4) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"5 t",False)
			Sendinput % "Each five (5) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"6 t",False)
			Sendinput % "Each six (6) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"7 t",False)
			Sendinput % "Each seven (7) tablets contains{Tab}" ShapeAndSize "{tab}" color "+{tab 2}"
		else if InStr(ServingSize,"1 sp",False)
			Sendinput % "Each (1) stick packet (" ShapeAndSize " g) contains{Tab}Blend"
		else if InStr(ServingSize,"2 sp",False)
			Sendinput % "Each two (2) stick packet (" ShapeAndSize " g) contains{Tab}Blend"
		else if InStr(ServingSize,"1 stick",False) || InStr(ServingSize,"1st",False)
			Sendinput % "Each (1)" SubStr(ServingSize, 2) " contains{Tab}Blend"
		else if InStr(ServingSize,"2 stick",False)
			Sendinput % "Each two (2)" SubStr(ServingSize, 2) " contains{Tab}Blend"
		else if InStr(ServingSize,"1 scoop",False) || InStr(ServingSize,"1s",False)
			Sendinput % "Each (1)" SubStr(ServingSize, 2) " contains{Tab}Blend"
		else if InStr(ServingSize,"2 scoop",False)  || InStr(ServingSize,"2s",False)
			Sendinput % "Each two (2)" SubStr(ServingSize, 2) " contains{Tab}Blend"
		else if InStr(ServingSize,"3 scoop",False)
			Sendinput % "Each three (3)" SubStr(ServingSize, 2) " contains{Tab}Blend"
		else if InStr(ServingSize,"4 scoop",False)
			Sendinput % "Each four (4)" SubStr(ServingSize, 2) "  contains{Tab}Blend"
		else if InStr(ServingSize,"5 scoop",False)
			Sendinput % "Each five (5)" SubStr(ServingSize, 2) "  contains{Tab}Blend"
		else if InStr(ServingSize,"1 s",False)
			Sendinput % "Each (1) scoop (" ShapeAndSize " g) contains{Tab}Blend"
		else if InStr(ServingSize,"2 s",False)
			Sendinput % "Each two (2) scoops (" ShapeAndSize " g) contains{Tab}Blend"
		else if InStr(ServingSize,"3 s",False)
			Sendinput % "Each three (3) scoops (" ShapeAndSize " g) contains{Tab}Blend"
		else if InStr(ServingSize,"4 s",False)
			Sendinput % "Each four (4) scoosp ("  ShapeAndSize " g)  contains{Tab}Blend"
		else if InStr(ServingSize,"5 s",False)
			Sendinput % "Each five (5) scoops (" ShapeAndSize " g)  contains{Tab}Blend"
		else if InStr(ServingSize,"",False)
			Sendinput {tab}

		If ShapeAndSize && !ServingSize
			sendraw, %ShapeAndSize%
		sendinput {Tab}
		; sendinput, {shiftdown}{tab}{shiftup}
		If Color && !!ServingSize
			sendinput, ^{a}%Color%
		sendinput, {shiftdown}{tab 2}{shiftup}
		; else
		Color:=
		ShapeSize:=
		Breaking.Point()
		sleep 900
		; clk(287, 578) ;click save
		return

	}

	HM_ReportOnly(){

		click 125,120 ;click 1st row

		clk(45, 65)
		winwaitactive, Edit Ingredient,,4
		Sendinput,{click 150,73}{tab}{right 11} ;arsenic
		Breaking.Point()
		winWaitClose, Edit Ingredient,,4
		click 125,140 ;click 2nd row
		clk(45, 65)
		winwaitactive, Edit Ingredient,,4
		Sendinput,{click 150,73}{tab}{right 167} ;lead
		Breaking.Point()
		click 390, 659	;click okay
		winWaitClose, Edit Ingredient,,4
		click 125,180 ;click 3rd row
		clk(45, 65)
		winwaitactive, Edit Ingredient,,4
		Sendinput,{click 150,73}{tab}{right 23} ;cadmium
		Breaking.Point()
		click 390, 659	;click okay
		winWaitClose, Edit Ingredient,,4
		click 125,200 ;click 4th row
		clk(45, 65)
		winwaitactive, Edit Ingredient,,4
		Sendinput,{click 150,73}{tab}{right 189} ;mercury
		Breaking.Point()

		click 390, 659	;click okay
		return
	}

	Dropdown_GenericIngredient(IterationCount:="",IngredientNote:=""){ ;; Generic List
		global
		GeneralCount:=IterationCount
		Click 150, 73
		sleep 50
		Sendinput,{tab}{Home}{right %Ingredient_List_Adjustment%}
		sleep 80
		if IngredientNote
		{
			Sendinput, {right %IngredientNoteDropDownCount%}{right %IterationCount%}
			sleep 405
			return
		}
		if GeneralCount=1
			Sendinput, {right 56}
		else if GeneralCount=2
			Sendinput, {right 60}
		else if GeneralCount=3
			Sendinput, {right 64}
		else if GeneralCount=4
			Sendinput, {right 68}
		else if GeneralCount=5
			Sendinput, {right 72}
		else if GeneralCount=6
			Sendinput, {right 76}
		else if GeneralCount=7
			Sendinput, {right 80}
		else if GeneralCount=8
			Sendinput, {right 81}
		else if GeneralCount=9
			Sendinput, {right 82}
		else if GeneralCount=10
			Sendinput, {right 83}
		else if GeneralCount=11
			Sendinput, {right 84}
		else if GeneralCount=12
			Sendinput, {right 85}
		else if GeneralCount=13
			Sendinput, {right 86}
		else if GeneralCount=14
			Sendinput, {right 87}
		else if GeneralCount=15
			Sendinput, {right 88}
		else if GeneralCount=16
			Sendinput, {right 89}
		else if GeneralCount=17
			Sendinput, {right 90}
		else if GeneralCount=18
			Sendinput, {right 91}
		else if GeneralCount=19
			Sendinput, {right 92}
		else if GeneralCount=20
			Sendinput, {right 93}
		else if GeneralCount=21
			Sendinput, {right 94}
		else if GeneralCount=22
			Sendinput, {right 95}
		else if GeneralCount=23
			Sendinput, {right 96}
		else if GeneralCount=24
			Sendinput, {right 97}
		else if GeneralCount=25
			Sendinput, {right 98}
		else if GeneralCount=26
			Sendinput, {right 99}
		else
			exit
		sleep 350
		return
	}


	IngredientsMenu() { ;; create a dropdown from CustomerMenu ini datafile
		global
		Ingredientfile:="U:\VQ_Helper\Ingredients.ini"
		try menu, IngredientsMenu, DeleteAll

		Loop, Read, %Ingredientfile%
		{
			If A_Index = 1
				Continue
			Ingredientslist := StrSplit(A_LoopReadLine, "=")
			Selection:= % Ingredientslist[1]
			If A_Index = 60
				Addbreak:="+BarBreak"
			Menu, IngredientsMenu, add, &%Selection%, IngredientsMenuhdr ;, %Addbreak%
			AddBreak:=
		}
		Menu, IngredientsMenu, Show,
		return

		IngredientsMenuhdr:
		sleep 200
		menuselection:=strsplit(A_thismenuitem, "&")
		InputVar:=MenuSelection[2]
		IniRead,IngredientPosition, Ingredients.ini, Ingredients, %InputVar%
		sleep 110


		this.Dropdown_IngredientSelect(Ingredientposition)
		menu, IngredientsMenu, DeleteAll
		; Nsub:=

		return
	}

}
return





;;----------------------------------------------------------
;;{{                 SpecTab class                         }}
class SpecTab {

	Table(){
		Global
		ShiftTable_X:=-355
		ShiftTable_Y:=200
		Try GUI, Spec_Table:destroy
		CoordMode, mouse, window
		ifwinnotactive, ahk_exe eln.exe
			winactivate, ahk_exe eln.exe
		sleep 200
		winGetPos, LMS_X, LMS_Y, LMS_w, LMS_h, A
		SpecTable_X:=LMS_X+10
		; SpecTable_X:=LMS_w+LMS_X+ShiftTable_X
		; SpecTable_Y:=LMS_Y+ShiftTable_Y
		SpecTable_Y:=LMS_Y+ShiftTable_Y+600
		CoordMode, mouse, screen
		SpecTab.CreateGUI()
		SpecTab.ModifyColumns()
		OnMessage(0x0201, "WM_Lbuttondown")
		If WinExist("ahk_exe eln.exe")
			SpecTab.ShowGUI()
		else
			MsgBox, %A_ThisMenuItemPos%, %SpecMsg%
		CoordMode, mouse, window
		return
	}

	ShowGUI(){
		global
		CoordMode, mouse, screen
		if WinExist("Select methods test"){
			ScreenEdge_X:=A_ScreenWidth-450
			ScreenEdge_Y:=A_Screenheight-180
			SpecTable_X:=LMS_X-450
			SpecTable_Y:=LMS_Y+ShiftTable_Y
		}
		else if Winactive("NuGenesis LMS"){
			ScreenEdge_X:=A_ScreenWidth-15
			ScreenEdge_Y:=A_Screenheight-180
		}
		sleep 200
		If (Name.Maxindex() <=1)
			return
		try GUI, Spec_Table:Show, x%SpecTable_X% y%SpecTable_Y% w350, %Product% Spec Table
		catch GUI, Spec_Table:Show, x%ScreenEdge_X% y%ScreenEdge_Y% w350, %Product% Spec Table
			CoordMode, mouse, window
		OnMessage(0x0201, "WM_Lbuttondown")
		return
	}
	ShowSpecMenu(){
		global
		send,{click}
		if !ClippedData {
			try Menu, SpecMenu, DeleteAll
			Fileread, ClippedData, ClippedExcelData.txt
			Clipboard:=ClippedData
			sleep 200
		}
		MouseGetPos, Mx, My
		Try Menu,SpecMenu,show
		return
	}

	CreateGUI(){
		global
		;SpecArray:=[]
		Try GUI, Spec_Table:destroy
		GUI, Spec_Table:Default
		try Menu, SpecMenu, DeleteAll
		i:=1
		if Table_height > 8
			Table_height :=12
		if !Table_height
			Table_height := 8
		Gui Spec_Table:+LastFound +Toolwindow +Owner +AlwaysOnTop -SysMenu +MinimizeBox
		GUI, Spec_Table:Font, s11 cBlack, Arial Narrow
		; GUI, Spec_Table:Add, ListView, x2 y0 w358 r%table_height% Grid checked altSubmit -hdr gSpec_Table, `t%Product%|`t%Name%|MinLimit|MaxLimit|Units|Percision|Description|Method
		; GUI, Spec_Table:Add, ListView, x2 y0 w358 r%table_height% Grid checked altSubmit -hdr gSpec_Table, `t%Product%|`t%Name%|MinLimit|MaxLimit|Units|Percision|Description|Method
		GUI, Spec_Table:Add, ListView, x2 y0 w358 r%table_height% Grid checked altSubmit -hdr gSpec_Table, `t%Product%|`t%Name%|MinLimit|MaxLimit|Units|Percision|Description|Method
		Gui, Spec_Table:Add, Button, X25 y+0 h18 w145 gAddSpecTableMethods, Add Methods
		Gui, Spec_Table:Add, Button, X+3 h18 w145 gAutoAddSpecs, Auto Specs
		OnMessage(0x0201, "WM_Lbuttondown")
		loop % Name.Maxindex(){
			if !(Requirement[A_index])
				continue
			LV_add(,""Name[A_index], Requirement[A_index], MinLimit[A_index],MaxLimit[A_index],Units[A_index],Percision[A_index],Description[A_index],Method[A_index])
			Menu, SpecMenu, Add, % "&" Name[A_Index] "`t " Requirement[A_Index], SpecMenuButton
			Test:= Name[A_index]
		}
		return

		AutoAddSpecs:
		ContinueToRun:=1
			winactivate, NuGenesis LMS
			loop % spectab.CountUncheckedRows()
				Key[1].Activate()
			continueToRun:=
			return
		AddSpecTableMethods:
		winactivate, Select methods tests
		click, 235, 72
		Send, ^a
		sleep 200
		Loop, Parse, MethodList, `n ; loop through each selected method
		{
			MethodListItem:=A_loopfield
			if inStr(MethodListItem,"/")
			{
				loop, parse, MethodListItem, /
					menu, ChooseMethodMenu, Add, %A_loopField%, ChooseMethodHdr
				menu, ChooseMethodMenu, show
				; This.SelectMethod(A_thismenuitem
				MethodListItem:=A_thismenuitem
			}

			{
				winactivate, Select methods tests
				sleep 200
				click, 235, 72
				Send, ^a
				click, 235, 72 ;click search bar
				sleep 100
				Sendinput, %MethodListItem%{enter} ; enter method to search
				sleep 300
				click 506, 341 ;move over
			}
		}
		return

		ChooseMethodHdr:
		Try Menu, ChooseMethodMenu, DeleteAll
		winactivate, Select methods tests
		sleep 200
		return

		SpecMenuButton:
		if A_ThisMenuItemPos
		{
			GUI, Spec_Table:Default
			; LV_ModifyCol(1, "Checkbox")
			LV_Modify(A_ThisMenuItemPos, "Check")
			Mousemove, %Mx%, %My%, 0
			Menu, SpecMenu, Check, %A_ThisMenuItem%
			SpecTab.GetRowText(A_ThisMenuItemPos)
			winactivate, ahk_exe eln.exe
			SpecTab.Autofill()
		}
		return
		; CoordMode, mouse, screen
	}
	FindRowNumber(InputTestName:=""){
		global
		GUI, Spec_Table:Default
		{
			Loop, % LV_GetCount()
			{
				LV_GetText(rowText,A_Index,1)
				if (rowText = InputTestName)
				{
					Try Menu, SpecMenu, Check, %A_ThisMenuItem%
					LV_Modify(A_Index, "Check")
					return A_Index
				}
			}
		}
		return 0
	}
ClickEmptyRequirements(){
	winactivate, NuGenesis LMS
		ImageFile := "U:\VQ_Helper\images\dash.PNG"
		ImageFile2 := "U:\VQ_Helper\images\emptyrdpnores.JPG"
		WinGetPos, WinX, WinY, WinWidth, WinHeight, NuGenesis LMS
		thirdwinwidth:=winwidth/3
		thirdwinHeight:=winHeight/3
		ImageSearch, FoundX, FoundY, %thirdwinwidth%, %thirdwinHeight%, %WinWidth%, %WinHeight%, *Trans96B4CE *140 %ImageFile%
		If ErrorLevel
			ImageSearch, FoundX, FoundY, %thirdwinwidth%, %thirdwinHeight%, %WinWidth%, %WinHeight%,  *190 %ImageFile2%
		if !FoundX
		{
			msgbox, No image found
			return
		}
			SendLevel 1
			cX:=foundX +10
			cY:=foundy +10
			Click, %cX%, %cY%
			MouseClick,middle, % cX, % cY, 1
		sendlevel 0
		return
	}


	CountUncheckedRows(){
		GUI, Spec_Table:Default
			; Loop, % LV_GetCount() {
				CheckedCount:=i:=0
        while i:=LV_GetNext(i, "Checked"){
					CheckedCount++
				}
			; }
			Total_Unchecked:=LV_GetCount() - CheckedCount
			GuiControl, ClipBar:Text, Iteration, %Total_Unchecked%
			return Total_Unchecked
		}



	ModifyColumns(){
		Global
		GUI, Spec_Table:Default
		LV_ModifyCol(1,130)
		LV_ModifyCol(2,120)
		LV_ModifyCol(3,0)
		LV_ModifyCol(4,0)
		LV_ModifyCol(5,0)
		LV_ModifyCol(6,0)
		LV_ModifyCol(7,0)
		LV_ModifyCol(8,90)
		LV_ModifyCol(9,0)
		LV_ModifyCol(2,"Sort")
		LV_ModifyCol(8,"Sort")
		; LV_Delete(Table_Height)
	}

	GetRowText(row:=""){
		global
		If !Row
			row:=A_EventInfo
		GUI, Spec_Table:Default
		LV_GetText(Name, 			row,1)
		LV_GetText(LabelClaim, 		row,2)
		LV_GetText(MinLimit, 		row,3)
		LV_GetText(MaxLimit, 		row,4)
		LV_GetText(Units, 			row,5)
		LV_GetText(Percision, 		row,6)
		LV_GetText(Description, 	row,7)
		LV_GetText(Method, 			row,8)
		GUI, Spec_Table:submit,NoHide
		winactivate, ahk_exe eln.exe
	}


	SelectMethod(InputMethod){
		winactivate, Select methods tests
		click, 235, 72
		Send, ^a
		click, 235, 72 ;click search bar
		sleep 100
		Sendinput, %InputMethod%{enter} ; enter method to search
		sleep 300
		click 506, 341 ;move over
	}

	;;      {{ METHOD TABLE}}
	Methods() {
		global
		TableMethodList:=
		FileRead, MethodList, MethodList.txt
		MethodRow:=[]
		winactivate, Select methods tests

		CoordMode, mouse, screen

		winGetPos, LMS_X, LMS_Y, LMS_w, LMS_h, A
		MethodTableX:=LMS_X+10
		MethodTableY:=LMS_Y+ShiftTable_Y+60
		CoordMode, mouse, window
		click, 235, 72
		Send, ^a
		gui, new, +HwndhGui, Methods List
		Loop, Read, Methods.ini
		{
			If A_Index = 1
				Continue
			MethodRow := StrSplit(A_LoopReadLine, "=")
			if !MethodRow[1]
				continue
			if (instr(MethodList,Methodrow[2]))
				TableMethodListSelect := "||"
			else
				TableMethodListSelect := "|"
			TableMethodList .=   MethodRow[1] . TableMethodListSelect
			MaxRows:=A_Index-2
		}
		Gui, Add, ListBox,  r%MaxRows% vListBox w300 glistviewhdlr multi, %TableMethodList%
		Gui, Add, Button, w100 gRunSelected, Add Methods
		try GUI, Show  ;x%MethodTableX% y%MethodTableY% w352

		return

		listviewhdlr:
		return


		RunSelected:
		Gui, submit
		sleep 300
		Loop, Parse, listbox, | ; loop through each selected method
		{
			Break.Point()
			winactivate, Select methods tests
			click, 235, 72
			Send, ^a
			IniRead,vOutput, Methods.ini, Methods, %A_LoopField%
			click, 235, 72 ;click search bar
			sleep 100
			Sendinput, %vOutput%{enter} ; enter method to search
			sleep 500
			click 506, 341 ;move over
			break.point()
			; sleep 500
		}
		break.point()
		return

		GuiEscape:
		Gui, Destroy
		return

		GuiClose:
		Gui, Destroy
		return


	}


	MethodsDropdown() {
		global
		winactivate, Select methods tests
		click, 235, 72
		Send, ^a
		Loop, Read, Methods.ini
		{
			If A_Index = 1
				Continue
			Methodmenu := StrSplit(A_LoopReadLine, "=")
			Selection:= % MethodMenu[1]
			Menu, Methodmenu, add, %Selection%, MethodsDropdownhdr
		}
		Menu, MethodMenu, Show
		return

		MethodsDropdownHdr:
		sleep 200
		InputVar:=A_ThisMenuItem
		IniRead,vOutput, Methods.ini, Methods, %InputVar%
		click, 235, 72 ;click search bar
		sleep 100
		Sendinput, %vOutput%{enter} ; enter method to search
		sleep 300
		click 506, 341 ;move over
		menu, Methodmenu, deleteAll
		This.MethodsDropdown()

		return
	}
	;; ____Copy spec template
	CopyPasteSpec(){
		global
		click
		try Menu,Menu, deleteAll
		Menu,Menu, add, Copy &Template, autofill
		Menu,Menu, add, Paste &Specs, Autofill
		Menu,Menu, add, Copy &Specs, Autofill
		return
	}
	CopySpecTemplate(DepartmentInput:=""){
		global
		MouseGetPos, premx, premy
		simpleclip:=1
		; Critical, On
		if winactive("NuGenesis LMS") && (DepartmentInput = "") {
			clipboard:=
			click
			Send, ^{c}
			clipwait, 5
			If !errorlevel
				clip.Department()
			; msgbox, %department%
		}
		else
			Department:=DepartmentInput

		click 102, 289 ;click copy into new spec link
		winwaitactive, Edit specification,,4
		if errorlevel
			return
		Breaking.Point()
		If (Department = "Analytical")
			SpecTab.Edit_Analytical_Copy()
		else If (Department = "ctPhysical")
			SpecTab.Edit_CoatedPhysical()
		else If (Department = "Physical")
			SpecTab.Edit_Physical()
		else if (Department = "Micro")
			SpecTab.Edit_Micro()
		else If (Department = "ctRetain")
			SpecTab.Edit_CoatedRetain()
		else If (Department = "Retain")
			SpecTab.Edit_Retain()
		; WinActivate, NuGenesis LMS
		Breaking.Point()
		premy:=PremY + 26
		WinActivate, NuGenesis LMS
		winwaitactive, NuGenesis LMS,,3
		if !errorlevel
			MouseMove, %Premx%, %Premy%,1
simpleclip:=
		; Critical, Off
		return
	}

	CopySpecs(){
		global
		winactivate, NuGenesis LMS
		; BlockInput, on
		; critical, On
		clipboard:=
		click 57, 715 ; edit Test
		; click 57, 750 ; edit results
		winwaitactive, Test Definition Editor,,2
		if errorlevel
			winactivate, Test Definition Editor
		; SimpleClip:=1
		click 418, 202
		Send, ^a
		Send, ^c
		; Clipwait,1
		Description:=Clipboard
		; StrReplace(Description, "eurofins", "intertek")
		iniwrite, %Description%, Settings.ini, CopiedSpecs, Description
		; MouseClick, left, 464, 532,2,0
		MouseClick, left, 464, 533,1,0
		MouseClick, left, 245, 489,1,0
		winactivate, Results Definition
		winWaitactive, Results Definition,,1
		if errorlevel
			winactivate, Results Definition
		click 282, 141 ; click row
		clipboard:=
		; msgbox, HERERE!!!!
		Send, ^c
		clipwait, 3
		Sendinput,{esc}
		clip.ParseSpecsTable()

		; blockinput off
		Sendinput, {esc}
		SimpleClip:=
		; Critical, Off
		tt(FullRequirements " " Percision,2000)
		; exit
		Return
	}


	PasteSpecs(){
		global
		tooltip,
		if winactive("NuGenesis LMS"){
			winactivate, NuGenesis LMS
			click 57, 715 ; edit Test
			winactivate, Test Definition Editor
			Breaking.Point()
			; winwaitactive, Test Definition Editor,,0.95
		}
		if winactive("Test Definition Editor"){
			click 418, 202 			;click description box
			SpecTab.TestDefinitionEditor(Description) ; the pre window
			MouseClick, left, 464, 533,1,0 ; Click scrollbar
			MouseClick, left, 245, 485,1,0  ; click results
			winactivate, Results Definition
			sleep 300
			Breaking.Point()
		}
		if winactive("Results Definition"){
			winactivate, Results Definition
			; winWaitactive, Results Definition,,0.75
			; if errorlevel
			click 84, 65  ;click Edit
		}
		; winwaitactive, Result Editor,,0.75
		; if errorlevel
		if winactive("Results Editor"){
			winactivate, Result Editor
			SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,Requirement)
			Breaking.Point()
		}
		return
	}

	AddMethod(MethodID){
		winactivate, NuGenesis LMS
		click 67, 562 ; Add Methods
		winwaitactive, Select methods tests,,0.65
		click 235, 69. 2 ; method search bar
		Sendinput, %MethodID%{enter}^{a}{click 506, 337}{click 851, 656} ; add test and hit okay
		sleep 200
		winactivate, NuGenesis LMS
		click 397, 591 ; click attrobutes
		return
	}
	Edit_Sample_Template_A(){
		global
		winactivate, Edit sample template
		Breaking.Point()
		Sendinput,{click 377, 82}{home}%Product%`,{space}{Shift down}I{Shift up}n{space}{Shift down}P{Shift up}rocess`,{space}{Shift down}A{Shift up}nalytical{tab 2}{Right 6}{tab}{right 6}{tab}{right}{enter}
		winWaitactive, NuGenesis LMS,,5
		Breaking.Point()
		if !errorlevel
			click, 73, 562
		return
	}

	EditSpecification_Analytical(){
		global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput, %Product%`,{space}{Shift down}I{Shift up}n{space}{Shift down}P{Shift up}rocess`,{space}{Shift down}A{Shift up}nalytical{tab 4}^a%Product%{tab}{enter}{tab}{space}{enter 2}{Tab}{right}{tab}{right 4}{tab}
		Breaking.Point()
		Sendinput,{right 6}{Tab 2}{Space}{tab 2}{right}{tab}{right}
		Breaking.Point()
		click, 340, 622 ;click okay
		Breaking.Point()
		winwaitactive, NuGenesis LMS, ,8
		if !ErrorLevel
			click, 88, 327 ; click add sample template
		Breaking.Point()
		winwaitactive, Edit sample template,, 8
		if !errorlevel
			SpecTab.Edit_Sample_Template_A()
		return
	}
	AutoInputTestDefinitionEditor(){
		Global
			click
			click, 57, 719 ;click Edit Test
			SelectedTestName:=
			winwaitactive, Test Definition Editor,, 2
			clipboard:=
			SimpleClip:=1
			send, {tab 1}^a^c ;copies the Test ID from test definition Editor
			ClipWait, 1
			SelectedTestName:=strreplace(Clipboard, " UPLC")
			ControlsetText, Edit4,%SelectedTestName%,ClipBar
			MatchingRow:=SpecTab.FindRowNumber(SelectedTestName)
			SpecTab.GetRowText(MatchingRow)
			if !MatchingRow
			{
				winactivate, Test Definition Editor
				MouseClick, left, 464, 532,2,0
				SpecTab.ShowSpecMenu()
				winactivate, Test Definition Editor
			}
			spectab.Autofill()
			preY+=26
			WinWaitActive, NuGenesis LMS,, 10
			if !errorlevel
				MouseMove, %preX%, %preY%, 1,
			return
		}
	AutoInputResultEditor(){
		Global
		SimpleClip:=1
		clipboard:=
		SelectedTestName:=
			If winactive("Results Definition"){
			click
			send, ^c
			clipwait,1
			ParseData:=Clipboard
					Loop, parse, ParseData, `t
						ParsedSpecs.insert(A_LoopField)
						TotalColumns:=ParsedSpecs.maxindex()//2
						AllowPrefixes:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Allow Prefixes") + TotalColumns],"`r`n")
						Clipped_ResultID:=Trim(ParsedSpecs[HasValue(ParsedSpecs, "Result Id") + TotalColumns],"`r`n")
						SelectedTestName:=Clipped_ResultID
				click, 84, 67 ;click Edit Test
				winwaitactive, Result Editor,, 2
				if Errorlevel
							return
			}
			else ;copies the Test ID from Result Editor
			{
			winactivate, Result Editor
			send, {tab 3}^a^c
			ClipWait, 1
			SelectedTestName:=Clipboard
			}
		ControlsetText, Edit4,%SelectedTestName%,ClipBar
			MatchingRow:=SpecTab.FindRowNumber(SelectedTestName)
		SpecTab.GetRowText(MatchingRow)
			; sleep 300
				winactivate, Result Editor
			; Send +{tab 3}
			if !MatchingRow
			{
				winactivate, Result Editor
				MouseClick, left, 464, 532,2,0
				; Spectab.PasteClipboardIntoSpec()
			; SpecTab.ShowSpecMenu()
				winactivate, Result Editor
		}
			SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,1)
			; spectab.Autofill()
			; preY+=26
			WinWaitActive, Results Definition,, 5
			if !errorlevel
				;MouseMove, %SpecTableMousePosX%, %SpecTableMousePosY%, 1,
			return
		}

	;; Run through all the menues to add
	AutoFill(){
		global
		; MouseGetPos, SaveX2, saveY2
		SimpleClip:=1
		If (Clipped_specs){
			clipped_Specs:=
			sleep 100
		}
		If winactive("NuGenesis LMS")
		{
			click, 57, 719 ;click Edit Test
			winwaitactive, Test Definition Editor,, 2
			sleep 200
			Breaking.Point()
		}
		If winactive("Test Definition Editor")
		{
			SpecTab.TestDefinitionEditor(Description) ; the pre window
			sleep 200
			winactivate, Test Definition Editor
			MouseClick, left, 464, 532,2,0 ;click scrollbar
			click 239, 246  ;results link
			sleep 200
			Breaking.Point()
			winwaitactive, Results,, 2
			sleep 100

		}
		if winactive("Results Definition") ;Selection window
		{
			winactivate, Results Definition
			If Method contains ICP-MS 231
			{
				MouseGetPos, SpecTableMouseX, SpecTableMouseY
				Sendinput,{click 217, 141} ;click first line

			}
			Sendinput,{click 80, 66} ;click edit
			Breaking.Point()
			winwaitactive, Result Editor,,4
			SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,1)
			; tooltip
			if SpecTableMouseX
			{
				mousemove, %SpecTableMouseX%, %SpecTableMouseY%
				SpecTableMouseX:=
			}
			return
		}
		If winactive("Result Editor") ;the editing window
		{
			Breaking.Point()
			SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,,1)
			return
		}
		; else
		simpleClip:=
			; mousemove, %saveX2%, %SaveY2%
			; return
		return
	}
PasteClipboardIntoSpec(){ 	;;//	for pasting clipboards into specs}}
					winactivate, Results Definition
					tooltip,
					Send,{click 80, 66} ;click edit
					sleep 200
					Breaking.Point()
					winwaitactive, Result Editor,,2
					if !errorlevel
						If (Clipped_specs){
							clipped_Specs:=
							sleep 100
						}
					If (MinLimit || MaxLimit) && !(FullRequirements)
						SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,1)
					else If (MinLimit || MaxLimit) && (FullRequirements)
						SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,FullRequirements)
					else If (!MinLimit && !MaxLimit)
		{
						; SpecTab.AutoInputResultEditor()
						inputbox, InputBoxString, inputstring, Min Limit,Max Limit,Units,Requirement
						ParsedString:=StrSplit(InputBoxString,",")
						MinLimit:=ParsedString[1]
						MaxLimit:=ParsedString[2]
						If ParsedString[3]
							FullRequirements:=ParsedString[3]
						Else
							FullRequirements:=1
						If ParsedString[4]
							Units:=ParsedString[4]
						Else
							Units:=""
						; inputbox, Units, Units, Units
						sleep 200
						Breaking.Point()
						SpecTab.ResultEditor(MinLimit,MaxLimit,Units,"",1,FullRequirements)
						minlimit:=
						maxlimit:=
						units:=
					}
					else
						return
				return
				}


	;;-------Fill In Test Specs
	ResultEditor(Min_Limit,Max_Limit,The_Units,The_Percision,UseLimitsBox:=0,CreateRequirements:=1,AllowPrefixesBox:=""){ ; 3rd window
		Global
		If (Clipped_specs){
			clipped_Specs:=
			sleep 100
		}
		simpleclip:=1
		winactivate, Result Editor
		click, 250, 140 ; click id box to orient
		Breaking.Point()
		Sendinput,{tab 2}%The_units%{tab}^{a}%The_Percision% ;{tab 5}
		; if (AllowPrefixesBox) || (AllowPrefixes=True)
		; TabAmount=3 ;{space}{tab 2}^{a}%Min_Limit%{tab}^a%Max_Limit%{tab 5}^a
		; else
		; TabAmount=2
		; AllowPrefixes = False
		; return
		; Sendinput, {tab 2}^{a}%Min_Limit%{tab}^a%Max_Limit%{tab 5}^a ;normal
		If (UseLimitsBox=1) || (UseLimits=True)
			click 100,406 ; click box
		; UseLimitBox:="{space}"
		; else
		; 	UseLimitBox:=
		Send, {tab 2}{space}
		sleep 200
		if !winexist("Result Editor"){
			Sendinput,{click 80, 66} ;click edit
			winwaitactive, Result Editor,,4
			SpecTab.ResultEditor(MinLimit,MaxLimit,Units,Percision,1,Clipped_Requirement)
			return
		}
		Sendinput, ^a%Min_Limit%{tab}^a%Max_Limit%{tab 5}^a ;normal
		if (Max_limit = "")&&(CreateRequirements=1){
			Sendinput, NLT %Min_Limit% %The_Units%
			return
		}
		else if (Min_limit = "<")&&(CreateRequirements=1){
			Sendinput, %min_limit%%Max_Limit% %The_Units%
			return
		}
		else if (Min_limit = "")&&(CreateRequirements=1){
			Sendinput, NMT %Max_Limit% %The_Units%
			return
		}
		Else If (CreateRequirements!=1){
			Sendinput, %CreateRequirements%
			return
		}
		else
			Sendinput, %Min_Limit% - %Max_Limit% %The_Units%
		if (The_Units)
			Mouseclick, left, 378, 667,1,0 ; click okay
		Breaking.Point()
		If Method contains ICP-MS 231
		{
			; WinActivate, Spec Table ahk_exe VQ_Helper.exe
			; sleep 200
			MouseMove, %mX%, %mY%, 0
			return
		}
		winwaitactive, Results Definition,, 5
		wingetpos, Results_X, Results_y, Results_w, Results_h, Results
		sleep 200
		Okay_x:=Results_W - 170
		Okay_y:=Results_H - 45
		mousemove, %Okay_x%, %Okay_y% ;Move mouse to Save/Okay
		Breaking.Point()
		; msgbox, %Results_w%  %Results_h%
		if ContinueToRun
			click
		WinWaitClose, Results Definition,, 8
		sleep 400
		; winwaitactive, Test Definition Editor,, 7
		if ContinueToRun
			click 342, 614
			; WinWaitActive, Test Definition Editor
		wingetpos, Test_X, Test_y, Test_w, Test_h, A
		Save_x:=test_W - 170
		Save_y:=test_H - 45
		Breaking.Point()
		mousemove, %Save_x%, %Save_y% ;Move mouse to Save/Okay
		sleep 300
		if ContinueToRun
			winwaitactive, NuGenesis LMS,,10
		sleep 300
		return
	}

	ChangeMicroLimit(NewLimit, ContinueToRun:=""){ ; 3rd window
		; Global
		winactivate, Result Editor
		click, 250, 140 ; click id box to orient
		Breaking.Point()
		Sendinput,{tab 12}
		send, ^{a}
		sendinput, %NewLimit%
		send, {tab 5}^{a}
		sendinput, NMT %NewLimit% CFU/g
		sleep 400
			Mouseclick, left, 378, 667,1,0 ; click okay
		Breaking.Point()
		if ContinueToRun
			Return
	winwaitactive, Results Definition,, 5
		wingetpos, Results_X, Results_y, Results_w, Results_h, Results
		sleep 200
		Okay_x:=Results_W - 170
		Okay_y:=Results_H - 45
		mousemove, %Okay_x%, %Okay_y% ;Move mouse to Save/Okay
		return
	}

	TestDefinitionEditor(The_Description){ ;,Department:=""){ ; 2nd window
		Global

		; TT(Selectedtestname,1000,1,-100,1)
		; SimpleClip:=
		if !(The_description) ; && !(Department)
		{
			MouseClick, left, 464, 532,2,0 	;click scrollbar
			click 239, 246 					;click results link
			return
		}
		else
		{
			winactivate, Test Definition Editor
			DescriptionRaw:=The_Description
			Trimmed_Description:=RTrim(DescriptionRaw, "`r`n")
			Click, 187, 200  ;click description box
			if Name contains Vitamin C
				Sendinput,{Home}{Delete 12}
			else
				Send, ^{a}
			sleep 100
			sendinput, %Trimmed_Description%
			if strLen(Trimmed_Description) > 100
				sleep 300
		}
	}
	toggleUseLimitsFromTheTest(){
		global
		ifwinactive, Results Definition
		{
			sendinput, {click 85, 69} ;click edit
		sleep 400
			; winwaitactive, Result Editor,,3
			; if errorlevel
				; sendinput, {click 85, 69} ;click edit
		}
	winactivate, Result Editor
		Send, {Click 511, 72}{Click 38, 409}{click 373, 657} ;Toggle Sse limits from test
		sleep 400
		lmsclick.OK() ;click okay
	return
	}

	TestDefinitionEditor_Stability(){ ;,Department:=""){ ; 2nd window
		Global
		if Winactive("NuGenesis LMS")
			mouseclick, left, 61, 715,1,0
		sleep 400
		winwaitactive, Test Definition Editor,,2
		if errorlevel
			winactivate, Test Definition Editor
		; CoordMode, Mouse, Screen
		; DescriptionRaw:=The_Description
		; Trimmed_Description:=RTrim(DescriptionRaw, "`r`n")
		mouseclick, left, 187, 200,1,0
		sleep 100  ;click description box
		sendinput, +{tab}^{a}Stability ;Stability Test Group
		sleep 200	;Orient
		mouseclick, left, 366, 280, 2,0 ;click department
		sleep 200
		sendinput, {down 15}{up}
		Breaking.Point()
		MouseClick, left, 464, 532,2,0 	;click scrollbar
		click 239, 246 					;click results link
		;Send,{shift down}{Tab 15}{Shift up}{enter}
	}

	Edit_Analytical_Copy(){
		Global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput,%Product%`,{space}{shift down}I{shift up}n{shift down}{space}P{shift up}rocess`,{space}{shift down}A{Shift up}nalytical{tab 3}^a{backspace}
		Breaking.Point()
		Sendinput,{tab}^a%Product%{tab 2}
		Breaking.Point()
		Sendinput,{Space}
		Breaking.Point()
		winwaitactive, Products List, , 2
		; if !errorlevel
		Sendinput,{enter 2}
		Breaking.Point()
		Sendinput,{tab}
		Sendinput,{right}
		sleep 300
		Breaking.Point()
		send, {enter}
		; click, 340, 622 ;click Save
		ifwinactive, Warning,
			exit
		winwaitactive, NuGenesis LMS, ,8
		if errorlevel
			exit
		Breaking.Point()
		click, 70, 518 ;edit sample method
		winwaitactive, Edit sample template,,8
		; if !errorlevel
		Breaking.Point()
		Sendinput,{tab}{^a}
		sendinput, % Product ",{space}{shift down}I{shift up}n{shift down}{space}P{shift up}rocess`,{space}{shift down}A{Shift up}nalytical"
		sleep 400
		Breaking.Point()
		send, {enter}
		return
	}
	Edit_Physical(){
		Global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput,%Product%`,{space}{shift down}I{shift up}n{shift down}{space}P{shift up}rocess`,{space}{shift down}P{shift up}hysical{tab 3}^a{backspace}
		Breaking.Point()
		Sendinput,{tab}^a%Product%{tab 2}
		Breaking.Point()
		Sendinput,{Space}
		Breaking.Point()
		winwaitactive, Products List, ,2
		if errorlevel
			exit
		Sendinput,{enter 2}
		Breaking.Point()
		Sendinput,{tab}
		Sendinput,{right}
		sleep 400
		Breaking.Point()
		Send,{enter}
		; click, 340, 622 ;click okay
		ifwinexist, Warning
			Exit
		Breaking.Point()
		winwaitactive, NuGenesis LMS, ,5
		if errorlevel
			exit
		sleep 300
		Breaking.Point()
		click, 70, 518 ;edit sample method
		winwaitactive, Edit sample template,,6
		; if !errorlevel
		Breaking.Point()
		Send,{tab}^{a}
		sendinput, % Product ",{space}{shift down}I{shift up}n{shift down}{space}P{shift up}rocess`,{space}{shift down}P{shift up}hysical"
		; sendinput, {enter}
		sleep 300
		Breaking.Point()
		send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; If !ErrorLevel
		; MouseMove, %premx%, %premy%, 0
		; click
		return
	}

	Edit_CoatedRetain(){
		global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Breaking.Point()
		Sendinput,%Product%`,{space}{shift down}C{shift up}oated`,{space}{shift down}R{shift up}etain{tab 4}^a%Product%{tab}{enter}{tab}{space}{Return 2}
		Sendinput,{tab}{right}
		Breaking.Point()
		Sendinput,{tab}{right}{tab 3} ;{left 4}
		sleep 200
		send, {enter}
		Breaking.Point()
		; click, 340, 622 ;click okay
		ifwinactive, Warning,
			exit
		winwaitactive, NuGenesis LMS, ,4
		if !errorlevel
			LMSclick.Edit_Sample_Template()
		Breaking.Point()
		Send,{tab}^{a}
		sendinput, % Product ",{space}{Shift down}C{shift up}oated`,{space}{shift down}R{shift up}etain"
		sleep 400
		Breaking.Point()
		send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; If !ErrorLevel
		; MouseMove, %premx%, %premy%, 0
		; click
		Breaking.Point()
		return
	}

	Edit_CoatedPhysical(){
		global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput,%Product%`,{space}{shift down}C{shift up}oated`,{space}{shift down}P{shift up}hysical{tab 4}^a%Product%{tab}{enter}{tab}{space}{Return 2}
		Sendinput,{tab}{right}
		Breaking.Point()
		Sendinput,{tab}{right}{tab} ;{left 4}
		sleep 400
		Breaking.Point()
		send, {enter}
		Breaking.Point()
		; click, 340, 622 ;click okay
		ifwinactive, Warning,
			exit
		winwaitactive, NuGenesis LMS,,4
		if !errorlevel
			LMSclick.Edit_Sample_Template()
		Breaking.Point()
		Sendinput,{tab}^{a}
		sendinput, % Product ",{space}{Shift down}C{shift up}oated`,{space}{shift down}P{shift up}hysical"
		sleep 300
		Breaking.Point()
		send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; If !ErrorLevel
		; MouseMove, %premx%, %premy%, 0
		; click
		return
	}

	Edit_Retain(){
		Global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput,%Product%`,{space}{shift down}I{shift up}n{space}{shift down}P{shift up}rocess`,{space}{shift down}R{shift up}etain{tab 4}^a%Product%{tab}{enter}{tab}{space}{Return 2}
		Breaking.Point()
		Sendinput,{tab}{right}
		Sendinput,{tab}{right}
		sleep 400
		Breaking.Point()
		send, {enter}
		; winWaitactive, Edit specification,, 4
		; if !errorlevel
		; click, 340, 622 ;click okay
		Breaking.Point()
		winwaitactive, NuGenesis LMS, ,5
		; if !errorlevel
		Breaking.Point()
		LMSclick.Edit_Sample_Template()
		Breaking.Point()
		Sendinput,{tab}^{a}
		sendinput, % Product ",{space}{shift down}I{shift up}n{space}{shift down}P{shift up}rocess`,{space}{shift down}R{shift up}etain"
		sleep 400
		Breaking.Point()
		Send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; If !ErrorLevel
		; MouseMove, %premx%, %premy%, 0
		; click
		return
	}
	Edit_Analytical(){
		Global
		Breaking.Point()
		If winactive("Edit sample template")
			SpecTab.Edit_Sample_Template_A()
		else If winexist("Edit specification")
		{
			winactivate,
			Breaking.Point()
			SpecTab.EditSpecification_Analytical()
		}
		return
	}

	Edit_Micro(){
		Global
		winactivate, Edit specification
		Sendinput,{click 376, 87}{home}
		Sendinput,%Product%`,{space}{shift down}F{shift up}inished`,{space}{shift down}M{shift up}icro{tab 4}^a%Product%{tab 2}
		Sendinput,{Space}
		winwaitactive, Products List,,2
		Sendinput,{enter 2}
		Sendinput,{tab}
		Breaking.Point()
		Sendinput,{right}{tab}  ; to select finish --> {left 2}
		send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; if !errorlevel
		Breaking.Point()
		LMSclick.Edit_Sample_Template()
		Breaking.Point()
		sendinput, {tab}^{a}
		sendinput, % Product "`,{space}{shift down}F{shift up}inished`,{space}{shift down}M{shift up}icro"
		sleep 400
		Breaking.Point()
		send, {enter}
		winwaitactive, NuGenesis LMS,,5
		; If !ErrorLevel
		; MouseMove, %premx%, %premy%, 0
		; click
		return
	}

	InsertDescription(){
		Global
		DescriptionRaw:=Description
		Description:=RTrim(DescriptionRaw, "`r`n")
		Sendinput,^a%Description%
		Return
	}

	AddPaeruginosa(){
		winactivate, ahk_exe eln.exe
		if winactive("Results Definition")
			clk(50, 65)  ;click add
		sleep 200
		winactivate, Results Editor
		sendinput, {tab 3}^{a}P. aeruginosa{tab 2}10g{tab 6}{space}{tab 3}0{tab 5}Absent/10g
		sleep 200
		send, {click 370,660}
		; return
	}



	HM_ReportOnly(){
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab 6}Report Only
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab 6}Report Only
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab 6}Report Only
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab 6}Report Only
		click 390, 659	;click okay
		return
	}

	HM_USP(){
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}15{tab 5}NMT 15 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,150 ;click 2nd row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}5{tab 5}NMT 5 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,190 ;click 3rd row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}5{tab 5}NMT 5 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,210 ;click 4th row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}15{tab 5}NMT 15 mcg/day
		click 390, 659	;click okay
		return
	}

	HM_Bloom(){
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}10{tab 5}NMT 10 mcg/day ;Arsenic
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,150 ;click 2nd row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}10{tab 5}NMT 10 mcg/day ;Lead
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,190 ;click 3rd row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}6{tab 5}NMT 6 mcg/day ;Cadmium
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		; click 125,210 ;click 4th row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 3}20{tab 5}NMT 20 mcg/day ;Murcury
		click 390, 659	;click okay
		return
	}

	HM_Custom(){
		global HM_Units, Arsenic_Limit, Arsenic_Requirement, Lead_Limit, Lead_Requirement, Cadmium_Limit, Cadmium_Requirement, Mercury_Limit, Mercury_Requirement,HM_Lower_Limit
		if !HM_Lower_Limit
			HM_Lower_Limit:=""
		winactivate, Result Definition
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,2
		Sendinput,{tab 5}%HM_Units%{tab 7}{space}{tab 2}%HM_Lower_Limit%{tab}%Arsenic_Limit%{tab 5}%Arsenic_Requirement%%A_space%%HM_Units%
		click 390, 659	;click okay
		Breaking.Point()
		winWaitClose, Result Editor,,2
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,2
		Sendinput,{tab 5}%HM_Units%{tab 7}{space}{tab 2}%HM_Lower_Limit%{tab}%Lead_Limit%{tab 5}%Lead_Requirement%%A_space%%HM_Units%
		click 390, 659	;click okay
		winWaitClose, Result Editor,,2
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,2
		Sendinput,{tab 5}%HM_Units%{tab 7}{space}{tab 2}%HM_Lower_Limit%{tab}%Cadmium_Limit%{tab 5}%Cadmium_Requirement%%A_space%%HM_Units%
		click 390, 659	;click okay
		winWaitClose, Result Editor,,2
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,2
		Sendinput,{tab 5}%HM_Units%{tab 7}{space}{tab 2}%HM_Lower_Limit%{tab}%Mercury_Limit%{tab 5}%Mercury_Requirement%%A_space%%HM_Units%
		click 390, 659	;click okay
		return
	}

	HM_Canada(){
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab}9.8{tab 5}NMT 9.8 mcg/day
		click 390, 659	;click okay
		Breaking.Point()
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab}9.8{tab 5}NMT 9.8 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab}6.3{tab 5}NMT 6.3 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 5}mcg/day{tab 7}{space}{tab 2}0{tab}20.3{tab 5}NMT 20.3 mcg/day
		click 390, 659	;click okay
		return
	}

	HM_Prop65(){
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winactivate, Result Editor
		winwaitactive, Result Editor,,4
		Sendinput,{tab 12}{space}{tab 2}^{a}
		;Sendinput,{tab 5}mcg/day{tab 7}{tab 2}^a
		Sendinput, 0{tab}^{a}
		sleep 100
		Sendinput, 9.999{tab 5}^a<10 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		winactivate, Result Definition
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 12}{space}{tab 2}^a
		;Sendinput,{tab 5}mcg/day{tab 7}{tab 2}^a
		Sendinput, 0{tab}^a
		sleep 100
		Sendinput, 0.499{tab 5}^a<0.5 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		winactivate, Result Definition
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 12}{space}{tab 2}^a
		;Sendinput,{tab 5}mcg/day{tab 7}{tab 2}^a
		Sendinput, 0{tab}^a
		sleep 100
		Sendinput, 4.099{tab 5}^a<4.1 mcg/day
		click 390, 659	;click okay
		winWaitClose, Result Editor,,4
		winactivate, Result Definition
		click 125,130 ;click 1st row
		click 80,70 ;Edit
		Breaking.Point()
		winwaitactive, Result Editor,,4
		Sendinput,{tab 12}{space}{tab 2}^a
		;Sendinput,{tab 5}mcg/day{tab 7}{tab 2}^a
		Sendinput, 0{tab}^a
		sleep 100
		Sendinput, 0.299{tab 5}^a<0.3 mcg/day
		click 390, 659	;click okay
		return
	}
	;; __Remove tests from spec__
	NewSpecVersion(Prompt:=""){
		Global
		MouseGetPos, m_x, m_Y
		SelectionStatus:=
		VersionStatus:=
		ifwinactive, NuGenesis LMS
			Clip.ParseMainSpecsTopTable()
		sleep 200
		{
			while (VersionStatus="Done" || VersionType = "Modifiable" ) { ;|| VersionStatus = "Removed") {
				sendinput,{down}
				sleep 200
				Clip.ParseMainSpecsTopTable()
			}
			winwaitactive, NuGenesis LMS,,2
			if !errorlevel
				click, 69, 249 ; new version
			WinWait, Edit specification,, 2
			if errorlevel
			{
				if winexist("Delete specification") || winexist("Lock specification") || winexist("Approve specification") || winactive("Delete specification"){
					; MsgBox, 4, , Do you want exit? (Press YES or NO) `n [Error 1668],5
					; ifmsgbox Timeout
					; {
					winactivate,
					sendinput, {n}
					sleep 200
					M_y+=26
					click %m_x% %m_Y% 1
					; return
					; }
					; IfMsgBox No
					; {
					; winactivate,
					; sendinput, {n}
					; sleep 200
					; return
					; }
					; IfMsgBox Yes
					; exit
				}
				; msgbox,, IDK
			}
			Breaking.Point()
		}
		Ifwinexist, Edit specification
		{
			winactivate, Edit specification
			click, 393, 177 ; click description
			Breaking.Point()
			send, ^{a}
			sleep 100
			Breaking.Point()
			sendinput, %NewSpecVersionDescriptionText%
			if Prompt
				MsgBox, 4, , Do you want to continue? (Press YES or NO) `n [Error 1699], 5
			IfMsgBox Timeout
				winactivate, Edit specification
			IfMsgBox No
				exit
			IfMsgBox Yes
				winactivate, Edit specification
			sleep 300
			Breaking.Point()
			click, 331, 617 ;click Save
			if Product
			{
				iniwrite, New, FixedSpecs.ini, Status, %Product%
				SelectionStatus:="New"
			}
			winwaitactive, NuGenesis LMS,,4
			if errorlevel
				return
			; M_y+=26
			; click %m_x% %m_Y% 1
		}
		return
	}

	RemoveTestSpec(){
		global
		if winactive("NuGenesis LMS")
			click, 63, 754   ; edit results
		winwaitactive, Results Definition,,3
		if errorlevel
			MsgBox, 4, , Do you want to continue?  (Press YES or NO) `n [Error 1718]
		IfMsgBox No
			exit
		if (!OkayButton_X || !OkayButton_Y){
			wingetpos, ResultsDefinition_X, ResultsDefinition_y, ResultsDefinition_w, ResultsDefinition_h, Results Definition
			OkayButton_X:=ResultsDefinition_w - 175
			OkayButton_Y:=ResultsDefinition_H - 40
		}
		click 111, 95 ;sort sequence to apply filter
		sleep 200
		Breaking.Point()
		click 128, 65 ; Remove button
		winwaitactive, Delete results,, 2
		if errorlevel
		{
			sleep 200
			winactivate, Results Definition
			click 128, 65
		}
		sendinput, {enter}
		Breaking.Point()
		; if winactive("NuGenesis LMS")
		; exit
		; else
		winwaitactive, Results Definition,,3
		if Errorlevel
			return
		click %OkayButton_X%,  %OkayButton_Y%
		if Product
			iniwrite, Removed, FixedSpecs.ini, Status, %Product%
		; SelectionStatus:="Removed"
		; winwaitactive, NuGenesis LMS,,6
		; if errorlevel
		; exit
		return
	}
	ApproveSpecVersion(){
		global
		; wingetpos, Nugenesis_X, Nugenesis_y, Nugenesis_w, Nugenesis_h, NuGenesis LMS
		; ScrollBarTop_X:=Nugenesis_W - 5
		; TopListItem_X:=Nugenesis_W - 50
		; ScrollBarTop_Y:=Nugenesis_y + 190
		winwaitactive, NuGenesis LMS,,3
		if !errorlevel
			Click, 76,269 ;click Approve Specification
		winwaitactive, Approve specification,,3
		if errorlevel
			if winexist("Delete specification") || winexist("Lock specification") { ; || winexist("Approve specification") ; || !winexist("Approve specification")
				WinActivate
				sendinput, {n}
				return
			}
		if winactive("Approve specification")
			sendinput, {y}
		if Product
			iniwrite, Done, FixedSpecs.ini, Status, %Product%
		SelectionStatus:=
		; sleep 2000
		; CoordMode, mouse, screen
		; Click, %TopListItem_X%, %ScrollBarTop_Y%
		; CoordMode, mouse, Window
	}


	RemoveAndApproveTestSpec(){
		global
		SelectionStatus:=
		MouseGetPos, m_x, m_Y
		SelectionStatus:= Clip.ParseMainSpecsTopTable()
		sleep 350
		If (SelectionStatus="Removed"){
			this.ApproveSpecVersion()
			M_Y+=26
			click %m_x% %m_Y% 1
			return
		}
		else If (SelectionStatus="NotStarted") {
			this.NewSpecVersion()
			winwaitactive, NuGenesis LMS,,3
			if !errorlevel
				this.RemoveTestSpec()
			return
		}
		else If (SelectionStatus="New")
			this.RemoveTestSpec()
		else If (SelectionStatus="Done"){
			sendinput, {down}
			Clip.ParseMainSpecsTopTable()
		}
		else
			return
		return
	}

	FullRemoveTest(){
		global
		sleep 200
		MouseGetPos, m_x, m_Y
		click, 62, 717 ; click View test
		sleep 200
		if winactive("NuGenesis LMS")
		{
			CoordMode, mouse, screen
			click, %m_x%, %m_y%
			sendinput, {down}
			CoordMode, mouse, window
			return
		}
		if winactive("Test Definition Viewer")
			send, {escape}
		sleep 200
		this.NewSpecVersion()
		if winexist("Delete specification") || winexist("Edit specification") || winexist("Lock specification") || winexist("Approve specification")
			exit
		if winactive("Delete specification")
			exit
		sleep 700

		sleep 500
		this.RemoveTestSpec()

		sleep 700
		Breaking.Point()
		if !winactive("NuGenesis LMS")
			sleep 500
		; my:= my
		; Mousemove, %mx%, %my%,0
		this.ApproveSpecVersion()
		; TT("Done")
		return
	}





}




WM_Lbuttondown:
Spec_Table:
	if (A_GuiEvent = "NORMAL" ){
		; if (A_GuiEvent := "I" ){
		DESCRIPTION:=""
		MouseGetPos, mX, mY, mWin
		Sendinput,{space}
		SpecTab.GetRowText()
		; tt(ShowVariables,1000)
		winactivate, ahk_exe eln.exe
		SpecTab.AutoFill()
	}
Return


Spec_TableGuiClose:
	GUI, Spec_Table:destroy
	; coordmode, mouse, window
return

;;----------------------------------------------------------
;;{{                   WorkTab Class                      }}
Class WorkTab {
	NewTestRequestLink(){
		Global
		winactivate, Register new samples
		MouseClick, Left, 351, 621
		if !Product
			exit
		winwaitactive, Edit request,,3
		if errorlevel
			breaking.point(1)
		sendinput, {tab 4}%Product%{tab 2}{space}{enter 2}{tab}{down}{tab 2}%batch%
		if !Batch
			exit
		MouseClick, Left, 356, 619

		return
	}
	registerNewSamples(){
		global
		mx:=

		my:=
		If Coated = "ERROR"
			Coated:=
		If Lot = "ERROR"
			lot:=

		ControlGetText, Iteration, Edit5, ClipBar
		ifwinactive, Register new samples
			MouseGetPos, mx, my
		click 2
		sleep 200
		winwaitactive, Edit sample (Field Configuration,, 2
		if errorlevel
			breaking.point(1)
		Sendinput, {tab 2}{right}{tab}^{a}%Batch%{tab} ; switch to Availible and enter batch
			; Iteration:=CustomerPosition
			Breaking.Point()
			WorkTab.Dropdown_CustomerSelect(Iteration)
		Ifwinactive, Edit sample (Field Configuration: F`, Micro)
		{
			Send, ^{a}%Lot%
			if Coated
				Send, {tab}^{a}%Coated%
			; Sendinput, {tab 2}
		}
		Ifwinactive, Edit sample (Field Configuration: CT`,
			{
				Send, {tab}^{a}%Coated%
			}
			Breaking.Point()
			Send, {enter} ; hit okay

			winwaitactive, Register new samples,,2
			if errorlevel
			breaking.point(1)
			mousemove, mx, my,0

			return
		}

		SaveCustomerList(){
			; if !(InputString)
			; Inputstring:=Clipboard
			Clipboard:=strReplace(StrReplace(Clipboard, "`r`n","`n"),"`t","") ; Normalize line endings
			lines := StrSplit(Clipboard, "`n") ; Split the string into lines
			; CustomerString:=
			totalLines := lines.Length() -1
			half := FLOOR(totalLines/2)
			halfAmount:=-(Half)

			Try FileDelete, Customers.ini
			CustomersString:=
			Loop, % totalLines
			{
				if (A_Index=1){
					CustomersString:="[Customers]"
					continue
				}
				line := lines[A_Index]
				if (A_Index <= half)
				lineNumber := A_Index - 2
				else
				{
					Linenumber:= halfAmount
					HalfAmount+=1
				}
				CustomersString.="`n"Line "=" lineNumber
			}
			sleep 300
			FileAppend,% Trim(CustomersString), Customers.ini
			return CustomersString
		}

		CustomerMenu() { ;; create a dropdown from CustomerMenu ini datafile
			global
			try menu, CustomersMenu, DeleteAll
			Loop, Read, %CustomerListPath%
			{
				MenuItems := StrSplit(A_LoopReadLine, "=")
				Selection:= % MenuItems[1]
				Count:= % MenuItems[2]
				Menu, CustomersMenu, add, %Count% `t &%Selection%, CustomerMenu ;%Addbreak%
				; AddBreak:=
			}
			Menu, CustomersMenu, Show,
			return

			CustomerMenu:
			sleep 200
			MenuSelection:=[]
			; InputVar:=StrReplace(A_ThisMenuItem,"&",)
			MenuSelection:=StrSplit(A_ThisMenuItem, "&")
			inputVar:=MenuSelection[2]
			; InputVar:=A_ThisMenuItem
			IniRead,CustomerPosition, %CustomerListPath%, Customers, %InputVar%
			sleep 30
			; if winactive("Edit sample (Field Configuration: F, Micro)")
			; send, {tab 9}
			; else
			; send, {tab 7}
			sleep 20
			menu, CustomersMenu, DeleteAll
			if CustomerPosition > 0
			customerPosition+=1
			if CustomerPosition < 0
			customerPosition-=1
			Iteration:=CustomerPosition
			GuiControl,ClipBar:Text, Iteration, %Iteration%
			GUI, ClipBar:submit,NoHide
			iniwrite, %Iteration%, Settings.ini, SavedVariables, Iteration
			; iniwrite, %Iteration%, Settings.ini, SavedVariables, CustomerPosition
			winactivate, Edit sample
			Ifwinactive, Edit sample (Field Configuration: F`, Micro)
				send, {Click 253, 138}{tab 4}
			Ifwinactive, Edit sample (Field Configuration: CT
				send, {Click 253, 138}{tab 2}
			else
				send, {Click 253, 138}{tab}
			sleep 200
			this.Dropdown_CustomerSelect(CustomerPosition)
			Nsub:=

			return
		}

		Dropdown_CustomerSelect(A_ShipTo){
			; Critical, on
			; msgbox, shipto %A_shipto% `niteration: %iteration% `nAbsselection: %Absselection% `n customerPosition:  %CustomerPosition%
			; setkeydelay, -1,-1
			AbsSelection:=Abs(A_ShipTo)-1
			if (a_shipto = "-1")
			; msgbox %A_shipto% -1
			Sendinput,{end}
			else if (a_shipTo = 1)
			; msgbox %A_shipto% 1
			Sendinput,{home}
			else if (a_ShipTo > 1)
			; msgbox %A_shipto% >1
			Sendinput,{home}{right %A_ShipTo%}{right}
			; Sendinput,{home}{right}{right %A_ShipTo%}  previously used one
			else if (a_ShipTo < 1)
			; msgbox %A_shipto% <1
			Sendinput,{end}{left %Absselection%}{left}
			; Sendinput,{end}{left}{left %Absselection%}   previously used one
			else if (a_ShipTo = "")
			; msgbox %A_shipto% blank
			worktab.CustomerMenu()
			else if (a_ShipTo = 0)
			; msgbox %A_shipto% 0
			worktab.CustomerMenu()
			else
			{
				; setkeydelay, 0, 0
				; critical, off
				msgbox, This is before the exit
				; exit
			}
			; else
			; if (a_shipto > 175) || (absselection > 175)
			; 	sleep 500
			; if winactive("Edit sample `(Field Configuration:")
			; critical, off
			; setkeydelay, 0, 0
			return
		}
		EditRequest(){
			Global
			winactivate, Edit request
			mouseclick, left, 369, 139,1,0
			sendinput, {tab 2}%Product%{tab 2}{space}
			WinWaitActive, Products List, ,1
			if ErrorLevel
				Return
			sendinput, {enter 2}
			winwaitclose, Products List,,1
			if ErrorLevel
				Return
			Send, {tab}{down}{tab 2}
			if !Batch
				return
			sendinput, %batch%
			Mousemove, 356, 619,1
			return
		}

		NewRequest(){
			global
			department:= ; Clip()
			Clipboard:=
			winactivate, NuGenesis LMS
			click
			; Send, ^c
			clip()
			click 64, 300 ;click Assign To New rewuest link
			Breaking.Point()
			winwaitactive, Edit request,,3
			; if !Errorlevel
			winactivate, Edit request,
			click 238, 622 ;pick test
			Breaking.Point()
			winwaitactive, Select tests for request,,3
			; if !Errorlevel
			click, right, 264, 590 ; click to show filer
			Breaking.Point()
			Sendinput,{click, right, 264, 590}{up}{enter}
			click, 97, 125 ; click filter
			Sendinput, %Department%{enter}{tab 2}
			Breaking.Point()
			Sendinput, %product%{enter}
			Sendinput,{tab}{ctrldown}{a}{ctrlup}
			input, , V T3,{Lbutton}{enter}
			click 504, 338 ; click arrow
			winactivate, Select tests for request
			click, right, 264, 590 ; click to clear filter
			Sendinput,{up}{enter}
			winactivate, Select tests for request
			; winactivate, Select tests for request
			Breaking.Point()
			click 854, 657 ; click okay
			winwaitclose, Select tests for request,,3
			if !Errorlevel
			winWaitactive, Edit request,,8
			Breaking.Point()
			While GetKeyState("Lbutton", "P")
				Ifwinnotactive, Edit request
					reload
			winactivate, Edit request
			Breaking.Point()
			Send,{tab}{enter}
			; tooltip,

			return
		}



		CorrectTestResults(Checkbox_Toggle:=0,LoopThrough:=""){
			global
			Thread, Interrupt, 0
			Breaking.Preable()
			MouseGetPos, xpos, ypos
			if LoopThrough
			{
				MouseGetPos, xpos, ypos
				Breaking.Point()
				loop 20,
				{
					blockinput, on
					Breaking.Point()
					click
					click 843, 202, 2
					if Checkbox_Toggle ;Contains Toggle
					Send,{tab}{Space}{tab}{Space}
					else
						Sendinput,{tab}{tab}
					;Breaking.Point()
					Send,{tab 10}^a
					if (Iteration != 0)
					Send, %Iteration%
					ypos:=ypos+26
					mousemove, xpos, ypos,0
					; Breaking.Point()
					if GetKeyState("Lbutton","P"){
						; sleep 500
						exit
					}
				}
				; Breaking.Point()
				click
			}
			Breaking.Point()
			if GetKeyState("Lbutton","P"){
				; sleep 500
				Exit
			}
			click
			click 843, 202, 2
			if Checkbox_Toggle { ;Contains Toggle
				Sendinput,{tab}{Space}{tab}{Space}
				mousemove, xpos, ypos+26,0
			}
			else
				Sendinput,{tab}{tab}
			; if keep_running = n ;another signal to stop
			; return
			;Breaking.Point()
			Sendinput,{tab 10}^a
			Breaking.Point()
			if ((Iteration != 0) && !(Iteration > 8) && !(Iteration < 0))
			Send, %iteration%
			;if !Checkbox_Toggle ; Not Contains toggle
			mousemove, xpos, ypos+26,0
			blockinput, off
			Breaking.Point()
			return
		}


		Click_EditResults(){
			ifwinnotactive, NuGenesis LMS
				winactivate
			Sendinput,{click 63, 754} ;edit results
			return
		}
		AddTestDescription(Text){
			Sendinput,{click 305, 294}{end}%TEXT%{click 330, 617}
			sleep 300
		}

		SelectTestSample(){
			global
			Click
			MouseGetPos, mx, my
			; blockinput on
			ifwinactive, Select tests for request: R
				send, {click 647, 75} ;click assign Samples
			winwaitactive, Select samples for test,,2
			if Errorlevel
			breaking.point(1)
			sleep 400
			send {click 467, 71} ;Click Filter button
			Breaking.Point()
			sendinput, {click 248, 68}{up} ;click dropdown then
			send, {click 212, 188} ; select Batches
			Send, {click 136, 119}
			send, ^{a} ;flick filter box
			if winactive("Select samples for test: Organoleptic Sensory Test") || winactive("Select samples for test: Average Capsule Weight") || winactive("Select samples for test: Average Tablet Weight")
			Department:="Physical"
			else if winactive("Select samples for test: Microbiological") || if winactive("Select samples for test: Pathogens")
			Department:="Micro"
			else
				Department:=""
			sendinput, %Department%{enter} ; send department
			Breaking.Point()
			sendinput, {click 504, 324} ;move over
			sleep 300
			Breaking.Point()
			sendinput, {click, 849, 661} ;click okay
			; sendinput, {click, 849, 661, 0} ;move okay
			Breaking.Point()
			; tt(Department,3000)
			winwaitactive, Select tests for request,,3
			sleep 200
			mousemove, mx, my+26,0
			; blockinput off

			return
		}

		CopyProductRotation(){
			global
			clipboard:=
			send, ^c
			clipwait, 1
			sleep 400
			filename:= Product ".txt"
			FileDelete, %FileName%
			FileAppend, %Clipboard%, %Filename%
			LMSwb:=ComObjactive("Excel.Application")
			Rotation:=lmswb.activeSheet.Range("A:A").Find(Product).offset(0,5)
			Rotation.Value:=Product ".txt"
			;iniwrite %Clipboard%, Settings.ini, Rotations, %Product%
		}

		PasteProductRotation(){
			global
			filename:= "Rotations\" Product ".txt"
			FileRead, Clipboard, %Filename%
			; iniread Clipboard,Settings.ini, Rotations, %Product%
			LMSwb:=ComObjactive("Excel.Application")
			RotationDone:=lmswb.activeSheet.Range("A:A").Find(Product).offset(0,6)
			RotationDone.Value:="1"
			sleep 200
			send, ^v
		}
	}



; ----------------------------------------------------
;;{{                  LMS Click                     }}


	class LMSclick {
		OK(sleeptime:="1"){
			global
			MouseGetPos, LMSmX, LMSmy, mWin
			; WinGetPos, LMSWinX, LMSWinY, LMSWinWidth, LMSWinHeight, A
			sleep %Sleeptime%
			If GetKeyState("Shift","P"){
				Send, {enter}
				return
			}
			else if winactive("NuGenesis LMS"){
				LMS.SaveCode()
				return
				}
			else if winactive("Results Definition"){
					wingetpos, Results_X, Results_y, Results_w, Results_h, Results
					sleep 200
					Okay_x:=Results_W - 170
					Okay_y:=Results_H - 45
					click, %Okay_x%, %Okay_y%
				; send, {Click 971, 473}
				return
			}
			else If winactive("Result Editor")
			clk(370,660)
			else if winactive("Edit Ingredient")
			clk(265, 561)
			else if winactive("Result Entry")
			clk(1028, 860)
			else if winexist("Delete Test") {
				winactivate, Delete Test
				clk(229, 136)
			}
			else if winactive("Results Definition")
			clk(951, 751)
			else if winactive("Microsoft Excel Security Notice")
			Send, !y
			else if winactive("Reason For Change")
			clk(241, 236)
			else if winactive("New Document"){
				LMS.Savecode()
				clk(415, 360)
			}
			else if winactive("Edit specification")
			clk(323, 621)
			else if winactive("Edit Ingredient")
			clk(277, 557)
			else If winactive("Test Definition Editor ")
			Click 341, 618
			else if winactive("Reason for Change")
			clk(170, 331)
			else if winexist("Error") {
				winactivate
				clk(148, 104)
				winactivate, Register new samples
				clk(181, 104, 2)
				Send, %product%{enter}
			}
			else if winexist("Change Configuration")
			clk(131, 296,"Change Configuration")
			Else
				Send, {enter}
			return
		}

		esc(){
			if winexist("Change Configuration")
			click 342, 296
			Else
				Send,{esc}
			sleep 500
			return
		}

		Save(){
			if winactive("Edit Product")
			click 275, 578
			else If winactive("Composition")
			click 1236, 888
			else If winactive("Test Definition Editor ")
			Click 341, 618
			Else
				return
		}
		edit(){
			if winactive("Results Definition"){
				click 78,63
				winwait, Result Editor,,1
				if errorlevel
					return
					clk(90, 409,,1,,0)
					sleep 300
					clk(510, 579,2,,0)
					clk(283, 321,,2,,0)
			}
			else
				click 84, 65
			return
		}
		Add(){
			if winactive("Formulation")
			click, 73, 280
			else
				click 50, 65
			return
		}
		Requests_tab(){
			click 865,83
			return
		}
		MyWork_Tab(){
			click 761,44
			return
		}
		Mywork_Searchbar(){
			click 500,127, 2
			return
		}
		Edit_Composition(){
			clk(90, 440,,,"NuGenesis LMS")
			return
		}
		Edit_Test(){
			ifwinnotactive, NuGenesis LMS
				winactivate, NuGenesis LMS
			if (Tab = Specs)
				clk(68, 718)
			return
		}
		Edit_Results(){
			ifwinnotactive, NuGenesis LMS
				winactivate, NuGenesis LMS
			if (Tab = Specs)
				clk(68, 718)
			return
		}
		Enter_Results(){
			winactivate, NuGenesis LMS
			clk(64, 755)
			return
		}

		Orient(){
			if winactive("Result Entry")
			click 843, 202, 2
			else if winactive("Test Definition Editor")
			Click, 187, 200

			Else
				return
		}
		Copy_Spec_Template(){
			winactivate, NuGenesis LMS
			click 102, 289 ;copy into new spec
			winWaitactive, Edit specification, ,2
			if ErrorLevel
			winactivate, Edit specification
			; click 317, 83
			return
		}
		New_Sample_Template(){
			winactivate, NuGenesis LMS
			click 103, 325
			return
		}

		Edit_Sample_Template(){
			winactivate, NuGenesis LMS
			click 70, 518
			winwaitactive, Edit sample template,, 3
			return
		}

		Password(){
			if winExist("Login"){
				winactivate
				Sendinput, mmignin{tab}kilgore7744{enter}
				return
			}
			Else If winexist("Sign :"){
				winactivate,
				Sendinput,{tab 2}{right 2}{tab 2}mmignin{tab}kilgore7744{enter}
				return
			}
			else if winexist("windows Security"){
				winactivate,
				Sendinput, kilgore7744{enter}
			}
			else if winexist("CredentialUIBroker.exe"){
				winactivate,
				Sendinput, kilgore7744{enter}
			}

			else
				return
			return
		}
	}







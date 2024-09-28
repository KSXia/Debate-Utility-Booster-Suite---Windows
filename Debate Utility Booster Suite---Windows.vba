' ===Debate Utility Booster Suite - Windows - v1.4.0==
' Updated on 2024-09-27.
' https://github.com/KSXia/Debate-Utility-Booster-Suite---Windows
' Thanks to Truf for creating and providing his Verbatim macros, upon which many of these macros and sub procedures are built upon! Macros in the Debate Utility Booster Suite built upon macros or code that Truf wrote have more specific attribution in their header(s). You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/

' ---Standardize Highlighting With Exceptions Macro v2.0.2---
' Updated on 2024-08-21.
' https://github.com/KSXia/Verbatim-Standardize-Highlighting-With-Exceptions-Macro
' Based on Verbatim 6.0.0's "UniHighlightWithException" function.
Sub StandardizeHighlightingWithExceptions()
	Dim ExceptionColors() As Variant
	
	' ---USER CUSTOMIZATION---
	' <<SET THE HIGHLIGHTING COLORS THAT SHOULD NOT BE STANDARDIZED HERE!>>
	' Add the names of highlighting colors that you want to exempt from standardization to the list in the ExceptionColors array. Make sure that the name of every highlighting color is in quotation marks and that each term is separated by commas.
	' NOTE: This macro does NOT automatically exempt the highlighting color you have set to be exempted in the Verbatim settings. You MUST MANUALLY enter the highlighting colors you would like to exempt into this list.
	'
	' These are the names of the highlighting colors in the each row of the highlighting color selection menu, listed from left to right:
	' First row: Yellow, Bright Green, Turquoise, Pink, Blue
	' Second row: Red, Dark Blue, Teal, Green, Violet
	' Third row: Dark Red, Dark Yellow, Dark Gray, Light Gray, Black
	' MAKE SURE TO USE THIS EXACT CAPITALIZATION AND SPELLING!
	'
	' If you are using gray highlighting, you are likely using the color Light Gray
	'
	' Warning: There needs to be at least one hightlighting color listed in the ExceptionColors array for this macro to work.
	ExceptionColors = Array("Light Gray", "Pink")
	
	' ---INITIAL SETUP---
	Dim r As Range
	Set r = ActiveDocument.Range
	
	Dim GreatestIndex As Integer
	GreatestIndex = UBound(ExceptionColors) - LBound(ExceptionColors)
	
	' ---CONVERT HIGHLIGHTING COLOR NAMES TO VBA INDEXES---
	Dim ExceptionEnums() as Long
	ReDim ExceptionEnums(0 To GreatestIndex) As Long
	For CurrentIndex = 0 to GreatestIndex Step +1
		Select Case ExceptionColors(CurrentIndex)
			Case Is = "None"
				ExceptionEnums(CurrentIndex) = wdNoHighlight
			Case Is = "Black"
				ExceptionEnums(CurrentIndex) = wdBlack
			Case Is = "Blue"
				ExceptionEnums(CurrentIndex) = wdBlue
			Case Is = "Bright Green"
				ExceptionEnums(CurrentIndex) = wdBrightGreen
			Case Is = "Dark Blue"
				ExceptionEnums(CurrentIndex) = wdDarkBlue
			Case Is = "Dark Red"
				ExceptionEnums(CurrentIndex) = wdDarkRed
			Case Is = "Dark Yellow"
				ExceptionEnums(CurrentIndex) = wdDarkYellow
			Case Is = "Light Gray"
				ExceptionEnums(CurrentIndex) = wdGray25
			Case Is = "Dark Gray"
				ExceptionEnums(CurrentIndex) = wdGray50
			Case Is = "Green"
				ExceptionEnums(CurrentIndex) = wdGreen
			Case Is = "Pink"
				ExceptionEnums(CurrentIndex) = wdPink
			Case Is = "Red"
				ExceptionEnums(CurrentIndex) = wdRed
			Case Is = "Teal"
				ExceptionEnums(CurrentIndex) = wdTeal
			Case Is = "Turquoise"
				ExceptionEnums(CurrentIndex) = wdTurquoise
			Case Is = "Violet"
				ExceptionEnums(CurrentIndex) = wdViolet
			Case Is = "White"
				ExceptionEnums(CurrentIndex) = wdWhite
			Case Is = "Yellow"
				ExceptionEnums(CurrentIndex) = wdYellow
			Case Else
				ExceptionEnums(CurrentIndex) = wdNoHighlight
		End Select
	Next CurrentIndex
	
	' ---MORE SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---REHIGHLIGHTING---
	With r.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Highlight = True
		.Replacement.Highlight = True
		.Text = ""
		.Replacement.Text = ""
		.Forward = True
		.Wrap = wdFindStop
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = False
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		
		Do While .Execute(Forward:=True) = True
			' Check if the color of the current word is one of the exceptions
			Dim IsException As Boolean
			IsException = False
			Dim i
			For i = LBound(ExceptionEnums) To UBound(ExceptionEnums)
				If r.HighlightColorIndex = ExceptionEnums(i) Then
					IsException = True
				End If
			Next I

			If IsException Then
				' If the color of the current word is an exception:
				r.Collapse Direction:=wdCollapseEnd
			Else
				' If the color of the current word is not an exception:
				' Set the highlighting of the current word to the default highlighting color
				r.HighlightColorIndex = Options.DefaultHighlightColorIndex
			End If
		Loop
		
		.ClearFormatting
		.Replacement.ClearFormatting
	End With
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub

' ---Argument Numberer v1.0.3---
' Updated on 2024-09-03.
' https://github.com/KSXia/Verbatim-Argument-Numberer/
' Based on Verbatim 6.0.0's "AutoNumberTags" function.
Sub NumberArguments()
	Dim NumberPlaceholder as String
	Dim TemplateToNumber as String
	Dim ResetArgumentNumberAtPocket As Boolean
	Dim ResetArgumentNumberAtHat As Boolean
	Dim ResetArgumentNumberAtBlock As Boolean
	
	' ---USER CUSTOMIZATION---
	' Set the NumberPlaceholder to the character that you want the number to replace.
	NumberPlaceholder = "x"
	
	' Set the TemplateToNumber your numbering template, with the character you set as the NumberPlaceholder in place of where the number should go.
	' In your document, you must put the TemplateToNumber at the beginning of any tag or analytic you want this macro to number.
	' WARNING: The NumberPlaceholder MUST not be repeated in the template. The character used as the NumberPlaceholder MUST only show up once in the TemplateToNumber.
	TemplateToNumber = "[x]"
	
	' Set the header types that the argument number should reset at.
	' If you want the argument number to reset at a certain header type, set the corresponding variable to True.
	' If you do NOT want the argument number to reset at a certain header type, set the corresponding variable to False.
	ResetArgumentNumberAtPocket = True
	ResetArgumentNumberAtHat = True
	ResetArgumentNumberAtBlock = False
	
	' ---INITIAL VARIABLE SETUP---
	Dim TemplateLength As Integer
	TemplateLength = Len(TemplateToNumber)
	
	' The following code for numbering arguments is based on Verbatim 6.0.0's "Auto Number Tags" function.
	Dim p As Paragraph
	Dim CurrentArgumentNumber As Long
	
	' ---PROCESS TO NUMBER ARGUMENTS---
	' Loop through each paragraph and insert the number if the numbering template is present at the start of the paragraph.
	' Reset the numbering on any specified larger heading.
	For Each p In ActiveDocument.Paragraphs
		Select Case p.OutlineLevel
			Case Is = 1
				If ResetArgumentNumberAtPocket = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 2
				If ResetArgumentNumberAtHat = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 3
				If ResetArgumentNumberAtBlock = True Then
					CurrentArgumentNumber = 0
				End If
			Case Is = 4
				If Len(p.Range.Text) >= TemplateLength Then
					Dim IsTheNumberingTemplatePresent As Boolean
					IsTheNumberingTemplatePresent = True
					Dim i As Integer
					For i = 1 to TemplateLength Step 1
						' Going character-by-character, compare the characters at the start of the paragraph with the characters in the TemplateToNumber to see if they are the same.
						If p.Range.Characters(i) <> Mid(TemplateToNumber, i, 1) Then
							IsTheNumberingTemplatePresent = False
						End If
					Next i
					If IsTheNumberingTemplatePresent = True Then
						CurrentArgumentNumber = CurrentArgumentNumber + 1
						Dim j As Integer
						For j = 1 to TemplateLength Step 1
							If p.Range.Characters(j) = NumberPlaceholder Then
								p.Range.Characters(j) = CurrentArgumentNumber
							End If
						Next j
					End If
				End If
		End Select
	Next p
	' End of code based on Verbatim 6.0.0's functions.
End Sub

' ---Read Doc Creator v2.3.1---
' Updated on 2024-09-27.
' This macro consists of 6 sub procedures.
' https://github.com/KSXia/Verbatim-Read-Doc-Creator
' Thanks to Truf for creating and providing the original code for activating invisibility mode! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/

' Sub procedure 1 of 6: Read Doc Creator Core
Sub CreateReadDoc(EnableInvisibilityMode As Boolean, EnableFastInvisibilityMode As Boolean)
	Dim CopyFormattedTitle As Boolean
	Dim AutomaticallySaveReadDoc As Boolean
	Dim AutomaticallyCloseSavedReadDoc As Boolean
	Dim DeleteStyles As Boolean
	Dim StylesToDelete() As Variant
	Dim DeleteLinkedCharacterStyles As Boolean
	Dim LinkedCharacterStylesToDelete() As Variant
	Dim DeleteForReferenceHighlightingInInvisibilityMode As Boolean
	Dim DeleteForReferenceCardHighlightingInNormalMode As Boolean
	Dim ForReferenceHighlightingColor As String
	Dim ReadDocNamePrefix As String
	Dim ReadDocNameSuffix As String
	
	' ---USER CUSTOMIZATION---
	' <<CUSTOMIZE THE SAVING MECHANISMS HERE!>>
	' If CopyFormattedTitle is set to True, this macro will copy the formatted name of the read doc into the clipboard.
	CopyFormattedTitle = False
	
	' If AutomaticallySaveReadDoc is set to True, this macro will automatically save the read doc.
	' WARNING: This feature to automatically save the read doc has LIMITED COMPATIBILITY! It might not work on MacOS.
	AutomaticallySaveReadDoc = True
	
	' If this macro is set to automatically save the read doc and AutomaticallyCloseSavedReadDoc is set to True, the read doc will automatically be closed after it is saved.
	AutomaticallyCloseSavedReadDoc = False
	
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to the list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Undertag")
	
	' If DeleteStyles is set to True, the styles listed in the StylesToDelete array will be deleted. If DeleteStyles is set to False, the styles listed in the StylesToDelete array will not be deleted.
	' If you want to disable the deletion of the styles listed in the StylesToDelete array, set DeleteStyles to False.
	DeleteStyles = True
	
	' <<SET THE LINKED CHARACTER STYLES TO DELETE HERE!>>
	' A linked style will either apply the style to the entire paragraph or a selection of words depending on what you have selected. If you have clicked on a paragraph and have selected no text or have selected the entire paragraph, it will apply the paragraph variant of the style. If you have selected a subset of the paragraph, it will apply the character variant of the style to your selection. The options in this section control whether this macro will delete the instances of character variants of linked styles and which linked styles this macro will operate on.
	
	' If DeleteLinkedCharacterStyles is set to True, the character variants of the linked styles listed in the LinkedCharacterStylesToDelete array will be deleted. If DeleteLinkedCharacterStyles is set to False, they will not be deleted.
	DeleteLinkedCharacterStyles = False
	
	' Add the names of linked styles that you want to delete the character variant of to the list in the LinkedCharacterStylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no character variants of linked styles will be deleted.
	LinkedCharacterStylesToDelete = Array()
	
	' <<SET WHETHER TO DELETE HIGHLIGHTED TEXT IN "For Reference" CARDS HERE!>>
	' If DeleteForReferenceCardsForInvisibilityMode is set to True, text highlighted in your "For Reference" highlighting color (which is set in the ForReferenceHighlightingColor option below) will be deleted when the read doc is set to have invisibility mode activated.
	DeleteForReferenceHighlightingInInvisibilityMode = False
	' If DeleteForReferenceCardsForNormalMode is set to True, text highlighted in your "For Reference" highlighting color (which is set in the ForReferenceHighlightingColor option below) will be deleted when the read doc is not set to have invisibility mode activated.
	DeleteForReferenceCardHighlightingInNormalMode = False
	
	' <<SET THE COLOR YOU USE FOR "For Reference" CARDS HERE!>>
	' Set ForReferenceHighlightingColor to the name of the highlighting color you use for "For Reference" cards.
	' WARNING: This highlighting color MUST ONLY be used for "For Reference" cards and nothing that you are reading! If this is not the case, DISABLE the function to delete highlighting for "For Reference" cards by setting DeleteForReferenceHighlightingInInvisibilityMode and DeleteForReferenceCardHighlightingInNormalMode to False.
	'
	' These are the names of the highlighting colors in the each row of the highlighting color selection menu, listed from left to right:
	' First row: Yellow, Bright Green, Turquoise, Pink, Blue
	' Second row: Red, Dark Blue, Teal, Green, Violet
	' Third row: Dark Red, Dark Yellow, Dark Gray, Light Gray, Black
	' MAKE SURE TO USE THIS EXACT CAPITALIZATION AND SPELLING!
	ForReferenceHighlightingColor = "Light Gray"
	
	' <<SET HOW THE READ DOC IS NAMED HERE!>>
	' Set ReadDocNamePrefix to the prefix you want to add to the read doc name.
	' Make sure there are quotation marks around the prefix you want to insert into the read doc name!
	' If you do not want to insert a prefix into the read doc name, put nothing in-between the quotation marks. If you do this, you MUST have a suffix for the read doc name.
	ReadDocNamePrefix = ""
	
	' Set ReadDocNameSuffix to the suffix you want to add to the read doc name.
	' Make sure there are quotation marks around the suffix you want to insert into the read doc name!
	' If you do not want to insert a suffix into the read doc name, put nothing in-between the quotation marks. If you do this, you MUST have a prefix for the read doc name.
	ReadDocNameSuffix = " [R]"
	
	' ---CHECK VALIDITY OF USER CONFIGURATION---
	' Check if there is either a prefix or suffix for the read doc name.
	If AutomaticallySaveReadDoc = True And ReadDocNamePrefix = "" And ReadDocNameSuffix = "" Then
		' If there is neither a prefix nor suffix for the read doc name:
		MsgBox "You have not set a suffix or prefix to add to the read doc name. Please set one in the macro settings and try again.", Title:="Error in Creating Read Doc"
		Exit Sub
	End If
	
	' ---INITIAL VARIABLE SETUP---
	Dim OriginalDoc As Document
	' Assign the original document to a variable.
	Set OriginalDoc = ActiveDocument
	
	' Check if the original document has previously been saved.
	If OriginalDoc.Path = "" Then
		' If the original document has not been previously saved:
		MsgBox "The current document must be saved at least once. Please save the current document and try again.", Title:="Error in Creating Read Doc"
		Exit Sub
	End If
	
	' Assign the original document name to a variable.
	Dim OriginalDocName As String
	OriginalDocName = OriginalDoc.Name
	
	' ---INITIAL GENERAL SETUP---
	' Disable screen updating for faster execution.
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---VARIABLE SETUP---
	Dim ReadDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the read doc.
	Set ReadDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	' ---STYLE DELETION SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present.
	On Error Resume Next
	
	' ---PRE-PROCESSING FOR STYLE DELETION---
	' Use Find and Replace to replace paragraph marks in the character variants of linked styles set for deletion with paragraph marks in Tag style.
	' This ensures all paragraph marks in lines or paragraphs that have character variants of linked styles set to be delted are in Tag style so they do not get deleted in the style deletion stage of this macro.
	' Otherwise, lines ending in character variants of linked styles set to be delted may have their paragraph mark deleted and have the following line be merged into them, which can mess up the formatting of the line.
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleNameToProcessIndex As Integer
		For CurrentLinkedCharacterStyleNameToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) = LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) & " Char"
		Next CurrentLinkedCharacterStyleNameToProcessIndex
		
		Dim CurrentLinkedCharacterStyleToProcessIndex As Integer
		For CurrentLinkedCharacterStyleToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToProcess As Style
			
			Set LinkedCharacterStyleToProcess = ReadDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToProcessIndex))
			
			With ReadDoc.Content.Find
				.ClearFormatting
				.Text = "^p"
				.Style = LinkedCharacterStyleToProcess
				.Replacement.ClearFormatting
				.Replacement.Text = "^p"
				.Replacement.Style = "Tag Char"
				.Format = True
				' Ensure various checks are disabled to have the search properly function.
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete.
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToProcessIndex
	End If
	
	' ---STYLE DELETION---
	If DeleteStyles = True Then
		Dim CurrentStyleToDeleteIndex As Integer
		For CurrentStyleToDeleteIndex = 0 To GreatestStyleIndex Step 1
			Dim StyleToDelete As Style
			
		' Specify the style to be deleted and delete it.
			Set StyleToDelete = ReadDoc.Styles(StylesToDelete(CurrentStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the specified style and delete it.
			With ReadDoc.Content.Find
				.ClearFormatting
				.Style = StyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization.
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete.
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentStyleToDeleteIndex
	End If
	
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleToDeleteIndex As Integer
		For CurrentLinkedCharacterStyleToDeleteIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToDelete As Style
			
			' Specify the linked style to delete the character variants of.
			Set LinkedCharacterStyleToDelete = ReadDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the character variants of the specified linked style and delete it.
			With ReadDoc.Content.Find
				.ClearFormatting
				.Style = LinkedCharacterStyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization.
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete.
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToDeleteIndex
	End If
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts.
	On Error GoTo 0
	
	' ---DELETE HIGHLIGHTED WORDS IN "For Reference" CARDS---
	If EnableInvisibilityMode = False And DeleteForReferenceCardHighlightingInNormalMode Then
		Call DeleteForReferenceCardHighlighting(ReadDoc, ForReferenceHighlightingColor)
	ElseIf EnableInvisibilityMode = True And DeleteForReferenceHighlightingInInvisibilityMode Then
		Call DeleteForReferenceCardHighlighting(ReadDoc, ForReferenceHighlightingColor)
	End If
	
	' ---DESTRUCTIVE INVISIBILITY MODE---
	If EnableInvisibilityMode And EnableFastInvisibilityMode Then
		Call EnableDestructiveInvisibilityMode(ReadDoc, True)
	ElseIf EnableInvisibilityMode Then
		Call EnableDestructiveInvisibilityMode(ReadDoc, False)
	End If
	
	' ---READ DOCUMENT TITLE COPIER---
	If CopyFormattedTitle = True Then
		Dim ClipboardText As DataObject
		
		' Set a variable to be the name of the read doc.
		Dim ReadDocName As String
		ReadDocName = ReadDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & ReadDocNameSuffix
		
		' Put the formatted name of the read doc into the clipboard.
		Set ClipboardText = New DataObject
		ClipboardText.SetText ReadDocName
		ClipboardText.PutInClipboard
	End If
	
	' ---SAVING THE READ DOC---
	If AutomaticallySaveReadDoc = True Then
		Dim SavePath As String
		SavePath = OriginalDoc.Path & "\" & ReadDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & ReadDocNameSuffix & ".docx"
		ReadDoc.SaveAs2 Filename:=SavePath, FileFormat:=wdFormatDocumentDefault
		
		If AutomaticallyCloseSavedReadDoc Then
			ReadDoc.Close SaveChanges:=wdSaveChanges
			MsgBox "The read doc is saved at " & SavePath, Title="Successfully Created and Saved Read Doc"
		End If
	End If
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts.
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub

' Sub procedure 2 of 6: Invisibility Mode Enabler
' Thanks to Truf for creating and providing the original code for activating invisibility mode! This sub procedure is based on Truf's "InvisibilityOn" and "InvisibilityOnFast" sub procedures. You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub EnableDestructiveInvisibilityMode(TargetDoc As Document, UseFastMode As Boolean)
	' Move the cursor to the beginning of the document.
	TargetDoc.Content.Select
	Selection.HomeKey Unit:=wdStory
	
	' Replace all paragraph marks with highlighted and bolded paragraph marks.
	With TargetDoc.Content.Find
		.ClearFormatting
		.MatchWildcards = False
		.Text = "^p"
		.Replacement.ClearFormatting
		.Replacement.Text = "^p"
		.Replacement.Style = "Underline"
		.Replacement.Font.Bold = True
		.Replacement.Highlight = True
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Normal" text.
	With TargetDoc.Content.Find
		.ClearFormatting
		.Text = ""
		.Style = "Normal"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Highlight = False
		.Font.Bold = False
		.Replacement.ClearFormatting
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Underline" text.
	With TargetDoc.Content.Find
		.ClearFormatting
		.Text = ""
		.Style = "Underline"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Highlight = False
		.Replacement.ClearFormatting
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Emphasis" text.
	With TargetDoc.Content.Find
		.ClearFormatting
		.Text = ""
		.Style = "Emphasis"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Highlight = False
		.Replacement.ClearFormatting
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove extra spaces between paragraph marks.
	With TargetDoc.Content.Find
		.ClearFormatting
		.Text = "^p ^p"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Replacement.ClearFormatting
		.Replacement.Text = ""
		.Replacement.Highlight = False
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove consecutive spaces in non-highlighted text.
	With TargetDoc.Content.Find
		.ClearFormatting
		.MatchWildcards = True
		.Text = "( ){2,}"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Highlight = False
		.Replacement.ClearFormatting
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove spaces at the beginning of paragraphs.
	With TargetDoc.Content.Find
		.ClearFormatting
		.MatchWildcards = False
		.Text = "^p "
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Replacement.ClearFormatting
		.Replacement.Text = "^p"
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove empty paragraphs by replacing consecutive paragraph marks with a single paragraph mark.
	With TargetDoc.Content.Find
		.ClearFormatting
		.MatchWildcards = True
		.Text = "^13{2,}"
		.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
		.Replacement.ClearFormatting
		.Replacement.Text = "^p"
		.Execute Replace:=wdReplaceAll
	End With
	
	If Not UseFastMode = True Then
		Dim DoesParagraphContainHighlighting As Boolean
		
		Dim ParagraphsToSkip() As Boolean
		
		Dim CharacterStylesToIgnore() As Variant
		CharacterStylesToIgnore = Array("Cite", "Tag Char", "Analytic Char")
		
		Dim GreatestParagraphIndex As Long
		GreatestParagraphIndex = 0
		Dim ParagraphToCount As Paragraph
		For Each ParagraphToCount In TargetDoc.Paragraphs
			GreatestParagraphIndex = GreatestParagraphIndex + 1
		Next ParagraphToCount
		
		ReDim ParagraphsToSkip(1 To GreatestParagraphIndex) As Boolean
		
		Dim GreatestCharacterStyleToIgnoreIndex As Integer
		GreatestCharacterStyleToIgnoreIndex = UBound(CharacterStylesToIgnore) - LBound(CharacterStylesToIgnore)
		
		On Error Resume Next
		
		Dim CurrentCharacterStyleToIgnoreIndex As Integer
		For CurrentCharacterStyleToIgnoreIndex = 0 To GreatestCharacterStyleToIgnoreIndex Step 1
			Dim CharacterStyleToIgnore As Style
			Set CharacterStyleToIgnore = TargetDoc.Styles(CharacterStylesToIgnore(CurrentCharacterStyleToIgnoreIndex))
			
			Dim RangeToScanForCharacterStyle As Range
			Set RangeToScanForCharacterStyle = TargetDoc.Range
			With RangeToScanForCharacterStyle.Find
				.ClearFormatting
				.Text = ""
				.Style = CharacterStyleToIgnore
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				.Wrap = wdFindStop
				Do While .Execute(Forward:=True) = True
					Dim CurrentParagraphIndex As Long
					CurrentParagraphIndex = TargetDoc.Range(0, RangeToScanForCharacterStyle.Paragraphs(1).Range.End).Paragraphs.Count
					ParagraphsToSkip(CurrentParagraphIndex) = True
				Loop
			End With
		Next CurrentCharacterStyleToIgnoreIndex
		
		Dim ParagraphStylesToIgnore() As Variant
		ParagraphStylesToIgnore = Array("Analytic", "Heading 4,Tag", "Heading 3,Block", "Heading 2,Hat", "Heading 1,Pocket", "Undertag")
		
		Dim ParagraphIndexToScanForStyles As Long
		ParagraphIndexToScanForStyles = 0
		Dim CurrentParagraph As Paragraph
		For Each CurrentParagraph In TargetDoc.Paragraphs
			ParagraphIndexToScanForStyles = ParagraphIndexToScanForStyles + 1
			Dim CurrentParagraphStyleToIgnore As Variant
			For Each CurrentParagraphStyleToIgnore In ParagraphStylesToIgnore
				If CurrentParagraph.Style = CurrentParagraphStyleToIgnore Then
					ParagraphsToSkip(ParagraphIndexToScanForStyles) = True
					Exit For
				End If
			Next CurrentParagraphStyleToIgnore
		Next CurrentParagraph
		
		On Error GoTo 0
		
		Dim ParagraphToInspect As Paragraph
		Dim RangeOfParagraphToInspect As Range
		
		Dim IsParagraphHighlighted() As Boolean
		ReDim IsParagraphHighlighted(1 To GreatestParagraphIndex) As Boolean
		
		Dim IsParagraphChecked() As Boolean
		ReDim IsParagraphChecked(1 To GreatestParagraphIndex) As Boolean
		
		Dim ParagraphIndex As Long
		ParagraphIndex = 1
		' Remove line breaks surrounded on both sides by highlighted text.
		For Each ParagraphToInspect In TargetDoc.Paragraphs
			If ParagraphIndex = GreatestParagraphIndex Then
				Exit For
			End If
		
			If ParagraphsToSkip(ParagraphIndex) = False And ParagraphsToSkip(ParagraphIndex+1) = False Then
				Set RangeOfParagraphToInspect = ParagraphToInspect.Range
				RangeOfParagraphToInspect.MoveEnd wdCharacter, -1 ' Ignore the paragraph mark.
				
				Dim CharacterIndexToInspect As Long
				' Check if the current paragraph contains highlighted text.
					DoesParagraphContainHighlighting = False
					If IsParagraphHighlighted(ParagraphIndex) = True Then
						DoesParagraphContainHighlighting = True
					ElseIf IsParagraphChecked(ParagraphIndex) = False Then
						If RangeOfParagraphToInspect.HighlightColorIndex <> wdNoHighlight Then
							DoesParagraphContainHighlighting = True
						End If
					End If
				
				If DoesParagraphContainHighlighting = True Then
					' Check if the next paragraph exists and contains highlighted text.
					Dim DoesFollowingParagraphContainHighlighting As Boolean
					DoesFollowingParagraphContainHighlighting = False
					If Not ParagraphToInspect.Next Is Nothing Then
						Dim RangeOfFollowingParagraphToInspectForHighlighting As Range
						Set RangeOfFollowingParagraphToInspectForHighlighting = ParagraphToInspect.Next.Range
						RangeOfFollowingParagraphToInspectForHighlighting.MoveEnd wdCharacter, -1 ' Ignore the paragraph mark.
						
						If RangeOfFollowingParagraphToInspectForHighlighting.HighlightColorIndex <> wdNoHighlight Then
							DoesFollowingParagraphContainHighlighting = True
								IsParagraphHighlighted(ParagraphIndex+1) = True
						End If
						IsParagraphChecked(ParagraphIndex+1) = True
					End If
					
					' If both paragraphs contain highlighted text, join them.
					If DoesParagraphContainHighlighting = True And DoesFollowingParagraphContainHighlighting = True Then
						RangeOfParagraphToInspect.InsertAfter " " ' Insert a space after the current paragraph.
						ParagraphToInspect.Range.Characters.Last.Delete ' Delete the paragraph mark
						
						' Remove any consecutive non-highlighted spaces the inserted space may have formed.
						With TargetDoc.Content.Find
							.ClearFormatting
							.MatchWildcards = True
							.Text = "( ){2,}"
							.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
							.Highlight = False
							.Replacement.ClearFormatting
							.Replacement.Text = " "
							.Execute Replace:=wdReplaceAll
						End With
					End If
				End If
			End If
			ParagraphIndex = ParagraphIndex + 1
		Next ParagraphToInspect
	End If
	
	' Clean up modified find and replace settings.
	TargetDoc.Content.Find.ClearFormatting
	TargetDoc.Content.Find.MatchWildcards = False
	TargetDoc.Content.Find.Replacement.ClearFormatting
	
	' Suppress grammar check and spell check.
	TargetDoc.ShowGrammaticalErrors = False
	TargetDoc.ShowSpellingErrors = False
End Sub

' Sub procedure 3 of 6: Delete Highlighting in "For Reference" Cards
Sub DeleteForReferenceCardHighlighting(TargetDoc As Document, ForReferenceHighlightingColor As String)
	Dim ForReferenceHighlightingColorEnum As Long
	' The following code for converting highlighting color name to enum is a modified version of Verbatim 6.0.0's "Standardize Highlighting With Exception" functon.
	Select Case ForReferenceHighlightingColor
		' Common highlighting colors:
		Case Is = "Turquoise"
			ForReferenceHighlightingColorEnum = wdTurquoise
		Case Is = "Bright Green"
			ForReferenceHighlightingColorEnum = wdBrightGreen
		Case Is = "Yellow"
			ForReferenceHighlightingColorEnum = wdYellow
		
		' Common rehighlighting colors:
		Case Is = "Pink"
			ForReferenceHighlightingColorEnum = wdPink
		Case Is = "Red"
			ForReferenceHighlightingColorEnum = wdRed
		
		' Common "For Reference" highlighting colors:
		Case Is = "Light Gray"
			ForReferenceHighlightingColorEnum = wdGray25
		Case Is = "Dark Gray"
			ForReferenceHighlightingColorEnum = wdGray50
		
		' Other high-contrast highlighting color(s):
		Case Is = "Dark Yellow"
			ForReferenceHighlightingColorEnum = wdDarkYellow
		
		' Other highlighting colors:
		Case Is = "Blue"
			ForReferenceHighlightingColorEnum = wdBlue
		Case Is = "Dark Blue"
			ForReferenceHighlightingColorEnum = wdDarkBlue
		Case Is = "Teal"
			ForReferenceHighlightingColorEnum = wdTeal
		Case Is = "Green"
			ForReferenceHighlightingColorEnum = wdGreen
		Case Is = "Dark Red"
			ForReferenceHighlightingColorEnum = wdDarkRed
		Case Is = "Violet"
			ForReferenceHighlightingColorEnum = wdViolet
		Case Is = "Black"
			ForReferenceHighlightingColorEnum = wdBlack
		Case Is = "White"
			ForReferenceHighlightingColorEnum = wdWhite
		
		' No highlighting color:
		Case Is = "None"
			Exit Sub
		
		' Other cases:
		Case Else
			' If the highlighting color name is not a name of any of Word's highlighting colors:
			' ForReferenceHighlightingColorEnum = wdNoHighlight
			Exit Sub
	End Select
	' End of code based on Verbatim 6.0.0's functions.
	
	With TargetDoc.Content
		With .Find
			.ClearFormatting
			.Text = ""
			.Highlight = True
			.Replacement.ClearFormatting
			.Replacement.Text = ""
			.Format = True
			' Disable checks in the find process for optimization.
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			' Modify the search process settings.
			.Forward = True
			.Wrap = wdFindStop
			End With
			' Delete all text with the "For Reference" highlighting color.
			Do While .Find.Execute = True
				If .HighlightColorIndex = ForReferenceHighlightingColorEnum Then .Delete
			Loop
	End With
End Sub

' Sub procedure 4 of 6: Trigger for Read Doc Creator
Sub CreateNormalReadDoc()
	Call CreateReadDoc(False, False)
End Sub

' Sub procedure 5 of 6: Trigger for Read Doc Creator
Sub CreateReadDocWithInvisibilityMode()
	Call CreateReadDoc(True, False)
End Sub

' Sub procedure 6 of 6: Trigger for Read Doc Creator
Sub CreateReadDocWithFastInvisibilityMode()
	Call CreateReadDoc(True, True)
End Sub
' <<END Read Doc Creator>>

' ---Send Doc Creator v3.0.0---
' Updated on 2024-08-23.
' https://github.com/KSXia/Verbatim-Send-Doc-Creator
' Thanks to Truf for creating and providing the original "Create Send Doc" macro this macro is based on! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub CreateSendDoc()
	Dim CopyFormattedTitle As Boolean
	Dim AutomaticallySaveSendDoc As Boolean
	Dim AutomaticallyCloseSavedSendDoc As Boolean
	Dim DeleteStyles As Boolean
	Dim StylesToDelete() As Variant
	Dim DeleteLinkedCharacterStyles As Boolean
	Dim LinkedCharacterStylesToDelete() As Variant
	Dim SendDocNamePrefix As String
	Dim SendDocNameSuffix As String
	
	' ---USER CUSTOMIZATION---
	' <<CUSTOMIZE THE SAVING MECHANISMS HERE!>>
	' If CopyFormattedTitle is set to True, this macro will copy the formatted name of the send doc into the clipboard.
	CopyFormattedTitle = False
	
	' If AutomaticallySaveSendDoc is set to True, this macro will automatically save the send doc.
	' WARNING: This feature to automatically save the send doc has LIMITED COMPATIBILITY! It might not work on MacOS.
	AutomaticallySaveSendDoc = True
	
	' If this macro is set to automatically save the send doc and AutomaticallyCloseSavedSendDoc is set to True, the send doc will automatically be closed after it is saved.
	AutomaticallyCloseSavedSendDoc = True
	
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to the list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Undertag", "Analytic")
	
	' If DeleteStyles is set to True, the styles listed in the StylesToDelete array will be deleted. If DeleteStyles is set to False, the styles listed in the StylesToDelete array will not be deleted.
	' If you want to disable the deletion of the styles listed in the StylesToDelete array, set DeleteStyles to False.
	DeleteStyles = True
	
	' <<SET THE LINKED CHARACTER STYLES TO DELETE HERE!>>
	' A linked style will either apply the style to the entire paragraph or a selection of words depending on what you have selected. If you have clicked on a paragraph and have selected no text or have selected the entire paragraph, it will apply the paragraph variant of the style. If you have selected a subset of the paragraph, it will apply the character variant of the style to your selection. The options in this section control whether this macro will delete the instances of character variants of linked styles and which linked styles this macro will operate on.
	
	' If DeleteLinkedCharacterStyles is set to True, the character variants of the linked styles listed in the LinkedCharacterStylesToDelete array will be deleted. If DeleteLinkedCharacterStyles is set to False, they will not be deleted.
	DeleteLinkedCharacterStyles = True
	
	' Add the names of linked styles that you want to delete the character variant of to the list in the LinkedCharacterStylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no character variants of linked styles will be deleted.
	LinkedCharacterStylesToDelete = Array("Analytic")
	
	' <<SET HOW THE SEND DOC IS NAMED HERE!>>
	' Set SendDocNamePrefix to the prefix you want to add to the send doc name.
	' Make sure there are quotation marks around the prefix you want to insert into the send doc name!
	' If you do not want to insert a prefix into the send doc name, put nothing in-between the quotation marks. If you do this, you MUST have a suffix for the send doc name.
	SendDocNamePrefix = ""
	
	' Set SendDocNameSuffix to the suffix you want to add to the send doc name.
	' Make sure there are quotation marks around the suffix you want to insert into the send doc name!
	' If you do not want to insert a suffix into the send doc name, put nothing in-between the quotation marks. If you do this, you MUST have a prefix for the send doc name.
	SendDocNameSuffix = " [S]"
	
	' ---CHECK VALIDITY OF USER CONFIGURATION---
	' Check if there is either a prefix or suffix for the send doc name
	If SendDocNamePrefix = "" And SendDocNameSuffix = "" Then
		' If there is neither a prefix nor suffix for the send doc name:
		MsgBox "You have not set a suffix or prefix to add to the send doc name. Please set one in the macro settings and try again.", Title:="Error in Creating Send Doc"
		Exit Sub
	End If
	
	' ---INITIAL VARIABLE SETUP---
	Dim OriginalDoc As Document
	' Assign the original document to a variable
	Set OriginalDoc = ActiveDocument
	
	' Check if the original document has previously been saved
	If OriginalDoc.Path = "" Then
		' If the original document has not been previously saved:
		MsgBox "The current document must be saved at least once. Please save the current document and try again.", Title:="Error in Creating Send Doc"
		Exit Sub
	End If
	
	' Assign the original document name to a variable
	Dim OriginalDocName As String
	OriginalDocName = OriginalDoc.Name
	
	' ---INITIAL GENERAL SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---VARIABLE SETUP---
	Dim SendDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the send doc
	Set SendDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	Dim GreatestLinkedCharacterStyleIndex As Integer
	GreatestLinkedCharacterStyleIndex = UBound(LinkedCharacterStylesToDelete) - LBound(LinkedCharacterStylesToDelete)
	
	' ---STYLE DELETION SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present
	On Error Resume Next
	
	' ---PRE-PROCESSING FOR STYLE DELETION---
	' Use Find and Replace to replace paragraph marks in the character variants of linked styles set for deletion with paragraph marks in Tag style.
	' This ensures all paragraph marks in lines or paragraphs that have character variants of linked styles set to be delted are in Tag style so they do not get deleted in the style deletion stage of this macro.
	' Otherwise, lines ending in character variants of linked styles set to be delted may have their paragraph mark deleted and have the following line be merged into them, which can mess up the formatting of the line.
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleNameToProcessIndex As Integer
		For CurrentLinkedCharacterStyleNameToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) = LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) & " Char"
		Next CurrentLinkedCharacterStyleNameToProcessIndex
		
		Dim CurrentLinkedCharacterStyleToProcessIndex As Integer
		For CurrentLinkedCharacterStyleToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToProcess As Style
			
			Set LinkedCharacterStyleToProcess = SendDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToProcessIndex))
			
			With SendDoc.Content.Find
				.ClearFormatting
				.Text = "^p"
				.Style = LinkedCharacterStyleToProcess
				.Replacement.ClearFormatting
				.Replacement.Text = "^p"
				.Replacement.Style = "Tag Char"
				.Format = True
				' Ensure various checks are disabled to have the search properly function
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToProcessIndex
	End If
	
	' ---STYLE DELETION---
	If DeleteStyles = True Then
		Dim CurrentStyleToDeleteIndex As Integer
		For CurrentStyleToDeleteIndex = 0 to GreatestStyleIndex Step 1
			Dim StyleToDelete As Style
			
			' Specify the style to be deleted
			Set StyleToDelete = SendDoc.Styles(StylesToDelete(CurrentStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the specified style and delete it
			With SendDoc.Content.Find
				.ClearFormatting
				.Style = StyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentStyleToDeleteIndex
	End If
	
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleToDeleteIndex As Integer
		For CurrentLinkedCharacterStyleToDeleteIndex = 0 to GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToDelete As Style
			
			' Specify the linked style to delete the character variants of
			Set LinkedCharacterStyleToDelete = SendDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the character variants of the specified linked style and delete it
			With SendDoc.Content.Find
				.ClearFormatting
				.Style = LinkedCharacterStyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToDeleteIndex
	End If
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts
	On Error GoTo 0
	
	' ---SEND DOCUMENT TITLE COPIER---
	If CopyFormattedTitle = True Then
		Dim ClipboardText As DataObject
		
		' Set a variable to be the name of the send doc
		Dim SendDocName As String
		SendDocName = SendDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & SendDocNameSuffix
		
		' Put the name of the send doc into the clipboard
		Set ClipboardText = New DataObject
		ClipboardText.SetText SendDocName
		ClipboardText.PutInClipboard
	End If
	
	' ---SAVING THE SEND DOC---
	If AutomaticallySaveSendDoc = True Then
		Dim SavePath As String
		SavePath = OriginalDoc.Path & "\" & SendDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & SendDocNameSuffix & ".docx"
		SendDoc.SaveAs2 Filename:=SavePath, FileFormat:=wdFormatDocumentDefault
		
		If AutomaticallyCloseSavedSendDoc = True Then
			SendDoc.Close SaveChanges:=wdSaveChanges
			MsgBox "The send doc is saved at " & SavePath, Title="Successfully Created and Saved Send Doc"
		End If
	End If
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub

' ---Format As For Reference Card Macro - Stable Edition - v1.1.0---
' Updated on 2024-09-01.
' https://github.com/KSXia/Verbatim-Format-As-For-Reference-Card-Macro---Stable-Edition
' Thanks to Truf for creating and providing his "ForReference" macro, which the idea for this macro came from! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub FormatAsForReferenceCard()
	' Check if any text is selected.
	If Selection.Type = wdSelectionIP Then
		MsgBox "You have not selected any text." & vbNewLine & "Please select the text you want" & vbNewLine & "to format as a ""For Reference"" card.", Title:="Error in Formatting as" & vbNewLine & "a ""For Reference"" Card"
        Exit Sub
	End If
	
	Dim SelectionRange As Range
	Set SelectionRange = Selection.Range
	
	If Len(SelectionRange.Text) > 1 Then
		' Record the user's default highlight color.
		Dim UserDefaultHighlightColor As Long
		UserDefaultHighlightColor = Options.DefaultHighlightColorIndex
		
		' Set the default highlight color to the "For Reference" highlight color.
		Options.DefaultHighlightColorIndex = wdGray25
		
		' Find all highlighted characters and replace their highlight color with the default highlight color, which should be set to the "For Reference" highlight color.
		With SelectionRange.Find
			' Specify find criteria.
			.ClearFormatting
			.MatchWildcards = True
			.Text = "*"
			.Highlight = True
			
			' Ensure other find options are set to their defaults.
			.MatchCase = False
			.MatchWholeWord = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.MatchPrefix = False
			.MatchSuffix = False
			.MatchPhrase = False
			
			' Specify replacement criteria.
			.Replacement.ClearFormatting
			.Replacement.Text = ""
			.Replacement.Highlight = True
			
			' Set execution properties.
			.Format = True
			.Forward = True
			.Wrap = wdFindStop
			.Execute Replace:=wdReplaceAll
		End With
		
		' Reset the default highlight color back to the user's default highlight color.
		Options.DefaultHighlightColorIndex = UserDefaultHighlightColor
	ElseIf Len(SelectionRange.Text) = 1 Then
		If SelectionRange.HighlightColorIndex <> wdNoHighlight Then
			SelectionRange.HighlightColorIndex = wdGray25
		End If
	End If
End Sub
' <<END Debate Utility Booster Suite>>
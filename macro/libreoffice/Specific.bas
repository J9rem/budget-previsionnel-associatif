Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' constantes
Public Const Nom_de_Fichier_Par_Defaut As String = "InCitu_Budget_Previsionnel_Associatif_LibreOffice"
Public Const BackupDefaultExtension As String = ".ods"
Public Const intoOds As Boolean = True

Public ODialog
Public publicDoc

' types
Type Financement
    Nom As String
    TypeFinancement As Integer ' Index in TypeFinancements
    Valeur As Double
    Statut As Integer ' 0 = empty
    BaseCell As Range
End Type

Type FinancementComplet
    Financements() As Financement
    Status As Boolean
End Type

' utils
Public Function choisirNomFicherASauvegarderSansMacro(ByRef FilePath As String) As Boolean
    
    Dim Adresse_dossier_courant As String
    Dim Default_File_Name As String
    Dim Fichier_De_Sauvegarde As String
    
    ' Default FileName
    Default_File_Name = Nom_de_Fichier_Par_Defaut & "_" & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hh-mm")
    ' Default
    FilePath = ""
    
    ' Changement de dossier
    Adresse_dossier_courant =ThisWorkbook.path
    ChDir (Adresse_dossier_courant)
    
    ' Fenêtre pour demander le nom du fichier de sauvegarde
    On Error Resume Next
    ' InitialFileName, FileFilter, FiltrerIndex, Title
    Fichier_De_Sauvegarde = GetSaveAsFilename( _
        Default_File_Name, _
        Array(Array("Libre Office (*.ods)","*.ods"),Array("Excel (*.xls)","*.xls")), _
        "Choisir le fichier à exporter", _
        Adresse_dossier_courant)
    On Error GoTo 0
    If Fichier_De_Sauvegarde = "" Or Fichier_De_Sauvegarde = Empty Or Fichier_De_Sauvegarde = "Faux" Or Fichier_De_Sauvegarde = "False" Then
        choisirNomFicherASauvegarderSansMacro = False
    Else
        FilePath = Fichier_De_Sauvegarde
        choisirNomFicherASauvegarderSansMacro = True
    End If

End Function

Public function GetSaveAsFilename(DefaultFileName as string, Filters, Title as String, MainDir as String) as String
	Dim oFilePicker 
	Dim sFilePickerArgs 
	Dim Index As integer
	Dim files
	
	GetSaveAsFilename = ""
	sFilePickerArgs = Array(_
    	com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_AUTOEXTENSION )    
	
	oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.OfficeFilePicker")
	oFilePicker.initialize(sFilePickerArgs())
    oFilePicker.setMultiSelectionMode(false)
	oFilePicker.DisplayDirectory = ConvertToURL(MainDir)
    oFilePicker.setDefaultName(DefaultFileName)
	oFilePicker.setTitle (Title)
	For Index = LBound(Filters) To UBound(Filters)
		oFilePicker.appendFilter(Filters(Index)(0), Filters(Index)(1) )
	Next Index
	if (oFilePicker.execute) Then
		files = oFilePicker.getFiles()
		if UBound(files) > -1 AND len(files(0)) > 0 Then
			GetSaveAsFilename = ConvertFromUrl(files(0))
		End If
	End If
End Function

Public Function saveWorkBookAsCopyNoMacro(FilePath As String, FileName As String, Ext As String, ByRef OpenedWorkbook As Workbook) As Boolean

    Dim DestFolder As String
    Dim wb As Workbook
    Dim doc 
    
    saveWorkBookAsCopyNoMacro = False
    
    DestFolder = Left(FilePath, Len(FilePath) - Len(FileName))
    
    ChDir (DestFolder)
    ThisWorkbook.SaveCopyAs FileName:=FilePath
    
    On Error Resume Next
    Set Wb = Workbooks.Open(FileName:=DestFolder & FileName, ReadOnly:=False)
    returnToCurrentPath
    Kill FilePath
    On Error GoTo 0
    Wb.SaveAs FileName:=FilePath
    If Not (Wb Is Nothing) Then
    	OpenedWorkbook = Wb
    	removeShapes OpenedWorkbook
	    If FindDoc(doc,FileName) Then
	    	RemoveMacro(doc)
	    	Wb.save
	    	doc.close(true)
            If FileExists(DestFolder & FileName) Then
                saveWorkBookAsCopyNoMacro = True
       		End If
	    End If
    End If
End Function

Public Function FindDoc(ByRef Doc, FileName as String) As Boolean
	Dim openDoc
	FindDoc = False
	For Each openDoc In StarDesktop.Components
	    If (openDoc.Title = FileName) Then
	    	Doc = openDoc
	    	FindDoc = True
	    End If
	Next
End Function

Public Function RemoveMacro(doc) as boolean
	Dim libraries
	Dim names
	Dim namesLevel2
	Dim currentName as string
	Dim currentSubName as string
	Dim currentLib
	Dim index as integer
	Dim index2 as integer
	
	libraries = doc.BasicLibraries
	names = libraries.ElementNames
	For index = LBound(names) To Ubound(names)
		currentName = names(index)
		currentLib = libraries.getByName(currentName)
		If currentName = "Standard" Then
			namesLevel2 = currentLib.ElementNames
			For index2 = LBound(namesLevel2) To Ubound(namesLevel2)
				currentSubName = namesLevel2(index2)
				currentLib.removeByName(currentSubName)
			Next			
		Else
			libraries.removeLibrary(currentName)
		End IF
	Next
End Function

Sub DeleteFile(FileName As String, FolderName As String)
	Dim Doc
	
	If FindDoc(Doc, FileName) Then
		Doc.close(true)
	End If
    ChDir (FolderName)
    Kill FolderName & FileName
End Sub

Public Sub SaveCopyAs(Wb As Workbook, FileName As String, FolderName As String)
    Dim Spacer as String
    ChDir (FolderName)
    If Right(FolderName,1) = "\" Then
    Spacer = ""
    Else
    	If Right(FolderName,1) = "/" Then
    		Spacer = ""
    	Else
    		If InStr(FolderName,"\") Then
    			Spacer = "\"
    		Else
    			Spacer = "/"
    		End If
    	End If
    End If
    Wb.SaveCopyAs FolderName & Spacer & FileName
    Wb.Save
End Sub

Public Function removeShapes(Wb As Workbook) As Boolean
    ' do nothing not compatible
    removeShapes = True
End Function

Public Function FindOpenedWorkBook(FileName As String, ByRef OpenedWorkbook As Workbook) As Boolean

    Dim Index As Integer
    
    Set OpenedWorkbook = Nothing
    For Index = 1 To Workbooks.Count
        If Workbooks(Index).Name = FileName Then
            Set OpenedWorkbook = Workbooks(Index)
        End If
    Next Index
    If Not OpenedWorkbook Is Nothing Then
        FindOpenedWorkBook = True
    Else
        FindOpenedWorkBook = False
    End If
End Function

Public Function TypeFinancementsFromWb(wb As Workbook)

    TypeFinancementsFromWb = TypeFinancements()
    
End Function

Public Sub CleanLineStylesForBudget(oCellRange, BaseCell As Range, HeadCell As Range, IsHeader As Boolean)
	Dim oNoLine As New com.sun.star.table.BorderLine2
	With oNoLine
		.Color = 0
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.NONE
		.LineWidth = 0
		.OuterLineWidth = 0
	End With
	If Not IsHeader Then
		oCellRange.BottomBorder = oNoLine
	
		If HeadCell.Row <> BaseCell.Row - 1 Then
			oCellRange.TopBorder = oNoLine
		End If
	End If
End Sub

Public Sub DefineLineStylesForBudget( _
		oCellRange, _
        BaseCell As Range, _
        HeadCell As Range, _
        IsHeader As Boolean _
    )
	Dim oSheet
    Dim oCellRange2
	Dim oLine As New com.sun.star.table.BorderLine2
    
    oSheet = ThisComponent.Sheets.getByName(BaseCell.Worksheet.Name)
    
	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 26
		.OuterLineWidth = 26
	End With
	oCellRange.LeftBorder = oLine
	oCellRange.RightBorder = oLine

	If IsHeader Then
		oCellRange2 = oSheet.getCellByPosition(BaseCell.Column-1,BaseCell.Row-2)
		oCellRange2.BottomBorder = oLine
	End If
	If IsHeader Or HeadCell.Row = BaseCell.Row - 1 Then
		oCellRange.TopBorder = oLine
	End If
End Sub

Public Sub SetFormatForBudget(BaseCell As Range, HeadCell As Range, IsHeader As Boolean)

    Dim IndexBis As Integer
    Dim oSheet
    Dim oCellRange
    Dim oFormat As Long
    
    oSheet = ThisComponent.Sheets.getByName(BaseCell.Worksheet.Name)
    
    For IndexBis = 1 To 3
    	oCellRange = oSheet.getCellByPosition(BaseCell.Column+IndexBis-2,BaseCell.Row-1)
		CleanLineStylesForBudget oCellRange, BaseCell, HeadCell, IsHeader
		DefineLineStylesForBudget oCellRange, BaseCell, HeadCell, IsHeader
    	oCellRange.CharFontStyleName = ""
    	oCellRange.CharFontPitch = 2
    	oCellRange.CharFontCharSet = -1
    	oCellRange.CharFontFamily = 5
    	oCellRange.CharFontName = "Calibri"
		If IsHeader Then
    		oCellRange.CharColor = RGB(255,255,255)
    		oCellRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
    		oCellRange.CellBackColor = RGB(164,164,164)
		Else
    		oCellRange.CharColor = -1
    		oCellRange.CharWeight = 100
    		oCellRange.CellBackColor = -1
		End If
    	oCellRange.CharHeight = 8
		
		If IndexBis = 1 Or (IndexBis = 3 And IsHeader) Then
			oCellRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
		Else
			oCellRange.HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
		End If
			oCellRange.VertJustify  = com.sun.star.table.CellVertJustify.TOP
		If IndexBis = 3 Then
			oFormat = CellSetNumberFormat("# ##0,00"" €""",ThisComponent)
			oCellRange.NumberFormat = oFormat
		Else
			oCellRange.NumberFormat = 0
		End If
    	
    Next IndexBis
End Sub

Public Function CellSetNumberFormat(stNumberFormat As String, oDoc As Object) AS Long
	DIM aLocale	AS New com.sun.star.lang.Locale
	DIM oNumberFormats AS Object
	DIM loFormatKey	AS Long
	oNumberFormats = oDoc.getNumberFormats()
	loFormatKey = oNumberFormats.queryKey(stNumberFormat, aLocale, FALSE)
	IF loFormatKey = -1 THEN loFormatKey = oNumberFormats.addNew(stNumberFormat, aLocale)
	CellSetNumberFormat = loFormatKey
End Function

Public Sub Validate_Click(document)
    Dim Nom As String
    Dim TypeFinancement As Integer
    Dim FinancementFantome As FinancementComplet
    Dim TypesFinancements() As Strings
    Dim parentWindow
    Dim currentType As String
    Dim context
    
    Dim CurrentNBChantier As Integer
    Dim wb As Workbook
    
    TypesFinancements = TypeFinancementsFromWb(ThisWorkbook)
    context = document.Source.AccessibleContext
    parentWindow= context.AccessibleParent
    currentType = GetType(parentWindow)
    
    Nom = ""
    On Error Resume Next
    Nom = publicDoc.Source.Text
    On Error GoTo 0
    
    If CurrentType <> "" Then
    	TypeFinancement = FindTypeFinancementIndex(CurrentType)
    Else
        TypeFinancement = 0
    End If
    
 	oDialog.endExecute()
    If Nom = "" Or Nom = Empty Then
        MsgBox "Le champ texte NOM ne peut être vide !"
        Exit Sub
    End If
    If TypeFinancement = 0 Then
        MsgBox "Un type de financement doit être choisit !"
        Exit Sub
    End If
    
    Set wb = ThisWorkbook
    
    ' Current NB
    CurrentNBChantier = GetNbChantiers(wb)
    
    If CurrentNBChantier < 1 Then
        Exit Sub
    End If
    
    FinancementFantome.Status = False
    Chantiers_Financements_Add_One wb, CurrentNBChantier, FinancementFantome, Nom, TypeFinancement
    
End Sub

Public Sub OpenUserForm()

	Dim oLib
	Dim oLibDlg
	
	DialogLibraries.loadLibrary("Standard")
	oLib = DialogLibraries.getByName("Standard")
	oLibDlg = oLib.getByName("UserForm1")
	oDialog = CreateUnoDialog(oLibDlg)
	oDialog.execute()
	
End Sub

Public Function GetType(parentWindow) As String

	Dim internalWindows
	Dim currentWindow
	Dim state as Boolean
	Dim curName as String
	
	GetType = ""
	
	internalWindows = parentWindow.Windows
	For Each currentWindow In internalWindows
		On Error Resume Next
		state = currentWindow.State
		curName = currentWindow.AccessibleContext.AccessibleName
		On Error Goto 0
		If state And Left(curName,5) <> "Barre" Then
			GetType = curName
			Exit Function
		End If
	Next currentWindow

End Function

Public Sub SaveDoc(document)
	publicDoc = document
End Sub

Public Sub EnleverBordures(CurrentCell As Range)

    Dim oSheet
    Dim oCellRange
	Dim oLine As New com.sun.star.table.BorderLine2

	oSheet = ThisComponent.Sheets.getByName(CurrentCell.Worksheet.Name)
    oCellRange = oSheet.getCellByPosition(CurrentCell.Column-1,CurrentCell.Row-1)

	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.NONE
		.LineWidth = 0
		.OuterLineWidth = 0
	End With
	
	oCellRange.LeftBorder = oLine
	oCellRange.RightBorder = oLine
	oCellRange.BottomBorder = oLine
	oCellRange.TopBorder = oLine
End Sub

Public Sub DefinirFormatPourChantier( _
        CurrentCell As Range, _
        AddTopBorder As Boolean, _
        AddBottomBorder As Boolean, _
		Bold As Boolean, _
		Italic As Boolean, _
		BlueColor As Boolean, _
		CurrencyFormat As Boolean _
    )
	
    Dim oCellRange
	Dim oFormat As Long
	Dim oLine As New com.sun.star.table.BorderLine2
	Dim oLineThin As New com.sun.star.table.BorderLine2
    Dim oSheet
    
    oSheet = ThisComponent.Sheets.getByName(CurrentCell.Worksheet.Name)
    oCellRange = oSheet.getCellByPosition(CurrentCell.Column-1,CurrentCell.Row-1)
    
	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 75
		.OuterLineWidth = 75
	End With
	With oLineThin
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 5
		.OuterLineWidth = 5
	End With
	oCellRange.LeftBorder = oLine
	oCellRange.RightBorder = oLine
    
    If AddTopBorder Then
		oCellRange.TopBorder = oLine
    Else
		oCellRange.TopBorder = oLineThin
    End If
    If AddBottomBorder Then
		oCellRange.BottomBorder = oLine
    Else
		oCellRange.BottomBorder = oLineThin
    End If

    If Italic Then
		oCellRange.CharFontStyleName = "Italic"
    Else
		oCellRange.CharFontStyleName = ""
    End If
	oCellRange.CharFontPitch = 2
	oCellRange.CharFontCharSet = -1
	oCellRange.CharFontFamily = 5
	oCellRange.CharFontName = "Arial"
	oCellRange.CharColor = RGB(0,0,0)
	If Bold Then
		oCellRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
    Else
		oCellRange.CharWeight = com.sun.star.awt.FontWeight.NORMAL
    End If
	oCellRange.CellBackColor = -1
	oCellRange.CharHeight = 8
	oCellRange.HoriJustify = com.sun.star.table.CellHoriJustify.STANDARD
	oCellRange.VertJustify  = com.sun.star.table.CellVertJustify.TOP
	If CurrencyFormat Then
		oFormat = CellSetNumberFormat("# ##0,00"" €""",ThisComponent)
		oCellRange.NumberFormat = oFormat
	End If
	oCellRange.CellBackColor = -1
End Sub

Public Sub CopieLogo(oldWorkbook As Workbook, NewWorkbook As Workbook, Name As String)

   ' Do Nothing

End Sub

Public Sub formatChargeCell(CurrentCell As Range, NoBorderOnRightAndLeft As Boolean)

    Dim IndexBis
    Dim oSheet
    Dim oCellRange
    Dim oFormat As Long
	Dim oLine As New com.sun.star.table.BorderLine2
	Dim oNoLine As New com.sun.star.table.BorderLine2

	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 26
		.OuterLineWidth = 26
	End With
	With oNoLine
		.Color = 0
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.NONE
		.LineWidth = 0
		.OuterLineWidth = 0
	End With
    
    oSheet = ThisComponent.Sheets.getByName(CurrentCell.Worksheet.Name)
    For IndexBis = 1 To (ColumnOfSecondPartInCharge + NBCatOfCharges * 2)
	    oCellRange = oSheet.getCellByPosition(CurrentCell.Column+IndexBis-2,CurrentCell.Row-1)
	    
		If IndexBis = (ColumnOfSecondPartInCharge - 1) Then
			oCellRange.TopBorder = oNoLine
			oCellRange.BottomBorder= oNoLine
		Else
			oCellRange.TopBorder = oLine
			oCellRange.BottomBorder= oLine
		End If
	    
	    If NoBorderOnRightAndLeft Then
			oCellRange.LeftBorder = oNoLine
			oCellRange.RightBorder = oNoLine
	    Else
			oCellRange.LeftBorder = oLine
			oCellRange.RightBorder = oLine
	    End If

		oCellRange.CharFontStyleName = ""
    	oCellRange.CharFontPitch = 2
    	oCellRange.CharFontCharSet = -1
    	oCellRange.CharFontFamily = 5
    	oCellRange.CharFontName = "Calibri"
		oCellRange.CharColor = -1
		oCellRange.CharWeight = 100
		oCellRange.CellBackColor = -1
    	oCellRange.CharHeight = 8
		oCellRange.HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
		oCellRange.VertJustify  = com.sun.star.table.CellVertJustify.TOP
		If IndexBis > 1 _
			And IndexBis <> (ColumnOfSecondPartInCharge - 1) _
			And IndexBis <> ColumnOfSecondPartInCharge  Then
			If IndexBis = 6 Then
				oFormat = CellSetNumberFormat("0"" ""%",ThisComponent)
				oCellRange.NumberFormat = oFormat
			Else
				oFormat = CellSetNumberFormat("# ##0,00"" €""",ThisComponent)
				oCellRange.NumberFormat = oFormat
			End If
		Else
			oCellRange.NumberFormat = 0
		End If
	Next IndexBis

End Sub

Public Sub AddBottomBorder(CurrentCell As Range)

    Dim oSheet
    Dim oCellRange
	Dim oLine As New com.sun.star.table.BorderLine2

	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 26
		.OuterLineWidth = 26
	End With

	oSheet = ThisComponent.Sheets.getByName(CurrentCell.Worksheet.Name)
	oCellRange = oSheet.getCellByPosition(CurrentCell.Column-1,CurrentCell.Row-1)
	oCellRange.BottomBorder = oLine
End Sub

Public Sub FormatFinancementCells(BaseCell As Range)

    Dim Index As Integer
    Dim oSheet
    Dim oCellRange
	Dim oLine As New com.sun.star.table.BorderLine2

	With oLine
		.Color = RGB(0,0,0)
		.InnerLineWidth = 0
		.LineDistance = 0
		.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
		.LineWidth = 26
		.OuterLineWidth = 26
	End With

	oSheet = ThisComponent.Sheets.getByName(BaseCell.Worksheet.Name)
    For Index = 2 To 3
		oCellRange = oSheet.getCellByPosition(BaseCell.Column+Index-2,BaseCell.Row-1)
		oCellRange.CharWeight = com.sun.star.awt.FontWeight.BOLD ' bold ?
		oCellRange.CharColor = RGB(255,255,255) ' white ?
		oCellRange.CellBackColor = RGB(200,200,200) ' grey ?
		oCellRange.TopBorder = oLine
		oCellRange.BottomBorder = oLine
    Next Index
End Sub

Public Sub DefinirFormatConditionnelPourLesDossier( _
        SetOfRange As SetOfRange, _
        NBChantiers As Integer _
    )

    Dim CurrentCells As Range
    Dim FirstCellAddress As String
    Dim oCellAddress
	Dim oCondition(2) As New com.sun.star.beans.PropertyValue
	Dim oCondFormat
    Dim oRange
    Dim oSheet

    Set CurrentCells = Range( _
        SetOfRange.HeadCell.Cells(2, 1), _
        SetOfRange.EndCell.Cells(1, 3 + NBChantiers) _
    )
	
	oSheet = ThisComponent.Sheets.getByName(CurrentCells.Worksheet.Name)
	oRange = oSheet.getCellRangeByName(CurrentCells.Address(False,False,xlA1,False))
	oCondFormat = oRange.ConditionalFormat
	oCondFormat.clear()

	oCellAddress = oSheet.getCellByPosition( _
		CurrentCells.Cells(1,1).Column - 1, _
		CurrentCells.Cells(1,1).Row - 1_
	).getCellAddress()
	FirstCellAddress = CurrentCells.Cells(1,1).Cells(2,1).Address(False,False,xlA1,False)
	oCondFormat.addNew(CreerConditionProps( _
		"=SI(" & FirstCellAddress & "=DOSSIER_OK;VRAI();FAUX()", _
		oCellAddress, _
		"CondDossierOK", _
		0, 255, 0 _
	))
	oCondFormat.addNew(CreerConditionProps( _
		"=SI(" & FirstCellAddress & "=DOSSIER_FAVORABLE_ISSUE_INCERTAINE;VRAI();FAUX()", _
		oCellAddress, _
		"CondDossierFavorableIssueIncertaine", _
		0, 204, 255 _
	))
	oCondFormat.addNew(CreerConditionProps( _
		"=SI(" & FirstCellAddress & "=DOSSIER_INCERTAIN;VRAI();FAUX()", _
		oCellAddress, _
		"CondDossierIncertain", _
		255, 204, 0 _
	))
	oCondFormat.addNew(CreerConditionProps( _
		"=SI(" & FirstCellAddress & "=DOSSIER_NON_DEPOSE;VRAI();FAUX()", _
		oCellAddress, _
		"CondDossierNonDepose", _
		255, 255, 0 _
	))
	FirstCellAddress = CurrentCells.Cells(1,1).Address(False,False,xlA1,False)
	oCondFormat.addNew(CreerConditionProps( _
		"=MOD(LIGNE(" & FirstCellAddress & ");2)", _
		oCellAddress, _
		"CondAlternanceLigne", _
		204, 204, 204 _
	))

	oRange.ConditionalFormat = oCondFormat
End Sub

Public Function CreerConditionProps( _
		Formula As String, _
		oCellAddress, _
		StyleName As String , _
		Red As Integer, _
		Green As Integer, _
		Blue As Integer _
	)
	Dim oCondition(3) As New com.sun.star.beans.PropertyValue

	CreerStyle StyleName, Red, Green, Blue

	oCondition(0).Name = "Operator"
	oCondition(0).Value = com.sun.star.sheet.ConditionOperator.FORMULA
	oCondition(1).Name = "Formula1"
	oCondition(1).Value = Formula
	oCondition(2).Name = "StyleName"
	oCondition(2).Value = StyleName
	oCondition(3).Name = "SourcePosition"
	oCondition(3).Value = oCellAddress

	CreerConditionProps = oCondition
End Function

Public Function CreerStyle(Name As String, Red As Integer, Green As Integer, Blue As Integer)

	Dim oStyles
	Dim newStyle

	If Len(Name) = 0  Then
		CreerStyle = Null
		Exit Function
	End If

	oStyles = ThisComponent.StyleFamilies.getByName("CellStyles")
	If Not oStyles.HasByName(Name) Then
		newStyle = ThisComponent.createInstance("com.sun.star.style.CellStyle")
		If oStyles.HasByName("Par défaut") Then
			newStyle.ParentStyle = "Par défaut"
		Else
			If oStyles.HasByName("By default") Then
				newStyle.ParentStyle = "By default"
			Else
				If oStyles.HasByName("default") Then
					newStyle.ParentStyle = "default"
				End If
			End If
		End If
		oStyles.insertByName(Name,newStyle)
		newStyle = oStyles.getByName(Name)
		newStyle.setPropertyValue("CellBackColor", RGB(Red,Green,Blue))
	End If
	CreerStyle = oStyles.getByName(Name)
End Function

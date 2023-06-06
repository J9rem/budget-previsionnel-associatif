Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' constantes
Public Const Nom_de_Fichier_Par_Defaut As String = "InCitu_Budget_Previsionnel_Associatif_v1_11_LibreOffice"
Public Const Nom_de_Fichier_Par_Defaut_xls As String = "InCitu_Budget_Previsionnel_Associatif_v1_11_LibreOffice.xls"
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
    Dim Fichier_De_Sauvegarde As String
    
    ' Default
    FilePath = ""
    
    ' Changement de dossier
    Adresse_dossier_courant =ThisWorkbook.path
    ChDir (Adresse_dossier_courant)
    
    ' Fenêtre pour demander le nom du fichier de sauvegarde
    On Error Resume Next
    ' InitialFileName, FileFilter, FiltrerIndex, Title
    Fichier_De_Sauvegarde = GetSaveAsFilename( _
        Nom_de_Fichier_Par_Defaut, _
        Array(Array("Libre Office (*.ods)","*.ods")), _
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
    On Error GoTo 0
    returnToCurrentPath
    Kill FilePath
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

Public Sub SetFormatForBudget(BaseCell As Range, HeadCell As Range)

    Dim IndexBis As Integer
    Dim oSheet
    Dim oCellRange
    
    oSheet = ThisComponent.Sheets.getByName(BaseCell.Worksheet.Name)
    
    For IndexBis = 1 To 3
    	oCellRange = oSheet.getCellByPosition(BaseCell.Column+IndexBis-2,BaseCell.Row-1)
    	oCellRange.CharFontStyleName = ""
    	oCellRange.CharFontPitch = 2
    	oCellRange.CharFontCharSet = -1
    	oCellRange.CharFontFamily = 5
    	oCellRange.CharFontName = "Calibri"
    	oCellRange.CharColor = -1
    	oCellRange.CellBackColor = -1
    	oCellRange.CharHeight = 8
    	oCellRange.CharWeight = 100
    	oCellRange.LeftBorder.Color = 0
		oCellRange.LeftBorder.InnerLineWidth = 0
		oCellRange.LeftBorder.OuterLineWidth = 26
		oCellRange.LeftBorder.LineDistance = 0
		oCellRange.LeftBorder.LineStyle = 0
		oCellRange.LeftBorder.LineWidth = 26
		oCellRange.RightBorder.Color = 0
		oCellRange.RightBorder.InnerLineWidth = 0
		oCellRange.RightBorder.OuterLineWidth = 26
		oCellRange.RightBorder.LineDistance = 0
		oCellRange.RightBorder.LineStyle = 0
		oCellRange.RightBorder.LineWidth = 26
		oCellRange.TopBorder.Color = 0
		oCellRange.TopBorder.InnerLineWidth = 0
		oCellRange.TopBorder.OuterLineWidth = 26
		oCellRange.TopBorder.LineDistance = 0
		oCellRange.TopBorder.LineStyle = 0
		oCellRange.TopBorder.LineWidth = 26
		oCellRange.BottomBorder.Color = 0
		oCellRange.BottomBorder.InnerLineWidth = 0
		oCellRange.BottomBorder.OuterLineWidth = 26
		oCellRange.BottomBorder.LineDistance = 0
		oCellRange.BottomBorder.LineStyle = 0
		oCellRange.BottomBorder.LineWidth = 26
		
    	
    Next IndexBis
End Sub

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
        MsgBox "Le nom ne peut être vide !"
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
    AjoutFinancement wb, CurrentNBChantier, FinancementFantome, Nom, TypeFinancement
    
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

Public Sub DefinirBordures(CurrentCell As Range, AddTopBorder As Boolean)

    Dim oSheet
    Dim oCellRange
    
    oSheet = ThisComponent.Sheets.getByName(CurrentCell.Worksheet.Name)
    oCellRange = oSheet.getCellByPosition(CurrentCell.Column-1,CurrentCell.Row-1)
    
   	oCellRange.LeftBorder.Color = 0
	oCellRange.LeftBorder.InnerLineWidth = 0
	oCellRange.LeftBorder.OuterLineWidth = 26
	oCellRange.LeftBorder.LineDistance = 0
	oCellRange.LeftBorder.LineStyle = 0
	oCellRange.LeftBorder.LineWidth = 26
	oCellRange.RightBorder.Color = 0
	oCellRange.RightBorder.InnerLineWidth = 0
	oCellRange.RightBorder.OuterLineWidth = 26
	oCellRange.RightBorder.LineDistance = 0
	oCellRange.RightBorder.LineStyle = 0
	oCellRange.RightBorder.LineWidth = 26
	oCellRange.TopBorder.Color = 0
	oCellRange.TopBorder.InnerLineWidth = 0
    If AddTopBorder Then
		oCellRange.TopBorder.OuterLineWidth = 40
		oCellRange.TopBorder.LineWidth = 40
    Else
		oCellRange.TopBorder.OuterLineWidth = 26
		oCellRange.TopBorder.LineWidth = 26
    End If
	oCellRange.TopBorder.LineDistance = 0
	oCellRange.TopBorder.LineStyle = 0
	oCellRange.BottomBorder.Color = 0
	oCellRange.BottomBorder.InnerLineWidth = 0
	oCellRange.BottomBorder.OuterLineWidth = 26
	oCellRange.BottomBorder.LineDistance = 0
	oCellRange.BottomBorder.LineStyle = 0
	oCellRange.BottomBorder.LineWidth = 26
		
End Sub

Public Sub CopieLogo(oldWorkbook As Workbook, NewWorkbook As Workbook, Name As String)

   ' Do Nothing

End Sub

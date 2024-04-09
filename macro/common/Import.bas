Attribute VB_Name = "Import"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

Public Function choisirFichierAImporter(ByRef FilePath) As Boolean

    Dim Fichier_De_Sauvegarde
    ' FileFilter, FiltrerIndex, Title
    On Error Resume Next
    Fichier_De_Sauvegarde = Application.GetOpenFilename( _
        "Fichiers compatibles (*.xlsx;*.xls;*.xlsm;*.ods)," _
        & "*.xlsx;*.xls;*.xlsm;*.ods," _
        & "Excel 2003-2007 (*.xls),*.xls," _
        & "Excel (*.xlsx),*.xlsx," _
        & "Excel avec macro (*.xlsm),*.xlsm," _
        & "Libre Office (*.ods),*.ods", _
        0, _
        T_Choose_File_To_Import _
    )
    On Error GoTo 0
    If Fichier_De_Sauvegarde = "" _
        Or Fichier_De_Sauvegarde = Empty _
        Or Fichier_De_Sauvegarde = "Faux" _
        Or Fichier_De_Sauvegarde = "False" Then
        choisirFichierAImporter = False
    Else
        FilePath = Fichier_De_Sauvegarde
        choisirFichierAImporter = True
    End If

End Function

Public Function ImportSheets(oldWorkbook As Workbook, NewWorkbook As Workbook) As ListOfCptResult

    Dim DefSheetNames As Variant
    Dim Formula As String
    Dim FormulaCell As Range
    Dim Index As Integer
    Dim ListOfCptResult As ListOfCptResult
    Dim NewWs As Worksheet
    Dim PageName As String
    Dim Suffix As String
    Dim WithReal As Boolean
    Dim ws As Worksheet

    ListOfCptResult = ImportSheets_Init_ListOfCptResult()
    
    DefSheetNames = DefaultSheetsNames()
    
    For Each ws In oldWorkbook.Worksheets
        PageName = ws.Name
        If Not inArray(PageName, DefSheetNames) Then
            If Not (CptResult_IsValidatedPageName(PageName)) Then
                If Not FindWorkSheet(NewWorkbook, NewWs, PageName) Then
                    ' Create the new sheet
                    Set NewWs = AddWorksheetAtEnd(NewWorkbook, PageName)
                End If
                replaceContentFromWorksheet NewWs, ws
            Else
                ' import CptResultPartial
                If CptResult_IsReal(PageName) Then
                    Suffix = Mid(PageName, Len(Nom_Feuille_CptResult_Real_prefix) + 1)
                    WithReal = True
                Else
                    Suffix = Mid(PageName, Len(Nom_Feuille_CptResult_prefix) + 1)
                    WithReal = False
                End If
                If Suffix <> Nom_Feuille_CptResult_suffix Then
                    ' extract formula
                    Set FormulaCell = CptResult_GetFormulaCell(ws)
                    If Not (FormulaCell Is Nothing) Then
                        Formula = FormulaCell.Value
                        If Formula <> "" Then
                            ListOfCptResult = ImportSheets_Update_ListOfCptResult( _
                                ListOfCptResult, Suffix, Formula, WithReal _
                            )
                        End If
                    End If
                End If
            End If
        End If
    Next ws
    
    removeCrossRef ThisWorkbook, oldWorkbook
    
    NewWorkbook.Activate
    NewWorkbook.Worksheets(1).Activate

    ImportSheets = ListOfCptResult

End Function

Public Function ImportSheets_Init_ListOfCptResult() As ListOfCptResult

    Dim Formula() As String
    Dim ListOfCptResult As ListOfCptResult
    Dim Suffix() As String
    Dim WithReal() As Boolean

    ReDim Suffix(0 To 0)
    Suffix(0) = ""
    ReDim Formula(0 To 0)
    Formula(0) = ""
    ReDim WithReal(0 To 0)
    WithReal(0) = False

    ListOfCptResult.Formula = Formula
    ListOfCptResult.Suffix = Suffix
    ListOfCptResult.WithReal = WithReal

    ImportSheets_Init_ListOfCptResult = ListOfCptResult
End Function

Public Sub ImportSheets_Create_ListOfCptResult( _
    wb As Workbook, _
    ListOfCptResult As ListOfCptResult _
)
    
    Dim FormulaArr() As String
    Dim Index As Integer
    Dim SuffixArr() As String
    Dim WithRealArr() As Boolean
    
    SuffixArr = ListOfCptResult.Suffix
    WithRealArr = ListOfCptResult.WithReal
    FormulaArr = ListOfCptResult.Formula

    For Index = 1 To UBound(SuffixArr)
        CptResult_View_ForOneOrSeveralChantiers_Create_With_Name _
            wb, _
            FormulaArr(Index), _
            SuffixArr(Index), _
            WithRealArr(Index), _
            False, _
            True
    Next Index

End Sub

Public Function ImportSheets_Update_ListOfCptResult( _
    ListOfCptResult As ListOfCptResult, _
    Suffix As String, _
    Formula As String, _
    WithReal As Boolean _
) As ListOfCptResult

    Dim FormulaArr() As String
    Dim FormulaNewArr() As String
    Dim ListOfCptResultInternal As ListOfCptResult
    Dim Index As Integer
    Dim NBElem As Integer
    Dim ShouldAppend As Boolean
    Dim SuffixArr() As String
    Dim SuffixNewArr() As String
    Dim WithRealArr() As Boolean
    Dim WithRealNewArr() As Boolean

    ShouldAppend = True
    SuffixArr = ListOfCptResult.Suffix
    WithRealArr = ListOfCptResult.WithReal
    FormulaArr = ListOfCptResult.Formula
    ListOfCptResultInternal = ListOfCptResult
    NBElem = UBound(SuffixArr)
    If NBElem > 0 Then
        For Index = 1 To NBElem
            If SuffixArr(Index) = Suffix Then
                ShouldAppend = False
                If WithReal Then
                    WithRealArr(Index) = True
                    ListOfCptResultInternal.WithReal = WithRealArr
                End If
            End If
        Next Index
    End If

    If ShouldAppend Then
        ReDim FormulaNewArr(1 To (NBElem + 1))
        ReDim SuffixNewArr(1 To (NBElem + 1))
        ReDim WithRealNewArr(1 To (NBElem + 1))
        
        For Index = 1 To NBElem
            FormulaNewArr(Index) = FormulaArr(Index)
            SuffixNewArr(Index) = SuffixArr(Index)
            WithRealNewArr(Index) = WithRealArr(Index)
        Next Index
        FormulaNewArr(NBElem + 1) = Formula
        SuffixNewArr(NBElem + 1) = Suffix
        WithRealNewArr(NBElem + 1) = WithReal

        ListOfCptResultInternal.Formula = FormulaNewArr
        ListOfCptResultInternal.Suffix = SuffixNewArr
        ListOfCptResultInternal.WithReal = WithRealNewArr
    End If

    ImportSheets_Update_ListOfCptResult = ListOfCptResultInternal
End Function

Public Sub replaceContentFromWorksheet( _
        newWorksheet As Worksheet, _
        oldWorksheet As Worksheet, _
        Optional AsValue As Boolean = False _
    )

    ' clear previous content
    On Error Resume Next
    newWorksheet.Parent.Activate
    newWorksheet.Activate
    newWorksheet.Cells.Clear
    
    ' copy new content
    oldWorksheet.Parent.Activate
    oldWorksheet.Activate
    oldWorksheet.Cells.Select
    oldWorksheet.Cells.Copy
    
    ' paste content
    newWorksheet.Parent.Activate
    newWorksheet.Activate
    With newWorksheet.Cells
        .Select
        If AsValue Then
            .PasteSpecial (xlPasteValuesAndNumberFormats)
            .PasteSpecial (xlPasteFormats)
            .PasteSpecial (xlPasteColumnWidths)
            .PasteSpecial (xlPasteValues)
        Else
            .PasteSpecial (xlPasteAll)
        End If
    
    End With
    On Error GoTo 0
    
End Sub

Public Function importData(FileName As String) As Boolean

    Dim Informations As Informations
    Dim ListOfCptResult As ListOfCptResult
    Dim oldWorkbook As Workbook
    Dim PreviousNBSalarie As Integer
    Dim PreviousNBChantiers As Integer
    Dim PreviousRevision As WbRevision
    
    importData = False
    PreviousRevision = getDefaultWbRevision()
    
    SetSilent
    
    ' open woorkbook
    If Not openWbSafe(oldWorkbook, FileName) Then
        GoTo FinImportData
    End If
    
    ' Init default value before extraction from previous
    PreviousRevision = getPrevious(oldWorkbook, PreviousNBSalarie, PreviousNBChantiers, PreviousRevision)
    If PreviousRevision.Error Then
        GoTo FinImportData
    End If
    
    ' Copie du logo
    CopieLogo oldWorkbook, ThisWorkbook, Nom_Feuille_Cout_J_Salaire
    CopieLogo oldWorkbook, ThisWorkbook, Nom_Feuille_Personnel
    ' copie des onglets avant la copie des donnees pour eviter les erreurs
    ImportSheets oldWorkbook, ThisWorkbook
     
    ' copy data
    prepareFichier ThisWorkbook, PreviousNBSalarie, PreviousNBChantiers
    CopyPreviousValues oldWorkbook, ThisWorkbook, PreviousRevision
    On Error Resume Next
    CptResult_Update_ForASheet ThisWorkbook, Nom_Feuille_CptResult_Real_prefix & Nom_Feuille_CptResult_suffix
    On Error GoTo 0
    
    ' copie des onglets avant la copie des donnees pour eviter les autres erreurs
    ListOfCptResult = ImportSheets(oldWorkbook, ThisWorkbook)

    ' import et mise a jour des Comptes Resultats Partiels, s'il y en a
    ImportSheets_Create_ListOfCptResult ThisWorkbook, ListOfCptResult
    
    ' save file
    ThisWorkbook.Save
    importData = True
    
FinImportData:
    ' reset config
    Application.DisplayAlerts = True
    SetActive
End Function
Public Function extraireInfos(wb As Workbook) As Informations

    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim Result As Informations
    
    Result = getDefaultInformations()
    
    On Error Resume Next
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Informations)
    On Error GoTo 0
    If CurrentSheet Is Nothing Then
        GoTo FinFunction
    End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Annees)
        If Not BaseCell Is Nothing Then
            Result.Annee = BaseCell.Cells(1, 2).Value
            Result.AnneeFormula = Common_GetFormula(BaseCell.Cells(1, 2))
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Convention_Collective)
        If Not BaseCell Is Nothing Then
            Result.ConventionCollective = BaseCell.Cells(1, 2).Value
            Result.ConventionCollective = Common_GetFormula(BaseCell.Cells(1, 2))
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBConges)
        If Not BaseCell Is Nothing Then
            Result.NBConges = BaseCell.Cells(1, 2).Value
            Result.NBCongesFormula = Common_GetFormula(BaseCell.Cells(1, 2))
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBRTT)
        If Not BaseCell Is Nothing Then
            Result.NBRTT = BaseCell.Cells(1, 2).Value
            Result.NBRTTFormula = Common_GetFormula(BaseCell.Cells(1, 2))
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NB_Jours_speciaux)
        If Not BaseCell Is Nothing Then
            Result.NBJoursSpeciaux = BaseCell.Cells(1, 2).Value
            Result.NBJoursSpeciauxFormula = Common_GetFormula(BaseCell.Cells(1, 2))
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Pentecote)
        If Not BaseCell Is Nothing Then
            If BaseCell.Cells(1, 2).Value = "Oui" Then
                Result.Pentecote = True
            Else
                Result.Pentecote = False
            End If
        End If
FinFunction:
    extraireInfos = Result
End Function

Public Sub importerInfos(wb As Workbook, Informations As Informations)

    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    
    On Error Resume Next
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Informations)
    On Error GoTo 0
    If CurrentSheet Is Nothing Then
        GoTo FinSub
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Annees)
    If Not BaseCell Is Nothing Then
        Common_SetFormula BaseCell.Cells(1, 2), Informations.Annee, Informations.AnneeFormula
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Convention_Collective)
    If Not BaseCell Is Nothing Then
        If Left(Informations.ConventionCollective, 1) = "=" Then
            BaseCell.Cells(1, 2).Formula = Informations.ConventionCollective
        Else
            BaseCell.Cells(1, 2).Value = Informations.ConventionCollective
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBConges)
    If Not BaseCell Is Nothing Then
        Common_SetFormula BaseCell.Cells(1, 2), Informations.NBConges, Informations.NBCongesFormula, True
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBRTT)
    If Not BaseCell Is Nothing Then
        Common_SetFormula BaseCell.Cells(1, 2), Informations.NBRTT, Informations.NBRTTFormula, True
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NB_Jours_speciaux)
    If Not BaseCell Is Nothing Then
        Common_SetFormula BaseCell.Cells(1, 2), Informations.NBJoursSpeciaux, Informations.NBJoursSpeciauxFormula, True
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Pentecote)
    If Not BaseCell Is Nothing Then
        If Informations.Pentecote Then
            BaseCell.Cells(1, 2).Value = "Oui"
        Else
            BaseCell.Cells(1, 2).Value = "Non"
        End If
    End If
FinSub:
End Sub

Public Sub CopyPreviousValues(oldWorkbook As Workbook, NewWorkbook As Workbook, PreviousRevision As WbRevision)

    Dim Data As Data
    
    If PreviousRevision.Error Then
        Exit Sub
    End If
    
    If PreviousRevision.Majeure > 0 Then
        Data = Extract_Data_From_Table(oldWorkbook, PreviousRevision)
    Else
        Data = Extract_Data_From_Revision_0(oldWorkbook, PreviousRevision)
    End If
        
    Chantiers_Import NewWorkbook, Data

End Sub

' Extract data from a workbook
' @param Workbook wb
' @param WbRevision Revision
' @param Boolean OnlyChantiersFinancements
' @param Boolean WithProvisions
' @return Data Data
Public Function Extract_Data_From_Table( _
        wb As Workbook, _
        Revision As WbRevision, _
        Optional OnlyChantiersFinancements As Boolean = False, _
        Optional WithProvisions As Boolean = True _
    ) As Data

    Dim BaseCell As Range
    Dim BaseCellChantier As Range
    Dim ChantierSheet As Worksheet
    Dim CurrentSheet As Worksheet
    Dim Data As Data
    Dim DonneesSalarie As DonneesSalarie
    Dim DonneesSalaries() As DonneesSalarie
    Dim NBChantiers As Integer
    Dim NBSalaries As Integer
    
    Data = getDefaultData()
    DonneesSalarie = getDefaultDonneesSalarie()
    
    ReDim DonneesSalaries(0 To 0)
    DonneesSalaries(0) = DonneesSalarie
    Data.Salaries = DonneesSalaries
    
    Data.Informations = extraireInfos(wb)
    
    NBSalaries = GetNbSalaries(wb)
    
    If NBSalaries > 0 Then
        
        On Error Resume Next
        Set CurrentSheet = wb.Worksheets(Nom_Feuille_Personnel)
        On Error GoTo 0
        If CurrentSheet Is Nothing Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Personnel)
        Else
            Set BaseCell = CurrentSheet.Range("A:A").Find(T_FirstName)
            If BaseCell Is Nothing Then
                MsgBox Replace(T_NotFoundFirstName, "%PageName%", Nom_Feuille_Personnel)
            Else
                NBChantiers = 0
                On Error Resume Next
                Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
                On Error GoTo 0
                If ChantierSheet Is Nothing Then
                    MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers)
                Else
                    Set BaseCellChantier = Common_FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
                    If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.Value, Len("Chantier")) <> "Chantier" Then
                        Set BaseCellChantier = Nothing
                    Else
                        NBChantiers = GetNbChantiers(wb)
                    End If
                    
                    If Not OnlyChantiersFinancements Then
                        Data = Extract_Salaries(Data, BaseCell, BaseCellChantier, NBSalaries, NBChantiers)
                    End If
                    
                    If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                        Data.Chantiers = Chantiers_Depenses_Extract(BaseCellChantier, NBSalaries, NBChantiers).Chantiers
                        Data.Chantiers = Chantiers_Names_Extract(BaseCellChantier, NBChantiers, Data).Chantiers
                        Data.Chantiers = Chantiers_Financements_Extract(BaseCellChantier, NBChantiers, Data).Chantiers
                    End If
                    
                End If
            End If
        End If
    End If
    If Not OnlyChantiersFinancements Then
        Data = Charges_Extract(wb, Data, Revision)
    End If
    If WithProvisions Then
        Data = Provisions_Extract(wb, Data, Revision)
    End If
    
    Extract_Data_From_Table = Data

End Function
Public Function Extract_Data_From_Revision_0(oldWorkbook As Workbook, Revision As WbRevision) As Data

    Dim BaseCell As Range
    Dim BaseCellChantier As Range
    Dim ChantierSheet As Worksheet
    Dim Data As Data
    Dim DonneesSalarie As DonneesSalarie
    Dim DonneesSalaries() As DonneesSalarie
    Dim NBChantiers As Integer
    Dim NBSalariesAndRange As NBAndRange
    Dim NBSalaries As Integer
    Dim NBJoursTot As Double
    
    Data = getDefaultData()
    DonneesSalarie = getDefaultDonneesSalarie()
    
    ReDim DonneesSalaries(0 To 0)
    DonneesSalaries(0) = DonneesSalarie
    Data.Salaries = DonneesSalaries
    
    NBSalariesAndRange = GetNbSalariesV0(oldWorkbook)
    
    If NBSalariesAndRange.NB > 0 Then
        NBSalaries = NBSalariesAndRange.NB
        Set BaseCell = NBSalariesAndRange.Range
        NBChantiers = 0
        On Error Resume Next
        Set ChantierSheet = oldWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
        On Error GoTo 0
        If ChantierSheet Is Nothing Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers)
        Else
            Set BaseCellChantier = Common_FindNextNotEmpty(ChantierSheet.Cells(2, 1), False)
            If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.Value, Len("Chantier")) <> "Chantier" Then
                Set BaseCellChantier = Nothing
            Else
                NBChantiers = GetNbChantiers(oldWorkbook, 2)
            End If
            
            NBJoursTot = BaseCell.Worksheet.Cells(1, 7).EntireColumn.Find("Nb jours travaillables").Cells(1, 2).Value
            ' NBJoursTot = BaseCell.Cells(1 + NBSalaries + 1, 8).Value
            
            Data = Extract_Salaries(Data, BaseCell, BaseCellChantier, NBSalaries, NBChantiers, True, NBJoursTot)
            
            If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                Set BaseCell = BaseCellChantier.Cells(4 + NBSalaries, 0)
                Data.Chantiers = Chantiers_Depenses_Extract(BaseCellChantier, NBSalaries, NBChantiers, BaseCell).Chantiers
                Data.Chantiers = Chantiers_Names_Extract(BaseCellChantier, NBChantiers, Data).Chantiers
                Data.Chantiers = Chantiers_Financements_Extract(BaseCellChantier, NBChantiers, Data, ForV0:=True).Chantiers
            End If
            
        End If
    End If
    Data = Charges_Extract(oldWorkbook, Data, Revision)
    
    Extract_Data_From_Revision_0 = Data


End Function

Public Function Extract_Salaries( _
        Data As Data, _
        BaseCell As Range, _
        BaseCellChantier As Range, _
        NBSalaries As Integer, _
        NBChantiers As Integer, _
        Optional IsV0 As Boolean = False, _
        Optional NBJoursTot As Integer = 0 _
    ) As Data
    
    Dim BaseCellChantierReal As Range
    Dim CurrentRange As Range
    Dim DonneesSalarie As DonneesSalarie
    Dim DonneesSalaries() As DonneesSalarie
    Dim Index As Integer
    Dim IndexChantiers As Integer

    ReDim DonneesSalaries(1 To NBSalaries)
    For Index = 1 To NBSalaries
        DonneesSalarie = getDefaultDonneesSalarie()
        DonneesSalarie.Erreur = False
        DonneesSalarie.Prenom = BaseCell.Cells(1 + Index, 1).Value
        If IsV0 Then
            DonneesSalarie.Nom = ""
            DonneesSalarie.TauxDeTempsDeTravail = WorksheetFunction.Round(BaseCell.Cells(1 + Index, 2).Value / NBJoursTot, 2)
            DonneesSalarie.MasseSalarialeAnnuelle = BaseCell.Cells(1 + NBSalaries + 5 + Index, 3).Value
            DonneesSalarie.TauxOperateur = BaseCell.Cells(1 + Index, 3).Value
        Else
            DonneesSalarie.Nom = BaseCell.Cells(1 + Index, 2).Value
            ' ----
            Set CurrentRange = BaseCell.Cells(1 + Index, 3)
            DonneesSalarie.TauxDeTempsDeTravail = CurrentRange.Value
            DonneesSalarie.TauxDeTempsDeTravailFormula = Common_GetFormula(CurrentRange)
            ' -----
            Set CurrentRange = BaseCell.Cells(1 + Index, 4)
            DonneesSalarie.MasseSalarialeAnnuelle = CurrentRange.Value
            DonneesSalarie.MasseSalarialeAnnuelleFormula = Common_GetFormula(CurrentRange)
            ' -----
            Set CurrentRange = BaseCell.Cells(1 + Index, 5)
            DonneesSalarie.TauxOperateur = CurrentRange.Value
            DonneesSalarie.TauxOperateurFormula = Common_GetFormula(CurrentRange)
        End If
        If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
            DonneesSalarie.JoursChantiers = geDefaultJoursChantiers(NBChantiers)
            DonneesSalarie.JoursChantiersReal = geDefaultJoursChantiers(NBChantiers)
            DonneesSalarie.JoursChantiersFormula = geDefaultJoursChantiersStr(NBChantiers)
            DonneesSalarie.JoursChantiersFormulaReal = geDefaultJoursChantiersStr(NBChantiers)
            Set BaseCellChantierReal = Common_getBaseCellChantierRealFromBaseCellChantier(BaseCellChantier)
            For IndexChantiers = 1 To NBChantiers
                If IsV0 Then
                    DonneesSalarie.JoursChantiers(IndexChantiers) = BaseCellChantier.Cells(3 + Index, IndexChantiers).Value
                    DonneesSalarie.JoursChantiersFormula(IndexChantiers) = ""
                    DonneesSalarie.JoursChantiersReal(IndexChantiers) = 0
                    DonneesSalarie.JoursChantiersFormulaReal(IndexChantiers) = ""
                Else
                    Set CurrentRange = BaseCellChantier.Cells(4 + Index, IndexChantiers)
                    DonneesSalarie.JoursChantiers(IndexChantiers) = CurrentRange.Value
                    DonneesSalarie.JoursChantiersFormula(IndexChantiers) = Common_GetFormula(CurrentRange)
                    If Not (BaseCellChantierReal Is Nothing) Then
                        Set CurrentRange = BaseCellChantierReal.Cells(4 + Index, IndexChantiers)
                        DonneesSalarie.JoursChantiersReal(IndexChantiers) = CurrentRange.Value
                        DonneesSalarie.JoursChantiersFormulaReal(IndexChantiers) = Common_GetFormula(CurrentRange)
                    Else
                        DonneesSalarie.JoursChantiersReal(IndexChantiers) = 0
                        DonneesSalarie.JoursChantiersFormulaReal(IndexChantiers) = ""
                    End If
                End If
            Next IndexChantiers
        End If
        DonneesSalaries(Index) = DonneesSalarie
    Next Index
    Data.Salaries = DonneesSalaries
    Extract_Salaries = Data
End Function
Public Function getPrevious(wb As Workbook, ByRef PreviousNBSalarie As Integer, ByRef PreviousNBChantiers As Integer, PreviousRevision As WbRevision) As WbRevision

    ' Init default value before extraction from previous
    
    PreviousNBSalarie = 0
    PreviousNBChantiers = 0
    
    PreviousRevision = DetecteVersion(wb)
    If Not PreviousRevision.Error Then
        If PreviousRevision.Majeure = 0 Then
            PreviousNBSalarie = GetNbSalariesV0(wb).NB
            PreviousNBChantiers = GetNbChantiers(wb, 2)
        Else
            PreviousNBSalarie = GetNbSalaries(wb)
            PreviousNBChantiers = GetNbChantiers(wb)
        End If
    End If
    getPrevious = PreviousRevision
End Function

Public Function prepareFichier(wb As Workbook, PreviousNBSalarie As Integer, PreviousNBChantiers As Integer) As Boolean

    Dim NBSalariesInWorkingWk As Integer
    Dim NBChantiersInWorkingWk As Integer
    
    ' Change NB Salaries
    NBSalariesInWorkingWk = GetNbSalaries(wb)
    If PreviousNBSalarie > 0 And NBSalariesInWorkingWk > 0 Then
        ChangeSalaries wb, NBSalariesInWorkingWk, PreviousNBSalarie
    End If
    
    ' Add Chantiers
    NBChantiersInWorkingWk = GetNbChantiers(wb)
    If PreviousNBChantiers > 0 And NBChantiersInWorkingWk > 0 Then
        ChangeChantiers wb, NBChantiersInWorkingWk, PreviousNBChantiers
        ChangeChantiersReel wb, NBChantiersInWorkingWk, PreviousNBChantiers
    End If
    
    prepareFichier = True
End Function

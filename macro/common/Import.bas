Attribute VB_Name = "Import"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

Public Function choisirFichierAImporter(ByRef FilePath) As Boolean

    Dim Fichier_De_Sauvegarde
    ' FileFilter, FiltrerIndex, Title
    On Error Resume Next
    Fichier_De_Sauvegarde = Application.GetOpenFilename( _
        "Fichiers compatibles (*.xlsx;*.xls;*.ods;*.xlsm),*.xlsx;*.xls;*.ods;*.xlsm,Excel 2003-2007 (*.xls),*.xls,Excel (*.xlsx),*.xlsx, Libre Office (*.ods),*.ods, Excel avec macro (*.xlsm),*.xlsm", _
        0, _
        "Choisir le fichier à importer" _
    )
    On Error GoTo 0
    If Fichier_De_Sauvegarde = "" Or Fichier_De_Sauvegarde = Empty Or Fichier_De_Sauvegarde = "Faux" Or Fichier_De_Sauvegarde = "False" Then
        choisirFichierAImporter = False
    Else
        FilePath = Fichier_De_Sauvegarde
        choisirFichierAImporter = True
    End If

End Function

Public Sub ImportSheets(oldWorkbook As Workbook, NewWorkbook As Workbook)

    Dim Index As Integer
    Dim ws As Worksheet
    Dim NewWs As Worksheet
    Dim DefSheetNames As Variant
    
    DefSheetNames = DefaultSheetsNames()
    
    For Each ws In oldWorkbook.Worksheets
        If Not inArray(ws.Name, DefSheetNames) Then
            If Not FindWorkSheet(NewWorkbook, NewWs, ws.Name) Then
                ' Create the new sheet
                Set NewWs = AddWorksheetAtEnd(NewWorkbook, ws.Name)
            End If
            replaceContentFromWorksheet NewWs, ws
        End If
    Next ws
    
    removeCrossRef ThisWorkbook, oldWorkbook
    
    NewWorkbook.Activate
    NewWorkbook.Worksheets(1).Activate

End Sub

Public Sub replaceContentFromWorksheet(newWorksheet As Worksheet, oldWorksheet As Worksheet)

    ' clear previous content
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
        .PasteSpecial (xlPasteAll)
    
    End With
    
End Sub

Public Function importData(FileName As String) As Boolean

    Dim Informations As Informations
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
    ' copie des onglets avant la copie des données pour éviter les erreurs
    ImportSheets oldWorkbook, ThisWorkbook
     
    ' copy data
    prepareFichier ThisWorkbook, PreviousNBSalarie, PreviousNBChantiers
    CopyPreviousValues oldWorkbook, ThisWorkbook, PreviousRevision
    MettreAJourBudgetGlobal ThisWorkbook
    
    ' copie des onglets avant la copie des données pour éviter les autres erreurs
    ImportSheets oldWorkbook, ThisWorkbook
    
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
            Result.Annee = BaseCell.Cells(1, 2).value
            If BaseCell.Cells(1, 2).HasFormula = True Then
                Result.AnneeFormula = BaseCell.Cells(1, 2).Formula
            End If
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Convention_Collective)
        If Not BaseCell Is Nothing Then
            Result.ConventionCollective = BaseCell.Cells(1, 2).value
            If BaseCell.Cells(1, 2).HasFormula = True Then
                Result.ConventionCollective = BaseCell.Cells(1, 2).Formula
            End If
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBConges)
        If Not BaseCell Is Nothing Then
            Result.NBConges = BaseCell.Cells(1, 2).value
            If BaseCell.Cells(1, 2).HasFormula = True Then
                Result.NBCongesFormula = BaseCell.Cells(1, 2).Formula
            End If
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBRTT)
        If Not BaseCell Is Nothing Then
            Result.NBRTT = BaseCell.Cells(1, 2).value
            If BaseCell.Cells(1, 2).HasFormula = True Then
                Result.NBRTTFormula = BaseCell.Cells(1, 2).Formula
            End If
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NB_Jours_speciaux)
        If Not BaseCell Is Nothing Then
            Result.NBJoursSpeciaux = BaseCell.Cells(1, 2).value
            If BaseCell.Cells(1, 2).HasFormula = True Then
                Result.NBJoursSpeciauxFormula = BaseCell.Cells(1, 2).Formula
            End If
        End If
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Pentecote)
        If Not BaseCell Is Nothing Then
            If BaseCell.Cells(1, 2).value = "Oui" Then
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
        If Informations.AnneeFormula = "" Then
            BaseCell.Cells(1, 2).value = Informations.Annee
        Else
            BaseCell.Cells(1, 2).Formula = Informations.AnneeFormula
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Convention_Collective)
    If Not BaseCell Is Nothing Then
        If Left(Informations.ConventionCollective, 1) = "=" Then
            BaseCell.Cells(1, 2).Formula = Informations.ConventionCollective
        Else
            BaseCell.Cells(1, 2).value = Informations.ConventionCollective
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBConges)
    If Not BaseCell Is Nothing Then
        If Informations.NBCongesFormula = "" Then
            If Informations.NBConges = 0 Then
                BaseCell.Cells(1, 2).value = ""
            Else
                BaseCell.Cells(1, 2).value = Informations.NBConges
            End If
        Else
            BaseCell.Cells(1, 2).Formula = Informations.NBCongesFormula
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NBRTT)
    If Not BaseCell Is Nothing Then
        If Informations.NBRTTFormula = "" Then
            If Informations.NBRTT = 0 Then
                BaseCell.Cells(1, 2).value = ""
            Else
                BaseCell.Cells(1, 2).value = Informations.NBRTT
            End If
        Else
            BaseCell.Cells(1, 2).Formula = Informations.NBRTTFormula
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_NB_Jours_speciaux)
    If Not BaseCell Is Nothing Then
        If Informations.NBJoursSpeciauxFormula = "" Then
            If Informations.NBJoursSpeciaux = 0 Then
                BaseCell.Cells(1, 2).value = ""
            Else
                BaseCell.Cells(1, 2).value = Informations.NBJoursSpeciaux
            End If
        Else
            BaseCell.Cells(1, 2).Formula = Informations.NBJoursSpeciauxFormula
        End If
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Pentecote)
    If Not BaseCell Is Nothing Then
        If Informations.Pentecote Then
            BaseCell.Cells(1, 2).value = "Oui"
        Else
            BaseCell.Cells(1, 2).value = "Non"
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
        Data = extraireDonneesVersion1(oldWorkbook, PreviousRevision)
    Else
        Data = extraireDonneesVersion0(oldWorkbook, PreviousRevision)
    End If
        
    insererDonnees NewWorkbook, Data

End Sub

Public Function extraireDonneesVersion1(oldWorkbook As Workbook, Revision As WbRevision) As Data
    Dim DonneesSalarie As DonneesSalarie
    Dim Data As Data
    Dim NBSalaries As Integer
    Dim NBChantiers As Integer
    Dim Index As Integer
    Dim IndexChantiers As Integer
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim BaseCellChantier As Range
    Dim DonneesSalaries() As DonneesSalarie
    Dim JoursChantiers() As Double
    
    Data = getDefaultData(Data)
    DonneesSalarie = getDefaultDonneesSalarie(DonneesSalarie)
    
    ReDim DonneesSalaries(0 To 0)
    DonneesSalaries(0) = DonneesSalarie
    Data.Salaries = DonneesSalaries
    
    Data.Informations = extraireInfos(oldWorkbook)
    
    NBSalaries = GetNbSalaries(oldWorkbook)
    
    If NBSalaries > 0 Then
        
        On Error Resume Next
        Set CurrentSheet = oldWorkbook.Worksheets(Nom_Feuille_Personnel)
        On Error GoTo 0
        If CurrentSheet Is Nothing Then
            MsgBox "'" & Nom_Feuille_Personnel & "' n'a pas été trouvée"
        Else
            Set BaseCell = CurrentSheet.Range("A:A").Find("Prénom")
            If BaseCell Is Nothing Then
                MsgBox "'Prénom' non trouvé dans '" & Nom_Feuille_Personnel & "' !"
            Else
                NBChantiers = 0
                On Error Resume Next
                Set ChantierSheet = oldWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
                On Error GoTo 0
                If ChantierSheet Is Nothing Then
                    MsgBox "'" & Nom_Feuille_Budget_chantiers & "' n'a pas été trouvée"
                Else
                    Set BaseCellChantier = ChantierSheet.Cells(3, 1).End(xlToRight)
                    If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.value, Len("Chantier")) <> "Chantier" Then
                        Set BaseCellChantier = Nothing
                    Else
                        NBChantiers = GetNbChantiers(oldWorkbook)
                    End If
                    
                    ReDim DonneesSalaries(1 To NBSalaries)
                    For Index = 1 To NBSalaries
                        DonneesSalarie = getDefaultDonneesSalarie(DonneesSalarie)
                        DonneesSalarie.Erreur = False
                        DonneesSalarie.Prenom = BaseCell.Cells(1 + Index, 1).value
                        DonneesSalarie.Nom = BaseCell.Cells(1 + Index, 2).value
                        DonneesSalarie.TauxDeTempsDeTravail = BaseCell.Cells(1 + Index, 3).value
                        If BaseCell.Cells(1 + Index, 3).HasFormula = True Then
                            DonneesSalarie.TauxDeTempsDeTravailFormula = BaseCell.Cells(1 + Index, 3).Formula
                        End If
                        DonneesSalarie.MasseSalarialeAnnuelle = BaseCell.Cells(1 + Index, 4).value
                        If BaseCell.Cells(1 + Index, 4).HasFormula = True Then
                            DonneesSalarie.MasseSalarialeAnnuelleFormula = BaseCell.Cells(1 + Index, 4).Formula
                        End If
                        DonneesSalarie.TauxOperateur = BaseCell.Cells(1 + Index, 5).value
                        If BaseCell.Cells(1 + Index, 5).HasFormula = True Then
                            DonneesSalarie.TauxOperateurFormula = BaseCell.Cells(1 + Index, 5).Formula
                        End If
                        If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                            ReDim JoursChantiers(1 To NBChantiers)
                            For IndexChantiers = 1 To NBChantiers
                                JoursChantiers(IndexChantiers) = BaseCellChantier.Cells(4 + Index, IndexChantiers).value
                            Next IndexChantiers
                            DonneesSalarie.JoursChantiers = JoursChantiers
                        End If
                        DonneesSalaries(Index) = DonneesSalarie
                    Next Index
                    Data.Salaries = DonneesSalaries
                    
                    If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                        Data.Chantiers = extraireDepensesChantier(BaseCellChantier, NBSalaries, NBChantiers).Chantiers
                        Data.Chantiers = extraireFinancementChantier(BaseCellChantier, NBChantiers, Data).Chantiers
                    End If
                    
                End If
            End If
        End If
    End If
    Data = extraireCharges(oldWorkbook, Data, Revision)
    
    extraireDonneesVersion1 = Data

End Function
Public Function extraireDonneesVersion0(oldWorkbook As Workbook, Revision As WbRevision) As Data

    Dim DonneesSalarie As DonneesSalarie
    Dim Data As Data
    Dim NBSalariesAndRange As NBAndRange
    Dim NBSalaries As Integer
    Dim NBChantiers As Integer
    Dim Index As Integer
    Dim IndexChantiers As Integer
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim BaseCellChantier As Range
    Dim NBJoursTot As Double
    Dim DonneesSalaries() As DonneesSalarie
    Dim JoursChantiers() As Double
    
    Data = getDefaultData(Data)
    DonneesSalarie = getDefaultDonneesSalarie(DonneesSalarie)
    
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
            MsgBox "'" & Nom_Feuille_Budget_chantiers & "' n'a pas été trouvée"
        Else
            Set BaseCellChantier = ChantierSheet.Cells(2, 1).End(xlToRight)
            If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.value, Len("Chantier")) <> "Chantier" Then
                Set BaseCellChantier = Nothing
            Else
                NBChantiers = GetNbChantiers(oldWorkbook, 2)
            End If
            
            ReDim DonneesSalaries(1 To NBSalaries)
            NBJoursTot = BaseCell.Worksheet.Cells(1, 7).EntireColumn.Find("Nb jours travaillables").Cells(1, 2).value
            ' NBJoursTot = BaseCell.Cells(1 + NBSalaries + 1, 8).Value
            For Index = 1 To NBSalaries
                DonneesSalarie = getDefaultDonneesSalarie(DonneesSalarie)
                DonneesSalarie.Erreur = False
                DonneesSalarie.Prenom = BaseCell.Cells(1 + Index, 1).value
                DonneesSalarie.Nom = ""
                DonneesSalarie.TauxDeTempsDeTravail = WorksheetFunction.Round(BaseCell.Cells(1 + Index, 2).value / NBJoursTot, 2)
                DonneesSalarie.MasseSalarialeAnnuelle = BaseCell.Cells(1 + NBSalaries + 5 + Index, 3).value
                DonneesSalarie.TauxOperateur = BaseCell.Cells(1 + Index, 3).value
                If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                    ReDim JoursChantiers(1 To NBChantiers)
                    For IndexChantiers = 1 To NBChantiers
                        JoursChantiers(IndexChantiers) = BaseCellChantier.Cells(3 + Index, IndexChantiers).value
                    Next IndexChantiers
                    DonneesSalarie.JoursChantiers = JoursChantiers
                End If
                DonneesSalaries(Index) = DonneesSalarie
            Next Index
            Data.Salaries = DonneesSalaries
            
            If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                Set BaseCell = BaseCellChantier.Cells(4 + NBSalaries, 0)
                Data.Chantiers = extraireDepensesChantier(BaseCellChantier, NBSalaries, NBChantiers, BaseCell).Chantiers
                Data.Chantiers = extraireFinancementChantier(BaseCellChantier, NBChantiers, Data, ForV0:=True).Chantiers
            End If
            
        End If
    End If
    Data = extraireCharges(oldWorkbook, Data, Revision)
    
    extraireDonneesVersion0 = Data


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
    End If
    
    prepareFichier = True
End Function

Attribute VB_Name = "Process"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

Public Sub ChangeSalaries(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)

    If FinalNB < 1 Then
        Exit Sub
    End If
    
    ChangeNBSalarieDansPersonnel wb, PreviousNB, FinalNB
    CoutJSalaires_Salaries_ChangeNB wb, PreviousNB, FinalNB
    Chantiers_Salaries_ChangeNB wb, PreviousNB, FinalNB, False
    Chantiers_Salaries_ChangeNB wb, PreviousNB, FinalNB, True

End Sub

Public Sub ChangeChantiers(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)

    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim EndRange As Range
    Dim Index As Integer
    Dim NBSalaries As Integer
    Dim SetOfRange As SetOfRange
    Dim StartRange As Range
    
    If FinalNB < 1 Then
        Exit Sub
    End If
    
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If
    Set BaseCell = Common_FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
    If BaseCell.Column > 1000 Then
        Exit Sub
    End If
    If Left(BaseCell.Value, Len("Chantier")) <> "Chantier" Then
        Exit Sub
    End If
    
    If FinalNB > PreviousNB Then
        BaseCell.Cells(1, 1).Worksheet.Activate
        BaseCell.Cells(1, PreviousNB).EntireColumn.Select
        BaseCell.Cells(1, PreviousNB).EntireColumn.Copy
        Range(BaseCell.Cells(1, PreviousNB + 1).EntireColumn, BaseCell.Cells(1, FinalNB).EntireColumn).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        BaseCell.Cells(1, PreviousNB).EntireColumn.Copy
        Range(BaseCell.Cells(1, PreviousNB + 1).EntireColumn, BaseCell.Cells(1, FinalNB).EntireColumn).PasteSpecial _
            Paste:=xlAll
        ' Clear contents
        For Index = PreviousNB + 1 To FinalNB
            BaseCell.Cells(2, Index).Value = "xx"
        Next Index
        NBSalaries = GetNbSalaries(wb)
        If NBSalaries > 0 Then
            ' empty first part for time for salarie
            Set StartRange = BaseCell.Cells(5, PreviousNB + 1)
            Set EndRange = BaseCell.Cells(5 + NBSalaries - 1, 1)
            Range(StartRange, EndRange.Cells(1, FinalNB)).ClearContents

            ' empty depenses for salarie
            SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheet, Nothing)
            If SetOfRange.Status Then
                Range( _
                SetOfRange.HeadCell.Cells(2, PreviousNB + 3), _
                SetOfRange.ResultCell.Cells(0, FinalNB + 1) _
                ).ClearContents
            Else
                ' Backup
                Set StartRange = EndRange.Cells(3 + NBSalaries, PreviousNB + 1)
                Set EndRange = Common_FindNextNotEmpty(StartRange.EntireRow.Cells(1, 2), True).EntireRow.Cells(0, BaseCell.Cells(1, FinalNB).Column)
                Range(StartRange, EndRange).ClearContents
            End If
        End If
        Chantiers_UpdateSums wb, BaseCell
    Else
        If FinalNB < PreviousNB Then
            Range(BaseCell.Cells(1, FinalNB + 1).EntireColumn, BaseCell.Cells(1, PreviousNB).EntireColumn).Delete Shift:=xlToLeft
        End If
    End If
    

End Sub

Public Sub ChangeChantiersReel(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)

    Dim BaseCell As Range
    Dim BaseCellReal As Range
    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim EndRange As Range
    Dim Index As Integer
    Dim IndexLevel2 As Integer
    Dim NBSalaries As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfRangeF As SetOfRange
    Dim StartRange As Range
    
    If FinalNB < 1 Then
        Exit Sub
    End If
    
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If
    Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    If ChantierSheetReal Is Nothing Then
        Exit Sub
    End If
    Set BaseCell = Common_FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
    If BaseCell.Column > 1000 Then
        Exit Sub
    End If
    If Left(BaseCell.Value, Len("Chantier")) <> "Chantier" Then
        Exit Sub
    End If
    Set BaseCellReal = Common_FindNextNotEmpty(ChantierSheetReal.Cells(3, 1), False)
    If BaseCellReal.Column > 1000 Then
        Exit Sub
    End If
    If Left(BaseCellReal.Value, Len("Chantier")) <> "Chantier" Then
        Exit Sub
    End If
    
    If FinalNB > PreviousNB Then
        BaseCellReal.Cells(1, 1).Worksheet.Activate
        BaseCellReal.Cells(1, 1 + 3 * (PreviousNB - 1)).EntireColumn.Select
        Range( _
            BaseCellReal.Cells(1, 1 + 3 * (PreviousNB - 1)).EntireColumn, _
            BaseCellReal.Cells(1, 3 + 3 * (PreviousNB - 1)).EntireColumn _
            ).Copy
        Range( _
            BaseCellReal.Cells(1, 1 + 3 * PreviousNB).EntireColumn, _
            BaseCellReal.Cells(1, 3 + 3 * (FinalNB - 1)).EntireColumn _
            ).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range( _
            BaseCellReal.Cells(1, 1 + 3 * (PreviousNB - 1)).EntireColumn, _
            BaseCellReal.Cells(1, 3 + 3 * (PreviousNB - 1)).EntireColumn _
            ).Copy
        Range( _
            BaseCellReal.Cells(1, 1 + 3 * PreviousNB).EntireColumn, _
            BaseCellReal.Cells(1, 3 + 3 * (FinalNB - 1)).EntireColumn _
        ).PasteSpecial _
            Paste:=xlAll

        ' update contents
        NBSalaries = GetNbSalaries(wb)
        SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheetReal, Nothing)
        SetOfRangeF = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
        For Index = (PreviousNB + 1) To FinalNB
            ' title
            BaseCellReal.Cells(1, 1 + 3 * (Index - 1)).Formula = "=" & _
                CleanAddress(BaseCell.Cells(1, Index).address(False, False, xlA1, True))
            ' name
            BaseCellReal.Cells(2, 1 + 3 * (Index - 1)).Formula = "=" & _
                CleanAddress(BaseCell.Cells(2, Index).address(False, False, xlA1, True))
            If NBSalaries > 0 Then
                ' empty first part for time for salarie
                Set StartRange = BaseCellReal.Cells(5, 2 + 3 * (Index - 1))
                Set EndRange = BaseCellReal.Cells(5 + NBSalaries - 1, 2 + 3 * (Index - 1))
                Range(StartRange, EndRange).ClearContents

                ' update formula
                For IndexLevel2 = 1 To NBSalaries
                    BaseCellReal.Cells(4 + IndexLevel2, 1 + 3 * (Index - 1)).Formula = "=" & _
                        CleanAddress(BaseCell.Cells(4 + IndexLevel2, Index).address(False, False, xlA1, True))
                Next IndexLevel2
                
                ' charges indirectes
                For IndexLevel2 = 1 To 4
                    BaseCellReal.Cells(6 + 2 * NBSalaries + IndexLevel2, 1 + 3 * (Index - 1)).Formula = "=" & _
                        CleanAddress(BaseCell.Cells(6 + 2 * NBSalaries + IndexLevel2, Index).address(False, False, xlA1, True))
                Next IndexLevel2

                ' empty depenses for salarie
                If SetOfRange.Status Then
                    Range( _
                        SetOfRange.HeadCell.Cells(2, 3 + 3 * PreviousNB), _
                        SetOfRange.ResultCell.Cells(0, 2 + 3 * PreviousNB) _
                    ).ClearContents
                End If
                
                ' depenses
                For IndexLevel2 = (SetOfRange.HeadCell.Row - BaseCellReal.Row + 1) To (SetOfRange.ResultCell.Row - BaseCellReal.Row)
                    BaseCellReal.Cells(IndexLevel2, 1 + 3 * (Index - 1)).Formula = "=" & _
                        CleanAddress(BaseCell.Cells(IndexLevel2, Index).address(False, False, xlA1, True))
                Next IndexLevel2
            End If

            ' financements
            If SetOfRangeF.Status And SetOfRangeF.StatusReal Then
                For IndexLevel2 = 2 To (SetOfRangeF.ResultCellReal.Row - SetOfRangeF.HeadCellReal.Row)
                    If SetOfRangeF.HeadCellReal.Cells(IndexLevel2, 2).Value = "Statut" Then
                        SetOfRangeF.HeadCellReal.Cells(IndexLevel2, 3 + 3 * (Index - 1)).Formula = "=" _
                            & "IF(" _
                            & CleanAddress(SetOfRangeF.HeadCell.Cells(IndexLevel2, 2 + Index).address(False, False, xlA1, True)) _
                            & "="""",""""," _
                            & CleanAddress(SetOfRangeF.HeadCell.Cells(IndexLevel2, 2 + Index).address(False, False, xlA1, True)) _
                            & ")"
                    Else
                        SetOfRangeF.HeadCellReal.Cells(IndexLevel2, 3 + 3 * (Index - 1)).Formula = "=" _
                            & CleanAddress(SetOfRangeF.HeadCell.Cells(IndexLevel2, 2 + Index).address(False, False, xlA1, True))
                        SetOfRangeF.HeadCellReal.Cells(IndexLevel2, 4 + 3 * (Index - 1)).Value = ""
                    End If
                Next IndexLevel2
                SetOfRangeF.ResultCellReal.Cells(2, 2 + 3 * (Index - 1)).Formula = "=" _
                    & CleanAddress(SetOfRangeF.ResultCell.Cells(2, 1 + Index).address(False, False, xlA1, True))
                SetOfRangeF.ResultCellReal.Cells(2, 3 + 3 * (Index - 1)).Value = ""
                For IndexLevel2 = 9 To 11
                    SetOfRangeF.ResultCellReal.Cells(IndexLevel2, 2 + 3 * (Index - 1)).Formula = "=" _
                        & CleanAddress(SetOfRangeF.ResultCell.Cells(IndexLevel2, 1 + Index).address(False, False, xlA1, True))
                    SetOfRangeF.ResultCellReal.Cells(IndexLevel2, 3 + 3 * (Index - 1)).Value = ""
                Next IndexLevel2
            End If
        Next Index
    Else
        If FinalNB < PreviousNB Then
            Range( _
                BaseCellReal.Cells(1, 1 + 3 * FinalNB).EntireColumn, _
                BaseCellReal.Cells(1, 3 + 3 * (PreviousNB - 1)).EntireColumn _
            ).Delete Shift:=xlToLeft
        End If
    End If

    Chantiers_UpdateSumsReal wb, BaseCellReal

End Sub

Public Sub ChangeUnChantier(Delta As Integer)

    Dim CurrentNBChantier As Integer
    Dim CurrentWs As Worksheet
    Dim NBToRemove As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set CurrentWs = wb.ActiveSheet
    
    SetSilent
    
    ' Current NB
    CurrentNBChantier = GetNbChantiers(wb)
    
    If Delta < 0 And (CurrentNBChantier + Delta) < 1 Then
        GoTo FinSub
    End If
    
    ChangeChantiers wb, CurrentNBChantier, CurrentNBChantier + Delta
    ChangeChantiersReel wb, CurrentNBChantier, CurrentNBChantier + Delta
    
FinSub:
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Cells(1, 1).Select
    Next 'Ws
    CurrentWs.Activate
    
    SetActive

End Sub

Public Sub ChangeUnSalarie(Delta As Integer)

    Dim CurrentNBSalaries As Integer
    Dim CurrentWs As Worksheet
    Dim NBToRemove As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set CurrentWs = wb.ActiveSheet
    
    SetSilent
    
    ' Current NB
    CurrentNBSalaries = GetNbSalaries(wb)
    
    If Delta < 0 And (CurrentNBSalaries + Delta) < 1 Then
        GoTo FinSub
    End If
    
    ChangeSalaries wb, CurrentNBSalaries, CurrentNBSalaries + Delta
    
FinSub:
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Cells(1, 1).Select
    Next 'Ws
    
    SetActive
    CurrentWs.Activate

End Sub

Public Sub ChangeNBSalarieDansPersonnel(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)

    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim RealFinalNB As Integer
    Dim endR As Range
    
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Personnel)
    If CurrentSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Personnel)
        Exit Sub
    End If
    
    Set BaseCell = CurrentSheet.Range("A:A").Find(T_FirstName)
    If BaseCell Is Nothing Then
        MsgBox Replace(T_NotFoundFirstName, "%PageName%", Nom_Feuille_Personnel)
        Exit Sub
    End If
    
    If FinalNB <= 1 Then
        RealFinalNB = 2
    Else
        RealFinalNB = FinalNB
    End If
    
    Set endR = Common_FindNextNotEmpty(BaseCell, False)
    
    If PreviousNB > RealFinalNB Then
        Common_RemoveRows BaseCell, PreviousNB, RealFinalNB, 1
    Else
        If PreviousNB < FinalNB Then
            Common_InsertRows BaseCell, PreviousNB, FinalNB, True, 1
        End If
    End If
    
    If FinalNB <= 1 And PreviousNB > 1 Then
        Range(BaseCell.Cells(3, 1), endR.Cells(3, 1)).ClearContents
    End If

End Sub

Public Sub Chantiers_And_Personal_Import_Salaries( _
        Data As Data, _
        BaseCell As Range, _
        BaseCellChantier As Range, _
        BaseCellChantierReal As Range, _
        NBSalaries As Integer, _
        NBChantiers As Integer _
    )
    
    Dim DonneesSalarie As DonneesSalarie
    Dim Index As Integer
    Dim IndexChantier As Integer
    Dim IndexTab As Integer

    Index = 1
    For IndexTab = LBound(Data.Salaries) To UBound(Data.Salaries)
        DonneesSalarie = Data.Salaries(IndexTab)
        
        If Not DonneesSalarie.Erreur And Index <= NBSalaries Then
            BaseCell.Cells(1 + Index, 1).Value = DonneesSalarie.Prenom
            BaseCell.Cells(1 + Index, 2).Value = DonneesSalarie.Nom
            Common_SetFormula _
                BaseCell.Cells(1 + Index, 3), _
                DonneesSalarie.TauxDeTempsDeTravail, _
                DonneesSalarie.TauxDeTempsDeTravailFormula
            
            Common_SetFormula _
                BaseCell.Cells(1 + Index, 4), _
                DonneesSalarie.MasseSalarialeAnnuelle, _
                DonneesSalarie.MasseSalarialeAnnuelleFormula
            
            Common_SetFormula _
                BaseCell.Cells(1 + Index, 5), _
                DonneesSalarie.TauxOperateur, _
                DonneesSalarie.TauxOperateurFormula

            If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                For IndexChantier = 1 To WorksheetFunction.Min(NBChantiers, UBound(DonneesSalarie.JoursChantiers))
                    Common_SetFormula _
                        BaseCellChantier.Cells(4 + Index, IndexChantier), _
                        DonneesSalarie.JoursChantiers(IndexChantier), _
                        DonneesSalarie.JoursChantiersFormula(IndexChantier), _
                        True
                    If Not (BaseCellChantierReal Is Nothing) Then
                        Common_SetFormula _
                            BaseCellChantierReal.Cells(4 + Index, 2 + 3 * (IndexChantier - 1)), _
                            DonneesSalarie.JoursChantiersReal(IndexChantier), _
                            DonneesSalarie.JoursChantiersFormulaReal(IndexChantier), _
                            True
                    End If
                Next IndexChantier
            End If
            Index = Index + 1
        End If
    Next IndexTab
End Sub

Public Function Chantiers_BaseCell_Get( _
    ChantierSheet As Worksheet _
    ) As Range
    
    Dim BaseCellChantier As Range
    
    Set BaseCellChantier = Nothing

    If Not (ChantierSheet Is Nothing) Then
        Set BaseCellChantier = Common_FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
        If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.Value, Len("Chantier")) <> "Chantier" Then
            Set BaseCellChantier = Nothing
        End If
    End If

    Set Chantiers_BaseCell_Get = BaseCellChantier
End Function

Public Function Chantiers_Depenses_Extract( _
        BaseCellChantier As Range, _
        NBSalaries As Integer, _
        NBChantiers As Integer, _
        Optional BaseCell As Range _
    ) As SetOfChantiers
        
    Dim BaseCellChantierReal As Range
    Dim BaseCellLocal As Range
    Dim Chantiers() As Chantier
    Dim ChantierSheetReal As Worksheet
    Dim ChantierTmp As Chantier
    Dim ChantierTmp1 As Chantier
    Dim CurrentRange As Range
    Dim DepensesTmp1() As DepenseChantier
    Dim DepenseTmp As DepenseChantier
    Dim IndexChantiers As Integer
    Dim IndexDepense As Integer
    Dim NBDepenses As Integer
    Dim NewFormatForAutofinancement As Integer
    Dim SetOfChantiers As SetOfChantiers
    Dim SetOfRange As SetOfRange
    Dim TestedValue As String
    
    Set BaseCellChantierReal = Common_getBaseCellChantierRealFromBaseCellChantier(BaseCellChantier)
    
    ' Depenses
    If BaseCellChantierReal Is Nothing Then
        Set ChantierSheetReal = Nothing
    Else
        Set ChantierSheetReal = BaseCellChantierReal.Worksheet
    End If
    SetOfRange = Chantiers_Depenses_SetOfRange_Get(BaseCellChantier.Worksheet, ChantierSheetReal)
    If SetOfRange.Status Then
        Set BaseCell = SetOfRange.HeadCell.Cells(2, 2)
    Else
        ' Backup
        Set BaseCell = BaseCellChantier.Cells(7 + 2 * NBSalaries, 1).EntireRow.Cells(1, 2)
    End If
    NBDepenses = Range(BaseCell, Common_FindNextNotEmpty(BaseCell, True).Cells(0, 1)).Rows.Count
    
    SetOfChantiers = Common_getDefaultSetOfChantiers(NBChantiers, NBDepenses)

    For IndexDepense = 1 To NBDepenses
        Chantiers_Depenses_Extract_Name SetOfChantiers, 1, IndexDepense, BaseCell.Cells(IndexDepense, 1).Value
    Next IndexDepense
    
    For IndexChantiers = 1 To NBChantiers
        Chantiers = SetOfChantiers.Chantiers
        ChantierTmp = Chantiers(IndexChantiers)
        ChantierTmp1 = Chantiers(1)
        DepensesTmp1 = ChantierTmp1.Depenses
        ChantierTmp.Nom = BaseCellChantier.Cells(2, IndexChantiers).Value
        For IndexDepense = 1 To NBDepenses
            If IndexChantiers > 1 Then
                DepenseTmp = DepensesTmp1(IndexDepense)
                Chantiers_Depenses_Extract_Name SetOfChantiers, IndexChantiers, IndexDepense, DepenseTmp.Nom
            End If
            Set CurrentRange = BaseCell.Cells(IndexDepense, IndexChantiers + 1)
            Chantiers_Depenses_Extract_Value SetOfChantiers, IndexChantiers, IndexDepense, CurrentRange
            Chantiers_Depenses_Extract_BaseCell SetOfChantiers, IndexChantiers, IndexDepense, CurrentRange
            If SetOfRange.StatusReal Then
                Set CurrentRange = SetOfRange.HeadCellReal.Cells(1 + IndexDepense, 4 + 3 * (IndexChantiers - 1))
                Chantiers_Depenses_Extract_Value SetOfChantiers, IndexChantiers, IndexDepense, CurrentRange, True
                Chantiers_Depenses_Extract_BaseCell SetOfChantiers, IndexChantiers, IndexDepense, CurrentRange, True
            End If
        Next IndexDepense
    Next IndexChantiers
    
    ' Autofinancements
    
    Set BaseCellLocal = BaseCellChantier.Worksheet.Cells(1, 2).EntireColumn.Find(Label_Autofinancement_Structure)
    If Not (BaseCellLocal Is Nothing) Then
        TestedValue = BaseCellLocal.Cells(-3, 1).Value
        If TestedValue = Label_Total_Financements Then
            NewFormatForAutofinancement = 2
        Else
            TestedValue = BaseCellLocal.Cells(6, 1).Value
            If TestedValue = Label_Autofinancement_Structure_Previous Then
                NewFormatForAutofinancement = 1
            Else
                NewFormatForAutofinancement = 0
            End If
        End If
        Chantiers = SetOfChantiers.Chantiers
        For IndexChantiers = 1 To NBChantiers
            ChantierTmp = Chantiers(IndexChantiers)
            ChantierTmp.AutoFinancementStructure = BaseCellLocal.Cells(1, 1 + IndexChantiers).Value
            ChantierTmp.AutoFinancementStructureFormula = Common_GetFormula(BaseCellLocal.Cells(1, 1 + IndexChantiers))
            If NewFormatForAutofinancement > 1 Then
                ChantierTmp.AutoFinancementAutres = BaseCellLocal.Cells(-2, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementAutresFormula = Common_GetFormula(BaseCellLocal.Cells(-2, 1 + IndexChantiers))
                ChantierTmp.AutoFinancementStructureAnneesPrecedentes = BaseCellLocal.Cells(5, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementStructureAnneesPrecedentesFormula = Common_GetFormula(BaseCellLocal.Cells(5, 1 + IndexChantiers))
                ChantierTmp.AutoFinancementAutresAnneesPrecedentes = BaseCellLocal.Cells(4, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementAutresAnneesPrecedentesFormula = Common_GetFormula(BaseCellLocal.Cells(4, 1 + IndexChantiers))
                ChantierTmp.CAanneesPrecedentes = BaseCellLocal.Cells(6, 1 + IndexChantiers).Value
                ChantierTmp.CAanneesPrecedentesFormula = Common_GetFormula(BaseCellLocal.Cells(6, 1 + IndexChantiers))
            Else
                ChantierTmp.AutoFinancementAutres = BaseCellLocal.Cells(2, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementAutresFormula = ""
                ChantierTmp.AutoFinancementStructureAnneesPrecedentesFormula = ""
                ChantierTmp.AutoFinancementAutresAnneesPrecedentesFormula = ""
                ChantierTmp.CAanneesPrecedentesFormula = ""
                If NewFormatForAutofinancement > 0 Then
                    ChantierTmp.AutoFinancementStructureAnneesPrecedentes = BaseCellLocal.Cells(6, 1 + IndexChantiers).Value
                    ChantierTmp.AutoFinancementAutresAnneesPrecedentes = BaseCellLocal.Cells(7, 1 + IndexChantiers).Value
                    ChantierTmp.CAanneesPrecedentes = BaseCellLocal.Cells(8, 1 + IndexChantiers).Value
                End If
            End If
            Chantiers(IndexChantiers) = ChantierTmp
        Next IndexChantiers
        SetOfChantiers.Chantiers = Chantiers
    End If
    
    Chantiers_Depenses_Extract = SetOfChantiers

End Function

Public Sub Chantiers_Depenses_Extract_BaseCell( _
        SetOfChantiers As SetOfChantiers, _
        IdxChantiers As Integer, _
        IdxDepense As Integer, _
        newRange As Range, _
        Optional IsReal As Boolean = False _
    )
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim DepensesTmp() As DepenseChantier
    Dim TmpDepense As DepenseChantier
    
    Chantiers = SetOfChantiers.Chantiers
    ChantierTmp = Chantiers(IdxChantiers)
    DepensesTmp = ChantierTmp.Depenses
    TmpDepense = DepensesTmp(IdxDepense)
    If IsReal Then
        Set TmpDepense.BaseCellReal = newRange.Cells(1, 0)
    Else
        Set TmpDepense.BaseCell = newRange
    End If
    DepensesTmp(IdxDepense) = TmpDepense
    ChantierTmp.Depenses = DepensesTmp
    Chantiers(IdxChantiers) = ChantierTmp
    SetOfChantiers.Chantiers = Chantiers
End Sub

Public Sub Chantiers_Depenses_Extract_Name(SetOfChantiers As SetOfChantiers, IdxChantiers As Integer, IdxDepense As Integer, newName As String)
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim DepensesTmp() As DepenseChantier
    Dim TmpDepense As DepenseChantier
    
    Chantiers = SetOfChantiers.Chantiers
    ChantierTmp = Chantiers(IdxChantiers)
    DepensesTmp = ChantierTmp.Depenses
    TmpDepense = DepensesTmp(IdxDepense)
    TmpDepense.Nom = newName
    DepensesTmp(IdxDepense) = TmpDepense
    ChantierTmp.Depenses = DepensesTmp
    Chantiers(IdxChantiers) = ChantierTmp
    SetOfChantiers.Chantiers = Chantiers
End Sub

Public Sub Chantiers_Depenses_Extract_Value( _
        SetOfChantiers As SetOfChantiers, _
        IdxChantiers As Integer, _
        IdxDepense As Integer, _
        CurrentCell As Range, _
        Optional IsReal As Boolean = False _
    )
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim DepensesTmp() As DepenseChantier
    Dim TmpDepense As DepenseChantier
    
    Chantiers = SetOfChantiers.Chantiers
    ChantierTmp = Chantiers(IdxChantiers)
    DepensesTmp = ChantierTmp.Depenses
    TmpDepense = DepensesTmp(IdxDepense)
    If IsReal Then
        TmpDepense.ValeurReal = CurrentCell.Value
        TmpDepense.FormulaReal = Common_GetFormula(CurrentCell)
    Else
        TmpDepense.Valeur = CurrentCell.Value
        TmpDepense.Formula = Common_GetFormula(CurrentCell)
    End If
    DepensesTmp(IdxDepense) = TmpDepense
    ChantierTmp.Depenses = DepensesTmp
    Chantiers(IdxChantiers) = ChantierTmp
    SetOfChantiers.Chantiers = Chantiers
End Sub

Public Function Chantiers_Depenses_Insert_One() As Integer
    
    Dim ChantierSheet As Worksheet
    Dim NBChantiers As Integer
    Dim NBSalariesAndCat As Integer
    Dim Previous As Integer
    Dim SetOfRange As SetOfRange
    
    Chantiers_Depenses_Insert_One = 0
    Set ChantierSheet = ThisWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Function
    End If
    SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheet, Nothing)
    If Not SetOfRange.Status Then
        Exit Function
    End If
    
    SetSilent

    NBChantiers = GetNbChantiers(ThisWorkbook)
    Chantiers_Depenses_Insert_One = NBChantiers
    Previous = SetOfRange.ResultCell.Row - SetOfRange.HeadCell.Row - 1

    Common_InsertRows _
        SetOfRange.HeadCell, _
        Previous, _
        Previous + 1, _
        False, _
        3 + NBChantiers, _
        False

    SetOfRange.ResultCell.Cells(0, 1).Value = "650 - Autre"
    Range( _
        SetOfRange.ResultCell.Cells(0, 2), _
        SetOfRange.ResultCell.Cells(0, 1 + NBChantiers) _
    ).ClearContents

    ' SetOfRange.EndCell = Cout_Journalier cell
    NBSalariesAndCat = SetOfRange.HeadCell.Row - SetOfRange.EndCell.Row
    Common_UpdateSumsByColumn _
        Range( _
            SetOfRange.EndCell.Cells(2, 2), _
            SetOfRange.ResultCell.Cells(0, 1 + NBChantiers) _
        ), _
        SetOfRange.ResultCell.Cells(1, 2), _
        Previous + NBSalariesAndCat
    
    SetActive
    
End Function

Public Sub Chantiers_Depenses_Insert_Real_One(NBChantiers As Integer)
    
    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim ColIndex As Integer
    Dim FormulaForSum As String
    Dim Index As Integer
    Dim NBSalariesAndCat As Integer
    Dim Previous As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfRangeReal As SetOfRange
    
    If NBChantiers < 1 Then
        Exit Sub
    End If
    Set ChantierSheet = ThisWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If
    Set ChantierSheetReal = ThisWorkbook.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    If ChantierSheetReal Is Nothing Then
        Exit Sub
    End If
    SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheet, ChantierSheetReal)
    If Not SetOfRange.Status Or Not SetOfRange.StatusReal Then
        Exit Sub
    End If
    
    SetSilent

    Previous = SetOfRange.ResultCellReal.Row - SetOfRange.HeadCellReal.Row - 1

    Common_InsertRows _
        SetOfRange.HeadCellReal, _
        Previous, _
        Previous + 1, _
        False, _
        3 * NBChantiers + 1 + NBExtraCols, _
        False

    ' SetOfRange.EndCellReal = Cout_Journalier cell
    NBSalariesAndCat = SetOfRange.HeadCellReal.Row - SetOfRange.EndCellReal.Row

    FormulaForSum = "="
    SetOfRange.ResultCellReal.Cells(0, 1).Formula = "=" & _
        CleanAddress(SetOfRange.ResultCell.Cells(0, 1).address(False, False, xlA1, True))

    For Index = 1 To NBChantiers
        SetOfRange.ResultCellReal.Cells(0, (Index - 1) * 3 + 2).Formula = "=" & _
            CleanAddress(SetOfRange.ResultCell.Cells(0, Index + 1).address(False, False, xlA1, True))
        SetOfRange.ResultCellReal.Cells(0, (Index - 1) * 3 + 3).ClearContents
        If Index > 1 Then
            FormulaForSum = FormulaForSum & "+"
        End If
        FormulaForSum = FormulaForSum & SetOfRange.ResultCellReal.Cells(0, (Index - 1) * 3 + 3).address(False, False, xlA1, False)

        For ColIndex = 1 To 2
            Common_UpdateSumsByColumn _
                Range( _
                    SetOfRange.EndCellReal.Cells(2, 1 + (Index - 1) * 3 + ColIndex), _
                    SetOfRange.ResultCellReal.Cells(0, 1 + (Index - 1) * 3 + ColIndex) _
                ), _
                SetOfRange.ResultCellReal.Cells(1, 1 + (Index - 1) * 3 + ColIndex), _
                Previous + NBSalariesAndCat
        Next ColIndex
    Next Index
    SetOfRange.ResultCellReal.Cells(0, NBChantiers * 3 + 2).Formula = FormulaForSum
    
    SetActive
    
End Sub

Public Sub Chantiers_Depenses_Remove_One()

    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim CurrentRange As Range
    Dim CurrentWs As Worksheet
    Dim NewLine As Integer
    Dim NBChantiers As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfRangeReal As SetOfRange
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set CurrentWs = wb.ActiveSheet
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If

    SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheet, ChantierSheetReal)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    NBChantiers = GetNbChantiers(wb)

    NewLine = Common_InputBox_Get_Line_Between( _
        Replace(T_Delete_Object_Of_Line, "%objectName%", T_Depense), _
        Replace(T_Line_To_Delete_For_Object, "%objectName%", T_Depense), _
        SetOfRange.HeadCell.Row + 1, _
        SetOfRange.ResultCell.Row - 1 _
    )
    
    If NewLine = -1 Then
        ' Cancel button
        Exit Sub
    End If

    If NewLine = 0 Then
        MsgBox Replace( _
            Replace(T_Given_Line_Is_Not_Line_Of_Object, "%objectName%", T_Depense), _
            "d'la", _
            "d'une" _
        )
        Exit Sub
    End If
    
    SetSilent

    Set CurrentRange = SetOfRange.HeadCell.Cells(NewLine - SetOfRange.HeadCell.Row + 1, 1)

    If Not (ChantierSheetReal Is Nothing) Then
        If SetOfRange.StatusReal Then
            Common_RemoveRows _
                SetOfRange.HeadCellReal, _
                NewLine - SetOfRange.HeadCellReal.Row, _
                NewLine - SetOfRange.HeadCellReal.Row - 1, _
                1 + NBChantiers * 3 + NBExtraCols
        End If
    End If

    Common_RemoveRows _
        SetOfRange.HeadCell, _
        NewLine - SetOfRange.HeadCell.Row, _
        NewLine - SetOfRange.HeadCell.Row - 1, _
        1 + NBChantiers + NBExtraCols
    
    CurrentWs.Activate
    SetActive
End Sub

Public Function Chantiers_Depenses_SetOfRange_Get( _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet _
    ) As SetOfRange

    Dim SetOfRange As SetOfRange
    Dim SetOfRangeFinal As SetOfRange
    
    SetOfRangeFinal = Chantiers_Depenses_SetOfRange_Get_Internal(ChantierSheet)
    SetOfRangeFinal.StatusReal = False
    If Not (ChantierSheetReal Is Nothing) Then
        Set SetOfRangeFinal.ChantierSheetReal = ChantierSheetReal
        SetOfRange = Chantiers_Depenses_SetOfRange_Get_Internal(ChantierSheetReal)
        If SetOfRange.Status Then
            SetOfRangeFinal.StatusReal = SetOfRange.Status
            Set SetOfRangeFinal.EndCellReal = SetOfRange.EndCell
            Set SetOfRangeFinal.HeadCellReal = SetOfRange.HeadCell
            Set SetOfRangeFinal.ResultCellReal = SetOfRange.ResultCell
        End If
    End If

    Chantiers_Depenses_SetOfRange_Get = SetOfRangeFinal
End Function

' HeadCell = Line before first line, first col
' ResultCell = cell of line for sum, colum with "total"
' EndCell = First line of indirect charges, first col
Public Function Chantiers_Depenses_SetOfRange_Get_Internal( _
        ChantierSheet As Worksheet _
    ) As SetOfRange

    Dim BaseCell As Range
    Dim BaseCellValue As String
    Dim CoutJJournalierCell As Range
    Dim Index As Integer
    Dim NewFormatForCat As Boolean
    Dim SetOfRange As SetOfRange
    Dim StructureCell As Range

    SetOfRange.Status = False
    SetOfRange.StatusReal = False
    Set SetOfRange.ChantierSheet = ChantierSheet

    Set BaseCell = Common_FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
    If BaseCell.Column > 1000 Then
        Exit Function
    End If
    BaseCellValue = BaseCell.Value
    If Left(BaseCellValue, Len("Chantier")) <> "Chantier" Then
        Exit Function
    End If

    Set StructureCell = BaseCell.Cells(3, 0)
    Set CoutJJournalierCell = StructureCell
    Set BaseCell = StructureCell
    BaseCellValue = Trim(BaseCell.Value)
    While BaseCellValue <> Label_Cout_J_Journalier And BaseCellValue <> "TOTAL" And BaseCell.Row < 200
        Set BaseCell = BaseCell.Cells(2, 1)
        BaseCellValue = Trim(BaseCell.Value)
    Wend
    
    If BaseCellValue <> Label_Cout_J_Journalier Then
        Exit Function
    End If
    
    Set CoutJJournalierCell = BaseCell
    If CoutJJournalierCell.Row - StructureCell.Row - 1 < 2 Then
        Exit Function
    End If

    Set SetOfRange.HeadCell = CoutJJournalierCell.Cells( _
        CoutJJournalierCell.Row - StructureCell.Row - 1, _
        0)
    NewFormatForCat = True
    For Index = 1 To 1 + NBCatOfCharges
        If NewFormatForCat Then
            BaseCellValue = Trim(SetOfRange.HeadCell.Cells(Index, 1).Value)
            If BaseCellValue = "" Then
                NewFormatForCat = False
            Else
                If Index > 1 _
                    And Left(BaseCellValue, Len("Charges ")) <> "Charges " Then
                    NewFormatForCat = False
                End If
            End If
        End If
    Next Index
    If NewFormatForCat Then
        Set SetOfRange.HeadCell = SetOfRange.HeadCell.Cells(1 + NBCatOfCharges, 1)
    End If
    Set BaseCell = SetOfRange.HeadCell.Cells(2, 2)
    BaseCellValue = Trim(BaseCell.Value)
    While BaseCellValue <> "TOTAL" And BaseCell.Row < 200
        Set BaseCell = BaseCell.Cells(2, 1)
        BaseCellValue = Trim(BaseCell.Value)
    Wend
    If BaseCellValue <> "TOTAL" Then
        Exit Function
    End If

    Set SetOfRange.ResultCell = BaseCell
    Set SetOfRange.EndCell = CoutJJournalierCell
    SetOfRange.Status = True

    Chantiers_Depenses_SetOfRange_Get_Internal = SetOfRange
End Function

Public Function Chantiers_Financements_BaseCell_Get( _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet _
    ) As SetOfRange

    Dim SetOfRange As SetOfRange

    SetOfRange.Status = False
    SetOfRange.StatusReal = False
    Set SetOfRange.ChantierSheet = ChantierSheet
    Set SetOfRange.HeadCell = ChantierSheet.Cells(1, 1).EntireColumn.Find(Label_Type_Financeur)
    If ChantierSheetReal Is Nothing Then
        Set SetOfRange.ChantierSheetReal = Nothing
        Set SetOfRange.HeadCellReal = Nothing
    Else
        Set SetOfRange.ChantierSheetReal = ChantierSheetReal
        Set SetOfRange.HeadCellReal = ChantierSheetReal.Cells(1, 1).EntireColumn.Find(Label_Type_Financeur)
    End If
    If Not (SetOfRange.HeadCell Is Nothing) Then
        Set SetOfRange.EndCell = ChantierSheet.Cells(1, 2).EntireColumn.Find(Label_Total_Financements)
        If Not (SetOfRange.EndCell Is Nothing) Then
            Set SetOfRange.ResultCell = SetOfRange.EndCell
            Set SetOfRange.EndCell = SetOfRange.EndCell.Cells(0, 0)
            SetOfRange.Status = True
        Else
            Set SetOfRange.EndCell = ChantierSheet.Cells(1, 2).EntireColumn.Find(Label_Autofinancement_Structure)
            If Not (SetOfRange.EndCell Is Nothing) Then
                Set SetOfRange.ResultCell = SetOfRange.EndCell
                Set SetOfRange.EndCell = SetOfRange.EndCell.Cells(0, 0)
                SetOfRange.Status = True
            End If
        End If
    End If
    If Not (SetOfRange.HeadCellReal Is Nothing) Then
        Set SetOfRange.EndCellReal = ChantierSheetReal.Cells(1, 2).EntireColumn.Find(Label_Total_Financements)
        If Not (SetOfRange.EndCellReal Is Nothing) Then
            Set SetOfRange.ResultCellReal = SetOfRange.EndCellReal
            Set SetOfRange.EndCellReal = SetOfRange.EndCellReal.Cells(0, 0)
            SetOfRange.StatusReal = True
        End If
    End If
    Chantiers_Financements_BaseCell_Get = SetOfRange
End Function

Public Function Chantiers_Financements_BaseCell_Get_ForV0(BaseCellChantier As Range) As Range
    Dim BaseCell As Range
    Set BaseCell = BaseCellChantier.Cells(1, 0).EntireColumn.Find(Label_Autofinancement_Structure)
    If BaseCell Is Nothing Then
        GoTo FinFunctionAvecErreur
    End If
    Set BaseCell = BaseCell.Cells(1, 2)
    While Left(BaseCell.Value, Len("Chantier")) <> "Chantier" And BaseCell.Row > (BaseCellChantier.Row + 1)
        Set BaseCell = BaseCell.Cells(0, 1)
    Wend
    If Left(BaseCell.Value, Len("Chantier")) <> "Chantier" Then
        GoTo FinFunctionAvecErreur
    End If
    
    Set BaseCell = BaseCell.Cells(2, -1)
    Set Chantiers_Financements_BaseCell_Get_ForV0 = BaseCell
    Exit Function
FinFunctionAvecErreur:
    Set Chantiers_Financements_BaseCell_Get_ForV0 = BaseCellChantier
End Function

Public Sub Chantiers_Financements_Clear( _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet, _
        NBChantiers As Integer _
    )
    Dim Index As Integer
    Dim SetOfRange As SetOfRange
    
    SetOfRange = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
    If SetOfRange.Status Then
        If SetOfRange.EndCell.Row > SetOfRange.HeadCell.Row + 1 Then
            Range( _
                    SetOfRange.HeadCell.Cells(2, 1), _
                    SetOfRange.EndCell.Cells(1, 3 + NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
        End If
    End If
    If SetOfRange.StatusReal Then
        If SetOfRange.EndCellReal.Row > SetOfRange.HeadCellReal.Row + 1 Then
            Range( _
                    SetOfRange.HeadCellReal.Cells(2, 1), _
                    SetOfRange.EndCellReal.Cells(1, 3 + 3 * NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
        End If
    End If
End Sub

Public Function Chantiers_Financements_Extract( _
        BaseCellChantier As Range, _
        NBChantiers As Integer, _
        Data As Data, _
        Optional ForV0 As Boolean = False _
        ) As SetOfChantiers
    Dim BaseCell As Range
    Dim ChantierSheetReal As Worksheet
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim CurrentCell As Range
    Dim LocalCounter As Integer
    Dim FinancementTmp As Financement
    Dim FinancementTmp1 As Financement
    Dim FinancementsTmp() As Financement
    Dim FinancementsTmp1() As Financement
    Dim FoundCell As Range
    Dim IndexChantiers As Integer
    Dim IndexFinancement As Integer
    Dim IndexType As Integer
    Dim IndexTypeName As Integer
    Dim NBFinancements As Integer
    Dim SetOfChantiers As SetOfChantiers
    Dim SetOfRange As SetOfRange
    Dim TypesFinancements As Variant
    Dim TypesStatuts As Variant
    Dim wb As Workbook
    
    Set wb = BaseCellChantier.Worksheet.Parent
    TypesFinancements = TypeFinancementsFromWb(wb)
    TypesStatuts = TypeStatut()
    
    Chantiers = Data.Chantiers
    SetOfChantiers.Chantiers = Chantiers
    
    If ForV0 Then
        Set BaseCell = Chantiers_Financements_BaseCell_Get_ForV0(BaseCellChantier)
        If BaseCell.address = BaseCellChantier.address Then
            GoTo FinFunction
        End If
        If BaseCell Is Nothing Then
            GoTo FinFunction
        End If
        Set BaseCell = BaseCell.Cells(2, 1)
        Set FoundCell = BaseCell.Cells(1, 2).EntireColumn.Find(Label_Total_Financements)
        If Not (FoundCell Is Nothing) Then
            NBFinancements = FoundCell.Row - BaseCell.Row
        Else
            NBFinancements = -1
        End If
        Set FoundCell = BaseCell.Cells(1, 2).EntireColumn.Find(Label_Autofinancement_Structure)
        If NBFinancements < 0 Or (FoundCell.Row < BaseCell.Row + NBFinancements) Then
            NBFinancements = FoundCell.Row - BaseCell.Row
        End If
    Else
        On Error Resume Next
        Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
        On Error GoTo 0
        SetOfRange = Chantiers_Financements_BaseCell_Get(BaseCellChantier.Worksheet, ChantierSheetReal)
        If Not (SetOfRange.Status) Then
            GoTo FinFunction
        End If
        Set BaseCell = SetOfRange.HeadCell.Cells(2, 1)
        NBFinancements = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
    End If
    
    ' Filter only financement removing second lines if begining by "statut"
    LocalCounter = 0
    For IndexFinancement = 1 To NBFinancements
        If BaseCell.Cells(IndexFinancement, 2).Value <> "Statut" Then
            LocalCounter = LocalCounter + 1
        End If
    Next IndexFinancement
    NBFinancements = LocalCounter
    
    For IndexChantiers = 1 To NBChantiers
        ChantierTmp = Chantiers(IndexChantiers)
        FinancementsTmp = getDefaultFinancements(NBFinancements)
        ChantierTmp.Financements = FinancementsTmp
        Chantiers(IndexChantiers) = ChantierTmp
    Next IndexChantiers
    
    ' Extraction des types avec le chantier 1
    LocalCounter = 1
    ChantierTmp = Chantiers(1)
    FinancementsTmp1 = ChantierTmp.Financements
    For IndexFinancement = 1 To NBFinancements
        FinancementTmp1 = FinancementsTmp1(IndexFinancement)
        FinancementTmp1.Nom = BaseCell.Cells(LocalCounter, 2).Value
        IndexType = 0
        For IndexTypeName = 1 To UBound(TypesFinancements)
            If TypesFinancements(IndexTypeName) = BaseCell.Cells(LocalCounter, 1).Value Then
                IndexType = IndexTypeName
            End If
        Next IndexTypeName
        FinancementTmp1.TypeFinancement = IndexType
        If IndexType > 0 Then
            LocalCounter = LocalCounter + 1
        Else
            If ForV0 And FinancementTmp1.Nom <> "" Then
                If Trim(FinancementTmp1.Nom) = "Formations" Or _
                    Trim(FinancementTmp1.Nom) = "Prestations" Or _
                    Trim(FinancementTmp1.Nom) = "Cotisations" Then
                    FinancementTmp1.TypeFinancement = 0
                Else
                    FinancementTmp1.TypeFinancement = FindTypeFinancementIndex("Autres")
                End If
            End If
        End If
        FinancementsTmp1(IndexFinancement) = FinancementTmp1
        LocalCounter = LocalCounter + 1
    Next IndexFinancement
    ChantierTmp.Financements = FinancementsTmp1
    Chantiers(1) = ChantierTmp
    
    ' Extraction des valeurs
    For IndexChantiers = 1 To NBChantiers
        LocalCounter = 1
        ChantierTmp = Chantiers(IndexChantiers)
        For IndexFinancement = 1 To NBFinancements
            FinancementsTmp = ChantierTmp.Financements
            FinancementTmp = FinancementsTmp(IndexFinancement)
            FinancementTmp1 = FinancementsTmp1(IndexFinancement)
            ' recuperation du type depuis le chantier 1
            If IndexChantiers > 1 Then
                FinancementTmp.Nom = FinancementTmp1.Nom
                FinancementTmp.TypeFinancement = FinancementTmp1.TypeFinancement
            End If
            Set CurrentCell = BaseCell.Cells(LocalCounter, IndexChantiers + 2)
            FinancementTmp.Valeur = CurrentCell.Value
            FinancementTmp.Formula = Common_GetFormula(CurrentCell)
            Set FinancementTmp.BaseCell = CurrentCell
            If Not ForV0 Then
                If SetOfRange.StatusReal Then
                    Set CurrentCell = SetOfRange.HeadCellReal.Cells( _
                            1 + LocalCounter, _
                            1 + 3 * IndexChantiers _
                        )
                    FinancementTmp.ValeurReal = CurrentCell.Value
                    FinancementTmp.FormulaReal = Common_GetFormula(CurrentCell)
                    Set FinancementTmp.BaseCellReal = CurrentCell
                Else
                    FinancementTmp.ValeurReal = 0
                    Set FinancementTmp.BaseCellReal = Nothing
                End If
            Else
                FinancementTmp.ValeurReal = 0
                Set FinancementTmp.BaseCellReal = Nothing
            End If
            
            If FinancementTmp.TypeFinancement > 0 And Not ForV0 Then
                IndexType = 0
                For IndexTypeName = 1 To UBound(TypesStatuts)
                    If TypesStatuts(IndexTypeName) = BaseCell.Cells(LocalCounter + 1, IndexChantiers + 2).Value Then
                        IndexType = IndexTypeName
                    End If
                Next IndexTypeName
                FinancementTmp.Statut = IndexType
                LocalCounter = LocalCounter + 1
            Else
                FinancementTmp.Statut = 0
            End If
            LocalCounter = LocalCounter + 1
            FinancementsTmp(IndexFinancement) = FinancementTmp
            ChantierTmp.Financements = FinancementsTmp
        Next IndexFinancement
        Chantiers(IndexChantiers) = ChantierTmp
    Next IndexChantiers
    
    SetOfChantiers.Chantiers = Chantiers
    
FinFunction:
    Chantiers_Financements_Extract = SetOfChantiers

End Function

Public Sub Chantiers_Financements_Add_One(wb As Workbook, _
        NBChantiers As Integer, _
        NewFinancementInChantier As FinancementComplet, _
        Optional Nom As String = "", _
        Optional TypeFinancement As Integer = 0, _
        Optional RetirerLignesVides As Boolean = False)

    Dim SetOfRange As SetOfRange
    
    If Not (NewFinancementInChantier.Status) And Nom = "" Then
        ' EmptyChantier
        Set wb = ThisWorkbook
        OpenUserForm
        Exit Sub
    End If

    SetOfRange = Chantiers_Financements_Add_Prepare(wb, NBChantiers)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, wb, NBChantiers, NewFinancementInChantier, Nom, TypeFinancement)
    
    Chantiers_Totals_UpdateSumsAndFormat wb, SetOfRange, NBChantiers

End Sub

Public Sub Chantiers_Financements_Format_Add_ValidationDossier(CurrentRange As Range)
    
    With CurrentRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=STATUT_DOSSIER"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Public Function Chantiers_Financements_Add_Internal( _
        SetOfRange As SetOfRange, _
        wb As Workbook, _
        NBChantiers As Integer, _
        NewFinancementInChantier As FinancementComplet, _
        Optional Nom As String = "", _
        Optional TypeFinancement As Integer = 0 _
    ) As SetOfRange

    Dim BaseCell As Range
    Dim CurrentAddress As String
    Dim Index As Integer
    Dim IsEmptyRow As Boolean
    Dim NBLinesToClean As Integer
    Dim NBNewLines As Integer
    Dim NBRows As Integer
    Dim ShoudInsert As Boolean
    Dim TmpFinancement As Financement
    Dim TypeFinancementStr As String
    Dim ValueOfFirstCellOnCurrentLine As String
    Dim ValueOfSecondCellOnNextLine As String
    Dim WorkingRange As Range
    Dim WorkingRangeReal As Range
    
    TypeFinancementStr = Common_GetTypeFinancementStr(wb, TypeFinancement, NewFinancementInChantier)
    Chantiers_Financements_Add_Internal = SetOfRange

    NBRows = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
    Set BaseCell = SetOfRange.EndCell
    If NBRows = 0 Then
        ' insert just after HeadCell
        ShoudInsert = True
        Set BaseCell = SetOfRange.HeadCell
    Else
        ShoudInsert = False
        If TypeFinancementStr <> "" Then
            ' search existing similar TypeFinancement
            For Index = 1 To NBRows
                If Not ShoudInsert Then
                    Set WorkingRange = SetOfRange.HeadCell.Cells(1 + Index, 1)
                    ValueOfFirstCellOnCurrentLine = WorkingRange.Value
                    ValueOfSecondCellOnNextLine = WorkingRange.Cells(2, 2).Value
                    If ValueOfFirstCellOnCurrentLine = TypeFinancementStr _
                        And ValueOfSecondCellOnNextLine = "Statut" Then
                        Set BaseCell = WorkingRange.Cells(2, 1)
                        ShoudInsert = True
                    Else
                        If BaseCell.Row = SetOfRange.EndCell.Row _
                                And Chantiers_Financements_IsEmpty(WorkingRange, 2 + NBChantiers, False) Then
                            Set BaseCell = WorkingRange
                        End If
                    End If
                End If
            Next Index
        Else
            ShoudInsert = True
            ' search empty line
            For Index = 1 To NBRows
                If ShoudInsert Then
                    Set WorkingRange = SetOfRange.HeadCell.Cells(1 + Index, 1)
                    ValueOfFirstCellOnCurrentLine = WorkingRange.Value
                    ValueOfSecondCellOnNextLine = WorkingRange.Cells(2, 2).Value
                    If Chantiers_Financements_IsEmpty(WorkingRange, 2 + NBChantiers, False) _
                        And ValueOfSecondCellOnNextLine <> "Statut" Then
                        Set BaseCell = WorkingRange
                        ShoudInsert = False
                    End If
                End If
            Next Index
        End If
    End If

    IsEmptyRow = Chantiers_Financements_IsEmpty(BaseCell, 2 + NBChantiers, False)
    If ShoudInsert _
        Or TypeFinancementStr <> "" _
        Or Not IsEmptyRow _
        Then
        If TypeFinancementStr <> "" _
            And ( _
                ShoudInsert _
                Or Not IsEmptyRow _
            ) Then
            NBNewLines = 2
        Else
            NBNewLines = 1
            If TypeFinancementStr <> "" Then
                ' because insertion could choose one line too bottom
                Set BaseCell = BaseCell.Cells(0, 1)
            End If
        End If
        If TypeFinancementStr <> "" Then
            NBLinesToClean = 2
        Else
            NBLinesToClean = 1
        End If
        
        Common_InsertRows _
            SetOfRange.HeadCell, _
            BaseCell.Row - SetOfRange.HeadCell.Row, _
            BaseCell.Row - SetOfRange.HeadCell.Row + NBNewLines, _
            False, _
            1 + NBChantiers + NBExtraCols, _
            False
        
        Set WorkingRange = BaseCell.Cells(2, 1)
        ' Clean values
        Range(WorkingRange, WorkingRange.Cells(NBLinesToClean, 2 + NBChantiers)).Value = ""
        Range(WorkingRange, WorkingRange.Cells(NBLinesToClean, 3 + NBChantiers)).MergeCells = False

        ' real
        If SetOfRange.StatusReal Then
            Common_InsertRows _
                SetOfRange.HeadCellReal, _
                BaseCell.Row - SetOfRange.HeadCell.Row, _
                BaseCell.Row - SetOfRange.HeadCell.Row + NBNewLines, _
                False, _
                1 + 3 * NBChantiers + NBExtraCols, _
                False
            ' Clean values
            Set WorkingRangeReal = SetOfRange.HeadCellReal.Cells(WorkingRange.Row - SetOfRange.HeadCell.Row + 1, 1)
            ' Reset formula and clean other values
            Range(WorkingRangeReal, WorkingRangeReal.Cells(NBLinesToClean, 2)).Value = ""
            Range( _
                WorkingRangeReal.Cells(1, 1), _
                WorkingRangeReal.Cells(NBLinesToClean, 3 + 3 * NBChantiers) _
                ).MergeCells = False
            For Index = 1 To NBChantiers
                Chantiers_Financements_Real_Set _
                    SetOfRange, _
                    WorkingRange.Row - SetOfRange.HeadCell.Row, _
                    Index, _
                    "", _
                    (NBLinesToClean > 1)
            Next Index
        End If
    Else
        Set WorkingRange = BaseCell
        If SetOfRange.StatusReal Then
            Set WorkingRangeReal = SetOfRange.HeadCellReal.Cells(WorkingRange.Row - SetOfRange.HeadCell.Row + 1, 1)
        End If
    End If

    WorkingRange.Cells(1, NBChantiers + 3).Formula = "=SUM(" & _
        Range( _
            WorkingRange.Cells(1, 3), _
            WorkingRange.Cells(1, NBChantiers + 2) _
        ).address(False, False, xlA1) & _
    ")"

    If SetOfRange.StatusReal Then
        WorkingRangeReal.Cells(1, 3 * NBChantiers + 3).Formula = "=" & _
            WorkingRangeReal.Cells(1, 4).address(False, False, xlA1, False)
    End If

    If TypeFinancementStr <> "" Then
        WorkingRange.Cells(1, 1).Value = TypeFinancementStr
        WorkingRange.Cells(2, 2).Value = "Statut"
        WorkingRange.Cells(2, 3 + NBChantiers).Value = ""
        If SetOfRange.StatusReal Then
            WorkingRangeReal.Cells(1, 1).Value = TypeFinancementStr
            WorkingRangeReal.Cells(2, 2).Value = "Statut"
            WorkingRangeReal.Cells(2, 3 + 3 * NBChantiers).Value = ""
        End If
    End If

    If SetOfRange.StatusReal Then
        CurrentAddress = CleanAddress(WorkingRange.Cells(1, 2).address(False, False, xlA1, True))
        WorkingRangeReal.Cells(1, 2).Formula = "=IF(" & CurrentAddress & "="""",""""," & CurrentAddress & ")"
    End If

    If Not (NewFinancementInChantier.Status) Then
        WorkingRange.Cells(1, 2).Value = Nom
    Else
        TmpFinancement = NewFinancementInChantier.Financements(1)
        WorkingRange.Cells(1, 2).Value = TmpFinancement.Nom
        For Index = 1 To UBound(NewFinancementInChantier.Financements)
            TmpFinancement = NewFinancementInChantier.Financements(Index)
            Common_SetFormula _
                WorkingRange.Cells(1, 2 + Index), _
                TmpFinancement.Valeur, _
                TmpFinancement.Formula, _
                True
            If SetOfRange.StatusReal Then
                Common_SetFormula _
                    WorkingRangeReal.Cells(1, 3 * Index + 1), _
                    TmpFinancement.ValeurReal, _
                    TmpFinancement.FormulaReal, _
                    True
            End If
            If TypeFinancementStr <> "" Then
                If TmpFinancement.Statut <> 0 Then
                    WorkingRange.Cells(2, 2 + Index).Value = TypeStatut()(TmpFinancement.Statut)
                End If
            End If
        Next Index
    End If
    ' adjust and return
    Set Chantiers_Financements_Add_Internal.EndCell = SetOfRange.ResultCell.Cells(0, 0)
End Function

Public Function Chantiers_Financements_Add_Prepare( _
        wb As Workbook, _
        NBChantiers As Integer, _
        Optional RetirerLignesVides As Boolean = False _
    ) As SetOfRange

    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim SetOfRange As SetOfRange

    ' Default
    SetOfRange.Status = False
    SetOfRange.StatusReal = False
    Chantiers_Financements_Add_Prepare = SetOfRange

    
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers)
        Exit Function
    End If
    Set SetOfRange.ChantierSheet = ChantierSheet

    Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    If ChantierSheetReal Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers_realise)
        Exit Function
    End If
    Set SetOfRange.ChantierSheetReal = ChantierSheetReal
    Chantiers_Financements_Add_Prepare = SetOfRange

    SetOfRange = Chantiers_Financements_BaseCell_Get(SetOfRange.ChantierSheet, SetOfRange.ChantierSheetReal)
    Chantiers_Financements_Add_Prepare = SetOfRange
    If Not SetOfRange.Status Then
        Exit Function
    End If

    If RetirerLignesVides Then
        Chantiers_Financements_Empty_Lines_Remove SetOfRange, NBChantiers
        SetOfRange = Chantiers_Financements_BaseCell_Get(SetOfRange.ChantierSheet, SetOfRange.ChantierSheetReal)
        Chantiers_Financements_Add_Prepare = SetOfRange
    End If

End Function

Public Sub Chantiers_Financements_Empty_Lines_Remove( _
    SetOfRange As SetOfRange, _
    NBChantiers As Integer _
    )

    Dim CurrentIndex As Integer
    Dim FirstCellOfLine As Range
    Dim FirstCellOfLineReal As Range
    Dim NBRows As Integer
    Dim IndexLine As Integer
    Dim ValueOfFirstCellOfLine As String
    Dim ValueOfSecondCellOfLine As String
    Dim ValueOfSecondCellOfNextLine As String

    NBRows = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
    CurrentIndex = 1

    For IndexLine = 1 To NBRows
        Set FirstCellOfLine = SetOfRange.HeadCell.Cells(1 + CurrentIndex, 1)
        Set FirstCellOfLineReal = SetOfRange.HeadCellReal.Cells(1 + CurrentIndex, 1)
        ValueOfFirstCellOfLine = FirstCellOfLine.Value
        ValueOfSecondCellOfLine = FirstCellOfLine.Cells(1, 2).Value
        ValueOfSecondCellOfNextLine = FirstCellOfLine.Cells(2, 2).Value
        If ValueOfSecondCellOfLine = "" _
            And ValueOfFirstCellOfLine <> "" _
            And ValueOfFirstCellOfLine <> Empty _
            And ValueOfSecondCellOfNextLine = "Statut" Then
            ' two lines
            If Chantiers_Financements_IsEmpty(FirstCellOfLine.Cells(1, 3), NBChantiers, True) Then
                Range( _
                    FirstCellOfLine, _
                    FirstCellOfLine.Cells(2, 3 + NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
                Range( _
                    FirstCellOfLineReal, _
                    FirstCellOfLineReal.Cells(2, 3 + 3 * NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
                CurrentIndex = CurrentIndex - 2
            End If
        Else
            If ValueOfFirstCellOfLine = "" _
                And ValueOfSecondCellOfLine = "" _
                And ValueOfSecondCellOfNextLine <> "Statut" Then
                ' one line
                If Chantiers_Financements_IsEmpty(FirstCellOfLine.Cells(1, 3), NBChantiers, False) Then
                    Range( _
                        FirstCellOfLine, _
                        FirstCellOfLine.Cells(1, 3 + NBChantiers + NBExtraCols) _
                    ).Delete Shift:=xlUp
                    Range( _
                        FirstCellOfLineReal, _
                        FirstCellOfLineReal.Cells(1, 3 + 3 * NBChantiers + NBExtraCols) _
                    ).Delete Shift:=xlUp
                    CurrentIndex = CurrentIndex - 1
                End If
            End If
        End If
        CurrentIndex = CurrentIndex + 1
    Next IndexLine
End Sub

Public Sub Chantiers_Financements_Init(wb As Workbook, NBFinancements As Integer, Optional Init As Boolean = False)

    Dim NBChantiers As Integer
    Dim FinancementCompletTmp As FinancementComplet
    FinancementCompletTmp = getDefaultFinancementComplet()
    Dim FinancementTmp As Financement
    Dim TypesFinancements() As String
    Dim Index As Integer
    Dim IndexLoop As Integer
    Dim SetOfRange As SetOfRange
    
    TypesFinancements = TypeFinancementsFromWb(wb)
    FinancementCompletTmp.Status = False
    
    NBChantiers = GetNbChantiers(wb)
    If NBChantiers < 1 Then
        Exit Sub
    End If
    If NBFinancements < 0 Or (NBFinancements = 0 And Init) Then
        Exit Sub
    End If

    SetOfRange = Chantiers_Financements_Add_Prepare(wb, NBChantiers, Init)
    If Not SetOfRange.Status Then
        Exit Sub
    End If
    
    For Index = 1 To UBound(TypesFinancements)
        For IndexLoop = 1 To NBFinancements
            SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Client " & (IndexLoop + (Index - 1) * NBFinancements), Index)
        Next IndexLoop
    Next Index
    For IndexLoop = 1 To NBFinancements
        SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Formations", 0)
    Next IndexLoop
    For IndexLoop = 1 To NBFinancements
        SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Prestations", 0)
    Next IndexLoop
    For IndexLoop = 1 To NBFinancements
        SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Cotisations", 0)
    Next IndexLoop

    Chantiers_Totals_UpdateSumsAndFormat wb, SetOfRange, NBChantiers
End Sub

Public Function Chantiers_Financements_IsEmpty(FirstCell As Range, NBChantiers As Integer, TwoLines As Boolean) As Boolean
    
    Dim CurrentValue As String
    Dim Index As Integer
    
    Chantiers_Financements_IsEmpty = False
    For Index = 1 To NBChantiers
        CurrentValue = FirstCell.Cells(1, Index).Value
        If CurrentValue <> "" Or CurrentValue <> Empty Then
            Exit Function
        End If
        If TwoLines Then
            CurrentValue = FirstCell.Cells(2, Index).Value
            If CurrentValue <> "" Or CurrentValue <> Empty Then
                Exit Function
            End If
        End If
    Next Index

    Chantiers_Financements_IsEmpty = True
End Function

Public Sub Chantiers_Financements_Real_Set( _
    SetOfRange As SetOfRange, _
    NBFirstLineOfFinancement As Integer, _
    NBofChantier As Integer, _
    ValueToSet As String, _
    Optional IsTwoLines As Boolean = False _
    )

    Dim CurrentAddress As String
    Dim CurrentBaseCell As Range

    Set CurrentBaseCell = SetOfRange.HeadCellReal.Cells(1 + NBFirstLineOfFinancement, 3 + 3 * (NBofChantier - 1))
    CurrentBaseCell.Formula = "=" & CleanAddress(SetOfRange.HeadCell.Cells(1 + NBFirstLineOfFinancement, 2 + NBofChantier).address(False, False, xlA1, True))
    If ValueToSet = "" Then
        CurrentBaseCell.Cells(1, 2).Value = ""
    Else
        CurrentBaseCell.Cells(1, 2).Value = CDbl(ValueToSet)
    End If
    CurrentBaseCell.Cells(1, 3).Formula = "=" _
        & CleanAddress(CurrentBaseCell.Cells(1, 2).address(False, False, xlA1, True)) _
        & "/(" _
        & CleanAddress(CurrentBaseCell.address(False, False, xlA1, True)) _
        & "+1E-9)"
    If IsTwoLines Then
        CurrentAddress = CleanAddress(SetOfRange.HeadCell.Cells(2 + NBFirstLineOfFinancement, 2 + NBofChantier).address(False, False, xlA1, True))
        CurrentBaseCell.Cells(2, 1).Formula = "=IF(" & CurrentAddress & "="""",""""," & CurrentAddress & ")"
        ' remove value to prevent errors when merging
        CurrentBaseCell.Cells(2, 2).Value = ""
        CurrentBaseCell.Cells(2, 3).Value = ""
        Range(CurrentBaseCell.Cells(2, 1), CurrentBaseCell.Cells(2, 3)).MergeCells = True
    End If
End Sub

Public Sub Chantiers_Financements_Remove_One()

    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim CurrentRange As Range
    Dim CurrentWs As Worksheet
    Dim NewLine As Integer
    Dim NBChantiers As Integer
    Dim SetOfRange As SetOfRange
    Dim ValueToTest As String
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set CurrentWs = wb.ActiveSheet

    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If
    Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    If ChantierSheetReal Is Nothing Then
        Exit Sub
    End If

    SetOfRange = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    NBChantiers = GetNbChantiers(wb)

    NewLine = Common_InputBox_Get_Line_Between( _
        Replace(T_Delete_Object_Of_Line, "%objectName%", T_Financement), _
        Replace( _
            Replace(T_Line_To_Delete_For_Object, "%objectName%", T_Financement), _
            "de le", _
            "du" _
        ), _
        SetOfRange.HeadCell.Row + 1, _
        SetOfRange.EndCell.Row _
    )
    
    If NewLine = -1 Then
        ' Cancel button
        Exit Sub
    End If

    If NewLine = 0 Then
        MsgBox Replace( _
            Replace(T_Given_Line_Is_Not_Line_Of_Object, "%objectName%", T_Financement), _
            "d'le", _
            "d'un" _
        )
        Exit Sub
    End If
    
    SetSilent

    Set CurrentRange = SetOfRange.HeadCell.Cells(NewLine - SetOfRange.HeadCell.Row + 1, 1)
    ValueToTest = CurrentRange.Cells(1, 2).Value
    If ValueToTest = "Statut" Then
        Common_RemoveRows _
            SetOfRange.HeadCell, _
            NewLine - SetOfRange.HeadCell.Row, _
            NewLine - SetOfRange.HeadCell.Row - 2, _
            1 + NBChantiers + NBExtraCols
    Else
        ValueToTest = CurrentRange.Cells(2, 2).Value
        If ValueToTest = "Statut" Then
            If SetOfRange.StatusReal Then
                Common_RemoveRows _
                    SetOfRange.HeadCellReal, _
                    NewLine - SetOfRange.HeadCellReal.Row + 1, _
                    NewLine - SetOfRange.HeadCellReal.Row - 1, _
                    1 + 3 * NBChantiers + NBExtraCols
            End If
            Common_RemoveRows _
                SetOfRange.HeadCell, _
                NewLine - SetOfRange.HeadCell.Row + 1, _
                NewLine - SetOfRange.HeadCell.Row - 1, _
                1 + NBChantiers + NBExtraCols
        Else
            If SetOfRange.StatusReal Then
                Common_RemoveRows _
                    SetOfRange.HeadCellReal, _
                    NewLine - SetOfRange.HeadCellReal.Row, _
                    NewLine - SetOfRange.HeadCellReal.Row - 1, _
                    1 + 3 * NBChantiers + NBExtraCols
            End If
            Common_RemoveRows _
                SetOfRange.HeadCell, _
                NewLine - SetOfRange.HeadCell.Row, _
                NewLine - SetOfRange.HeadCell.Row - 1, _
                1 + NBChantiers + NBExtraCols
        End If
    End If

    Chantiers_Totals_UpdateSumsAndFormat wb, SetOfRange, NBChantiers, False

    CurrentWs.Activate
    SetActive
End Sub

Public Sub Chantiers_Financements_Totals_RenewFormula( _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet, _
        NBChantiers As Integer _
    )
    Dim Formula As String
    Dim FormulaReal1 As String
    Dim FormulaReal2 As String
    Dim IndexChantier As Integer
    Dim IndexLigne As Integer
    Dim NBRowsFinancements As Integer
    Dim SetOfRange As SetOfRange

    SetOfRange = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
    If SetOfRange.Status Then
        NBRowsFinancements = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
        For IndexChantier = 1 To NBChantiers
            Formula = "="
            FormulaReal1 = "="
            FormulaReal2 = "="
            For IndexLigne = 1 To NBRowsFinancements
                If SetOfRange.HeadCell.Cells(1 + IndexLigne, 2).Value <> "Statut" Then
                    If Formula <> "=" Then
                        Formula = Formula & "+"
                    End If
                    Formula = Formula & _
                        SetOfRange.HeadCell.Cells(1 + IndexLigne, 2 + IndexChantier) _
                            .address(False, False, xlA1, False)
                    If FormulaReal1 <> "=" Then
                        FormulaReal1 = FormulaReal1 & "+"
                    End If
                    FormulaReal1 = FormulaReal1 & _
                        SetOfRange.HeadCellReal.Cells(1 + IndexLigne, 3 * IndexChantier) _
                            .address(False, False, xlA1, False)
                    If FormulaReal2 <> "=" Then
                        FormulaReal2 = FormulaReal2 & "+"
                    End If
                    FormulaReal2 = FormulaReal2 & _
                        SetOfRange.HeadCellReal.Cells(1 + IndexLigne, 1 + 3 * IndexChantier) _
                            .address(False, False, xlA1, False)
                End If
            Next IndexLigne
            If Formula = "=" Then
                Formula = "=0"
            End If
            SetOfRange.ResultCell.Cells(1, 1 + IndexChantier).Formula = Formula
            If FormulaReal1 = "=" Then
                FormulaReal1 = "=0"
            End If
            SetOfRange.ResultCellReal.Cells(1, 2 + 3 * (IndexChantier - 1)).Formula = FormulaReal1
            If FormulaReal2 = "=" Then
                FormulaReal2 = "=0"
            End If
            SetOfRange.ResultCellReal.Cells(1, 3 + 3 * (IndexChantier - 1)).Formula = FormulaReal2
        Next IndexChantier
    End If
End Sub

Public Sub Chantiers_Format_Set( _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet, _
        NBChantiers As Integer, _
        Optional AddTopBorder As Boolean = True, _
        Optional AddBottomBorder As Boolean = True _
    )

    Dim ColumnIndex As Integer
    Dim CurrentArea As Range
    Dim IsEuros As Boolean
    Dim IsPercent As Boolean
    Dim IsReal As Boolean
    Dim LastColumnNumber As Integer
    Dim NBColumns As Integer
    Dim NBRows As Integer
    Dim RowIndex As Integer
    Dim SetLeftCol As Boolean
    Dim SetOfRange As SetOfRange
    Dim SetRightCol As Boolean
    Dim ValueOfSecondCellOfLine As String

    SetOfRange = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
    If Not SetOfRange.Status Then
        Exit Sub
    End If
    
    IsReal = Not (ChantierSheetReal Is Nothing)
    If IsReal Then
        If Not SetOfRange.StatusReal Then
            Exit Sub
        End If
        LastColumnNumber = 3 + 3 * NBChantiers
        Set CurrentArea = Range( _
            SetOfRange.HeadCellReal.Cells(2, 1), _
            SetOfRange.EndCellReal.Cells(1, LastColumnNumber) _
        )
    Else
        LastColumnNumber = 3 + NBChantiers
        Set CurrentArea = Range( _
            SetOfRange.HeadCell.Cells(2, 1), _
            SetOfRange.EndCell.Cells(1, LastColumnNumber) _
        )
    End If

    NBColumns = CurrentArea.Columns.Count
    NBRows = CurrentArea.Rows.Count

    For RowIndex = 1 To NBRows
        ValueOfSecondCellOfLine = CurrentArea.Cells(RowIndex, 2).Value
        For ColumnIndex = 1 To NBColumns
            If Not IsReal Then
                IsEuros = (ColumnIndex > 2)
                IsPercent = False
                SetLeftCol = True
                SetRightCol = True
            Else
                IsEuros = (ColumnIndex > 2 And ((ColumnIndex - 3) Mod 3) <> 2)
                IsPercent = (ColumnIndex > 2 And ((ColumnIndex - 3) Mod 3) = 2)
                SetLeftCol = ( _
                    ValueOfSecondCellOfLine = "Statut" _
                    Or ColumnIndex <= 2 _
                    Or ColumnIndex = NBColumns _
                    Or ((ColumnIndex - 3) Mod 3) = 0 _
                )
                SetRightCol = ( _
                    ValueOfSecondCellOfLine = "Statut" _
                    Or ColumnIndex <= 2 _
                    Or ColumnIndex = NBColumns _
                    Or ((ColumnIndex - 3) Mod 3) = 2 _
                )
            End If
            DefinirFormatPourChantier CurrentArea.Cells(RowIndex, ColumnIndex), _
                (AddTopBorder And RowIndex = 1), _
                (AddBottomBorder And RowIndex = NBRows), _
                (ColumnIndex = 2 Or ColumnIndex = NBColumns), _
                (ValueOfSecondCellOfLine = "Statut" And ColumnIndex <= 2), _
                (ColumnIndex = NBColumns), _
                IsEuros, _
                IsPercent, _
                SetLeftCol, _
                SetRightCol
            If IsReal _
                And ValueOfSecondCellOfLine = "Statut" _
                And ColumnIndex > 2 _
                And ColumnIndex < NBColumns _
                And ((ColumnIndex - 3) Mod 3) = 2 Then
                CurrentArea.Cells(RowIndex, ColumnIndex - 1).Value = ""
                CurrentArea.Cells(RowIndex, ColumnIndex).Value = ""
                Range( _
                    CurrentArea.Cells(RowIndex, ColumnIndex - 2), _
                    CurrentArea.Cells(RowIndex, ColumnIndex) _
                ).MergeCells = True
            End If
        Next ColumnIndex
        If ValueOfSecondCellOfLine = "Statut" Then
            Range( _
                CurrentArea.Cells(RowIndex, 1), _
                CurrentArea.Cells(RowIndex, 2) _
            ).Validation.Delete
            CurrentArea.Cells(RowIndex, LastColumnNumber).Validation.Delete
            Chantiers_Financements_Format_Add_ValidationDossier Range( _
                CurrentArea.Cells(RowIndex, 3), _
                CurrentArea.Cells(RowIndex, LastColumnNumber - 1) _
            )
        Else
            Range( _
                CurrentArea.Cells(RowIndex, 1), _
                CurrentArea.Cells(RowIndex, LastColumnNumber) _
            ).Validation.Delete
        End If
    Next RowIndex

    DefinirFormatConditionnelPourLesDossier SetOfRange, NBChantiers
End Sub

Public Sub Chantiers_Import(NewWorkbook As Workbook, Data As Data)
    Dim BaseCell As Range
    Dim BaseCellChantier As Range
    Dim BaseCellChantierReal As Range
    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim Chantiers() As Chantier
    Dim CurrentSheet As Worksheet
    Dim NBChantiers As Integer
    Dim NBSalaries As Integer

    importerInfos NewWorkbook, Data.Informations
    
    NBSalaries = GetNbSalaries(NewWorkbook)
    If NBSalaries > 0 Then
        Set CurrentSheet = NewWorkbook.Worksheets(Nom_Feuille_Personnel)
        If CurrentSheet Is Nothing Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Personnel)
        Else
            Set BaseCell = CurrentSheet.Range("A:A").Find(T_FirstName)
            If BaseCell Is Nothing Then
                MsgBox Replace(T_NotFoundFirstName, "%PageName%", Nom_Feuille_Personnel)
            Else
                On Error Resume Next
                Set ChantierSheet = NewWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
                Set ChantierSheetReal = NewWorkbook.Worksheets(Nom_Feuille_Budget_chantiers_realise)
                On Error GoTo 0
                NBChantiers = 0
                Set BaseCellChantier = Chantiers_BaseCell_Get(ChantierSheet)
                Set BaseCellChantierReal = Chantiers_BaseCell_Get(ChantierSheetReal)
                If Not (BaseCellChantier Is Nothing) Then
                    NBChantiers = GetNbChantiers(NewWorkbook)
                End If

                Chantiers_And_Personal_Import_Salaries Data, BaseCell, BaseCellChantier, BaseCellChantierReal, NBSalaries, NBChantiers
                
                If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) And UBound(Data.Chantiers) > 1 Then
                    ' depenses
                    Chantiers_Import_Title_And_Depenses Data, ChantierSheet, ChantierSheetReal, BaseCellChantier, NBSalaries, NBChantiers
                    
                    ' Financements
                    Chantiers_Import_Financements NewWorkbook, Data, ChantierSheet, ChantierSheetReal, NBChantiers
                    
                    ' Autofinancement
                    Chantiers_Import_Autofinancements Data, ChantierSheet, ChantierSheetReal
                End If
            End If
        End If
    End If
    
    ' Ajouter Charges
    Charges_Import NewWorkbook, Data

End Sub

Public Sub Chantiers_Import_Autofinancements( _
        Data As Data, _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet _
    )
    
    Dim Chantiers() As Chantier
    Dim IndexChantier As Integer
    Dim SetOfRange As SetOfRange
    Dim TmpChantier As Chantier

    Chantiers = Data.Chantiers

    Application.Calculate
    SetOfRange = Chantiers_Financements_BaseCell_Get(ChantierSheet, ChantierSheetReal)
    If SetOfRange.Status Then
        For IndexChantier = 1 To UBound(Chantiers)
            TmpChantier = Chantiers(IndexChantier)
            ' does not set AutoFinancementStructure because calculated !
            Common_SetFormula _
                SetOfRange.ResultCell.Cells(2, 1 + IndexChantier), _
                TmpChantier.AutoFinancementAutres, _
                TmpChantier.AutoFinancementAutresFormula
            Common_SetFormula _
                SetOfRange.ResultCell.Cells(10, 1 + IndexChantier), _
                TmpChantier.AutoFinancementStructureAnneesPrecedentes, _
                TmpChantier.AutoFinancementStructureAnneesPrecedentesFormula
            Common_SetFormula _
                SetOfRange.ResultCell.Cells(9, 1 + IndexChantier), _
                TmpChantier.AutoFinancementAutresAnneesPrecedentes, _
                TmpChantier.AutoFinancementAutresAnneesPrecedentesFormula
            Common_SetFormula _
                SetOfRange.ResultCell.Cells(11, 1 + IndexChantier), _
                TmpChantier.CAanneesPrecedentes, _
                TmpChantier.CAanneesPrecedentesFormula
        Next IndexChantier
    End If
End Sub

Public Sub Chantiers_Import_Depenses_ChangeNB(BaseCell As Range, NBSalaries As Integer, NewNBDepenses As Integer, NBChantiers As Integer)
    Dim PreviousNBDepenses As Integer
    PreviousNBDepenses = Range(BaseCell, Common_FindNextNotEmpty(BaseCell, True).Cells(0, 1)).Rows.Count
                    
    If PreviousNBDepenses > NewNBDepenses Then
        ' Remove Lines
        Range(BaseCell.Cells(NewNBDepenses + 1, 1).EntireRow, BaseCell.Cells(PreviousNBDepenses, 1).EntireRow).Delete _
            Shift:=xlShiftUp
    Else
        If PreviousNBDepenses < NewNBDepenses Then
            ' Insert Lines
            BaseCell.Cells(1, 1).Worksheet.Activate
            BaseCell.Cells(PreviousNBDepenses - 1, 1).EntireRow.Select
            BaseCell.Cells(PreviousNBDepenses - 1, 1).EntireRow.Copy
            Range(BaseCell.Cells(PreviousNBDepenses - 1, 1).EntireRow, BaseCell.Cells(NewNBDepenses - 1, 1).EntireRow).Insert _
                Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Copy Format
            BaseCell.EntireRow.Copy
            Range(BaseCell.EntireRow, BaseCell.Cells(NewNBDepenses, 1).EntireRow).PasteSpecial _
                Paste:=xlPasteFormats
                
            ' Copy formula for total
            BaseCell.Cells(1, 1 + NBChantiers + 1).Copy
            Range(BaseCell.Cells(PreviousNBDepenses - 1, 1 + NBChantiers + 1), _
                BaseCell.Cells(NewNBDepenses, 1 + NBChantiers + 1)).PasteSpecial _
                Paste:=xlPasteFormulas
            
        End If
    End If

End Sub

Public Sub Chantiers_Import_Financements( _
        NewWorkbook As Workbook, _
        Data As Data, _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet, _
        NBChantiers As Integer _
    )
    
    Dim Chantiers() As Chantier
    Dim FinancementCompletTmp As FinancementComplet
    Dim FinancementsTmp() As Financement
    Dim Financements() As Financement
    Dim Index As Integer
    Dim IndexChantier As Integer
    Dim SetOfRange As SetOfRange
    Dim TmpChantier As Chantier
    Dim TmpChantier1 As Chantier
    
    FinancementCompletTmp = getDefaultFinancementComplet()
    Chantiers = Data.Chantiers
    TmpChantier1 = Chantiers(1)

    Chantiers_Financements_Clear ChantierSheet, ChantierSheetReal, NBChantiers
    Financements = TmpChantier1.Financements
    If UBound(Chantiers) > 0 And UBound(Financements) > 0 Then
        ReDim FinancementsTmp(1 To UBound(Chantiers))
        SetOfRange = Chantiers_Financements_Add_Prepare(NewWorkbook, NBChantiers, False)
        If SetOfRange.Status Then
            For Index = 1 To UBound(Financements)
                For IndexChantier = 1 To UBound(Chantiers)
                    TmpChantier = Chantiers(IndexChantier)
                    Financements = TmpChantier.Financements
                    FinancementsTmp(IndexChantier) = Financements(Index)
                Next IndexChantier
                FinancementCompletTmp.Financements = FinancementsTmp
                FinancementCompletTmp.Status = True
                SetOfRange = Chantiers_Financements_Add_Internal(SetOfRange, NewWorkbook, NBChantiers, FinancementCompletTmp, "", 0)
            Next Index
            
            Chantiers_Totals_UpdateSumsAndFormat NewWorkbook, SetOfRange, NBChantiers
        End If
    End If
End Sub

Public Sub Chantiers_Import_Title_And_Depenses( _
        Data As Data, _
        ChantierSheet As Worksheet, _
        ChantierSheetReal As Worksheet, _
        BaseCellChantier As Range, _
        NBSalaries As Integer, _
        NBChantiers As Integer _
    )
    
    Dim BaseCell As Range
    Dim Chantiers() As Chantier
    Dim DepenseTmp As DepenseChantier
    Dim DepensesTmp() As DepenseChantier
    Dim Index As Integer
    Dim IndexChantier As Integer
    Dim SetOfRange As SetOfRange
    Dim StrAddress As String
    Dim TmpChantier As Chantier
    Dim TmpChantier1 As Chantier

    
    Chantiers = Data.Chantiers

    ' nom des depenses
    SetOfRange = Chantiers_Depenses_SetOfRange_Get(ChantierSheet, ChantierSheetReal)
    If SetOfRange.Status Then
        Set BaseCell = SetOfRange.HeadCell.Cells(2, 2)
    Else
        ' backup
        Set BaseCell = BaseCellChantier.Cells(7 + 2 * NBSalaries, 1).EntireRow.Cells(1, 2)
    End If
    TmpChantier = Chantiers(1)
    TmpChantier1 = Chantiers(1)
    DepensesTmp = TmpChantier1.Depenses

    Chantiers_Import_Depenses_ChangeNB BaseCell, NBSalaries, UBound(TmpChantier.Depenses), NBChantiers
    If SetOfRange.StatusReal Then
        Chantiers_Import_Depenses_ChangeNB SetOfRange.HeadCellReal.Cells(2, 2), NBSalaries, UBound(TmpChantier.Depenses), NBChantiers
    End If
    
    For Index = 1 To UBound(TmpChantier.Depenses)
        DepenseTmp = DepensesTmp(Index)
        If DepenseTmp.Nom = "0" Then
            BaseCell.Cells(Index, 1).Value = ""
        Else
            BaseCell.Cells(Index, 1).Value = DepenseTmp.Nom
        End If
        If SetOfRange.StatusReal Then
            StrAddress = CleanAddress(BaseCell.Cells(Index, 1).address(False, False, xlA1, True))
            SetOfRange.HeadCellReal.Cells(1 + Index, 2).FormulaLocal = _
                Replace("=SI(%ADR%="""";"""";%ADR%)", "%ADR%", StrAddress)
        End If
    Next Index
    
    For IndexChantier = 1 To WorksheetFunction.Min(NBChantiers, UBound(Chantiers))
        TmpChantier = Chantiers(IndexChantier)
        If (TmpChantier.Nom = "0") Or (TmpChantier.Nom = "") Then
            BaseCellChantier.Cells(2, IndexChantier).Value = "xx"
        Else
            BaseCellChantier.Cells(2, IndexChantier).Value = TmpChantier.Nom
        End If
        
        DepensesTmp = TmpChantier.Depenses
        For Index = 1 To UBound(DepensesTmp)
            DepenseTmp = DepensesTmp(Index)
            Common_SetFormula _
                BaseCell.Cells(Index, 1 + IndexChantier), _
                DepenseTmp.Valeur, _
                DepenseTmp.Formula, _
                True
            If SetOfRange.StatusReal Then
                StrAddress = CleanAddress(BaseCell.Cells(Index, 1 + IndexChantier).address(False, False, xlA1, True))
                SetOfRange.HeadCellReal.Cells(1 + Index, 3 * IndexChantier).FormulaLocal = _
                    Replace("=SI(%ADR%="""";0;%ADR%)", "%ADR%", StrAddress)
                    
                Common_SetFormula _
                    SetOfRange.HeadCellReal.Cells(1 + Index, 3 * IndexChantier + 1), _
                    DepenseTmp.ValeurReal, _
                    DepenseTmp.FormulaReal, _
                    True
            End If
        Next Index
    Next IndexChantier
End Sub

Public Function Chantiers_Names_Extract( _
        BaseCellChantier As Range, _
        NBChantiers As Integer, _
        Data As Data _
        ) As SetOfChantiers

    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim IndexChantiers As Integer
    Dim SetOfChantiers As SetOfChantiers

    Chantiers = Data.Chantiers
    SetOfChantiers.Chantiers = Chantiers

    For IndexChantiers = 1 To NBChantiers
        ChantierTmp = Chantiers(IndexChantiers)
        ChantierTmp.Nom = BaseCellChantier.Cells(2, IndexChantiers).Value
        Chantiers(IndexChantiers) = ChantierTmp
    Next IndexChantiers

    SetOfChantiers.Chantiers = Chantiers
    Chantiers_Names_Extract = SetOfChantiers
End Function

Public Sub Chantiers_Salaries_ChangeNB( _
        wb As Workbook, _
        PreviousNB As Integer, _
        FinalNB As Integer, _
        Optional IsRealSheet As Boolean = False _
    )
    Dim BaseCell As Range
    Dim CurrentChantier As Integer
    Dim CurrentSheet As Worksheet
    Dim CurrentSheetName As String
    Dim NBChantiers As Integer
    Dim NBExtraColsInternal As Integer
    Dim RealFinalNB As Integer
    Dim TmpRange As Range
    
    NBChantiers = GetNbChantiers(wb)
    If IsRealSheet Then
        CurrentSheetName = Nom_Feuille_Budget_chantiers_realise
        NBExtraColsInternal = 3 * NBChantiers + 1 + NBExtraCols
    Else
        CurrentSheetName = Nom_Feuille_Budget_chantiers
        NBExtraColsInternal = NBExtraCols
    End If

    Set CurrentSheet = wb.Worksheets(CurrentSheetName)
    If CurrentSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", CurrentSheetName)
        Exit Sub
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(T_WorkingPeople)
    If BaseCell Is Nothing Then
        Exit Sub
    End If
    If BaseCell.Cells(0, 2).Value <> "Structure" Then
        Exit Sub
    End If
    
    If FinalNB <= 1 Then
        RealFinalNB = 2
    Else
        RealFinalNB = FinalNB
    End If

    If PreviousNB < RealFinalNB Then
        Common_InsertRows BaseCell, PreviousNB, RealFinalNB, False, NBExtraColsInternal, True
        Set TmpRange = Common_InsertRows(BaseCell.Cells(1 + RealFinalNB + 2, 1), PreviousNB, RealFinalNB, False, NBExtraColsInternal, False)

        Common_UpdateSumsByColumn _
            Range( _
                BaseCell.Cells(1 + RealFinalNB + 3, 3), _
                BaseCell.Cells(1 + 2 * RealFinalNB + 2, TmpRange.Columns.Count) _
            ), _
            BaseCell.Cells(0, 3), _
            PreviousNB
    Else
        If PreviousNB > RealFinalNB Then
            Common_RemoveRows BaseCell, PreviousNB, RealFinalNB, NBExtraColsInternal
            Common_RemoveRows BaseCell.Cells(1 + RealFinalNB + 2, 1), PreviousNB, RealFinalNB, NBExtraColsInternal
        End If
    End If
    If FinalNB <= 1 And PreviousNB > 1 Then
        If IsRealSheet Then
            For CurrentChantier = 1 To NBChantiers
                BaseCell.Cells(3, 1 + 3 * CurrentChantier).ClearContents
            Next CurrentChantier
        Else
            Range(BaseCell.Cells(3, 3), BaseCell.Cells(3, 2 + NBChantiers)).ClearContents
        End If
    End If
    
End Sub

Public Sub Chantiers_Totals_UpdateSumsAndFormat( _
    wb As Workbook, _
    SetOfRange As SetOfRange, _
    NBChantiers As Integer, _
    Optional SetFormat As Boolean = True _
    )

    If SetOfRange.StatusReal Then
        ' Be carefull if position of cells changes for real
        Chantiers_UpdateSumsReal wb, SetOfRange.ChantierSheetReal.Cells(1, 3)
    End If
    Chantiers_Financements_Totals_RenewFormula SetOfRange.ChantierSheet, SetOfRange.ChantierSheetReal, NBChantiers
    If SetFormat Then
        Chantiers_Format_Set SetOfRange.ChantierSheet, Nothing, NBChantiers
        Chantiers_Format_Set SetOfRange.ChantierSheet, SetOfRange.ChantierSheetReal, NBChantiers
    End If
End Sub

Public Sub Chantiers_UpdateSums(wb As Workbook, BaseCell As Range)

    Dim FirstCellAddress As String
    Dim FoundCell As Range
    Dim NBChantiers As Integer
    Dim ResultCell As Range
    Dim RowIndex As Integer
    
    NBChantiers = GetNbChantiers(wb)

    Set FoundCell = BaseCell.Cells(1, 0).EntireColumn.Find(Label_Total_Financements)
    If FoundCell Is Nothing Then
        Exit Sub
    End If
    For RowIndex = 3 To (FoundCell.Row + 7 - BaseCell.Row)
        FirstCellAddress = BaseCell.Cells(RowIndex, 1).address(False, False, xlA1, False)
        Set ResultCell = BaseCell.Cells(RowIndex, 1 + NBChantiers)
        If Left(ResultCell.Formula, Len(FirstCellAddress) + 5) = ("=SUM(" & FirstCellAddress) Then
            ResultCell.Formula = "=SUM(" _
                & CleanAddress(Range( _
                        BaseCell.Cells(RowIndex, 1), _
                        BaseCell.Cells(RowIndex, NBChantiers) _
                    ).address(False, False, xlA1, False)) _
                & ")"
        End If
    Next RowIndex

End Sub

Public Sub Chantiers_UpdateSumsReal(wb As Workbook, BaseCellReal As Range)

    Dim ChantierIndex As Integer
    Dim FirstCellAddress As String
    Dim FormulaForSum As String
    Dim FoundCell As Range
    Dim NBChantiers As Integer
    Dim ResultCell As Range
    Dim RowIndex As Integer
    
    NBChantiers = GetNbChantiers(wb)

    Set FoundCell = BaseCellReal.Cells(1, 0).EntireColumn.Find(Label_Total_Financements)
    If FoundCell Is Nothing Then
        Exit Sub
    End If
    For RowIndex = 3 To (FoundCell.Row + 7 - BaseCellReal.Row)
        FirstCellAddress = BaseCellReal.Cells(RowIndex, 2).address(False, False, xlA1, False)
        Set ResultCell = BaseCellReal.Cells(RowIndex, 1 + 3 * NBChantiers)
        If Left(ResultCell.Formula, Len(FirstCellAddress) + 1) = ("=" & FirstCellAddress) Then
            FormulaForSum = "="
            For ChantierIndex = 1 To NBChantiers
                If ChantierIndex > 1 Then
                    FormulaForSum = FormulaForSum & "+"
                End If
                FormulaForSum = FormulaForSum & _
                    BaseCellReal.Cells(RowIndex, 2 + 3 * (ChantierIndex - 1)).address(False, False, xlA1, False)
            Next ChantierIndex
            ResultCell.Formula = FormulaForSum
        End If
    Next RowIndex

End Sub

Public Sub Charges_Add_One_From_Button()

    Dim ChargesSheet As Worksheet
    Dim CodeIndex As Integer
    Dim ExtractedValue As Integer
    Dim FormattedValue As String
    Dim Offset As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfCellsCategories As SetOfCellsCategories
    Dim Value
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set ChargesSheet = wb.Worksheets(Nom_Feuille_Charges)
    If ChargesSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Charges)
        Exit Sub
    End If

    Value = InputBox("Quel nom de charge ?", "Ajouter une ligne de charge", "650 - Autre")
    FormattedValue = Trim(Value)

    If FormattedValue = "" Then
        MsgBox T_Error_Not_Empty_Charge_Name
        Exit Sub
    End If

    ExtractedValue = CInt(Left(FormattedValue, 2))
    If ExtractedValue < 60 Or ExtractedValue > 68 Then
        MsgBox T_Error_Bad_Format_For_Charge_Name
        Exit Sub
    End If

    CodeIndex = FindTypeChargeIndexFromCode(ExtractedValue)
    If CodeIndex > 0 Then
        SetOfCellsCategories = Charges_Categories_GetRows(ChargesSheet)
        SetOfRange = Charges_Categories_GetNext_Cell(SetOfCellsCategories, ExtractedValue)
        If SetOfRange.Status Then
            If SetOfRange.EndCell.Cells(0, 1) = Empty Then
                Offset = -1
            Else
                Offset = 0
            End If
            Charges_Add_One_Line SetOfRange.EndCell.Cells(Offset, 1), False, FormattedValue, 0, 0, 0, 0, 1
            Charges_UpdateFormula SetOfRange
        Else
            MsgBox T_Error_Not_Possible_To_Found_Type
        End If
    Else
        MsgBox T_Error_Not_Possible_To_Associate_Line_To_Type
    End If

End Sub

Public Sub Charges_Add_One_Line( _
        BaseCell As Range, _
        Optional NoBorderOnRightAndLeft As Boolean = True, _
        Optional Name As String = "", _
        Optional PreviousN2YearValue As Double = 0, _
        Optional PreviousYearValue As Double = 0, _
        Optional CurrentYearValue As Double = 0, _
        Optional CurrentRealizedYearValue As Double = 0, _
        Optional Category As Integer = 0 _
    )
    
    Dim CurrentCell As Range
    Set CurrentCell = Charges_Add_Line_Insert(BaseCell)
    ' Add value
    CurrentCell.Cells(1, 1).Value = Name
    CurrentCell.Cells(1, 2).Value = PreviousN2YearValue
    CurrentCell.Cells(1, 3).Value = PreviousYearValue
    CurrentCell.Cells(1, 4).Value = CurrentYearValue
    CurrentCell.Cells(1, 5).Value = CurrentRealizedYearValue
    CurrentCell.Cells(1, ColumnOfSecondPartInCharge - 1).Value = ""
    If Category = 0 Then
        CurrentCell.Cells(1, ColumnOfSecondPartInCharge).Value = ""
    Else
        CurrentCell.Cells(1, ColumnOfSecondPartInCharge).Value = Category
    End If
    formatChargeCell CurrentCell, NoBorderOnRightAndLeft
End Sub

Public Function Charges_Add_Line_Insert(CurrentCell As Range) As Range

    Dim Offset As Integer
    Dim ColumnOffset As Integer

    ' insert line
    CurrentCell.Worksheet.Activate
    CurrentCell.Cells(1, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 1).Select
    CurrentCell.Cells(1, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 1).Copy
    Range(CurrentCell.Cells(2, 1), CurrentCell.Cells(2, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 1)).Insert _
        Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Copy Format
    Range(CurrentCell.Cells(1, ColumnOfSecondPartInCharge), CurrentCell.Cells(1, ColumnOfSecondPartInCharge + NBCatOfCharges * 2)).Copy
    Range(CurrentCell.Cells(2, ColumnOfSecondPartInCharge), CurrentCell.Cells(2, ColumnOfSecondPartInCharge + NBCatOfCharges * 2)).PasteSpecial Paste:=xlPasteFormats
    ' Create formulae
    For Offset = 0 To 1
        CurrentCell.Cells(2, ColumnOfSecondPartInCharge + NBCatOfCharges * Offset + 1).FormulaLocal = "=SI(" _
            & "ET(" _
                & CurrentCell.Cells(2, ColumnOfSecondPartInCharge).address(False, False, xlA1, False) & "<>2;" _
                & CurrentCell.Cells(2, ColumnOfSecondPartInCharge).address(False, False, xlA1, False) & "<>3" _
            & ");" _
            & CurrentCell.Cells(2, 4 + Offset).address(False, False, xlA1, False) _
            & ";0" _
        & ")"
        For ColumnOffset = 2 To 3
        CurrentCell.Cells(2, ColumnOfSecondPartInCharge + NBCatOfCharges * Offset + ColumnOffset).FormulaLocal = "=SI(" _
                & CurrentCell.Cells(2, ColumnOfSecondPartInCharge).address(False, False, xlA1, False) & "=" _
                & ColumnOffset & ";" _
                & CurrentCell.Cells(2, 4 + Offset).address(False, False, xlA1, False) _
                & ";0" _
            & ")"
        Next ColumnOffset
    Next Offset
    ' Validation for first cell
    With CurrentCell.Cells(2, ColumnOfSecondPartInCharge).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=VAL_CAT"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Set Charges_Add_Line_Insert = CurrentCell.Cells(2, 1)
End Function

Public Function Charges_Add_One(SetOfCellsCategories As SetOfCellsCategories, Charge As Charge) As SetOfRange

    Dim CatIndex As Integer
    Dim CurrentCell As Range
    Dim Offset As Integer
    Dim SetOfRange As SetOfRange
    Dim TypeCharge As TypeCharge

    SetOfRange.Status = False

    If Charge.IndexTypeCharge > 0 Then
        TypeCharge = TypesDeCharges().Values(Charge.IndexTypeCharge)
        If TypeCharge.Index > 0 Then
            SetOfRange = Charges_Categories_GetNext_Cell(SetOfCellsCategories, TypeCharge.Index)
            If SetOfRange.Status Then
                If SetOfRange.EndCell.Cells(0, 1) = Empty Then
                    Offset = -1
                Else
                    Offset = 0
                End If
                Charges_Add_One_Line _
                    SetOfRange.EndCell.Cells(Offset, 1), _
                    False, _
                    Charge.Nom, _
                    Charge.PreviousN2YearValue, _
                    Charge.PreviousYearValue, _
                    Charge.CurrentYearValue, _
                    Charge.CurrentRealizedYearValue, _
                    Charge.Category
            End If
        End If
    End If

    Charges_Add_One = SetOfRange
End Function

Public Function Charges_Categories_GetNext_Cell(SetOfCellsCategories As SetOfCellsCategories, IndexOfCat As Integer) As SetOfRange

    Dim CurrentCell As Range
    Dim CurrentCells() As Range
    Dim HeadCell As Range
    Dim Index As Integer
    Dim NextCell As Range
    Dim SetOfRange As SetOfRange

    SetOfRange.Status = False
    SetOfRange.StatusReal = False

    CurrentCells = SetOfCellsCategories.Cells
    Set HeadCell = CurrentCells(IndexOfCat)
    If Not (HeadCell Is Nothing) Then
        Set SetOfRange.HeadCell = HeadCell
        Set SetOfRange.ResultCell = SetOfCellsCategories.TotalCell
        Set NextCell = SetOfCellsCategories.TotalCell
        For Index = 60 To 68
            If Index <> IndexOfCat Then
                Set CurrentCell = CurrentCells(Index)
                If Not (CurrentCell Is Nothing) Then
                    If CurrentCell.Row > HeadCell.Row And CurrentCell.Row < NextCell.Row Then
                        Set NextCell = CurrentCell
                    End If
                End If
            End If
        Next Index
        Set SetOfRange.EndCell = NextCell
        SetOfRange.Status = True
    End If

    Charges_Categories_GetNext_Cell = SetOfRange
End Function

Public Function Charges_Categories_GetRows(ChargesSheet As Worksheet) As SetOfCellsCategories
    
    Dim CurrentCell As Range
    Dim CurrentCellValue As String
    Dim IndexCode As Integer
    Dim StartCell As Range
    Dim SetOfCellsCategories As SetOfCellsCategories

    Set CurrentCell = ChargesSheet.Cells(2, 1)
    CurrentCellValue = CurrentCell.Value
    While (CurrentCellValue = "" Or CurrentCellValue = Empty) And CurrentCell.Row < 1000
        Set CurrentCell = CurrentCell.Cells(2, 1)
        CurrentCellValue = CurrentCell.Value
    Wend

    If CurrentCell.Row < 1000 Then
        ' on eline before to be able to scan 60
        Set StartCell = CurrentCell.Cells(0, 1)

        ' Total Cell
        While Left(CurrentCellValue, 5) <> "TOTAL" And CurrentCell.Row < 1000
            Set CurrentCell = CurrentCell.Cells(2, 1)
            CurrentCellValue = CurrentCell.Value
        Wend
        If Left(CurrentCellValue, 5) = "TOTAL" Then
            Set SetOfCellsCategories.TotalCell = CurrentCell

            ' cells by categories
            For IndexCode = 60 To 68
                Set CurrentCell = StartCell
                CurrentCellValue = CurrentCell.Value
                While (Left(CurrentCellValue, 2) <> IndexCode _
                    Or Mid(CurrentCellValue, 3, 2) <> " -") _
                    And CurrentCell.Row < SetOfCellsCategories.TotalCell.Row
                    Set CurrentCell = CurrentCell.Cells(2, 1)
                    CurrentCellValue = CurrentCell.Value
                Wend
                If Left(CurrentCellValue, 2) = IndexCode Then
                    Set SetOfCellsCategories.Cells(IndexCode) = CurrentCell
                End If

            Next IndexCode
        End If
    End If

    Charges_Categories_GetRows = SetOfCellsCategories
End Function

Public Function Charges_Extract(wb As Workbook, Data As Data, Revision As WbRevision) As Data
    Dim ChargesSheet As Worksheet
    Dim CurrentCell As Range
    Dim CurrentIndexTypeCharge As Integer
    Dim Charges() As Charge
    Dim TmpCharge As Charge
    Dim Index As Integer
    Dim PreviousIndex As Integer
    Dim NBNewCharges As Integer
    Dim Has3Years As Boolean
    Dim HasRealValues As Boolean
    Dim SetOfCharges As SetOfCharges
    Dim Titles() As String
    Dim TitlesBaseColumn As Integer
    Dim TitlesRow As Integer
    ReDim Charges(0)
    ReDim Titles(1 To 3)

    On Error Resume Next
    Set ChargesSheet = wb.Worksheets(Nom_Feuille_Charges)
    On Error GoTo 0
    If ChargesSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Charges)
        GoTo FinFunction
    End If
    
    HasRealValues = False
    If (Revision.Majeure = 2 And Revision.Mineure > 1) Or Revision.Majeure > 2 Then
        HasRealValues = True
    End If

    ' Get Titles for categories for charges
    If Revision.Majeure <= 1 Then
        Titles(1) = "Cat. 1"
        Titles(2) = "Cat. 2"
        Titles(3) = "Cat. 3"
    Else
        If HasRealValues Then
            TitlesBaseColumn = ColumnOfSecondPartInCharge
        Else
            TitlesBaseColumn = 6
        End If
        TitlesRow = 3
        Titles(1) = ChargesSheet.Cells(TitlesRow, TitlesBaseColumn + 1).Value
        Titles(2) = ChargesSheet.Cells(TitlesRow, TitlesBaseColumn + 2).Value
        Titles(3) = ChargesSheet.Cells(TitlesRow, TitlesBaseColumn + 3).Value
    End If
    Data.TitlesForChargesCat = Titles
    
    Set CurrentCell = ChargesSheet.Cells(2, 1)
    While (CurrentCell.Value = "" Or CurrentCell.Value = Empty) And CurrentCell.Row < 1000
        Set CurrentCell = CurrentCell.Cells(2, 1)
    Wend
    
    CurrentIndexTypeCharge = FindTypeChargeIndex(CurrentCell.Value)
    
    If (Revision.Majeure = 1 And Revision.Mineure > 9) Or Revision.Majeure > 1 Then
        Has3Years = True
    Else
        Has3Years = False
    End If
    
    While CurrentIndexTypeCharge > 0
        ' Find NB new charges
        Index = 2
        While CurrentCell.Cells(Index, 1).Value <> "" And FindTypeChargeIndex(CurrentCell.Cells(Index, 1).Value) = 0
            Index = Index + 1
        Wend
        NBNewCharges = Index - 2
        If NBNewCharges > 0 Then
            PreviousIndex = UBound(Charges)
            If PreviousIndex < 0 Then
                PreviousIndex = 0
            End If
            If PreviousIndex = 0 Then
                Charges = Common_getChargesDefault(NBNewCharges).Charges
            Else
                SetOfCharges.Charges = Charges
                Charges = Common_getChargesDefaultPreserve(SetOfCharges, PreviousIndex + NBNewCharges).Charges
            End If
            For Index = 1 To NBNewCharges
                TmpCharge = getDefaultCharge()
                TmpCharge.Nom = CurrentCell.Cells(1 + Index, 1).Value
                TmpCharge.IndexTypeCharge = CurrentIndexTypeCharge
                If Has3Years Then
                    TmpCharge.CurrentYearValue = CurrentCell.Cells(1 + Index, 4).Value
                    TmpCharge.PreviousYearValue = CurrentCell.Cells(1 + Index, 3).Value
                    TmpCharge.PreviousN2YearValue = CurrentCell.Cells(1 + Index, 2).Value
                Else
                    TmpCharge.CurrentYearValue = CurrentCell.Cells(1 + Index, 3).Value
                    TmpCharge.PreviousYearValue = CurrentCell.Cells(1 + Index, 2).Value
                    TmpCharge.PreviousN2YearValue = 0
                End If
                If HasRealValues Then
                    TmpCharge.CurrentRealizedYearValue = CurrentCell.Cells(1 + Index, 5).Value
                Else
                    TmpCharge.CurrentRealizedYearValue = 0
                End If
                If HasRealValues Then
                    If CurrentCell.Cells(1 + Index, ColumnOfSecondPartInCharge).Value > 0 _
                        And CurrentCell.Cells(1 + Index, ColumnOfSecondPartInCharge).Value < 4 Then
                        TmpCharge.Category = CInt(CurrentCell.Cells(1 + Index, ColumnOfSecondPartInCharge).Value)
                    Else
                        TmpCharge.Category = 1
                    End If
                Else
                    If Revision.Majeure > 1 _
                        And CurrentCell.Cells(1 + Index, 6).Value > 0 _
                        And CurrentCell.Cells(1 + Index, 6).Value < 4 Then
                        TmpCharge.Category = CInt(CurrentCell.Cells(1 + Index, 6).Value)
                    Else
                        TmpCharge.Category = 1
                    End If
                End If
                Set TmpCharge.ChargeCell = CurrentCell.Cells(1 + Index, 1)
                Charges(PreviousIndex + Index) = TmpCharge
            Next Index
        End If
        
        Index = 2 + NBNewCharges
        While CurrentCell.Cells(Index, 1).Value = ""
            Index = Index + 1
        Wend
        
        Set CurrentCell = CurrentCell.Cells(Index, 1)
        CurrentIndexTypeCharge = FindTypeChargeIndex(CurrentCell.Value)
    
    Wend
    
    Data.Charges = Charges
    
FinFunction:
    Charges_Extract = Data
End Function

Public Sub Charges_Import(wb As Workbook, Data As Data)
    Dim Charge As Charge
    Dim Charges() As Charge
    Dim ChargesSheet As Worksheet
    Dim Index As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfCellsCategories As SetOfCellsCategories
    Dim Titles() As String
    Dim TitlesRow As Integer

    On Error Resume Next
    Set ChargesSheet = wb.Worksheets(Nom_Feuille_Charges)
    On Error GoTo 0
    If ChargesSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Charges)
        Exit Sub
    End If

    ' update titles for categories
    TitlesRow = 3
    Titles = Data.TitlesForChargesCat
    For Index = LBound(Titles) To UBound(Titles)
        ChargesSheet.Cells(TitlesRow, ColumnOfSecondPartInCharge + Index).Value = Titles(Index)
    Next Index

    SetOfCellsCategories = Charges_Categories_GetRows(ChargesSheet)
    Charges_Import_CleanCategory SetOfCellsCategories
    Charges = Data.Charges

    ' add charges
    For Index = 1 To UBound(Charges)
        Charge = Charges(Index)
        Charges_Add_One SetOfCellsCategories, Charge
    Next Index

    ' add empty lines and update sums
    For Index = 60 To 68
        SetOfRange = Charges_Categories_GetNext_Cell(SetOfCellsCategories, Index)
        If SetOfRange.Status Then
            Charges_Add_One_Line SetOfRange.EndCell.Cells(0, 1)

            Charges_UpdateFormula SetOfRange
        End If
    Next Index
End Sub

Public Sub Charges_Import_CleanCategory(SetOfCellsCategories As SetOfCellsCategories)

    Dim IndexOfCat As Integer
    Dim SetOfRange As SetOfRange

    For IndexOfCat = 60 To 68
        SetOfRange = Charges_Categories_GetNext_Cell(SetOfCellsCategories, IndexOfCat)
        If SetOfRange.Status Then
            ' clean previous
            SetOfRange.HeadCell.Cells(1, 2).Value = 0
            SetOfRange.HeadCell.Cells(1, 3).Value = 0
            SetOfRange.HeadCell.Cells(1, 4).Value = 0
            SetOfRange.HeadCell.Cells(1, ColumnOfSecondPartInCharge).Value = ""
            SetOfRange.HeadCell.Cells(1, ColumnOfSecondPartInCharge + 1).Value = 0
            SetOfRange.HeadCell.Cells(1, ColumnOfSecondPartInCharge + 2).Value = 0
            SetOfRange.HeadCell.Cells(1, ColumnOfSecondPartInCharge + 3).Value = 0
            If (SetOfRange.HeadCell.Row + 1) < SetOfRange.EndCell.Row Then
                Range(SetOfRange.HeadCell.Cells(2, 1), SetOfRange.EndCell.Cells(0, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 5)).Delete Shift:=xlShiftUp
            End If
        End If
    Next IndexOfCat
End Sub

Public Sub Charges_Remove_One()
    
    Dim ChargesSheet As Worksheet
    Dim CurrentCell As Range
    Dim CurrentCells() As Range
    Dim Index As Integer
    Dim IsOK As Boolean
    Dim MaxRow As Integer
    Dim MinRow As Integer
    Dim NewLine As Integer
    Dim SetOfCellsCategories As SetOfCellsCategories
    Dim SetOfRange As SetOfRange
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set ChargesSheet = wb.Worksheets(Nom_Feuille_Charges)
    If ChargesSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Charges)
        Exit Sub
    End If
    SetOfCellsCategories = Charges_Categories_GetRows(ChargesSheet)
    CurrentCells = SetOfCellsCategories.Cells

    MinRow = 1
    MaxRow = SetOfCellsCategories.TotalCell.Row - 1
    For Index = 60 To 68
        Set CurrentCell = CurrentCells(Index)
        If Not (CurrentCell Is Nothing) Then
            If MinRow = 1 Or CurrentCell.Row < MinRow Then
                MinRow = CurrentCell.Row
            End If
        End If
    Next Index

    NewLine = Common_InputBox_Get_Line_Between( _
        Replace(T_Delete_Object_Of_Line, "%objectName%", T_Charge), _
        Replace(T_Line_To_Delete_For_Object, "%objectName%", T_Charge), _
        MinRow + 1, _
        MaxRow _
    )

    If NewLine = -1 Then
        ' Cancel button
        Exit Sub
    End If

    If NewLine = 0 Then
        IsOK = False
    Else
        IsOK = True
        For Index = 60 To 68
            Set CurrentCell = CurrentCells(Index)
            If Not (CurrentCell Is Nothing) Then
                If CurrentCell.Row = NewLine Then
                    IsOK = False
                End If
            End If
        Next Index
    End If
    
    If Not IsOK Then
        MsgBox Replace( _
            Replace(T_Given_Line_Is_Not_Line_Of_Object, "%objectName%", T_Charge), _
            "d'la", _
            "d'une" _
        )
        Exit Sub
    End If

    SetSilent

    Range(ChargesSheet.Cells(NewLine, 1), ChargesSheet.Cells(NewLine, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 5)).Delete Shift:=xlShiftUp
    
    ' update sums
    For Index = 60 To 68
        SetOfRange = Charges_Categories_GetNext_Cell(SetOfCellsCategories, Index)
        If SetOfRange.Status Then
            Charges_UpdateFormula SetOfRange
        End If
    Next Index

    SetActive
End Sub

Public Sub Charges_UpdateFormula(SetOfRange As SetOfRange)

    Dim ColumnIndex As Integer
    Dim RowIndex As Integer

    If SetOfRange.Status Then
        For ColumnIndex = 2 To 5
            If (SetOfRange.HeadCell.Row + 1) < SetOfRange.EndCell.Row Then
                SetOfRange.HeadCell.Cells(1, ColumnIndex).Formula = "=SUM(" _
                    & Range( _
                        SetOfRange.HeadCell.Cells(2, ColumnIndex), _
                        SetOfRange.EndCell.Cells(0, ColumnIndex) _
                    ).address(False, False, xlA1) & ")"
            Else
                SetOfRange.HeadCell.Cells(1, ColumnIndex).Value = 0
            End If
        Next ColumnIndex
        If (SetOfRange.HeadCell.Row + 1) < SetOfRange.EndCell.Row Then
            For RowIndex = 1 To (SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row + 1)
                SetOfRange.HeadCell.Cells(RowIndex, 6).Formula = "=" _
                    & SetOfRange.HeadCell.Cells(RowIndex, 5).address(False, False, xlA1) _
                    & "/(" _
                        & SetOfRange.HeadCell.Cells(RowIndex, 4).address(False, False, xlA1) _
                    & "+1E-9)"
            Next RowIndex
        End If
        SetOfRange.HeadCell.Cells(1, ColumnOfSecondPartInCharge).Value = ""
        
        For ColumnIndex = (ColumnOfSecondPartInCharge + 1) To (ColumnOfSecondPartInCharge + NBCatOfCharges * 2)
            SetOfRange.HeadCell.Cells(1, ColumnIndex).Value = 0
        Next ColumnIndex
    End If
End Sub

Public Sub CoutJSalaires_Salaries_ChangeNB(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim RealFinalNB As Integer
    
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CurrentSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Cout_J_Salaire)
        Exit Sub
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find(T_FirstName)
    If BaseCell Is Nothing Then
        Exit Sub
    End If
    If BaseCell.Cells(-1, 1).Value <> Label_Cout_J_Salaire_Part_A Then
        Exit Sub
    End If
    
    If FinalNB <= 1 Then
        RealFinalNB = 2
    Else
        RealFinalNB = FinalNB
    End If

    If PreviousNB < RealFinalNB Then
        Common_InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            Common_RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
    ' Part B
    Set BaseCell = Common_FindNextNotEmpty(BaseCell.Cells(1 + RealFinalNB + 1, 1), True)
    If BaseCell.Value <> Label_Cout_J_Salaire_Part_B Then
        Exit Sub
    End If
    If BaseCell.Cells(3, 1).Value <> T_FirstName Then
        Exit Sub
    End If
    Set BaseCell = BaseCell.Cells(3, 1)
    
    If PreviousNB < RealFinalNB Then
        Common_InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            Common_RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
    ' Part D
    Set BaseCell = CurrentSheet.Range("A:A").Find("TOTAL")
    If BaseCell Is Nothing Then
        Exit Sub
    End If
    If BaseCell.Cells(5, 1).Value <> T_FirstName Then
        Exit Sub
    End If
    Set BaseCell = BaseCell.Cells(5, 1)

    If PreviousNB < RealFinalNB Then
        Common_InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            Common_RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
End Sub

Public Function GetNbSalaries(wb As Workbook)

    Dim CoutJSalaireSheet As Worksheet
    Dim BaseCell As Range
    Dim TmpRange As Range
    
    Set CoutJSalaireSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CoutJSalaireSheet Is Nothing Then
        GetNbSalaries = -1
        Exit Function
    End If
    Set BaseCell = CoutJSalaireSheet.Range("A:A").Find(T_FirstName)
    If BaseCell Is Nothing Then
        GetNbSalaries = -2
        Exit Function
    End If
    If BaseCell.Cells(-1, 1).Value <> Label_Cout_J_Salaire_Part_A Then
        GetNbSalaries = -3
        Exit Function
    End If
    ' TODO find dynamically the right row
    If BaseCell.Value <> T_FirstName Then
        GetNbSalaries = -4
        Exit Function
    End If
    If (BaseCell.Cells(2, 1).Formula <> "") And (BaseCell.Cells(3, 1).Formula = "") Then
        GetNbSalaries = -5
        Exit Function
    End If
    
    Set TmpRange = Common_FindNextNotEmpty(BaseCell.Cells(2, 1), True)
    If TmpRange.Value = T_FirstName Or TmpRange.Value = Label_Cout_J_Salaire_Part_B Then
        GetNbSalaries = -6
        Exit Function
    End If
    GetNbSalaries = TmpRange.Row - BaseCell.Row
    
End Function

Public Function GetNbSalariesV0(wb As Workbook) As NBAndRange

    Dim Result As NBAndRange
    Result = getDefaulNBAndRange()
    Dim CoutJSalaireSheet As Worksheet
    Dim BaseCell As Range
    Dim TmpRange As Range
    
    Set CoutJSalaireSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CoutJSalaireSheet Is Nothing Then
        Result.NB = -1
        GoTo FinFunction
    End If
    Set BaseCell = CoutJSalaireSheet.Range("A:A").Find("TOTAL Structure ")
    If BaseCell Is Nothing Then
        Result.NB = -2
        GoTo FinFunction
    End If
    Set TmpRange = BaseCell
    Set BaseCell = Common_FindNextNotEmpty(Common_FindNextNotEmpty(CoutJSalaireSheet.Cells(1, 1), True), True)
    If BaseCell.Value <> Label_Cout_J_Salaire_Part_A Then
        Result.NB = -3
        GoTo FinFunction
    End If
    Set BaseCell = BaseCell.Cells(3, 1)
    If BaseCell.Cells(1, 2).Value <> "Nb de jours de travail annuel" Then
        Result.NB = -4
        GoTo FinFunction
    End If
    
    Result.NB = TmpRange.Row - BaseCell.Row - 1
    Set Result.Range = BaseCell
    
FinFunction:
    GetNbSalariesV0 = Result
    
End Function

Public Function GetNbChantiers(wb As Workbook, Optional BaseRow As Integer = 3)

    Dim ChantierSheet As Worksheet
    Dim BaseCell As Range
    Dim Counter As Integer
    
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        GetNbChantiers = -1
        Exit Function
    End If
    Set BaseCell = Common_FindNextNotEmpty(ChantierSheet.Cells(BaseRow, 1), False)
    If BaseCell.Column > 1000 Then
        GetNbChantiers = -2
        Exit Function
    End If
    If Left(BaseCell.Value, Len("Chantier")) <> "Chantier" Then
        GetNbChantiers = -3
        Exit Function
    End If
    Counter = 1
    While Left(BaseCell.Cells(1, Counter).Value, Len("Chantier")) = "Chantier"
        Counter = Counter + 1
    Wend
    
    GetNbChantiers = Counter - 1
    
End Function

Attribute VB_Name = "Process"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

Public Sub CleanDepenses(BaseCell)
    Dim Anchor As String

    Anchor = "Total "

    ' remove others lines and leave one formatted
    While Left(BaseCell.Cells(2, 1).Value, Len(Anchor)) <> Anchor
        Range(BaseCell.Cells(2, 1), BaseCell.Cells(2, 3)).Delete Shift:=xlShiftUp
    Wend
End Sub

Public Function AddHeader(BaseCell As Range, CodeValue As Integer, CodeIndex As Integer) As Range
    Dim CurrentCell As Range
    Dim NomTypeCharge As TypeCharge

    Set CurrentCell = InsertLineAndFormat(BaseCell, BaseCell, True)
    CurrentCell.Value = CodeValue

    NomTypeCharge = TypesDeCharges().Values(CodeIndex)
    CurrentCell.Cells(1, 2).Value = NomTypeCharge.Nom
    CurrentCell.Cells(1, 3).Value = 0

    Set AddHeader = CurrentCell
End Function

Public Function AddDepensesDepuisCharges( _
        Data As Data, _
        BaseCell As Range, _
        HeadCell As Range, _
        IndexFound As Integer _
    )

    Dim Charges() As Charge
    Dim currentCharge As Charge
    Dim CurrentCell As Range
    Dim Index As Integer

    Set CurrentCell = BaseCell

    Charges = Data.Charges
    For Index = 1 To UBound(Charges)
        currentCharge = Charges(Index)
        
        If currentCharge.IndexTypeCharge = IndexFound Then
            
            Set CurrentCell = InsertLineAndFormat(CurrentCell, HeadCell, False)
            CurrentCell.Value = ""
            CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(currentCharge.ChargeCell.address(False, False, xlA1, True))
            ' Be carefull to the number of columns if a 'charges' coles is added
            CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(currentCharge.ChargeCell.Cells(1, 4).address(False, False, xlA1, True))
        End If
    Next Index
    Set AddDepensesDepuisCharges = CurrentCell
End Function

Public Function AddDepensesDepuisChantiers( _
        Data As Data, _
        BaseCell As Range, _
        HeadCell As Range, _
        CodeValue As Integer _
    )

    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim CurrentCell As Range
    Dim Depenses() As DepenseChantier
    Dim Depense As DepenseChantier
    Dim Index As Integer
    Dim NBChantiers As Integer

    Set CurrentCell = BaseCell

    Chantiers = Data.Chantiers
    Chantier = Chantiers(1)
    Depenses = Chantier.Depenses
    NBChantiers = UBound(Chantiers)

    For Index = 1 To UBound(Depenses)
        Depense = Depenses(Index)
        If Left(Depense.Nom, 2) = CStr(CodeValue) Then
            Set CurrentCell = InsertLineAndFormat(CurrentCell, HeadCell, False)
            CurrentCell.Value = ""
            CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 0).address(False, False, xlA1, True))
            CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True))
        End If
    Next Index
    Set AddDepensesDepuisChantiers = CurrentCell
End Function

Public Function AddDepenses(wb As Workbook, Data As Data, BaseCell As Range) As Range
    Dim CodeValue As Integer
    Dim CodeIndex As Integer
    Dim CurrentCell As Range
    Dim HeadCell As Range
    Dim StartCell As Range
    Dim TotalCell As Range

    Set TotalCell = BaseCell.Cells(2, 1)
    TotalCell.Cells(1, 3).Formula = "=0"

    Set CurrentCell = BaseCell

    For CodeValue = 60 To 69
        CodeIndex = FindTypeChargeIndexFromCode(CodeValue)
        If CodeIndex > 0 Then
            Set HeadCell = AddHeader(CurrentCell, CodeValue, CodeIndex)
            TotalCell.Cells(1, 3).Formula = TotalCell.Cells(1, 3).Formula _
                & "+" _
                & CleanAddress(HeadCell.Cells(1, 3).address(False, False, xlA1))
            Set CurrentCell = AddDepensesDepuisCharges(Data, HeadCell, HeadCell, CodeIndex)
            Set CurrentCell = AddDepensesDepuisChantiers(Data, CurrentCell, HeadCell, CodeValue)

            If CodeValue = 64 Then
                ' ajouter les d?penses de personnel
                Set CurrentCell = InsertLineAndFormat(CurrentCell, HeadCell, False)
                CurrentCell.Value = ""
                CurrentCell.Cells(1, 2).Value = "Rémunération des personnels"
                CurrentCell.Cells(1, 2).Font.Bold = True
                CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(SearchRangeForEmployeesSalary(wb).address(False, False, xlA1, True)) & "/1.5"
                Set CurrentCell = InsertLineAndFormat(CurrentCell, HeadCell, False)
                CurrentCell.Value = ""
                CurrentCell.Cells(1, 2).Value = "Charges sociales"
                CurrentCell.Cells(1, 2).Font.Bold = True
                CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(SearchRangeForEmployeesSalary(wb).address(False, False, xlA1, True)) & "-" & CleanAddress(CurrentCell.Cells(0, 3).address(False, False, xlA1, False))
            End If

            ' set sum
            If CurrentCell.Row > HeadCell.Row Then
                HeadCell.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCell.Cells(2, 3), CurrentCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
            End If
        End If
    Next CodeValue

    Set AddDepenses = CurrentCell
End Function

Public Function AddFinancements(wb As Workbook, Data As Data, StartCell As Range) As Boolean

    Dim BaseCell As Range
    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim Financement As Financement
    Dim HeadCell As Range
    Dim HeadCellFinancement As Range
    Dim Index As Integer
    Dim IndexTypeFinancement As Integer
    Dim NBChantiers As Integer
    Dim TypesFinancements() As String

    Set BaseCell = StartCell
    Set HeadCell = StartCell
    HeadCell.Cells(1, 3).Value = 0
    Chantiers = Data.Chantiers
    NBChantiers = UBound(Chantiers)
    Chantier = Chantiers(1)
    For Index = 1 To UBound(Chantier.Financements)
        Financement = Chantier.Financements(Index)
        If Financement.TypeFinancement = 0 Then
            Set BaseCell = InsertLineAndFormat(BaseCell, HeadCell, False)
            BaseCell.Cells(1, 2).Formula = "=" & CleanAddress(Financement.BaseCell.Cells(1, 0).address(False, False, xlA1, True))
            BaseCell.Cells(1, 3).Formula = "=" & CleanAddress(Financement.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True))
        End If
    Next Index
    
    ' remove others lines and leave one formatted
    While BaseCell.Cells(2, 1).Value = ""
        Range(BaseCell.Cells(2, 1), BaseCell.Cells(2, 3)).Delete Shift:=xlShiftUp
    Wend
    
    If BaseCell.Row > HeadCell.Row Then
        HeadCell.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCell.Cells(2, 3), BaseCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
    End If
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
    Next Index
    
    Set BaseCell = BaseCell.Cells(2, 1)
    
    If BaseCell.Value <> 74 Then
        AddFinancements = False
        Exit Function
    End If
    Set HeadCell = BaseCell
    HeadCell.Cells(1, 3).Formula = "=0"
    
    TypesFinancements = TypeFinancementsFromWb(wb)
    
    For IndexTypeFinancement = 1 To UBound(TypesFinancements)
        Set BaseCell = InsertLineAndFormat(BaseCell, HeadCell, False)
        BaseCell.Cells(1, 2).Value = TypesFinancements(IndexTypeFinancement)
        BaseCell.Cells(1, 3).Value = 0
        
        FormatFinancementCells BaseCell
        HeadCell.Cells(1, 3).Formula = HeadCell.Cells(1, 3).Formula & "+" & CleanAddress(BaseCell.Cells(1, 3).address(False, False, xlA1))
        Set HeadCellFinancement = BaseCell
        Chantiers = Data.Chantiers
        NBChantiers = UBound(Chantiers)
        Chantier = Chantiers(1)
        For Index = 1 To UBound(Chantier.Financements)
            Financement = Chantier.Financements(Index)
            If Financement.TypeFinancement = IndexTypeFinancement Then
                Set BaseCell = InsertLineAndFormat(BaseCell, HeadCellFinancement, False)
                BaseCell.Cells(1, 2).Formula = "=" & CleanAddress(Financement.BaseCell.Cells(1, 0).address(False, False, xlA1, True))
                BaseCell.Cells(1, 3).Formula = "=" & CleanAddress(Financement.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True))
            End If
        Next Index
        If BaseCell.Row > HeadCellFinancement.Row Then
            HeadCellFinancement.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCellFinancement.Cells(2, 3), BaseCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
        End If
    Next IndexTypeFinancement
    
    ' remove others lines and leave one formatted
    While BaseCell.Cells(2, 1).Value = ""
        Range(BaseCell.Cells(2, 1), BaseCell.Cells(2, 3)).Delete Shift:=xlShiftUp
    Wend
    
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
    Next Index
    AddFinancements = True
End Function

' Macro pour mettre à jour le budget update
Public Sub MettreAJourBudgetGlobal(wb As Workbook)

    Dim Data As Data
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim rev As WbRevision
        
    SetSilent
    
    rev = DetecteVersion(wb)
    Data = extraireDonneesVersion1(wb, rev)
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Budget_global)
    If CurrentSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Budget_global & "' n'a pas été trouvée"
        GoTo EndSub
    End If
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Budget_chantiers & "' n'a pas été trouvée"
        GoTo EndSub
    End If
    
    Set BaseCell = CurrentSheet.Cells(1, 1).EntireColumn.Find("Compte")
    If BaseCell Is Nothing Then
        GoTo EndSub
    End If
    Set BaseCell = BaseCell.Cells(2, 1)
    While BaseCell.Value = "" Or Len(BaseCell.Value) = 0 Or CInt(BaseCell.Value) < 60 Or CInt(BaseCell.Value) > 69
        Set BaseCell = BaseCell.Cells(2, 1)
    Wend
    
    Set BaseCell = BaseCell.Cells(0, 1)
    CleanDepenses BaseCell
    AddDepenses wb, Data, BaseCell
    
    ' Produits
    Set BaseCell = CurrentSheet.Cells(1, 1).EntireColumn.Find("Compte")
    If BaseCell Is Nothing Then
        GoTo EndSub
    End If
    Set BaseCell = BaseCell.Cells(1, 5)
    While BaseCell.Value = "" Or BaseCell.Value <> 70
        Set BaseCell = BaseCell.Cells(2, 1)
    Wend

    If Not AddFinancements(wb, Data, BaseCell) Then
        GoTo EndSub
    End If
    
    ' Egaliser la longueur des colonnes
    EgaliserLesColonnes CurrentSheet
    
EndSub:
    Application.DisplayAlerts = True
    SetActive
    BaseCell.EntireRow.Cells(1, 1).EntireColumn.Cells(1, 1).Select

End Sub

Public Function InsertLineAndFormat(BaseCell As Range, HeadCell As Range, IsHeader As Boolean) As Range

    Dim Index As Integer

    If (Not IsHeader) And BaseCell.Cells(2, 1).Value = "" Then
        Set BaseCell = BaseCell.Cells(2, 1)
    Else
        ' insert line
        BaseCell.Worksheet.Activate
        BaseCell.Select
        BaseCell.Copy
        Range(BaseCell.Cells(2, 1), BaseCell.Cells(2, 3)).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        For Index = 1 To 3
            BaseCell.Cells(2, Index).Value = ""
        Next Index
        
        Set BaseCell = BaseCell.Cells(2, 1)
        
    End If
    ' Format cell
    SetFormatForBudget BaseCell, HeadCell, IsHeader
    
    Set InsertLineAndFormat = BaseCell
End Function

Public Function SearchRangeForEmployeesSalary(wb As Workbook) As Range
    Dim CoutJSalaireSheet As Worksheet
    Dim BaseCell As Range
    
    Set BaseCell = Nothing
    
    Set CoutJSalaireSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CoutJSalaireSheet Is Nothing Then
        GoTo EndFunction
    End If
    
    Set BaseCell = CoutJSalaireSheet.Cells.Find("Masse salariale des " & Chr(10) & "opérateurs : ")
    If BaseCell Is Nothing Then
        GoTo EndFunction
    End If
    Set BaseCell = BaseCell.Cells(1, 2)
    
    
EndFunction:
    Set SearchRangeForEmployeesSalary = BaseCell
End Function
Public Sub EgaliserLesColonnes(ws As Worksheet)

    Dim EndFirstCol As Range
    Dim EndSecondCol As Range
    Dim Ecart As Integer
    Dim Index As Integer
    Dim BaseCell As Range
    Dim HeadCell
    
    Set EndFirstCol = ws.Cells.Find("Total Dépenses (1) + (2)")
    Set EndSecondCol = ws.Cells.Find("Total Financements (1) + (2)+ (3)")
    Ecart = EndFirstCol.Row - EndSecondCol.Row
    
    If Ecart > 0 Then
        Set BaseCell = ws.Cells(1, 5).EntireColumn.Find(75).Cells(0, 1)
    Else
        Set BaseCell = ws.Cells.Find("Total Dépenses (1)").Cells(0, 1)
        Ecart = -Ecart
    End If
    
    For Index = 1 To Ecart
        Set BaseCell = InsertLineAndFormat(BaseCell, BaseCell.Cells(-1, 1), False)
    Next Index
    
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
    Next Index
    
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
    Set BaseCell = CoutJSalaireSheet.Range("A:A").Find("Prénom")
    If BaseCell Is Nothing Then
        GetNbSalaries = -2
        Exit Function
    End If
    If BaseCell.Cells(-1, 1).Value <> Label_Cout_J_Salaire_Part_A Then
        GetNbSalaries = -3
        Exit Function
    End If
    ' TODO find dynamically the right row
    If BaseCell.Value <> "Prénom" Then
        GetNbSalaries = -4
        Exit Function
    End If
    If (BaseCell.Cells(2, 1).Formula <> "") And (BaseCell.Cells(3, 1).Formula = "") Then
        GetNbSalaries = -5
        Exit Function
    End If
    
    Set TmpRange = FindNextNotEmpty(BaseCell.Cells(2, 1), True)
    If TmpRange.Value = "Prénom" Or TmpRange.Value = Label_Cout_J_Salaire_Part_B Then
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
    Set BaseCell = FindNextNotEmpty(FindNextNotEmpty(CoutJSalaireSheet.Cells(1, 1), True), True)
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
    Set BaseCell = FindNextNotEmpty(ChantierSheet.Cells(BaseRow, 1), False)
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

Public Sub ChangeSalaries(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)

    If FinalNB < 1 Then
        Exit Sub
    End If
    
    ChangeNBSalarieDansPersonnel wb, PreviousNB, FinalNB
    ChangerNBSalariesDansCoutJSalaires wb, PreviousNB, FinalNB
    ChangeNBSalariesDansChantier wb, PreviousNB, FinalNB

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
    Set BaseCell = FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
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
        Range(BaseCell.Cells(1, PreviousNB).EntireColumn, BaseCell.Cells(1, FinalNB - 1).EntireColumn).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        BaseCell.Cells(1, FinalNB).EntireColumn.Copy
        Range(BaseCell.Cells(1, PreviousNB).EntireColumn, BaseCell.Cells(1, FinalNB - 1).EntireColumn).PasteSpecial _
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
            SetOfRange = GetCellsForDepense(ChantierSheet)
            If SetOfRange.Status Then
                Range( _
                SetOfRange.HeadCell.Cells(2, PreviousNB + 3), _
                SetOfRange.ResultCell.Cells(0, FinalNB + 1) _
                ).ClearContents
            Else
                ' Backup
                Set StartRange = EndRange.Cells(3 + NBSalaries, PreviousNB + 1)
                Set EndRange = FindNextNotEmpty(StartRange.EntireRow.Cells(1, 2), True).EntireRow.Cells(0, BaseCell.Cells(1, FinalNB).Column)
                Range(StartRange, EndRange).ClearContents
            End If
        End If
    Else
        If FinalNB < PreviousNB Then
            Range(BaseCell.Cells(1, FinalNB + 1).EntireColumn, BaseCell.Cells(1, PreviousNB).EntireColumn).Delete Shift:=xlToLeft
        End If
    End If
    

End Sub

Public Sub ChangeUnChantier(Delta As Integer)

    Dim CurrentNBChantier As Integer
    Dim CurrentWs As Worksheet
    Dim NBToRemove As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    SetSilent
    
    ' Current NB
    CurrentNBChantier = GetNbChantiers(wb)
    
    If Delta < 0 And (CurrentNBChantier + Delta) < 1 Then
        GoTo FinSub
    End If
    
    ChangeChantiers wb, CurrentNBChantier, CurrentNBChantier + Delta
    
FinSub:

    Set CurrentWs = wb.ActiveSheet
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Cells(1, 1).Select
    Next 'Ws
    CurrentWs.Activate
    
    SetActive

End Sub

Public Sub ChangerNBSalariesDansCoutJSalaires(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim RealFinalNB As Integer
    
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CurrentSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Cout_J_Salaire & "' n'a pas été trouvée"
        Exit Sub
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find("Prénom")
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
        InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
    ' Part B
    Set BaseCell = FindNextNotEmpty(BaseCell.Cells(1 + RealFinalNB + 1, 1), True)
    If BaseCell.Value <> Label_Cout_J_Salaire_Part_B Then
        Exit Sub
    End If
    If BaseCell.Cells(3, 1).Value <> "Prénom" Then
        Exit Sub
    End If
    Set BaseCell = BaseCell.Cells(3, 1)
    
    If PreviousNB < RealFinalNB Then
        InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
    ' Part D
    Set BaseCell = CurrentSheet.Range("A:A").Find("TOTAL")
    If BaseCell Is Nothing Then
        Exit Sub
    End If
    If BaseCell.Cells(5, 1).Value <> "Prénom" Then
        Exit Sub
    End If
    Set BaseCell = BaseCell.Cells(5, 1)

    If PreviousNB < RealFinalNB Then
        InsertRows BaseCell, PreviousNB, RealFinalNB
    Else
        If PreviousNB > RealFinalNB Then
            RemoveRows BaseCell, PreviousNB, RealFinalNB, 0, True
        End If
    End If
    
End Sub
Public Sub ChangeNBSalariesDansChantier(wb As Workbook, PreviousNB As Integer, FinalNB As Integer)
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim RealFinalNB As Integer
    Dim TmpRange As Range
    
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If CurrentSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Budget_chantiers & "' n'a pas été trouvée"
        Exit Sub
    End If
    Set BaseCell = CurrentSheet.Range("A:A").Find("Salarié")
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
        InsertRows BaseCell, PreviousNB, RealFinalNB, False, NBExtraCols, True
        Set TmpRange = InsertRows(BaseCell.Cells(1 + RealFinalNB + 2, 1), PreviousNB, RealFinalNB, False, NBExtraCols, False)

        UpdateSumsByColumn _
            Range( _
                BaseCell.Cells(1 + RealFinalNB + 3, 3), _
                BaseCell.Cells(1 + 2 * RealFinalNB + 2, TmpRange.Columns.Count) _
            ), _
            BaseCell.Cells(0, 3), _
            PreviousNB
    Else
        If PreviousNB > RealFinalNB Then
            RemoveRows BaseCell, PreviousNB, RealFinalNB, NBExtraCols
            RemoveRows BaseCell.Cells(1 + RealFinalNB + 2, 1), PreviousNB, RealFinalNB, NBExtraCols
        End If
    End If
    If FinalNB <= 1 And PreviousNB > 1 Then
        Range(BaseCell.Cells(3, 1), FindNextNotEmpty(BaseCell, False).Cells(3, 1)).ClearContents
        Range(BaseCell.Cells(3 + RealFinalNB + 2, 1), FindNextNotEmpty(BaseCell, False).Cells(3 + RealFinalNB + 2, 1)).ClearContents
    End If
    
End Sub

Public Function IsEmptyFunding(FirstCell As Range, NBChantiers As Integer, TwoLines As Boolean) As Boolean
    
    Dim CurrentValue As String
    Dim Index As Integer
    
    IsEmptyFunding = False
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

    IsEmptyFunding = True
End Function

Public Sub RetirerLignesVidesFunc( _
    SetOfRange As SetOfRange, _
    NBChantiers As Integer _
    )

    Dim CurrentIndex As Integer
    Dim FirstCellOfLine As Range
    Dim NBRows As Integer
    Dim IndexLine As Integer
    Dim ValueOfFirstCellOfLine As String
    Dim ValueOfSecondCellOfLine As String
    Dim ValueOfSecondCellOfNextLine As String

    NBRows = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
    CurrentIndex = 1

    For IndexLine = 1 To NBRows
        Set FirstCellOfLine = SetOfRange.HeadCell.Cells(1 + CurrentIndex, 1)
        ValueOfFirstCellOfLine = FirstCellOfLine.Value
        ValueOfSecondCellOfLine = FirstCellOfLine.Cells(1, 2).Value
        ValueOfSecondCellOfNextLine = FirstCellOfLine.Cells(2, 2).Value
        If ValueOfSecondCellOfLine = "" _
            And ValueOfFirstCellOfLine <> "" _
            And ValueOfFirstCellOfLine <> Empty _
            And ValueOfSecondCellOfNextLine = "Statut" Then
            ' two lines
            If IsEmptyFunding(FirstCellOfLine.Cells(1, 3), NBChantiers, True) Then
                Range( _
                    FirstCellOfLine, _
                    FirstCellOfLine.Cells(2, 3 + NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
                CurrentIndex = CurrentIndex - 2
            End If
        Else
            If ValueOfFirstCellOfLine = "" _
                And ValueOfSecondCellOfLine = "" _
                And ValueOfSecondCellOfNextLine <> "Statut" Then
                ' one line
                If IsEmptyFunding(FirstCellOfLine.Cells(1, 3), NBChantiers, False) Then
                    Range( _
                        FirstCellOfLine, _
                        FirstCellOfLine.Cells(1, 3 + NBChantiers + NBExtraCols) _
                    ).Delete Shift:=xlUp
                    CurrentIndex = CurrentIndex - 1
                End If
            End If
        End If
        CurrentIndex = CurrentIndex + 1
    Next IndexLine
End Sub

Public Function GetTypeFinancementStr( _
        wb As Workbook, _
        TypeFinancement As Integer, _
        NewFinancementInChantier As FinancementComplet _
    ) As String

    Dim Financements() As Financement
    Dim Financement As Financement
    Dim TypeFinancementsLocal() As String

    TypeFinancementsLocal = TypeFinancementsFromWb(wb)

    If (TypeFinancement <> 0) Then
        GetTypeFinancementStr = TypeFinancementsLocal(TypeFinancement)
    Else
        If NewFinancementInChantier.Status Then
            Financements = NewFinancementInChantier.Financements
            Financement = Financements(1)
            If Financement.TypeFinancement <> 0 Then
                GetTypeFinancementStr = TypeFinancementsLocal(Financement.TypeFinancement)
            Else
                GetTypeFinancementStr = ""
            End If
        Else
            GetTypeFinancementStr = ""
        End If
    End If
End Function

Public Function PrepareAjoutFinancement( _
        wb As Workbook, _
        NBChantiers As Integer, _
        Optional RetirerLignesVides As Boolean = False _
    ) As SetOfRange

    Dim ChantierSheet As Worksheet
    Dim SetOfRange As SetOfRange

    ' Default
    SetOfRange.Status = False
    PrepareAjoutFinancement = SetOfRange

    
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Budget_chantiers & "' n'a pas été trouvée"
        Exit Function
    End If
    Set SetOfRange.ChantierSheet = ChantierSheet
    PrepareAjoutFinancement = SetOfRange

    SetOfRange = GetCellsForFinancement(SetOfRange.ChantierSheet)
    PrepareAjoutFinancement = SetOfRange
    If Not SetOfRange.Status Then
        Exit Function
    End If

    If RetirerLignesVides Then
        RetirerLignesVidesFunc SetOfRange, NBChantiers
        SetOfRange = GetCellsForFinancement(SetOfRange.ChantierSheet)
        PrepareAjoutFinancement = SetOfRange
    End If

End Function

Public Function AjoutFinancementInternal( _
        SetOfRange As SetOfRange, _
        wb As Workbook, _
        NBChantiers As Integer, _
        NewFinancementInChantier As FinancementComplet, _
        Optional Nom As String = "", _
        Optional TypeFinancement As Integer = 0 _
    ) As SetOfRange

    Dim BaseCell As Range
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
    
    TypeFinancementStr = GetTypeFinancementStr(wb, TypeFinancement, NewFinancementInChantier)
    AjoutFinancementInternal = SetOfRange

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
                                And IsEmptyFunding(WorkingRange, 2 + NBChantiers, False) Then
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
                    If IsEmptyFunding(WorkingRange, 2 + NBChantiers, False) _
                        And ValueOfSecondCellOnNextLine <> "Statut" Then
                        Set BaseCell = WorkingRange
                        ShoudInsert = False
                    End If
                End If
            Next Index
        End If
    End If

    IsEmptyRow = IsEmptyFunding(BaseCell, 2 + NBChantiers, False)
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
        
        InsertRows _
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
    Else
        Set WorkingRange = BaseCell
    End If

    WorkingRange.Cells(1, NBChantiers + 3).Formula = "=SUM(" & _
        Range( _
            WorkingRange.Cells(1, 3), _
            WorkingRange.Cells(1, NBChantiers + 2) _
        ).address(False, False, xlA1) & _
    ")"

    If TypeFinancementStr <> "" Then
        WorkingRange.Cells(1, 1).Value = TypeFinancementStr
        WorkingRange.Cells(2, 2).Value = "Statut"
        WorkingRange.Cells(2, 3 + NBChantiers).Value = ""
    End If

    If Not (NewFinancementInChantier.Status) Then
        WorkingRange.Cells(1, 2).Value = Nom
    Else
        TmpFinancement = NewFinancementInChantier.Financements(1)
        WorkingRange.Cells(1, 2).Value = TmpFinancement.Nom
        For Index = 1 To UBound(NewFinancementInChantier.Financements)
            TmpFinancement = NewFinancementInChantier.Financements(Index)
            If TmpFinancement.Valeur <> 0 Then
                WorkingRange.Cells(1, 2 + Index) = TmpFinancement.Valeur
            End If
            If TypeFinancementStr <> "" Then
                If TmpFinancement.Statut <> 0 Then
                    WorkingRange.Cells(2, 2 + Index) = TypeStatut()(TmpFinancement.Statut)
                End If
            End If
        Next Index
    End If
    ' adjust and return
    Set AjoutFinancementInternal.EndCell = SetOfRange.ResultCell.Cells(0, 0)
End Function

Public Sub DefinirFormatChantiers( _
        ChantierSheet As Worksheet, _
        NBChantiers As Integer, _
        Optional AddTopBorder As Boolean = True, _
        Optional AddBottomBorder As Boolean = True _
    )

    Dim ColumnIndex As Integer
    Dim CurrentArea As Range
    Dim NBColumns As Integer
    Dim NBRows As Integer
    Dim RowIndex As Integer
    Dim SetOfRange As SetOfRange
    Dim ValueOfSecondCellOfLine As String

    SetOfRange = GetCellsForFinancement(ChantierSheet)
    If Not SetOfRange.Status Then
        Exit Sub
    End If
    
    Set CurrentArea = Range( _
        SetOfRange.HeadCell.Cells(2, 1), _
        SetOfRange.EndCell.Cells(1, 3 + NBChantiers) _
    )

    NBColumns = CurrentArea.Columns.Count
    NBRows = CurrentArea.Rows.Count

    For RowIndex = 1 To NBRows
        ValueOfSecondCellOfLine = CurrentArea.Cells(RowIndex, 2).Value
        For ColumnIndex = 1 To NBColumns
            DefinirFormatPourChantier CurrentArea.Cells(RowIndex, ColumnIndex), _
                (AddTopBorder And RowIndex = 1), _
                (AddBottomBorder And RowIndex = NBRows), _
                (ColumnIndex = 2 Or ColumnIndex = NBColumns), _
                (ValueOfSecondCellOfLine = "Statut" And ColumnIndex <= 2), _
                (ColumnIndex = NBColumns), _
                (ColumnIndex > 2)
        Next ColumnIndex
        If ValueOfSecondCellOfLine = "Statut" Then
            AddValidationDossier Range( _
                CurrentArea.Cells(RowIndex, 3), _
                CurrentArea.Cells(RowIndex, 2 + NBChantiers) _
            )
            Range( _
                CurrentArea.Cells(RowIndex, 1), _
                CurrentArea.Cells(RowIndex, 2) _
            ).Validation.Delete
            CurrentArea.Cells(RowIndex, 3 + NBChantiers).Validation.Delete
        Else
            Range( _
                CurrentArea.Cells(RowIndex, 1), _
                CurrentArea.Cells(RowIndex, 3 + NBChantiers) _
            ).Validation.Delete
        End If
    Next RowIndex

    DefinirFormatConditionnelPourLesDossier SetOfRange, NBChantiers
End Sub

Public Sub AjoutFinancement(wb As Workbook, _
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

    SetOfRange = PrepareAjoutFinancement(wb, NBChantiers)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    SetOfRange = AjoutFinancementInternal(SetOfRange, wb, NBChantiers, NewFinancementInChantier, Nom, TypeFinancement)
    
    RenewFormulaForTotalFinancement SetOfRange.ChantierSheet, NBChantiers
    DefinirFormatChantiers SetOfRange.ChantierSheet, NBChantiers

End Sub

Public Sub AddValidationDossier(CurrentRange As Range)
    
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

Public Function InsertRows( _
    BaseCell As Range, _
    PreviousNB As Integer, _
    FinalNB As Integer, _
    Optional AutoFitNext As Boolean = True, _
    Optional ExtraCols As Integer = 0, _
    Optional UpdateSum As Boolean = True) As Range

    Dim endOfRow As Range
    
    Set endOfRow = FindNextNotEmpty(BaseCell, False) ' To Right
    ' Insert Cells
    BaseCell.Worksheet.Activate
    BaseCell.Cells(1, 1).Select 'Force Selection
    Range(BaseCell.Cells(1 + PreviousNB, 1), endOfRow.Cells(1 + PreviousNB, 1 + ExtraCols)).Copy
    Range(BaseCell.Cells(1 + PreviousNB + 1, 1), endOfRow.Cells(1 + FinalNB, 1 + ExtraCols)).Insert _
        Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Copy Format
    Range(BaseCell.Cells(1 + PreviousNB, 1), endOfRow.Cells(1 + PreviousNB, 1 + ExtraCols)).Copy
    Range(BaseCell.Cells(1 + PreviousNB + 1, 1), endOfRow.Cells(1 + FinalNB, 1 + ExtraCols)).PasteSpecial _
        Paste:=xlPasteFormats
    If PreviousNB > 2 Then
        Range(BaseCell.Cells(1 + PreviousNB - 1, 1), endOfRow.Cells(1 + PreviousNB - 1, 1 + ExtraCols)).Copy
        Range(BaseCell.Cells(1 + PreviousNB, 1), endOfRow.Cells(1 + FinalNB, 1 + ExtraCols)).PasteSpecial _
            Paste:=xlPasteFormats
    End If
    
    ' Update Sums
    If UpdateSum Then
        UpdateSumsByColumn _
            Range(BaseCell.Cells(2, 1), endOfRow.Cells(1 + FinalNB, 1 + ExtraCols)), _
            BaseCell.Cells(1 + FinalNB + 1, 1), _
            PreviousNB
    End If
        
    ' Row AutoFit
    On Error Resume Next
    If AutoFitNext Then
        Range(BaseCell.Cells(2, 1).EntireRow, BaseCell.Cells(1 + FinalNB, 1).EntireRow).RowHeight = 18 ' Instead of AutoFit
        Range(BaseCell.Cells(1 + FinalNB + 1, 1).EntireRow, BaseCell.Cells(1 + FinalNB + FinalNB - PreviousNB, 1).EntireRow).AutoFit ' Instead of AutoFit
    Else
        Range(BaseCell.Cells(2, 1).EntireRow, BaseCell.Cells(1 + FinalNB, 1).EntireRow).AutoFit
    End If
    On Error GoTo 0
    Set InsertRows = Range(BaseCell.Cells(1 + FinalNB + 1, 1), endOfRow.Cells(1 + FinalNB + 1, 1))
    BaseCell.Select 'Force Selection
End Function

Public Sub UpdateSumsByColumn(BaseRange As Range, DestinationRange As Range, PreviousNB As Integer)
    
    Dim FormulaRelative As String
    Dim FormulaAbsolute As String
    Dim Index As Integer
    Dim NBColumns As Integer
    Dim NBRows As Integer
    Dim SanitizedPreviousNB As Integer

    NBColumns = BaseRange.Columns.Count
    NBRows = BaseRange.Rows.Count

    If PreviousNB < 1 Then
        SanitizedPreviousNB = 1
    Else
        If PreviousNB > NBRows Then
            SanitizedPreviousNB = NBRows
        Else
            SanitizedPreviousNB = PreviousNB
        End If
    End If

    For Index = 1 To NBColumns
        FormulaRelative = "=SUM(" & CleanAddress( _
                Range( _
                    BaseRange.Cells(1, Index), _
                    BaseRange.Cells(PreviousNB, Index) _
                ).address(False, False, xlA1, False) _
            ) & ")"
        FormulaAbsolute = "=SUM(" & CleanAddress( _
                Range( _
                    BaseRange.Cells(1, Index), _
                    BaseRange.Cells(PreviousNB, Index) _
                ).address(False, False, xlA1, False) _
            ) & ")"
        If DestinationRange.Cells(1, Index).Formula = FormulaRelative _
            Or DestinationRange.Cells(1, Index).Formula = FormulaAbsolute Then
            DestinationRange.Cells(1, Index).Formula = "=SUM(" & CleanAddress( _
                Range( _
                    BaseRange.Cells(1, Index), _
                    BaseRange.Cells(NBRows, Index) _
                ).address(False, False, xlA1, False) _
            ) & ")"
        End If
    Next Index
End Sub

Public Sub RemoveRows(BaseCell As Range, PreviousNB As Integer, FinalNB As Integer, Optional ExtraCols As Integer = 0, Optional AutoFitNext As Boolean = False)

    ' Remove Cells
    Range(BaseCell.Cells(1 + FinalNB + 1, 1), FindNextNotEmpty(BaseCell, False).Cells(1 + PreviousNB, 1 + ExtraCols)).Delete _
        Shift:=xlShiftUp
    
    ' Row AutoFit
    On Error Resume Next
    If AutoFitNext Then
        Range(BaseCell.Cells(1 + FinalNB + 1, 1).EntireRow, BaseCell.Cells(1 + FinalNB + FinalNB - PreviousNB, 1).EntireRow).AutoFit ' Instead of AutoFit
    End If
    On Error GoTo 0
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
        MsgBox "'" & Nom_Feuille_Personnel & "' n'a pas été trouvée"
        Exit Sub
    End If
    
    Set BaseCell = CurrentSheet.Range("A:A").Find("Prénom")
    If BaseCell Is Nothing Then
        MsgBox "'Prénom' non trouvé dans '" & Nom_Feuille_Personnel & "' !"
        Exit Sub
    End If
    
    If FinalNB <= 1 Then
        RealFinalNB = 2
    Else
        RealFinalNB = FinalNB
    End If
    
    Set endR = FindNextNotEmpty(BaseCell, False)
    
    If PreviousNB > RealFinalNB Then
        RemoveRows BaseCell, PreviousNB, RealFinalNB, 1
    Else
        If PreviousNB < FinalNB Then
            InsertRows BaseCell, PreviousNB, FinalNB, True, 1
        End If
    End If
    
    If FinalNB <= 1 And PreviousNB > 1 Then
        Range(BaseCell.Cells(3, 1), endR.Cells(3, 1)).ClearContents
    End If

End Sub

Public Function extraireDepensesChantier( _
        BaseCellChantier As Range, _
        NBSalaries As Integer, _
        NBChantiers As Integer, _
        Optional BaseCell As Range _
    ) As SetOfChantiers
        
    Dim Chantiers() As Chantier
    Dim SetOfChantiers As SetOfChantiers
    Dim IndexChantiers As Integer
    Dim IndexDepense As Integer
    Dim NBDepenses As Integer
    Dim NewFormatForAutofinancement As Integer
    Dim BaseCellLocal As Range
    Dim ChantierTmp As Chantier
    Dim ChantierTmp1 As Chantier
    Dim DepensesTmp1() As DepenseChantier
    Dim DepenseTmp As DepenseChantier
    Dim SetOfRange As SetOfRange
    Dim TestedValue As String
    
    ' Depenses
    If BaseCell Is Nothing Then
        SetOfRange = GetCellsForDepense(BaseCellChantier.Worksheet)
        If SetOfRange.Status Then
            Set BaseCell = SetOfRange.HeadCell.Cells(2, 2)
        Else
            ' Backup
            Set BaseCell = BaseCellChantier.Cells(7 + 2 * NBSalaries, 1).EntireRow.Cells(1, 2)
        End If
    End If
    NBDepenses = Range(BaseCell, FindNextNotEmpty(BaseCell, True).Cells(0, 1)).Rows.Count
    
    SetOfChantiers = getDefaultSetOfChantiers(NBChantiers, NBDepenses)

    For IndexDepense = 1 To NBDepenses
        updateNameDepense SetOfChantiers, 1, IndexDepense, BaseCell.Cells(IndexDepense, 1).Value
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
                updateNameDepense SetOfChantiers, IndexChantiers, IndexDepense, DepenseTmp.Nom
            End If
            updateValDepense SetOfChantiers, IndexChantiers, IndexDepense, BaseCell.Cells(IndexDepense, IndexChantiers + 1).Value
            updateBaseCellDepense SetOfChantiers, IndexChantiers, IndexDepense, BaseCell.Cells(IndexDepense, IndexChantiers + 1)
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
            If NewFormatForAutofinancement > 1 Then
                ChantierTmp.AutoFinancementAutres = BaseCellLocal.Cells(-2, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementStructureAnneesPrecedentes = BaseCellLocal.Cells(5, 1 + IndexChantiers).Value
                ChantierTmp.AutoFinancementAutresAnneesPrecedentes = BaseCellLocal.Cells(4, 1 + IndexChantiers).Value
                ChantierTmp.CAanneesPrecedentes = BaseCellLocal.Cells(6, 1 + IndexChantiers).Value
            Else
                ChantierTmp.AutoFinancementAutres = BaseCellLocal.Cells(2, 1 + IndexChantiers).Value
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
    
    extraireDepensesChantier = SetOfChantiers

End Function

Public Function extraireNomsChantier( _
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
    extraireNomsChantier = SetOfChantiers
End Function

Public Function extraireFinancementChantier( _
        BaseCellChantier As Range, _
        NBChantiers As Integer, _
        Data As Data, _
        Optional ForV0 As Boolean = False _
        ) As SetOfChantiers
    Dim Chantiers() As Chantier
    Dim SetOfChantiers As SetOfChantiers
    Dim ChantierTmp As Chantier
    Dim BaseCell As Range
    Dim FoundCell As Range
    Dim IndexChantiers As Integer
    Dim IndexFinancement As Integer
    Dim IndexType As Integer
    Dim NBFinancements As Integer
    Dim TypesFinancements As Variant
    Dim TypesStatuts As Variant
    Dim IndexTypeName As Integer
    Dim ColCounter As Integer
    Dim FinancementTmp As Financement
    Dim FinancementTmp1 As Financement
    Dim FinancementsTmp() As Financement
    Dim FinancementsTmp1() As Financement
    
    TypesFinancements = TypeFinancementsFromWb(BaseCellChantier.Worksheet.Parent)
    TypesStatuts = TypeStatut()
    
    Chantiers = Data.Chantiers
    SetOfChantiers.Chantiers = Chantiers
    
    If ForV0 Then
        Set BaseCell = TrouveBaseCellFinancementV0(BaseCellChantier)
        If BaseCell.address = BaseCellChantier.address Then
            GoTo FinFunction
        End If
    Else
        Set BaseCell = BaseCellChantier.EntireRow.Cells(1, 1).EntireColumn.Find(Label_Type_Financeur)
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
    ColCounter = 0
    For IndexFinancement = 1 To NBFinancements
        If BaseCell.Cells(IndexFinancement, 2).Value <> "Statut" Then
            ColCounter = ColCounter + 1
        End If
    Next IndexFinancement
    NBFinancements = ColCounter
    
    For IndexChantiers = 1 To NBChantiers
        ChantierTmp = Chantiers(IndexChantiers)
        FinancementsTmp = getDefaultFinancements(NBFinancements)
        ChantierTmp.Financements = FinancementsTmp
        Chantiers(IndexChantiers) = ChantierTmp
    Next IndexChantiers
    
    ' Extraction des types avec le chantier 1
    ColCounter = 1
    ChantierTmp = Chantiers(1)
    FinancementsTmp1 = ChantierTmp.Financements
    For IndexFinancement = 1 To NBFinancements
        FinancementTmp1 = FinancementsTmp1(IndexFinancement)
        FinancementTmp1.Nom = BaseCell.Cells(ColCounter, 2).Value
        IndexType = 0
        For IndexTypeName = 1 To UBound(TypesFinancements)
            If TypesFinancements(IndexTypeName) = BaseCell.Cells(ColCounter, 1).Value Then
                IndexType = IndexTypeName
            End If
        Next IndexTypeName
        FinancementTmp1.TypeFinancement = IndexType
        If IndexType > 0 Then
            ColCounter = ColCounter + 1
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
        ColCounter = ColCounter + 1
    Next IndexFinancement
    ChantierTmp.Financements = FinancementsTmp1
    Chantiers(1) = ChantierTmp
    
    ' Extraction des valeurs
    For IndexChantiers = 1 To NBChantiers
        ColCounter = 1
        ChantierTmp = Chantiers(IndexChantiers)
        For IndexFinancement = 1 To NBFinancements
            FinancementsTmp = ChantierTmp.Financements
            FinancementTmp = FinancementsTmp(IndexFinancement)
            FinancementTmp1 = FinancementsTmp1(IndexFinancement)
            ' récupération du type depuis le chantier 1
            If IndexChantiers > 1 Then
                FinancementTmp.Nom = FinancementTmp1.Nom
                FinancementTmp.TypeFinancement = FinancementTmp1.TypeFinancement
            End If
            FinancementTmp.Valeur = BaseCell.Cells(ColCounter, IndexChantiers + 2).Value
            Set FinancementTmp.BaseCell = BaseCell.Cells(ColCounter, IndexChantiers + 2)
            
            If FinancementTmp.TypeFinancement > 0 And Not ForV0 Then
                IndexType = 0
                For IndexTypeName = 1 To UBound(TypesStatuts)
                    If TypesStatuts(IndexTypeName) = BaseCell.Cells(ColCounter + 1, IndexChantiers + 2).Value Then
                        IndexType = IndexTypeName
                    End If
                Next IndexTypeName
                FinancementTmp.Statut = IndexType
                ColCounter = ColCounter + 1
            Else
                FinancementTmp.Statut = 0
            End If
            ColCounter = ColCounter + 1
            FinancementsTmp(IndexFinancement) = FinancementTmp
            ChantierTmp.Financements = FinancementsTmp
        Next IndexFinancement
        Chantiers(IndexChantiers) = ChantierTmp
    Next IndexChantiers
    
    SetOfChantiers.Chantiers = Chantiers
    
FinFunction:
    extraireFinancementChantier = SetOfChantiers

End Function
Public Function TrouveBaseCellFinancementV0(BaseCellChantier As Range) As Range
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
    Set TrouveBaseCellFinancementV0 = BaseCell
    Exit Function
FinFunctionAvecErreur:
    Set TrouveBaseCellFinancementV0 = BaseCellChantier
End Function

Public Function extraireCharges(wb As Workbook, Data As Data, Revision As WbRevision) As Data
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
        MsgBox "'" & Nom_Feuille_Charges & "' n'a pas été trouvée"
        GoTo FinFunction
    End If

    ' Get Titles for categories for charges
    If Revision.Majeure <= 1 Then
        Titles(1) = "Cat. 1"
        Titles(2) = "Cat. 2"
        Titles(3) = "Cat. 3"
    Else
        TitlesBaseColumn = 6
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
    HasRealValues = False
    If (Revision.Majeure = 2 And Revision.Mineure > 1) Or Revision.Majeure > 2 Then
        HasRealValues = True
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
                Charges = getChargesDefault(NBNewCharges).Charges
            Else
                SetOfCharges.Charges = Charges
                Charges = getChargesDefaultPreserve(SetOfCharges, PreviousIndex + NBNewCharges).Charges
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
    extraireCharges = Data
End Function

Public Sub ClearFinancements( _
        ChantierSheet As Worksheet, _
        NBChantiers As Integer _
    )
    Dim Index As Integer
    Dim SetOfRange As SetOfRange
    
    SetOfRange = GetCellsForFinancement(ChantierSheet)
    If SetOfRange.Status Then
        If SetOfRange.EndCell.Row > SetOfRange.HeadCell.Row + 1 Then
            Range( _
                    SetOfRange.HeadCell.Cells(2, 1), _
                    SetOfRange.EndCell.Cells(1, 3 + NBChantiers + NBExtraCols) _
                ).Delete Shift:=xlUp
        End If
    End If
End Sub

Public Sub insererDonnees(NewWorkbook As Workbook, Data As Data)
    Dim BaseCell As Range
    Dim BaseCellChantier As Range
    Dim ChantierSheet As Worksheet
    Dim Chantiers() As Chantier
    Dim CurrentSheet As Worksheet
    Dim DepenseTmp As DepenseChantier
    Dim DepensesTmp() As DepenseChantier
    Dim DonneesSalarie As DonneesSalarie
    Dim FinancementCompletTmp As FinancementComplet
    Dim Financements() As Financement
    Dim FinancementsTmp() As Financement
    Dim Index As Integer
    Dim IndexChantier As Integer
    Dim IndexTab As Integer
    Dim NBChantiers As Integer
    Dim NBSalaries As Integer
    Dim SetOfRange As SetOfRange
    Dim TauxAutre As Double
    Dim TmpChantier As Chantier
    Dim TmpChantier1 As Chantier
    Dim TotalCell As Range

    FinancementCompletTmp = getDefaultFinancementComplet()

    importerInfos NewWorkbook, Data.Informations
    
    NBSalaries = GetNbSalaries(NewWorkbook)
    If NBSalaries > 0 Then
        Set CurrentSheet = NewWorkbook.Worksheets(Nom_Feuille_Personnel)
        If CurrentSheet Is Nothing Then
            MsgBox "'" & Nom_Feuille_Personnel & "' n'a pas été trouvée"
        Else
            Set BaseCell = CurrentSheet.Range("A:A").Find("Prénom")
            If BaseCell Is Nothing Then
                MsgBox "'Prénom' non trouvé dans '" & Nom_Feuille_Personnel & "' !"
            Else
                On Error Resume Next
                Set ChantierSheet = NewWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
                On Error GoTo 0
                NBChantiers = 0
                If ChantierSheet Is Nothing Then
                    Set BaseCellChantier = Nothing
                Else
                    Set BaseCellChantier = FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
                    If BaseCellChantier.Column > 1000 Or Left(BaseCellChantier.Value, Len("Chantier")) <> "Chantier" Then
                        Set BaseCellChantier = Nothing
                    Else
                        NBChantiers = GetNbChantiers(NewWorkbook)
                    End If
                End If
                
                Index = 1
                For IndexTab = LBound(Data.Salaries) To UBound(Data.Salaries)
                    DonneesSalarie = Data.Salaries(IndexTab)
                    
                    If Not DonneesSalarie.Erreur And Index <= NBSalaries Then
                        BaseCell.Cells(1 + Index, 1).Value = DonneesSalarie.Prenom
                        BaseCell.Cells(1 + Index, 2).Value = DonneesSalarie.Nom
                        If DonneesSalarie.TauxDeTempsDeTravailFormula = "" Then
                            BaseCell.Cells(1 + Index, 3).Value = DonneesSalarie.TauxDeTempsDeTravail
                        Else
                            BaseCell.Cells(1 + Index, 3).Formula = DonneesSalarie.TauxDeTempsDeTravailFormula
                        End If
                        If DonneesSalarie.MasseSalarialeAnnuelleFormula = "" Then
                            BaseCell.Cells(1 + Index, 4).Value = DonneesSalarie.MasseSalarialeAnnuelle
                        Else
                            BaseCell.Cells(1 + Index, 4).Formula = DonneesSalarie.MasseSalarialeAnnuelleFormula
                        End If
                        If DonneesSalarie.TauxOperateurFormula = "" Then
                            BaseCell.Cells(1 + Index, 5).Value = DonneesSalarie.TauxOperateur
                        Else
                            BaseCell.Cells(1 + Index, 5).Formula = DonneesSalarie.TauxOperateurFormula
                        End If
                        If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) Then
                            For IndexChantier = 1 To WorksheetFunction.Min(NBChantiers, UBound(DonneesSalarie.JoursChantiers))
                                If CInt(DonneesSalarie.JoursChantiers(IndexChantier)) = 0 Or CStr(DonneesSalarie.JoursChantiers(IndexChantier)) = "" Then
                                    BaseCellChantier.Cells(4 + Index, IndexChantier).Value = ""
                                Else
                                    BaseCellChantier.Cells(4 + Index, IndexChantier).Value = DonneesSalarie.JoursChantiers(IndexChantier)
                                End If
                            Next IndexChantier
                        End If
                        Index = Index + 1
                    End If
                Next IndexTab
                If (Not BaseCellChantier Is Nothing) And (NBChantiers > 0) And UBound(Data.Chantiers) > 1 Then
                    ' nom des dépenses
                    SetOfRange = GetCellsForDepense(ChantierSheet)
                    If SetOfRange.Status Then
                        Set BaseCell = SetOfRange.HeadCell.Cells(2, 2)
                    Else
                        ' backup
                        Set BaseCell = BaseCellChantier.Cells(7 + 2 * NBSalaries, 1).EntireRow.Cells(1, 2)
                    End If
                    TmpChantier = Data.Chantiers(1)
                    TmpChantier1 = Data.Chantiers(1)
                    DepensesTmp = TmpChantier1.Depenses

                    ChangeDepenses BaseCell, NBSalaries, UBound(TmpChantier.Depenses), NBChantiers
                    
                    For Index = 1 To UBound(TmpChantier.Depenses)
                        DepenseTmp = DepensesTmp(Index)
                        If DepenseTmp.Nom = "0" Then
                            BaseCell.Cells(Index, 1).Value = ""
                        Else
                            BaseCell.Cells(Index, 1).Value = DepenseTmp.Nom
                        End If
                    Next Index
                    
                    Chantiers = Data.Chantiers
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
                            If DepenseTmp.Valeur = 0 Then
                                BaseCell.Cells(Index, 1 + IndexChantier).Value = ""
                            Else
                                BaseCell.Cells(Index, 1 + IndexChantier).Value = DepenseTmp.Valeur
                            End If
                        Next Index
                    Next IndexChantier
                    DepensesTmp = TmpChantier1.Depenses
                    Set TotalCell = BaseCell.Cells(UBound(DepensesTmp) + 1, 1)
                    
                    ' Financements
                    ClearFinancements ChantierSheet, NBChantiers
                    Financements = TmpChantier1.Financements
                    If UBound(Chantiers) > 0 And UBound(Financements) > 0 Then
                        ReDim FinancementsTmp(1 To UBound(Chantiers))
                        SetOfRange = PrepareAjoutFinancement(NewWorkbook, NBChantiers, False)
                        If SetOfRange.Status Then
                            For Index = 1 To UBound(Financements)
                                For IndexChantier = 1 To UBound(Chantiers)
                                    TmpChantier = Chantiers(IndexChantier)
                                    Financements = TmpChantier.Financements
                                    FinancementsTmp(IndexChantier) = Financements(Index)
                                Next IndexChantier
                                FinancementCompletTmp.Financements = FinancementsTmp
                                FinancementCompletTmp.Status = True
                                SetOfRange = AjoutFinancementInternal(SetOfRange, NewWorkbook, NBChantiers, FinancementCompletTmp, "", 0)
                            Next Index
                            RenewFormulaForTotalFinancement SetOfRange.ChantierSheet, NBChantiers
                            DefinirFormatChantiers SetOfRange.ChantierSheet, NBChantiers
                        End If
                    End If
                    
                    ' Autofinancement
                    Application.Calculate
                    SetOfRange = GetCellsForFinancement(ChantierSheet)
                    If SetOfRange.Status Then
                        For IndexChantier = 1 To UBound(Chantiers)
                            TmpChantier = Chantiers(IndexChantier)
                            ' does not set AutoFinancementStructure because calculated !
                            SetOfRange.ResultCell.Cells(2, 1 + IndexChantier).Value = TmpChantier.AutoFinancementAutres
                            SetOfRange.ResultCell.Cells(10, 1 + IndexChantier).Value = TmpChantier.AutoFinancementStructureAnneesPrecedentes
                            SetOfRange.ResultCell.Cells(9, 1 + IndexChantier).Value = TmpChantier.AutoFinancementAutresAnneesPrecedentes
                            SetOfRange.ResultCell.Cells(11, 1 + IndexChantier).Value = TmpChantier.CAanneesPrecedentes
                        Next IndexChantier
                    End If
                End If
            End If
        End If
    End If
    
    ' Ajouter Charges
    AjoutCharges NewWorkbook, Data

End Sub

Public Sub ChangeDepenses(BaseCell As Range, NBSalaries As Integer, NewNBDepenses As Integer, NBChantiers As Integer)
    Dim PreviousNBDepenses As Integer
    PreviousNBDepenses = Range(BaseCell, FindNextNotEmpty(BaseCell, True).Cells(0, 1)).Rows.Count
                    
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

Public Sub InitialiserLesFinancements(wb As Workbook, NBFinancements As Integer, Optional Init As Boolean = False)

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

    SetOfRange = PrepareAjoutFinancement(wb, NBChantiers, Init)
    If Not SetOfRange.Status Then
        Exit Sub
    End If
    
    For Index = 1 To UBound(TypesFinancements)
        For IndexLoop = 1 To NBFinancements
            SetOfRange = AjoutFinancementInternal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Client " & (IndexLoop + (Index - 1) * NBFinancements), Index)
        Next IndexLoop
    Next Index
    For IndexLoop = 1 To NBFinancements
        SetOfRange = AjoutFinancementInternal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Formations", 0)
    Next IndexLoop
    For IndexLoop = 1 To NBFinancements
        SetOfRange = AjoutFinancementInternal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Prestations", 0)
    Next IndexLoop
    For IndexLoop = 1 To NBFinancements
        SetOfRange = AjoutFinancementInternal(SetOfRange, wb, NBChantiers, FinancementCompletTmp, "Cotisations", 0)
    Next IndexLoop
    
    RenewFormulaForTotalFinancement SetOfRange.ChantierSheet, NBChantiers
    DefinirFormatChantiers SetOfRange.ChantierSheet, NBChantiers
End Sub

Public Function InsertNewLineForCharges(CurrentCell As Range) As Range

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
    Set InsertNewLineForCharges = CurrentCell.Cells(2, 1)
End Function

Public Function GetRowsOfCategoriesOfCharges(ChargesSheet As Worksheet) As SetOfCellsCategories
    
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

    GetRowsOfCategoriesOfCharges = SetOfCellsCategories
End Function

Public Function GetNextCell(SetOfCellsCategories As SetOfCellsCategories, IndexOfCat As Integer) As SetOfRange

    Dim CurrentCell As Range
    Dim CurrentCells() As Range
    Dim HeadCell As Range
    Dim Index As Integer
    Dim NextCell As Range
    Dim SetOfRange As SetOfRange

    SetOfRange.Status = False

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

    GetNextCell = SetOfRange
End Function

Public Function Ajout1Charge(SetOfCellsCategories As SetOfCellsCategories, Charge As Charge) As SetOfRange

    Dim CatIndex As Integer
    Dim CurrentCell As Range
    Dim Offset As Integer
    Dim SetOfRange As SetOfRange
    Dim TypeCharge As TypeCharge

    SetOfRange.Status = False

    If Charge.IndexTypeCharge > 0 Then
        TypeCharge = TypesDeCharges().Values(Charge.IndexTypeCharge)
        If TypeCharge.Index > 0 Then
            SetOfRange = GetNextCell(SetOfCellsCategories, TypeCharge.Index)
            If SetOfRange.Status Then
                If SetOfRange.EndCell.Cells(0, 1) = Empty Then
                    Offset = -1
                Else
                    Offset = 0
                End If
                Ajout1LigneCharge _
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

    Ajout1Charge = SetOfRange
End Function

Public Sub UpdateChargeFormula(SetOfRange As SetOfRange)

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

Public Sub Ajout1LigneCharge( _
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
    Set CurrentCell = InsertNewLineForCharges(BaseCell)
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

Public Sub AjoutCharges(wb As Workbook, Data As Data)
    Dim Charge As Charge
    Dim Charges() As Charge
    Dim ChargesSheet As Worksheet
    Dim Index As Integer
    Dim SetOfRange As SetOfRange
    Dim SetOfCellsCategories As SetOfCellsCategories
    Dim Titles() As String
    Dim TitlesBaseColumn As Integer
    Dim TitlesRow As Integer

    On Error Resume Next
    Set ChargesSheet = wb.Worksheets(Nom_Feuille_Charges)
    On Error GoTo 0
    If ChargesSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Charges & "' n'a pas été trouvée"
        Exit Sub
    End If

    ' update titles for categories
    TitlesBaseColumn = 6
    TitlesRow = 3
    Titles = Data.TitlesForChargesCat
    For Index = LBound(Titles) To UBound(Titles)
        ChargesSheet.Cells(TitlesRow, TitlesBaseColumn + Index).Value = Titles(Index)
    Next Index

    SetOfCellsCategories = GetRowsOfCategoriesOfCharges(ChargesSheet)
    CleanCategoryOfCharges SetOfCellsCategories
    Charges = Data.Charges

    ' add charges
    For Index = 1 To UBound(Charges)
        Charge = Charges(Index)
        Ajout1Charge SetOfCellsCategories, Charge
    Next Index

    ' add empty lines and update sums
    For Index = 60 To 68
        SetOfRange = GetNextCell(SetOfCellsCategories, Index)
        If SetOfRange.Status Then
            Ajout1LigneCharge SetOfRange.EndCell.Cells(0, 1)

            UpdateChargeFormula SetOfRange
        End If
    Next Index
End Sub

Public Sub CreerLigneCharge()

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
        MsgBox "'" & Nom_Feuille_Charges & "' n'a pas été trouvée"
        Exit Sub
    End If

    Value = InputBox("Quel nom de charge ?", "Ajouter une ligne de charge", "650 - Autre")
    FormattedValue = Trim(Value)

    If FormattedValue = "" Then
        MsgBox "Erreur : Le nom fourni pour la charge ne peut pas être vide"
        Exit Sub
    End If

    ExtractedValue = CInt(Left(FormattedValue, 2))
    If ExtractedValue < 60 Or ExtractedValue > 68 Then
        MsgBox "Erreur : les deux premiers caractères du nom doivent être compris entre 60 et 68 inclus."
        Exit Sub
    End If

    CodeIndex = FindTypeChargeIndexFromCode(ExtractedValue)
    If CodeIndex > 0 Then
        SetOfCellsCategories = GetRowsOfCategoriesOfCharges(ChargesSheet)
        SetOfRange = GetNextCell(SetOfCellsCategories, ExtractedValue)
        If SetOfRange.Status Then
            If SetOfRange.EndCell.Cells(0, 1) = Empty Then
                Offset = -1
            Else
                Offset = 0
            End If
            Ajout1LigneCharge SetOfRange.EndCell.Cells(Offset, 1), False, FormattedValue, 0, 0, 0, 0, 1
            UpdateChargeFormula SetOfRange
        Else
            MsgBox "Erreur : impossible de retrouver les différents types de paiement (60 à 68)"
        End If
    Else
        MsgBox "Erreur : impossible d'associer cette ligne à un type de paiement (entre 60 et 68)"
    End If

End Sub

Public Sub RetirerLigneCharge()
    
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
        MsgBox "'" & Nom_Feuille_Charges & "' n'a pas été trouvée"
        Exit Sub
    End If
    SetOfCellsCategories = GetRowsOfCategoriesOfCharges(ChargesSheet)
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

    NewLine = InputLineBetween( _
        "Supprimer la charge de la ligne ?", _
        "Ligne de charge à supprimer", _
        MinRow + 1, _
        MaxRow _
    )

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
        MsgBox "La ligne entrée n'est pas la ligne d'une charge"
        Exit Sub
    End If

    SetSilent

    Range(ChargesSheet.Cells(NewLine, 1), ChargesSheet.Cells(NewLine, ColumnOfSecondPartInCharge + NBCatOfCharges * 2 + 5)).Delete Shift:=xlShiftUp
    
    ' update sums
    For Index = 60 To 68
        SetOfRange = GetNextCell(SetOfCellsCategories, Index)
        If SetOfRange.Status Then
            UpdateChargeFormula SetOfRange
        End If
    Next Index

    SetActive
End Sub

Public Sub CleanCategoryOfCharges(SetOfCellsCategories As SetOfCellsCategories)

    Dim IndexOfCat As Integer
    Dim SetOfRange As SetOfRange

    For IndexOfCat = 60 To 68
        SetOfRange = GetNextCell(SetOfCellsCategories, IndexOfCat)
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

Public Function FindNextNotEmpty(BaseCell As Range, directionDown As Boolean) As Range

    Dim NB As Integer
    Dim CurrentRange As Range
    Dim NextRange As Range
    
    ' Init
    NB = 0
    Set CurrentRange = BaseCell
    
    If BaseCell.Value = "" Then
        While CurrentRange.Value = "" And NB < 1000
            If directionDown Then
                Set CurrentRange = CurrentRange.Cells(2, 1)
            Else
                Set CurrentRange = CurrentRange.Cells(1, 2)
            End If
            NB = NB + 1
        Wend
    Else
        Set NextRange = CurrentRange
        While NextRange.Value <> "" And NB < 1000
            Set CurrentRange = NextRange
            If directionDown Then
                Set NextRange = CurrentRange.Cells(2, 1)
            Else
                Set NextRange = CurrentRange.Cells(1, 2)
            End If
            NB = NB + 1
        Wend
    End If
    Set FindNextNotEmpty = CurrentRange

End Function

Public Sub updateNameDepense(SetOfChantiers As SetOfChantiers, IdxChantiers As Integer, IdxDepense As Integer, newName As String)
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

Public Sub updateValDepense(SetOfChantiers As SetOfChantiers, IdxChantiers As Integer, IdxDepense As Integer, newVal)
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim DepensesTmp() As DepenseChantier
    Dim TmpDepense As DepenseChantier
    
    Chantiers = SetOfChantiers.Chantiers
    ChantierTmp = Chantiers(IdxChantiers)
    DepensesTmp = ChantierTmp.Depenses
    TmpDepense = DepensesTmp(IdxDepense)
    TmpDepense.Valeur = newVal
    DepensesTmp(IdxDepense) = TmpDepense
    ChantierTmp.Depenses = DepensesTmp
    Chantiers(IdxChantiers) = ChantierTmp
    SetOfChantiers.Chantiers = Chantiers
End Sub

Public Sub updateBaseCellDepense(SetOfChantiers As SetOfChantiers, IdxChantiers As Integer, IdxDepense As Integer, newRange As Range)
    Dim Chantiers() As Chantier
    Dim ChantierTmp As Chantier
    Dim DepensesTmp() As DepenseChantier
    Dim TmpDepense As DepenseChantier
    
    Chantiers = SetOfChantiers.Chantiers
    ChantierTmp = Chantiers(IdxChantiers)
    DepensesTmp = ChantierTmp.Depenses
    TmpDepense = DepensesTmp(IdxDepense)
    Set TmpDepense.BaseCell = newRange
    DepensesTmp(IdxDepense) = TmpDepense
    ChantierTmp.Depenses = DepensesTmp
    Chantiers(IdxChantiers) = ChantierTmp
    SetOfChantiers.Chantiers = Chantiers
End Sub

Public Function getDefaultSetOfChantiers(NBChantiers As Integer, NbDefaultDepenses As Integer) As SetOfChantiers

    Dim newArray() As Chantier
    Dim SetOfChantiers As SetOfChantiers
    Dim idx As Integer
    
    ReDim newArray(1 To NBChantiers)
    
    For idx = 1 To NBChantiers
        newArray(idx) = getDefaultChantier(NbDefaultDepenses)
    Next idx
    SetOfChantiers.Chantiers = newArray
    getDefaultSetOfChantiers = SetOfChantiers

End Function

Public Function getChargesDefault(NB As Integer) As SetOfCharges

    Dim SetOfCharges As SetOfCharges
    Dim Charges() As Charge
    ReDim Charges(0)
    SetOfCharges.Charges = Charges
    
    getChargesDefault = getChargesDefaultPreserve(SetOfCharges, NB)
    
End Function

Public Function getChargesDefaultPreserve(PreviousSetOfCharges As SetOfCharges, NB As Integer) As SetOfCharges

    Dim PreviousCharges() As Charge
    Dim Charges() As Charge
    Dim SetOfCharges As SetOfCharges
    Dim Index As Integer
    ReDim Charges(1 To NB)
    
    PreviousCharges = PreviousSetOfCharges.Charges
    For Index = 1 To NB
        If Index <= UBound(PreviousCharges) Then
            Charges(Index) = PreviousCharges(Index)
        Else
            Charges(Index) = getDefaultCharge()
        End If
    Next Index
    
    SetOfCharges.Charges = Charges
    getChargesDefaultPreserve = SetOfCharges
    
End Function

Public Sub InsererUneDepenseInternal()
    
    Dim ChantierSheet As Worksheet
    Dim NBChantiers As Integer
    Dim NBSalariesAndCat As Integer
    Dim Previous As Integer
    Dim SetOfRange As SetOfRange
    
    Set ChantierSheet = ThisWorkbook.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If
    SetOfRange = GetCellsForDepense(ChantierSheet)
    If Not SetOfRange.Status Then
        Exit Sub
    End If
    
    SetSilent

    NBChantiers = GetNbChantiers(ThisWorkbook)
    Previous = SetOfRange.ResultCell.Row - SetOfRange.HeadCell.Row - 1

    InsertRows _
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
    UpdateSumsByColumn _
        Range( _
            SetOfRange.EndCell.Cells(2, 2), _
            SetOfRange.ResultCell.Cells(0, 1 + NBChantiers) _
        ), _
        SetOfRange.ResultCell.Cells(1, 2), _
        Previous + NBSalariesAndCat
    
    SetActive
    
End Sub

Public Sub RenewFormulaForTotalFinancement( _
        ChantierSheet As Worksheet, _
        NBChantiers As Integer _
    )
    Dim Formula As String
    Dim IndexChantier As Integer
    Dim IndexLigne As Integer
    Dim NBRowsFinancements As Integer
    Dim SetOfRange As SetOfRange

    SetOfRange = GetCellsForFinancement(ChantierSheet)
    If SetOfRange.Status Then
        NBRowsFinancements = SetOfRange.EndCell.Row - SetOfRange.HeadCell.Row
        For IndexChantier = 1 To NBChantiers
            Formula = "=0"
            For IndexLigne = 1 To NBRowsFinancements
                If SetOfRange.HeadCell.Cells(1 + IndexLigne, 2).Value <> "Statut" Then
                    Formula = Formula & "+" & _
                        SetOfRange.HeadCell.Cells(1 + IndexLigne, 2 + IndexChantier) _
                            .address(False, False, xlA1, False)
                End If
            Next IndexLigne
            SetOfRange.ResultCell.Cells(1, 1 + IndexChantier).Formula = Formula
        Next IndexChantier
    End If
End Sub
Public Function GetCellsForFinancement( _
        ChantierSheet As Worksheet _
    ) As SetOfRange

    Dim SetOfRange As SetOfRange

    SetOfRange.Status = False
    Set SetOfRange.ChantierSheet = ChantierSheet
    Set SetOfRange.HeadCell = ChantierSheet.Cells(1, 1).EntireColumn.Find(Label_Type_Financeur)
    If Not (SetOfRange.HeadCell Is Nothing) Then
        Set SetOfRange.EndCell = ChantierSheet.Cells(1, 2).EntireColumn.Find(Label_Total_Financements)
        If Not (SetOfRange.EndCell Is Nothing) Then
            Set SetOfRange.ResultCell = SetOfRange.EndCell
            Set SetOfRange.EndCell = SetOfRange.EndCell.Cells(0, 0)
            SetOfRange.Status = True
        End If
    End If
    GetCellsForFinancement = SetOfRange
End Function

Public Function GetCellsForDepense( _
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
    Set SetOfRange.ChantierSheet = ChantierSheet

    Set BaseCell = FindNextNotEmpty(ChantierSheet.Cells(3, 1), False)
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

    GetCellsForDepense = SetOfRange
End Function

Public Sub RetirerUnFinanceur()

    Dim ChantierSheet As Worksheet
    Dim CurrentRange As Range
    Dim NewLine As Integer
    Dim NBChantier As Integer
    Dim SetOfRange As SetOfRange
    Dim ValueToTest As String
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If

    SetOfRange = GetCellsForFinancement(ChantierSheet)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    NBChantier = GetNbChantiers(wb)

    NewLine = InputLineBetween( _
        "Supprimer le financement de la ligne ?", _
        "Ligne du financement à supprimer", _
        SetOfRange.HeadCell.Row + 1, _
        SetOfRange.EndCell.Row _
    )

    If NewLine = 0 Then
        MsgBox "La ligne entrée n'est pas la ligne d'un financement"
        Exit Sub
    End If
    
    SetSilent

    Set CurrentRange = SetOfRange.HeadCell.Cells(NewLine - SetOfRange.HeadCell.Row + 1, 1)
    ValueToTest = CurrentRange.Cells(1, 2).Value
    If ValueToTest = "Statut" Then
        RemoveRows _
            SetOfRange.HeadCell, _
            NewLine - SetOfRange.HeadCell.Row, _
            NewLine - SetOfRange.HeadCell.Row - 2, _
            1 + NBChantier + NBExtraCols
    Else
        ValueToTest = CurrentRange.Cells(2, 2).Value
        If ValueToTest = "Statut" Then
            RemoveRows _
                SetOfRange.HeadCell, _
                NewLine - SetOfRange.HeadCell.Row + 1, _
                NewLine - SetOfRange.HeadCell.Row - 1, _
                1 + NBChantier + NBExtraCols
        Else
            RemoveRows _
                SetOfRange.HeadCell, _
                NewLine - SetOfRange.HeadCell.Row, _
                NewLine - SetOfRange.HeadCell.Row - 1, _
                1 + NBChantier + NBExtraCols
        End If
    End If

    RenewFormulaForTotalFinancement ChantierSheet, NBChantier

    SetActive
End Sub

Public Sub RetirerUneDepense()

    Dim ChantierSheet As Worksheet
    Dim CurrentRange As Range
    Dim NewLine As Integer
    Dim NBChantier As Integer
    Dim SetOfRange As SetOfRange
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        Exit Sub
    End If

    SetOfRange = GetCellsForDepense(ChantierSheet)
    If Not SetOfRange.Status Then
        Exit Sub
    End If

    NBChantier = GetNbChantiers(wb)

    NewLine = InputLineBetween( _
        "Supprimer la dépense de la ligne ?", _
        "Ligne de la dépense à supprimer", _
        SetOfRange.HeadCell.Row + 1, _
        SetOfRange.ResultCell.Row - 1 _
    )

    If NewLine = 0 Then
        MsgBox "La ligne entrée n'est pas la ligne d'une dépense"
        Exit Sub
    End If
    
    SetSilent

    Set CurrentRange = SetOfRange.HeadCell.Cells(NewLine - SetOfRange.HeadCell.Row + 1, 1)

    RemoveRows _
        SetOfRange.HeadCell, _
        NewLine - SetOfRange.HeadCell.Row, _
        NewLine - SetOfRange.HeadCell.Row - 1, _
        1 + NBChantier + NBExtraCols

    SetActive
End Sub

Public Function InputLineBetween( _
    Message As String, _
    Title As String, _
    MinLine As Integer, _
    MaxLine As Integer _
    ) As Integer

    Dim FormatValue As Integer
    Dim Value
    Value = InputBox(Message, Title, MaxLine)

    InputLineBetween = 0
    If Value > 0 Then
        FormatValue = CInt(Value)
        If FormatValue <= MaxLine _
            And FormatValue >= MinLine Then
            InputLineBetween = FormatValue
        End If
    End If
End Function


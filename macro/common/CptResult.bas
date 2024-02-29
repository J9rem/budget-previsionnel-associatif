Attribute VB_Name = "CptResult"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

Public Function BudgetGlobal_Depenses_Add_From_Chantiers( _
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
            Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False)
            CurrentCell.Value = ""
            CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 0).address(False, False, xlA1, True))
            CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True))
        End If
    Next Index
    Set BudgetGlobal_Depenses_Add_From_Chantiers = CurrentCell
End Function

Public Function BudgetGlobal_Depenses_Add_From_Charges( _
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
            
            Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False)
            CurrentCell.Value = ""
            CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(currentCharge.ChargeCell.address(False, False, xlA1, True))
            ' Be carefull to the number of columns if a 'charges' coles is added
            CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(currentCharge.ChargeCell.Cells(1, 4).address(False, False, xlA1, True))
        End If
    Next Index
    Set BudgetGlobal_Depenses_Add_From_Charges = CurrentCell
End Function

Public Function BudgetGlobal_Depenses_Add_Header(BaseCell As Range, CodeValue As Integer, CodeIndex As Integer) As Range
    Dim CurrentCell As Range
    Dim NomTypeCharge As TypeCharge

    Set CurrentCell = BudgetGlobal_InsertLineAndFormat(BaseCell, BaseCell, True)
    CurrentCell.Value = CodeValue

    NomTypeCharge = TypesDeCharges().Values(CodeIndex)
    CurrentCell.Cells(1, 2).Value = NomTypeCharge.Nom
    CurrentCell.Cells(1, 3).Value = 0

    Set BudgetGlobal_Depenses_Add_Header = CurrentCell
End Function

Public Function BudgetGlobal_Depenses_Add(wb As Workbook, Data As Data, BaseCell As Range) As Range
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
            Set HeadCell = BudgetGlobal_Depenses_Add_Header(CurrentCell, CodeValue, CodeIndex)
            TotalCell.Cells(1, 3).Formula = TotalCell.Cells(1, 3).Formula _
                & "+" _
                & CleanAddress(HeadCell.Cells(1, 3).address(False, False, xlA1))
            Set CurrentCell = BudgetGlobal_Depenses_Add_From_Charges(Data, HeadCell, HeadCell, CodeIndex)
            Set CurrentCell = BudgetGlobal_Depenses_Add_From_Chantiers(Data, CurrentCell, HeadCell, CodeValue)

            If CodeValue = 64 Then
                ' ajouter les depenses de personnel
                Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False)
                CurrentCell.Value = ""
                CurrentCell.Cells(1, 2).Value = T_Salary
                CurrentCell.Cells(1, 2).Font.Bold = True
                CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                    BudgetGlobal_Depenses_SearchRangeForEmployeesSalary(wb).address(False, False, xlA1, True) _
                    ) & "/1.5"
                Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False)
                CurrentCell.Value = ""
                CurrentCell.Cells(1, 2).Value = "Charges sociales"
                CurrentCell.Cells(1, 2).Font.Bold = True
                CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                        BudgetGlobal_Depenses_SearchRangeForEmployeesSalary(wb).address(False, False, xlA1, True)) _
                        & "-" & CleanAddress(CurrentCell.Cells(0, 3).address(False, False, xlA1, False) _
                    )
            End If

            ' set sum
            If CurrentCell.Row > HeadCell.Row Then
                HeadCell.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCell.Cells(2, 3), CurrentCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
            End If
        End If
    Next CodeValue

    Set BudgetGlobal_Depenses_Add = CurrentCell
End Function

Public Sub BudgetGlobal_Depenses_Clean(BaseCell)
    Dim Anchor As String

    Anchor = "Total "

    ' remove others lines and leave one formatted
    While Left(BaseCell.Cells(2, 1).Value, Len(Anchor)) <> Anchor
        Range(BaseCell.Cells(2, 1), BaseCell.Cells(2, 3)).Delete Shift:=xlShiftUp
    Wend
End Sub

Public Function BudgetGlobal_Depenses_SearchRangeForEmployeesSalary(wb As Workbook) As Range
    Dim CoutJSalaireSheet As Worksheet
    Dim BaseCell As Range
    
    Set BaseCell = Nothing
    
    Set CoutJSalaireSheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    If CoutJSalaireSheet Is Nothing Then
        GoTo EndFunction
    End If
    
    Set BaseCell = CoutJSalaireSheet.Cells.Find(Replace(T_Amout_Salary_of_WorkingPeople, "%n%", Chr(10)))
    If BaseCell Is Nothing Then
        GoTo EndFunction
    End If
    Set BaseCell = BaseCell.Cells(1, 2)
    
    
EndFunction:
    Set BudgetGlobal_Depenses_SearchRangeForEmployeesSalary = BaseCell
End Function

Public Sub BudgetGlobal_EgaliserLesColonnes(ws As Worksheet)

    Dim EndFirstCol As Range
    Dim EndSecondCol As Range
    Dim Ecart As Integer
    Dim Index As Integer
    Dim BaseCell As Range
    Dim HeadCell
    
    Set EndFirstCol = ws.Cells.Find(T_Total_Charges & " (1) + (2)")
    Set EndSecondCol = ws.Cells.Find("Total Financements (1) + (2)+ (3)")
    Ecart = EndFirstCol.Row - EndSecondCol.Row
    
    If Ecart > 0 Then
        Set BaseCell = ws.Cells(1, 5).EntireColumn.Find(75).Cells(0, 1)
    Else
        Set BaseCell = ws.Cells.Find(T_Total_Charges & " (1)").Cells(0, 1)
        Ecart = -Ecart
    End If
    
    For Index = 1 To Ecart
        Set BaseCell = BudgetGlobal_InsertLineAndFormat(BaseCell, BaseCell.Cells(-1, 1), False)
    Next Index
    
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
    Next Index
    
End Sub

Public Function BudgetGlobal_Financements_Add(wb As Workbook, Data As Data, StartCell As Range) As Boolean

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
            Set BaseCell = BudgetGlobal_InsertLineAndFormat(BaseCell, HeadCell, False)
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
        BudgetGlobal_Financements_Add = False
        Exit Function
    End If
    Set HeadCell = BaseCell
    HeadCell.Cells(1, 3).Formula = "=0"
    
    TypesFinancements = TypeFinancementsFromWb(wb)
    
    For IndexTypeFinancement = 1 To UBound(TypesFinancements)
        Set BaseCell = BudgetGlobal_InsertLineAndFormat(BaseCell, HeadCell, False)
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
                Set BaseCell = BudgetGlobal_InsertLineAndFormat(BaseCell, HeadCellFinancement, False)
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
    BudgetGlobal_Financements_Add = True
End Function

Public Function BudgetGlobal_InsertLineAndFormat(BaseCell As Range, HeadCell As Range, IsHeader As Boolean) As Range

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
    
    Set BudgetGlobal_InsertLineAndFormat = BaseCell
End Function


' Macro pour mettre a jour le budget update
Public Sub CptResult_Update(wb As Workbook)

    Dim Data As Data
    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim rev As WbRevision
        
    SetSilent
    
    rev = DetecteVersion(wb)
    Data = Extract_Data_From_Table(wb, rev)
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_CptResult_prefix & Nom_Feuille_CptResult_suffix)
    If CurrentSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_CptResult_prefix & Nom_Feuille_CptResult_suffix)
        GoTo EndSub
    End If
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers)
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
    BudgetGlobal_Depenses_Clean BaseCell
    BudgetGlobal_Depenses_Add wb, Data, BaseCell
    
    ' Produits
    Set BaseCell = CurrentSheet.Cells(1, 1).EntireColumn.Find("Compte")
    If BaseCell Is Nothing Then
        GoTo EndSub
    End If
    Set BaseCell = BaseCell.Cells(1, 5)
    While BaseCell.Value = "" Or BaseCell.Value <> 70
        Set BaseCell = BaseCell.Cells(2, 1)
    Wend

    If Not BudgetGlobal_Financements_Add(wb, Data, BaseCell) Then
        GoTo EndSub
    End If
    
    ' Egaliser la longueur des colonnes
    BudgetGlobal_EgaliserLesColonnes CurrentSheet
    
EndSub:
    Application.DisplayAlerts = True
    SetActive
    BaseCell.EntireRow.Cells(1, 1).EntireColumn.Cells(1, 1).Select

End Sub


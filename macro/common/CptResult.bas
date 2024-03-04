Attribute VB_Name = "CptResult"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

' insert depenses from Chantiers
' @param Data Data
' @param Range BaseCell
' @param Range HeadCell
' @param Integer CodeValue
' @param Boolean IsReal
' @param Range BaseCellRelative
' @param Boolean IsGlobal
' @param Boolean TestReal
' @return Range NewCurrentCell
Public Function BudgetGlobal_Depenses_Add_From_Chantiers( _
        Data As Data, _
        BaseCell As Range, _
        HeadCell As Range, _
        CodeValue As Integer, _
        IsReal As Boolean, _
        BaseCellRelative As Range, _
        IsGlobal As Boolean, _
        TestReal As Boolean _
    )

    Dim CanAdd As Boolean
    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim CurrentBaseCell As Range
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
            CanAdd = IsGlobal
            If Not CanAdd Then
                If IsReal Or TestReal Then
                    CanAdd = (Depense.BaseCell.Cells(1, 1 + NBChantiers).Value <> 0) _
                        And (Depense.BaseCellReal.Cells(1, 1 + 3 * NBChantiers).Value <> 0)
                Else
                    CanAdd = (Depense.BaseCell.Cells(1, 1 + NBChantiers).Value <> 0)
                End If
            End If
            
            If CanAdd Then
                Set CurrentBaseCell = CurrentCell
                Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False)
                CurrentCell.Value = ""
                ' TODO manage not IsGlobal
                If Not IsReal Then
                    CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 0).address(False, False, xlA1, True))
                    CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(Depense.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True))
                Else
                    CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(Depense.BaseCellReal.Cells(1, 0).address(False, False, xlA1, True))
                    CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                        Depense.BaseCellReal.Cells(1, 1 + 3 * NBChantiers).address(False, False, xlA1, True) _
                    )
                    
                    ' percent part
                    Set CurrentBaseCell = BudgetGlobal_InsertLineAndFormat( _
                        CurrentBaseCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                        HeadCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                        False, _
                        True _
                    )
                    
                    CurrentBaseCell.Cells(1, 2).Formula = "=" & CleanAddress(CurrentCell.Cells(1, 2).address(False, False, xlA1, True))
                    CurrentBaseCell.Cells(1, 3).Formula = CptResult_GetFormulaForPercent( _
                            CurrentCell.Cells(1, 3), _
                            BaseCellRelative.Cells(CurrentCell.Row - BaseCell.Row + 1, 3) _
                        )
                End If
            End If
        End If
    Next Index
    Set BudgetGlobal_Depenses_Add_From_Chantiers = CurrentCell
End Function

' add a depense from a charge
' @param Data Data
' @param Range BaseCell
' @param Range HeadCell
' @param Range BaseCellRelative only for IsReal
' @param Boolean IsReal
' @param Boolean IsGlobal
' @param Boolean TestReal
' @return Range NewCurrenCell
Public Function BudgetGlobal_Depenses_Add_From_Charges( _
        Data As Data, _
        BaseCell As Range, _
        HeadCell As Range, _
        BaseCellRelative As Range, _
        IndexFound As Integer, _
        IsReal As Boolean, _
        IsGlobal As Boolean, _
        TestReal As Boolean _
    )

    Dim CanAdd As Boolean
    Dim Charges() As Charge
    Dim currentCharge As Charge
    Dim CurrentBaseCell As Range
    Dim CurrentCell As Range
    Dim Index As Integer

    Set CurrentCell = BaseCell

    Charges = Data.Charges
    For Index = 1 To UBound(Charges)
        currentCharge = Charges(Index)
        
        If currentCharge.IndexTypeCharge = IndexFound Then
            CanAdd = IsGlobal
            If Not CanAdd Then
                If IsReal Or TestReal Then
                    CanAdd = (currentCharge.ChargeCell.Cells(1, 4).Value <> 0) _
                        And (currentCharge.ChargeCell.Cells(1, 5).Value <> 0)
                Else
                    CanAdd = (currentCharge.ChargeCell.Cells(1, 4).Value <> 0)
                End If
            End If

            If CanAdd Then
            
                Set CurrentBaseCell = CurrentCell
                Set CurrentCell = BudgetGlobal_InsertLineAndFormat(CurrentBaseCell, HeadCell, False)
                CurrentCell.Value = ""
                CurrentCell.Cells(1, 2).Formula = "=" & CleanAddress(currentCharge.ChargeCell.address(False, False, xlA1, True))
                ' TODO manage not IsGlobal
                If Not IsReal Then
                    ' Be carefull to the number of columns if a 'charges' cols is added
                    CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress(currentCharge.ChargeCell.Cells(1, 4).address(False, False, xlA1, True))
                Else
                    ' Be carefull to the number of columns if a 'charges' cols is added
                    CurrentCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                        currentCharge.ChargeCell.Cells(1, 5).address(False, False, xlA1, True) _
                    )
                    ' percent part
                    Set CurrentBaseCell = BudgetGlobal_InsertLineAndFormat( _
                        CurrentBaseCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                        HeadCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                        False, _
                        True _
                    )
                    
                    CurrentBaseCell.Cells(1, 2).Formula = "=" & CleanAddress(CurrentCell.Cells(1, 2).address(False, False, xlA1, True))
                    CurrentBaseCell.Cells(1, 3).Formula = CptResult_GetFormulaForPercent( _
                            CurrentCell.Cells(1, 3), _
                            BaseCellRelative.Cells(CurrentCell.Row - BaseCell.Row + 1, 3) _
                        )
                End If

            End If
        End If
    Next Index
    Set BudgetGlobal_Depenses_Add_From_Charges = CurrentCell
End Function

Public Function BudgetGlobal_Depenses_Add_Header( _
        BaseCell As Range, _
        CodeValue As Integer, _
        CodeIndex As Integer, _
        Optional IsPercent As Boolean = False _
    ) As Range
    Dim CurrentCell As Range
    Dim NomTypeCharge As TypeCharge

    Set CurrentCell = BudgetGlobal_InsertLineAndFormat(BaseCell, BaseCell, True, IsPercent)
    CurrentCell.Value = CodeValue

    NomTypeCharge = TypesDeCharges().Values(CodeIndex)
    CurrentCell.Cells(1, 2).Value = NomTypeCharge.Nom
    CurrentCell.Cells(1, 3).Value = 0

    Set BudgetGlobal_Depenses_Add_Header = CurrentCell
End Function

' Function that add a depense and return CurrentCell
' @param Workbook wb
' @param Data Data
' @param Range BaseCell
' @param Range BaseCellRelative only for IsReal
' @param Boolean IsReal
' @param Boolean IsGlobal
' @param Boolean TestReal
' @param Integer() ChantiersToAdd
' @return Range CurrentCell
Public Function BudgetGlobal_Depenses_Add( _
        wb As Workbook, _
        Data As Data, _
        BaseCell As Range, _
        BaseCellRelative As Range, _
        IsReal As Boolean, _
        IsGlobal As Boolean, _
        TestReal As Boolean, _
        ChantiersToAdd _
    ) As Range

    Dim CodeValue As Integer
    Dim CodeIndex As Integer
    Dim CurrentCell As Range
    Dim FirstLineCell As Range
    Dim HeadCell As Range
    Dim HeadCellPercent As Range
    Dim SecondLineCell As Range
    Dim StartCell As Range
    Dim TmpBaseCellRelative As Range
    Dim TmpBaseCellRelative2 As Range
    Dim TotalCell As Range

    Set TotalCell = BaseCell.Cells(2, 1)
    TotalCell.Cells(1, 3).Formula = "=0"

    Set CurrentCell = BaseCell.Cells(1, 1)

    For CodeValue = 60 To 69
        CodeIndex = FindTypeChargeIndexFromCode(CodeValue)
        If CodeIndex > 0 Then
            Set HeadCell = BudgetGlobal_Depenses_Add_Header(CurrentCell, CodeValue, CodeIndex)
            TotalCell.Cells(1, 3).Formula = TotalCell.Cells(1, 3).Formula _
                & "+" _
                & CleanAddress(HeadCell.Cells(1, 3).address(False, False, xlA1))
            If IsReal Then
                ' percent part
                Set HeadCellPercent = BudgetGlobal_Depenses_Add_Header( _
                    CurrentCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                    CodeValue, _
                    CodeIndex, _
                    True _
                )
                Set TmpBaseCellRelative = BaseCellRelative.Cells(HeadCell.Row - BaseCell.Row + 1, 1)
                Set TmpBaseCellRelative2 = BaseCellRelative
            Else
                Set TmpBaseCellRelative = Nothing
                Set TmpBaseCellRelative2 = Nothing
            End If
            Set CurrentCell = BudgetGlobal_Depenses_Add_From_Charges( _
                Data, HeadCell, HeadCell, _
                TmpBaseCellRelative, _
                CodeIndex, IsReal, IsGlobal, TestReal _
            )
            Set CurrentCell = BudgetGlobal_Depenses_Add_From_Chantiers( _
                Data, CurrentCell, HeadCell, CodeValue, IsReal, TmpBaseCellRelative2, IsGlobal, TestReal _
            )

            If CodeValue = 64 Then
                ' ajouter les depenses de personnel
                Set FirstLineCell = CptResult_Charges_Personal_Add(wb, CurrentCell, HeadCell, T_Salary, Nothing, Nothing)
                Set SecondLineCell = CptResult_Charges_Personal_Add(wb, FirstLineCell, HeadCell, "Charges sociales", FirstLineCell, Nothing)
                If IsReal Then
                    ' percent part
                    CptResult_Charges_Personal_Add _
                        wb, _
                        CurrentCell.Cells(1, 1 + Offset_NB_Cols_For_Percent_In_CptResultReal), _
                        HeadCell.Cells(1, 1 + Offset_NB_Cols_For_Percent_In_CptResultReal), _
                        T_Salary, Nothing, _
                        BaseCellRelative.Cells(HeadCell.Row - BaseCell.Row + 1, 1)
                    CptResult_Charges_Personal_Add _
                        wb, _
                        FirstLineCell.Cells(1, 1 + Offset_NB_Cols_For_Percent_In_CptResultReal), _
                        HeadCell.Cells(1, 1 + Offset_NB_Cols_For_Percent_In_CptResultReal), _
                        "Charges sociales", _
                        FirstLineCell.Cells(1, 1 + Offset_NB_Cols_For_Percent_In_CptResultReal), _
                        BaseCellRelative.Cells(HeadCell.Row - BaseCell.Row + 1, 1)
                End If
                Set CurrentCell = SecondLineCell
            End If

            ' set sum
            If CurrentCell.Row > HeadCell.Row Then
                HeadCell.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCell.Cells(2, 3), CurrentCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
                If IsReal Then
                    ' percent part
                    HeadCellPercent.Cells(1, 3).Formula = CptResult_GetFormulaForPercent( _
                        HeadCell.Cells(1, 3), _
                        BaseCellRelative.Cells(HeadCell.Row - BaseCell.Row + 1, 3) _
                    )
                End If
            End If
        End If
    Next CodeValue

    Set BudgetGlobal_Depenses_Add = CurrentCell
End Function

Public Sub BudgetGlobal_Depenses_Clean(BaseCell As Range, IsReal As Boolean)
    Dim Anchor As String

    Anchor = "Total "

    ' remove others lines and leave one formatted
    While Left(BaseCell.Cells(2, 1).Value, Len(Anchor)) <> Anchor
        CptResult_Clean_Lines BaseCell.Cells(2, 1), BaseCell.Cells(2, 3), IsReal
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

Public Sub BudgetGlobal_EgaliserLesColonnes(ws As Worksheet, IsReal As Boolean)

    Dim EndFirstCol As Range
    Dim EndSecondCol As Range
    Dim Ecart As Integer
    Dim Index As Integer
    Dim BaseCell As Range
    Dim BaseCellTmp As Range
    
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
        Set BaseCellTmp = BudgetGlobal_InsertLineAndFormat(BaseCell, BaseCell.Cells(-1, 1), False)
        ' manage percent
        If IsReal Then
            BudgetGlobal_InsertLineAndFormat _
                BaseCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                BaseCell.Cells(-1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
                False, _
                True
        End If
        Set BaseCell = BaseCellTmp
    Next Index
    
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
    Next Index
    
End Sub

' define financement
' @param Workbook wb
' @param Data Data
' @param RangeStartCell
' @param Range BaseCellRelative only for IsReal
' @param Boolean IsReal
' @param Boolean IsGlobal
' @param Boolean TestReal
' @param Integer() ChantiersToAdd
' @return Boolean All is right
Public Function BudgetGlobal_Financements_Add( _
        wb As Workbook, _
        Data As Data, _
        StartCell As Range, _
        BaseCellRelative As Range, _
        IsReal As Boolean, _
        IsGlobal As Boolean, _
        TestReal As Boolean, _
        ChantiersToAdd _
    ) As Boolean

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
            Set BaseCell = CptResult_Add_A_LineOfFinancement( _
                BaseCell, HeadCell, StartCell, IsReal, BaseCellRelative, Financement, NBChantiers)
        End If
    Next Index
    
    ' remove others lines and leave one formatted
    While BaseCell.Cells(2, 1).Value = ""
        CptResult_Clean_Lines BaseCell.Cells(2, 1), BaseCell.Cells(2, 3), IsReal
    Wend
    
    If BaseCell.Row > HeadCell.Row Then
        HeadCell.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCell.Cells(2, 3), BaseCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
        If IsReal Then
            HeadCell.Cells(1, 3 + Offset_NB_Cols_For_Percent_In_CptResultReal).Formula = _
                CptResult_GetFormulaForPercent( _
                    HeadCell.Cells(1, 3), _
                    BaseCellRelative.Cells(HeadCell.Row - StartCell.Row + 2, 3) _
                )
        End If
    End If
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
        If IsReal Then
            AddBottomBorder BaseCell.Cells(1, Index + Offset_NB_Cols_For_Percent_In_CptResultReal)
        End If
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
        
        If IsReal Then
            FormatFinancementCells BudgetGlobal_InsertLineAndFormat_Percent( _
                BaseCell, HeadCell, False, StartCell, BaseCellRelative.Cells(2, 1))
        End If

        ' TODO manage IsGlobal
        
        HeadCell.Cells(1, 3).Formula = HeadCell.Cells(1, 3).Formula _
            & "+" _
            & CleanAddress(BaseCell.Cells(1, 3).address(False, False, xlA1))
        Set HeadCellFinancement = BaseCell
        Chantiers = Data.Chantiers
        NBChantiers = UBound(Chantiers)
        Chantier = Chantiers(1)
        For Index = 1 To UBound(Chantier.Financements)
            Financement = Chantier.Financements(Index)
            If Financement.TypeFinancement = IndexTypeFinancement Then
                Set BaseCell = CptResult_Add_A_LineOfFinancement( _
                    BaseCell, HeadCellFinancement, StartCell, IsReal, BaseCellRelative, Financement, NBChantiers)
            End If
        Next Index
        If BaseCell.Row > HeadCellFinancement.Row Then
            HeadCellFinancement.Cells(1, 3).Formula = "=SUM(" & CleanAddress(Range(HeadCellFinancement.Cells(2, 3), BaseCell.Cells(1, 3)).address(False, False, xlA1)) & ")"
        End If
    Next IndexTypeFinancement
    
    ' remove others lines and leave one formatted
    While BaseCell.Cells(2, 1).Value = ""
        CptResult_Clean_Lines BaseCell.Cells(2, 1), BaseCell.Cells(2, 3), IsReal
    Wend
    
    For Index = 1 To 3
        AddBottomBorder BaseCell.Cells(1, Index)
        If IsReal Then
            AddBottomBorder BaseCell.Cells(1, Index + Offset_NB_Cols_For_Percent_In_CptResultReal)
        End If
    Next Index
    BudgetGlobal_Financements_Add = True
End Function

Public Function BudgetGlobal_InsertLineAndFormat( _
        BaseCellParam As Range, _
        HeadCell As Range, _
        IsHeader As Boolean, _
        Optional IsPercent As Boolean = False _
    ) As Range

    Dim BaseCell As Range
    Dim Index As Integer
    
    Set BaseCell = BaseCellParam.Cells(1, 1)

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
    SetFormatForBudget BaseCell, HeadCell, IsHeader, IsPercent
    
    Set BudgetGlobal_InsertLineAndFormat = BaseCell
End Function

' add line and format for percent
' @param Range BaseCell
' @param Range HeadCell
' @param Boolean IsHeader
' @param Range BaseCellRelative
' @param Range StartCell
' @return Range NewBaseCellPercent
Public Function BudgetGlobal_InsertLineAndFormat_Percent( _
        BaseCell As Range, _
        HeadCell As Range, _
        IsHeader As Boolean, _
        StartCell As Range, _
        BaseCellRelative As Range _
    ) As Range

    Dim BaseCellPercent As Range

    Set BaseCellPercent = BudgetGlobal_InsertLineAndFormat( _
            BaseCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
            HeadCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
            IsHeader, _
            True _
        )
    If BaseCell.Value = "" Then
        BaseCellPercent.Value = ""
    End If
    BaseCellPercent.Cells(1, 2).Formula = "=" & CleanAddress( _
        BaseCell.Cells(BaseCellPercent.Row - BaseCell.Row + 1, 2).address(False, False, xlA1, False) _
    )
    BaseCellPercent.Cells(1, 3).Formula = CptResult_GetFormulaForPercent( _
        BaseCell.Cells(BaseCellPercent.Row - BaseCell.Row + 1, 3), _
        BaseCellRelative.Cells(BaseCellPercent.Row - StartCell.Row + 1, 3) _
    )
    Set BudgetGlobal_InsertLineAndFormat_Percent = BaseCellPercent
End Function

' add a line of financement
' @param Range BaseCellParam
' @param Range HeadCell
' @param Range StartCell
' @param Boolean IsReal
' @param Range BaseCellRelative
' @param Financement Financement
' @param Integer NBChantiers
' @return Range NewBaseCell
Public Function CptResult_Add_A_LineOfFinancement( _
        BaseCellParam As Range, _
        HeadCell As Range, _
        StartCell As Range, _
        IsReal As Boolean, _
        BaseCellRelative As Range, _
        Financement As Financement, _
        NBChantiers As Integer _
    ) As Range

    Dim BaseCell As Range

    Set BaseCell = BaseCellParam
    Set BaseCell = BudgetGlobal_InsertLineAndFormat(BaseCell, HeadCell, False)
    If IsReal Then
        BaseCell.Cells(1, 2).Formula = "=" & CleanAddress( _
            Financement.BaseCellReal.Cells(1, 0).address(False, False, xlA1, True) _
        )
        BaseCell.Cells(1, 3).Formula = "=" & CleanAddress( _
            Financement.BaseCellReal.Cells(1, 1 + 3 * NBChantiers).address(False, False, xlA1, True) _
        )
        BudgetGlobal_InsertLineAndFormat_Percent _
            BaseCellParam, HeadCell, False, StartCell, BaseCellRelative.Cells(2, 1)
    Else
        BaseCell.Cells(1, 2).Formula = "=" & CleanAddress( _
            Financement.BaseCell.Cells(1, 0).address(False, False, xlA1, True) _
        )
        BaseCell.Cells(1, 3).Formula = "=" & CleanAddress( _
            Financement.BaseCell.Cells(1, 1 + NBChantiers).address(False, False, xlA1, True) _
        )
    End If
    Set CptResult_Add_A_LineOfFinancement = BaseCell
End Function

' add a personal depense for charge in CptResult
' @param Workbook wb
' @param Range CurrentCell
' @param Range HeadCell
' @param String Name
' @param Range FirstLineCell if second line, put Nothing otherwise
' @param Range HeadCellRelative if percent, put Nothing otherwise
' @return Range CurrentCell
Public Function CptResult_Charges_Personal_Add( _
    wb As Workbook, _
    CurrentCell As Range, _
    HeadCell As Range, _
    Name As String, _
    FirstLineCell As Range, _
    HeadCellRelative As Range _
) As Range

    Dim WorkingCell As Range

    Set WorkingCell = BudgetGlobal_InsertLineAndFormat(CurrentCell, HeadCell, False, Not (HeadCellRelative Is Nothing))
    WorkingCell.Value = ""
    WorkingCell.Cells(1, 2).Value = Name
    WorkingCell.Cells(1, 2).Font.Bold = True
    If HeadCellRelative Is Nothing Then
        If FirstLineCell Is Nothing Then
            WorkingCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                BudgetGlobal_Depenses_SearchRangeForEmployeesSalary(wb).address(False, False, xlA1, True) _
                ) & "/1.5"
        Else
            WorkingCell.Cells(1, 3).Formula = "=" & CleanAddress( _
                    FirstLineCell.Cells(1, 3).address(False, False, xlA1, False) _
                ) & "*0.5"
        End If
    Else
        WorkingCell.Cells(1, 3).Formula = _
            CptResult_GetFormulaForPercent( _
                WorkingCell.Cells(1, 3 - Offset_NB_Cols_For_Percent_In_CptResultReal), _
                HeadCellRelative.Cells(WorkingCell.Row - HeadCell.Row + 1, 3) _
            )
    End If
    Set CptResult_Charges_Personal_Add = WorkingCell
End Function

' clean lines and if needed percent lines
' @param Range FirstCell
' @param Range LastCell
' @param Boolean IsReal
Public Sub CptResult_Clean_Lines( _
        FirstCell As Range, _
        LastCell As Range, _
        IsReal As Boolean _
    )
    If IsReal Then
        ' clean before to preserve references
        Range( _
            FirstCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1), _
            LastCell.Cells(1, Offset_NB_Cols_For_Percent_In_CptResultReal + 1) _
        ).Delete Shift:=xlShiftUp
    End If
    Range(FirstCell, LastCell).Delete Shift:=xlShiftUp
End Sub

' Function to append a new value in array of integer
' @param Integer() Values
' @param Integer Value
' @return Integer() Values
Public Function CptResult_AppendInArray(Values, Value)

    Dim FormatedValue As Integer
    Dim Index As Integer
    Dim WorkingArray() As Integer

    ReDim WorkingArray(0 To 0)
    WorkingArray(0) = 0

    CptResult_AppendInArray = WorkingArray

    FormatedValue = CInt(Value)
    If Not IsArray(Values) Then
        Exit Function
    End If
    If Not inArray(FormatedValue, Values) Then
        ReDim WorkingArray(0 To (UBound(Values) + 1))
        For Index = 0 To UBound(Values)
            WorkingArray(Index) = Values(Index)
        Next Index
        WorkingArray(UBound(Values) + 1) = FormatedValue
    End If
    CptResult_AppendInArray = WorkingArray
End Function

' Function to get formula to calculate CptResult
' @param Range BaseCell in CptResult sheet
' @param Integer NBChantiers
' @return Integer() ChantiersToAdd Base 0 with ChantiersToAdd(0) = 0 if error
Public Function CptResult_GetChantiersToAdd(BaseCell As Range, NBChantiers As Integer)

    Dim CellWhereExpectedFormula As Range
    Dim ExtractedValue As String
    Dim OutputArray() As Integer

    ReDim OutputArray(0)
    OutputArray(0) = 0

    CptResult_GetChantiersToAdd = OutputArray

    If BaseCell Is Nothing Then
        Exit Function
    End If

    Set CellWhereExpectedFormula = BaseCell.Cells(-2, Offset_NB_Cols_For_Percent_In_CptResultReal)
    If CellWhereExpectedFormula.Value = "" Then
        Exit Function
    End If
    ExtractedValue = CStr(CellWhereExpectedFormula.Value)
    CptResult_GetChantiersToAdd = CptResult_ValidateFormula(ExtractedValue, NBChantiers)

End Function

Public Function CptResult_FindEndOfHeaderTable(BaseCell As Range) As Range

    Dim WorkingCell As Range

    If BaseCell Is Nothing Then
        Set CptResult_FindEndOfHeaderTable = Nothing
    End If
    Set WorkingCell = BaseCell.Cells(2, 1)
    While WorkingCell.Row < 1000 And ( _
            WorkingCell.Value = "" _
            Or Len(WorkingCell.Value) = 0 _
            Or CInt(WorkingCell.Value) < 60 _
            Or CInt(WorkingCell.Value) > 69 _
        )
        Set WorkingCell = WorkingCell.Cells(2, 1)
    Wend
    
    Set CptResult_FindEndOfHeaderTable = WorkingCell.Cells(0, 1)
End Function

' generat formula for percent
' @param Range BaseCell
' @param Range BaseCellRelative
' @return String
Public Function CptResult_GetFormulaForPercent( _
    BaseCell As Range, _
    BaseCellRelative As Range _
    ) As String

    CptResult_GetFormulaForPercent = "=" & CleanAddress( _
            BaseCell.address(False, False, xlA1, False) _
        ) & "/(" & CleanAddress( _
            BaseCellRelative.address(False, False, xlA1, True) _
        ) & "+1E-9)"
End Function

Public Function CptResult_IsReal(PageName As String) As Boolean

    CptResult_IsReal = (Left(PageName, Len(Nom_Feuille_CptResult_Real_prefix)) = Nom_Feuille_CptResult_Real_prefix)
End Function

Public Function CptResult_IsValidatedPageName(PageName As String) As Boolean

    CptResult_IsValidatedPageName = ( _
        Left(PageName, Len(Nom_Feuille_CptResult_prefix)) = Nom_Feuille_CptResult_prefix _
        Or CptResult_IsReal(PageName) _
    )
End Function

' Test if formula is validate and return clean one if asked
' @param String ExtractedFormula
' @param Integer NBChantiers
' @return Integer() ChantiersToAdd Base 0 with ChantiersToAdd(0) = 0 if error
Public Function CptResult_ValidateFormula( _
        ExtractedFormula As String, _
        NBChantiers As Integer _
    )

    Dim Index As Integer
    Dim IndexL2 As Integer
    Dim OutputArray() As Integer
    Dim TmpValue As String
    Dim SecondLevelValues() As String
    Dim Values() As String

    ReDim OutputArray(0)
    OutputArray(0) = 0

    CptResult_IsValidatedFormula = OutputArray

    If ExtractedFormula = "" Then
        Exit Function
    End If

    Values = Split(ExtractedFormula, ",")

    If Not IsArray(Values) Then
        Exit Function
    End If
    If UBound(Values) < 0 Then
        Exit Function
    End If

    ' -1 = all is right
    OutputArray(0) = -1

    For Index = 0 To UBound(Values)
        TmpValue = Trim(Values(Index))
        If InStr(TmpValue, "-") Then
            SecondLevelValues = Split(TmpValue, "-")
            If IsArray(SecondLevelValues) _
                And UBound(SecondLevelValues) = 1 _
                And CInt(SecondLevelValues(0)) <= CInt(SecondLevelValues(1)) _
                And CInt(SecondLevelValues(0)) >= 1 _
                And CInt(SecondLevelValues(1)) >= 1 _
                And CInt(SecondLevelValues(0)) <= NBChantiers Then
                If CInt(SecondLevelValues(1)) <= NBChantiers Then
                    For IndexL2 = CInt(SecondLevelValues(0)) To CInt(SecondLevelValues(1))
                        OutputArray = CptResult_AppendInArray(OutputArray, IndexL2)
                    End If
                Else
                    OutputArray(0) = 0
                    For IndexL2 = CInt(SecondLevelValues(0)) To NBChantiers
                        OutputArray = CptResult_AppendInArray(OutputArray, IndexL2)
                    End If
                End If
            Else
                OutputArray(0) = 0
            End If
        Else
            If CInt(TmpValue) >= 1 And CInt(TmpValue) <= NBChantiers Then
                OutputArray = CptResult_AppendInArray(OutputArray, CInt(TmpValue))
            Else
                OutputArray(0) = 0
            End If
        End If
    Next Index

    CptResult_IsValidatedFormula = OutputArray
End Function


' Macro pour mettre a jour le budget update
Public Sub CptResult_Update(wb As Workbook)

    Dim CurrentActiveSheet As Worksheet

    Set CurrentActiveSheet = wb.ActiveSheet

    CptResult_Update_ForASheet wb, CurrentActiveSheet.Name

End Sub

' Function that update content of CptResult
' @param Workbook wb
' @param String PageName
' @param Boolean TestReal = False
' @return Boolean False If Error
Public Function CptResult_Update_ForASheet( _
        wb As Workbook, _
        PageName As String, _
        Optional TestReal As Boolean = False _
    ) As Boolean

    Dim BaseCell As Range
    Dim ChantierSheet As Worksheet
    Dim ChantierSheetReal As Worksheet
    Dim ChantiersToAdd() As Integer
    Dim CurrentActiveSheet As Worksheet
    Dim CurrentSheet As Worksheet
    Dim Data As Data
    Dim EndOfHeaderCell As Range
    Dim EndOfHeaderCellRelative As Range
    Dim IsGlobal As Boolean
    Dim IsReal As Boolean
    Dim NBChantiers As Integer
    Dim RelativeSheet As Worksheet
    Dim rev As WbRevision
    Dim Suffix As String
        
    SetSilent

    CptResult_Update_ForASheet = False
    Set CurrentActiveSheet = wb.ActiveSheet

    If Not CptResult_IsValidatedPageName(PageName) Then
        GoTo EndCptResultUpdateForASheet
    End If

    IsReal = CptResult_IsReal(PageName)
    If IsReal Then
        Suffix = Mid(PageName, Len(Nom_Feuille_CptResult_Real_prefix) + 1)
        Set RelativeSheet = wb.Worksheets(Nom_Feuille_CptResult_prefix & Suffix)
        If RelativeSheet Is Nothing Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_CptResult_prefix & Suffix)
            GoTo EndCptResultUpdateForASheet
        End If
        ' update relative sheet
        If Not CptResult_Update_ForASheet(wb, Nom_Feuille_CptResult_prefix & Suffix, True) Then
            GoTo EndCptResultUpdateForASheet
        End If
        SetSilent
    Else
        Suffix = Mid(PageName, Len(Nom_Feuille_CptResult_prefix) + 1)
        Set RelativeSheet = Nothing
    End If
    IsGlobal = (Suffix = Nom_Feuille_CptResult_suffix)

    rev = DetecteVersion(wb)
    NBChantiers = GetNbChantiers(wb)
    Data = Extract_Data_From_Table(wb, rev)

    ' === find sheets ====
    Set CurrentSheet = wb.Worksheets(PageName)
    If CurrentSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", PageName)
        GoTo EndCptResultUpdateForASheet
    End If
    Set ChantierSheet = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    If ChantierSheet Is Nothing Then
        MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers)
        GoTo EndCptResultUpdateForASheet
    End If
    If IsReal Then
        Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
        If ChantierSheetReal Is Nothing Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Budget_chantiers_realise)
            GoTo EndCptResultUpdateForASheet
        End If
    Else
        Set ChantierSheetReal = Nothing
    End If
    
    Set BaseCell = CurrentSheet.Cells(1, 1).EntireColumn.Find("Compte")
    Set EndOfHeaderCell = CptResult_FindEndOfHeaderTable(BaseCell)
    If EndOfHeaderCell Is Nothing Then
        GoTo EndCptResultUpdateForASheet
    End If
    If IsReal Then
        Set EndOfHeaderCellRelative = CptResult_FindEndOfHeaderTable( _
            RelativeSheet.Cells(1, 1).EntireColumn.Find("Compte") _
        )
        If EndOfHeaderCellRelative Is Nothing Then
            GoTo EndCptResultUpdateForASheet
        End If
    Else
        Set EndOfHeaderCellRelative = Nothing
    End If

    If IsGlobal Then
        ReDim ChantiersToAdd(0 To 0)
        ChantiersToAdd(0) = 0
    Else
        ChantiersToAdd = CptResult_GetChantiersToAdd(BaseCell, NBChantiers)
        If UBound(ChantiersToAdd) = 0 Then
            If Not IsReal Then
                MsgBox Replace( _
                    T_Error_Formula_In_CptResult, _
                    "%adr%", _
                    BaseCell.Cells(-2, Offset_NB_Cols_For_Percent_In_CptResultReal).address(False, False, xlA1, True) _
                )
            End If
            ReDim ChantiersToAdd(0 To 1)
            ChantiersToAdd(0) = -1
            ChantiersToAdd(1) = 1
        End If
        ' TODO define rate for charges based on worked days for chantiers / NB tot worked
    End If

    BudgetGlobal_Depenses_Clean EndOfHeaderCell, IsReal
    BudgetGlobal_Depenses_Add _
        wb, _
        Data, _
        EndOfHeaderCell, _
        EndOfHeaderCellRelative, _
        IsReal, _
        IsGlobal, _
        TestReal, _
        ChantiersToAdd
    
    ' Produits
    Set EndOfHeaderCell = BaseCell.Cells(1, 5)
    While EndOfHeaderCell.Value = "" Or EndOfHeaderCell.Value <> 70
        Set EndOfHeaderCell = EndOfHeaderCell.Cells(2, 1)
    Wend
    
    If IsReal Then
        Set EndOfHeaderCellRelative = EndOfHeaderCellRelative.Cells(1, 5)
    End If

    If Not BudgetGlobal_Financements_Add( _
        wb, Data, EndOfHeaderCell, EndOfHeaderCellRelative, _
        IsReal, IsGlobal, TestReal, ChantiersToAdd) Then
        GoTo EndCptResultUpdateForASheet
    End If
    
    ' Egaliser la longueur des colonnes
    BudgetGlobal_EgaliserLesColonnes CurrentSheet, IsReal
    CptResult_Update_ForASheet = True
    
EndCptResultUpdateForASheet:
    CurrentActiveSheet.Activate
    CurrentActiveSheet.Cells(1, 1).Select
    Application.DisplayAlerts = True
    SetActive

End Function


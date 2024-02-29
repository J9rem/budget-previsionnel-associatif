Attribute VB_Name = "Common"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

Public Function Common_FindNextNotEmpty(BaseCell As Range, directionDown As Boolean) As Range

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
    Set Common_FindNextNotEmpty = CurrentRange

End Function

Public Function Common_getBaseCellChantierRealFromBaseCellChantier(BaseCellChantier As Range) As Range

    Dim BaseCellChantierReal As Range
    Dim ChantierSheetReal As Worksheet
    Dim wb As Workbook

    Set Common_getBaseCellChantierRealFromBaseCellChantier = Nothing

    Set wb = BaseCellChantier.Worksheet.Parent
    On Error Resume Next
    Set ChantierSheetReal = wb.Worksheets(Nom_Feuille_Budget_chantiers_realise)
    On Error GoTo 0
    If Not (ChantierSheetReal Is Nothing) Then
        Set BaseCellChantierReal = Common_FindNextNotEmpty(ChantierSheetReal.Cells(3, 1), False)
        If BaseCellChantier.Column <= 1000 _
            And Left(BaseCellChantierReal.Value, Len("Chantier")) = "Chantier" Then
            Set Common_getBaseCellChantierRealFromBaseCellChantier = BaseCellChantierReal
        End If
    End If
End Function

Public Function Common_getChargesDefault(NB As Integer) As SetOfCharges

    Dim SetOfCharges As SetOfCharges
    Dim Charges() As Charge
    ReDim Charges(0)
    SetOfCharges.Charges = Charges
    
    Common_getChargesDefault = Common_getChargesDefaultPreserve(SetOfCharges, NB)
    
End Function

Public Function Common_getChargesDefaultPreserve(PreviousSetOfCharges As SetOfCharges, NB As Integer) As SetOfCharges

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
    Common_getChargesDefaultPreserve = SetOfCharges
    
End Function

Public Function Common_getDefaultSetOfChantiers(NBChantiers As Integer, NbDefaultDepenses As Integer) As SetOfChantiers

    Dim newArray() As Chantier
    Dim SetOfChantiers As SetOfChantiers
    Dim idx As Integer
    
    ReDim newArray(1 To NBChantiers)
    
    For idx = 1 To NBChantiers
        newArray(idx) = getDefaultChantier(NbDefaultDepenses)
    Next idx
    SetOfChantiers.Chantiers = newArray
    Common_getDefaultSetOfChantiers = SetOfChantiers

End Function

Public Function Common_GetFormula(CurrentCell As Range) As String
    
    If CurrentCell.HasFormula = True Then
        Common_GetFormula = CurrentCell.Formula
    Else
        Common_GetFormula = ""
    End If
End Function

Public Function Common_GetTypeFinancementStr( _
        wb As Workbook, _
        TypeFinancement As Integer, _
        NewFinancementInChantier As FinancementComplet _
    ) As String

    Dim Financements() As Financement
    Dim Financement As Financement
    Dim TypeFinancementsLocal() As String

    TypeFinancementsLocal = TypeFinancementsFromWb(wb)

    If (TypeFinancement <> 0) Then
        Common_GetTypeFinancementStr = TypeFinancementsLocal(TypeFinancement)
    Else
        If NewFinancementInChantier.Status Then
            Financements = NewFinancementInChantier.Financements
            Financement = Financements(1)
            If Financement.TypeFinancement <> 0 Then
                Common_GetTypeFinancementStr = TypeFinancementsLocal(Financement.TypeFinancement)
            Else
                Common_GetTypeFinancementStr = ""
            End If
        Else
            Common_GetTypeFinancementStr = ""
        End If
    End If
End Function

Public Function Common_InputBox_Get_Line_Between( _
    Message As String, _
    Title As String, _
    MinLine As Integer, _
    MaxLine As Integer _
    ) As Integer

    Dim FormatValue As Integer
    Dim Value
    Value = InputBox(Message, Title, MaxLine)

    Common_InputBox_Get_Line_Between = 0
    If Value <> "" Then
        If Value > 0 Then
            FormatValue = CInt(Value)
            If FormatValue <= MaxLine _
                And FormatValue >= MinLine Then
                Common_InputBox_Get_Line_Between = FormatValue
            End If
        End If
    Else
        Common_InputBox_Get_Line_Between = -1
    End If
End Function

Public Function Common_InsertRows( _
    BaseCell As Range, _
    PreviousNB As Integer, _
    FinalNB As Integer, _
    Optional AutoFitNext As Boolean = True, _
    Optional ExtraCols As Integer = 0, _
    Optional UpdateSum As Boolean = True) As Range

    Dim endOfRow As Range
    
    Set endOfRow = Common_FindNextNotEmpty(BaseCell, False) ' To Right
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
        Common_UpdateSumsByColumn _
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
    Set Common_InsertRows = Range(BaseCell.Cells(1 + FinalNB + 1, 1), endOfRow.Cells(1 + FinalNB + 1, 1))
    BaseCell.Select 'Force Selection
End Function

Public Function Common_IsFormula(wRange As Range) As Boolean

    Dim Cell As Range

    If wRange Is Nothing Then
        Common_IsFormula = False
    Else
        Set Cell = wRange.Cells(1, 1)
        Common_IsFormula = (Left(Cell.Formula, 1) = "=" And Cell.Formula <> Cell.Value)
    End If
End Function

Public Sub Common_RemoveRows(BaseCell As Range, PreviousNB As Integer, FinalNB As Integer, Optional ExtraCols As Integer = 0, Optional AutoFitNext As Boolean = False)

    ' Remove Cells
    Range(BaseCell.Cells(1 + FinalNB + 1, 1), Common_FindNextNotEmpty(BaseCell, False).Cells(1 + PreviousNB, 1 + ExtraCols)).Delete _
        Shift:=xlShiftUp
    
    ' Row AutoFit
    On Error Resume Next
    If AutoFitNext Then
        Range(BaseCell.Cells(1 + FinalNB + 1, 1).EntireRow, BaseCell.Cells(1 + FinalNB + FinalNB - PreviousNB, 1).EntireRow).AutoFit ' Instead of AutoFit
    End If
    On Error GoTo 0
End Sub

Public Sub Common_SetFormula( _
        CurrentCell As Range, _
        CurrentValue, _
        CurrentFormula As String, _
        Optional SetEmptyStrIfNul As Boolean = False _
    )
    If CurrentFormula <> "" Then
        On Error GoTo ImportValueInsteadOfFormula
        CurrentCell.Formula = CurrentFormula
        On Error GoTo 0
        Exit Sub
    End If
ImportValueInsteadOfFormula:
    On Error GoTo 0
    If SetEmptyStrIfNul _
        And ( _
            CInt(CurrentValue) = 0 _
            Or CStr(CurrentValue) = "" _
        ) Then
        CurrentCell.Value = ""
    Else
        CurrentCell.Value = CurrentValue
    End If
End Sub

Public Sub Common_UpdateSumsByColumn(BaseRange As Range, DestinationRange As Range, PreviousNB As Integer)
    
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


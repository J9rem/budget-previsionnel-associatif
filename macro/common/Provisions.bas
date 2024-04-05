Attribute VB_Name = "Provisions"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

' clean content of Provision sheet
' @param Worksheet ProvisionsSheet
' @return Boolean allIsRight
Public Function Provisions_Clean_Sheet(ProvisionsSheet As Worksheet) As Boolean

    Dim FinanciersLines() As String
    Dim FinanciersLinesRaw As String
    Dim NBYears As Integer

    ' Default False
    Provisions_Clean_Sheet = False
    
    NBYears = Provisions_Years_getNb(ProvisionsSheet)
    If NBYears > 0 Then
        FinanciersLinesRaw = Provisions_Financiers_Get_Lines(ProvisionsSheet, NBYears)
        If FinanciersLinesRaw <> "" Then
            FinanciersLines = Split(FinanciersLinesRaw, ",")

            ' delete lines (margin of 5 lines)
            Range( _
                ProvisionsSheet.Cells(5, 1), _
                ProvisionsSheet.Cells(CInt(FinanciersLines(UBound(FinanciersLines))) + NBYears + 5, 1) _
            ).EntireRow.Delete Shift:=xlUp

            ' All is right
            Provisions_Clean_Sheet = True
        End If
    End If

End Function

' find similar Financier in Data
' same name and european
' @param Data Data
' @return Data
Public Function Provisions_Data_Update_Index(Data As Data) As Data

    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim Financement As Financement
    Dim FinancementFound As Integer
    Dim Financements() As Financement
    Dim IndexFinancement As Integer
    Dim IndexProvision As Integer
    Dim Provision As Provision
    Dim Provisions() As Provision
    Dim ProvisionsNames() As String

    Provisions = Data.Provisions

    If UBound(Provisions) > 0 Then
        ' Extract ProvisionsNames
        ReDim ProvisionsNames(1 To UBound(Provisions))

        For IndexProvision = 1 To UBound(Provisions)
            Provision = Provisions(IndexProvision)
            ProvisionsNames(IndexProvision) = Provision.NomDuFinanceur
        Next IndexProvision

        ' Find similar financement in Chantiers
        Chantiers = Data.Chantiers
        If UBound(Chantiers) > 0 Then
            Chantier = Chantiers(1)
            Financements = Chantier.Financements
            FinancementFound = 0
            For IndexFinancement = 1 To UBound(Financements)
                If FinancementFound < 0 Then
                    Financement = Financements(IndexFinancement)
                    ' only check european funding
                    If Financement.TypeFinancement = 6 Then
                        IndexProvision = indexOfInArrayStr(Financement.Nom, ProvisionsNames)
                        If IndexProvision <> -1 Then
                            FinancementFound = IndexFinancement
                            Data = Provisions_Data_Update_Index_In_Financement( _
                                Data, _
                                IndexFinancement, _
                                IndexProvision _
                            )
                            Data = Provisions_Data_Update_Range_In_Provisions( _
                                Data, _
                                IndexFinancement, _
                                IndexProvision _
                            )
                        End If
                    End If
                End If
            Next IndexFinancement
        End If
    End If

    Provisions_Data_Update_Index = Data

End Function

' update IndexInProvisions In Financement for each Chantier
' @param Data Data
' @param Integer IndexFinancement
' @param Integer IndexProvision
' @return Data
Public Function Provisions_Data_Update_Index_In_Financement( _
        Data As Data, _
        IndexFinancement As Integer, _
        IndexProvision As Integer _
    ) As Data

    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim Financement As Financement
    Dim Financements() As Financement
    Dim IndexChantier As Integer

    ' get
    Chantiers = Data.Chantiers

    If UBound(Chantiers) > 0 Then

        For IndexChantier = 1 To UBound(Chantiers)
            ' get
            Chantier = Chantiers(IndexChantier)
            Financements = Chantier.Financements
            If UBound(Financements) > 0 Then
                ' get
                Financement = Financements(IndexFinancement)
                Financement.IndexInProvisions = IndexProvision
                ' set
                Financements(IndexFinancement) = Financement
            End If
            ' set
            Chantier.Financements = Financements
            Chantiers(IndexChantier) = Chantier
        Next IndexChantier

        ' set
        Data.Chantiers = Chantiers
    End If

    Provisions_Data_Update_Index_In_Financement = Data
End Function

' update Range for each Provision
' @param Data Data
' @param Integer IndexFinancement
' @param Integer IndexProvision
' @return Data
Public Function Provisions_Data_Update_Range_In_Provisions( _
        Data As Data, _
        IndexFinancement As Integer, _
        IndexProvision As Integer _
    ) As Data

    Dim Chantier As Chantier
    Dim Chantiers() As Chantier
    Dim Financement As Financement
    Dim Financements() As Financement
    Dim Provision As Provision
    Dim Provisions() As Provision
    Dim NBChantiers As Integer

    ' get
    Provisions = Data.Provisions

    Chantiers = Data.Chantiers
    NBChantiers = UBound(Chantiers)

    If NBChantiers > 0 And UBound(Provisions) > 0 Then
        Chantier = Chantiers(1)
        Financements = Chantier.Financements
        If UBound(Financements) > 0 Then
            Financement = Financements(IndexFinancement)

            ' get
            Provision = Provisions(IndexProvision)
            ' update range
            Set Provision.RangeForTitle = Financement.BaseCell.Cells(1, 2)
            Set RangeForLastYearWaitedValue = Financement.BaseCell.Cells(1, 3 + NBChantiers)
            Set RangeForLastYearPayedValue = Financement.BaseCell.Cells(1, 3 + 3 * NBChantiers)

            ' set
            Provisions(IndexProvision) = Provision
            Data.Provisions = Provisions
        End If
    End If

    Provisions_Data_Update_Range_In_Provisions = Data
End Function

Public Function Provisions_Extract(wb As Workbook, Data As Data, Revision As WbRevision) As Data

    Dim FinanciersLines() As String
    Dim FinanciersLinesRaw As String
    Dim FirstYear As Integer
    Dim Index As Integer
    Dim NBYears As Integer
    Dim Provisions() As Provision
    Dim ProvisionsSheet As Worksheet
    Dim ShouldHaveProvisions As Boolean

    ReDim Provisions(0)
    Data.Provisions = Provisions

    ' Define if provisions are waited
    ShouldHaveProvisions = False
    If (Revision.Majeure = 2 And Revision.Mineure > 2) Or Revision.Majeure > 2 Then
        ShouldHaveProvisions = True
    End If

    ' First get the right sheet
    On Error Resume Next
    Set ProvisionsSheet = wb.Worksheets(Nom_Feuille_Provisions)
    On Error GoTo 0
    If ProvisionsSheet Is Nothing Then
        ' only show error message if revision higher than 2.3
        If ShouldHaveProvisions Then
            MsgBox Replace(T_NotFoundPage, "%PageName%", Nom_Feuille_Provisions)
        End If
        GoTo FinFunctionProvisions
    End If

    NBYears = Provisions_Years_getNb(ProvisionsSheet)
    If NBYears = 0 Then
        GoTo FinFunctionProvisions
    End If
    FirstYear = CInt(ProvisionsSheet.Cells(4, 4).Value)

    FinanciersLinesRaw = Provisions_Financiers_Get_Lines(ProvisionsSheet, NBYears)
    If FinanciersLinesRaw = "" Then
        GoTo FinFunctionProvisions
    End If
    FinanciersLines = Split(FinanciersLinesRaw, ",")

    ReDim Provisions(1 To (UBound(FinanciersLines) + 1))
    For Index = 1 To UBound(Provisions)
        Provisions(Index) = Provisions_Extract_For_A_Financier( _
            ProvisionsSheet, _
            NBYears, _
            FirstYear, _
            CInt(FinanciersLines(Index - 1)) _
        )
    Next Index
    Data.Provisions = Provisions

    Data = Provisions_Data_Update_Index(Data)

FinFunctionProvisions:
    Provisions_Extract = Data
End Function


' extract provision for a financier
' @param Worksheet ws
' @param Integer NBYears
' @param Integer FirstYear
' @param Integer RowLine
' @return Provision
Public Function Provisions_Extract_For_A_Financier( _
        ws As Worksheet, _
        NBYears As Integer, _
        FirstYear As Integer, _
        RowLine As Integer _
    ) As Provision

    Dim BaseCell As Range
    Dim IndexColumn As Integer
    Dim IndexPayedValues As Integer
    Dim IndexRetrievalTenPercent As Integer
    Dim IndexYear As Integer
    Dim PayedValues() As Double
    Dim Provision As Provision
    Dim RetrievalTenPercent() As Double
    Dim RetrievalTenPercentFormula() As String
    Dim WaitedValues() As Double
    Dim WorkingCell As Range

    Set BaseCell = ws.Cells(RowLine, 1)

    Provision = getDefaultProvision(NBYears)
    
    ' Title
    Provision.NomDuFinanceur = BaseCell.Value
    ' Search base range default
    Set Provision.RangeForTitle = Nothing
    Provision.FirstYear = FirstYear

    Set Provision.RangeForLastYearWaitedValue = Nothing
    Set Provision.RangeForLastYearPayedValue = Nothing

    ' get
    PayedValues = Provision.PayedValues
    RetrievalTenPercent = Provision.RetrievalTenPercent
    RetrievalTenPercentFormula = Provision.RetrievalTenPercentFormula
    WaitedValues = Provision.WaitedValues

    IndexPayedValues = 1
    IndexRetrievalTenPercent = 1

    For IndexYear = 1 To NBYears

        WaitedValues(IndexYear) = CDbl(BaseCell.Cells(IndexYear, 3).Value)

        ' PayedValues
        For IndexColumn = IndexYear To NBYears
            PayedValues(IndexPayedValues) = CDbl(BaseCell.Cells(IndexYear, 3 + IndexColumn).Value)
            IndexPayedValues = IndexPayedValues + 1
        Next IndexColumn

        If IndexYear < NBYears Then
            ' RetrievalTenPercent
            For IndexColumn = (IndexYear + 1) To NBYears
                Set WorkingCell = BaseCell.Cells(IndexYear, 6 + 3 * NBYears + IndexColumn)
                RetrievalTenPercent(IndexRetrievalTenPercent) = CDbl(WorkingCell.Value)
                RetrievalTenPercentFormula(IndexRetrievalTenPercent) = Common_GetFormula(WorkingCell)
                IndexRetrievalTenPercent = IndexRetrievalTenPercent + 1
            Next IndexColumn
        End If

    Next IndexYear

    ' set
    Provision.PayedValues = PayedValues
    Provision.RetrievalTenPercent = RetrievalTenPercent
    Provision.RetrievalTenPercentFormula = RetrievalTenPercentFormula
    Provision.WaitedValues = WaitedValues
    
    Provisions_Extract_For_A_Financier = Provision
End Function

' search each line of a financier
' @param Worksheet ws
' @param Integer NBYears
' @return String coma separated lines
Public Function Provisions_Financiers_Get_Lines(ws As Worksheet, NBYears As Integer) As String

    Dim ContinueTest As Boolean
    Dim CurrentRange As Range
    Dim CurrentValue
    Dim Result As String

    ' init (value to also define errors)
    Result = ""

    Set CurrentRange = ws.Cells(5, 1)
    CurrentValue = CurrentRange.Value
    If CurrentRange.HasFormula = True Then
        ContinueTest = True
    Else
        ContinueTest = Not (CurrentValue = "" Or CurrentValue = Empty)
    End If
    While ContinueTest
        If Result <> "" Then
            Result = Result & ","
        End If
        Result = Result & CurrentRange.Row
        Set CurrentRange = CurrentRange.Cells(NBYears + 3, 1)
        If CurrentRange.HasFormula = True Then
            ContinueTest = True
        Else
            CurrentValue = CurrentRange.Value
            ContinueTest = Not (CurrentValue = "" Or CurrentValue = Empty)
        End If
    Wend

    Provisions_Financiers_Get_Lines = Result
End Function

' import Provisions
' @param Workbook wb
' @param Data data
Public Sub Provisions_Import(wb As Workbook, Data As Data)

    Dim ProvisionsSheet As Worksheet

    ' get Provisions Sheet
    On Error Resume Next
    Set ProvisionsSheet = wb.Worksheets(Nom_Feuille_Provisions)
    On Error GoTo 0
    If Not (ProvisionsSheet Is Nothing) Then
        ' clean Sheet
        If Provisions_Clean_Sheet(ProvisionsSheet) Then
            ' add new content
            Provisions_NewContent_Add ProvisionsSheet, Data
        End If
    End If
End Sub

' init content of a provision
' @param Provision Provision
' @param Integer NBYears
' @return Provision
Public Function Provisions_Init(Provision As Provision, NBYears As Integer) As Provision
    
    Dim Index As Integer
    Dim LengthForPayed As Integer
    Dim LengthForRetrieval As Integer
    Dim PayedValues() As Double
    Dim RetrievalTenPercent() As Double
    Dim RetrievalTenPercentFormula() As String
    Dim WaitedValues() As Double

    ' Initiate length for retrieval and payed
    LengthForPayed = 0
    LengthForRetrieval = 0
    For Index = 1 To NBYears
        LengthForPayed = LengthForPayed + (NBYears - Index + 1)
        LengthForRetrieval = LengthForRetrieval + (NBYears - Index)
    Next Index

    ' calculate sum of n element algebric
    ReDim PayedValues(1 To LengthForPayed)
    ReDim RetrievalTenPercent(1 To LengthForRetrieval)
    ReDim RetrievalTenPercentFormula(1 To LengthForRetrieval)
    ReDim WaitedValues(1 To NBYears)

    ' Init Values
    For Index = 1 To LengthForPayed
        PayedValues(Index) = 0
    Next Index
    For Index = 1 To LengthForRetrieval
        RetrievalTenPercent(Index) = 0
        RetrievalTenPercentFormula(Index) = ""
    Next Index
    
    Provision.NomDuFinanceur = ""
    Set Provision.RangeForTitle = Nothing
    Provision.NBYears = NBYears
    Provision.FirstYear = 2000
    Provision.WaitedValues = WaitedValues
    Set Provision.RangeForLastYearWaitedValue = Nothing
    Provision.PayedValues = PayedValues
    Provision.RetrievalTenPercent = RetrievalTenPercent
    Provision.RetrievalTenPercentFormula = RetrievalTenPercentFormula
    Set Provision.RangeForLastYearPayedValue = Nothing
    
    Provisions_Init = Provision

End Function

' extract value of current main year in first worskeet
' @param Workbook wb
' @param Integer MainYearValue
Public Function Provisions_Main_Year_Get(wb As Workbook) As Integer

    Dim CurrentSheet As Worksheet
    Dim BaseCell As Range

    ' default value
    Provisions_Main_Year_Get = 2024
    
    On Error Resume Next
    Set CurrentSheet = wb.Worksheets(Nom_Feuille_Informations)
    On Error GoTo 0
    If Not (CurrentSheet Is Nothing) Then
        Set BaseCell = CurrentSheet.Range("A:A").Find(Label_Annees)
        If Not BaseCell Is Nothing Then
            Provisions_Main_Year_Get = BaseCell.Cells(1, 2).Value
        End If
    End If
End Function

' add new content in Provisions sheet
' @param Worksheet ProvisionsSheet
' @param Data As Data
Public Sub Provisions_NewContent_Add(ProvisionsSheet As Worksheet, Data As Data)

    Dim CurrentStartCell As Range
    Dim FirstYear As Integer
    Dim Index As Integer
    Dim NBYears As Integer
    Dim Provision As Provision
    Dim Provisions() As Provision

    Provisions = Data.Provisions

    If UBound(Provisions) > 0 Then
        NBYears = Provisions_UpdateNBYears(ProvisionsSheet, Data)
        FirstYear = CInt(ProvisionsSheet.Cells(4, 4).Value)
        Set CurrentStartCell = ProvisionsSheet.Cells(5, 1)
        For Index = 1 To UBound(Provisions)
            Provision = Provisions(Index)
            Set CurrentStartCell = Provisions_Provision_Add(CurrentStartCell, Provision, NBYears, FirstYear)
        Next Index
    End If

End Sub

' add content of a provision and return next start cell
' @param Range CurrentStartCell
' @param Provision Provision
' @param Integer FirstYear
' @param Integer NBYears
' @return Range NextCurrentStartCell
Public Function Provisions_Provision_Add( _
        CurrentStartCell As Range, _
        Provision As Provision, _
        NBYears As Integer, _
        FirstYear As Integer _
    ) As Range

    Dim Index As Integer
    Dim Index2 As Integer
    Dim RangeForTitle As Range
    Dim NextCurrentStartCell As Range
    Dim PayedValue As Double
    Dim PayedValues() As Double
    Dim WaitedValue As Double
    Dim WaitedValues() As Double
    Dim WorkingYear As Integer
    Dim WorkingYear2 As Integer

    Set NextCurrentStartCell = Nothing
    If Not (CurrentStartCell Is Nothing) Then

        ' Add title
        Set RangeForTitle = Provision.RangeForTitle
        If RangeForTitle Is Nothing Then
            CurrentStartCell.Cells(1, 1).Value = Provision.NomDuFinanceur
        Else
            CurrentStartCell.Cells(1, 1).Formula = "=SIERREUR(" _
                    & CleanAddress(RangeForTitle.address(True, True, xlA1, True)) _
                    & ";""" & Provision.NomDuFinanceur & """" _
                & ")"
        End If
        Specific_Provisions_Theme_Set _
            CurrentStartCell.Cells(1, 1), _
            False, "lightGrey", False

        ' add compta
        For Index = 1 To NBYears
            CurrentStartCell.Cells(Index, 2).Value = FirstYear + Index - 1
            Specific_Provisions_Theme_Set _
                CurrentStartCell.Cells(Index, 2), _
                False, "middleGrey", False
        Next Index

        ' add content
        PayedValues = Provision.PayedValues
        WaitedValues = Provision.WaitedValues
        For Index = 1 To NBYears
            WorkingYear = FirstYear + Index - 1
            If WorkingYear < Provision.FirstYear _
                Or WorkingYear > (Provision.FirstYear + Provision.NBYears - 1) Then
                ' add to receive
                CurrentStartCell.Cells(Index, 3).Value = 0
                ' add payments
                For Index2 = 1 To NBYears
                    If Index2 < Index Then
                        CurrentStartCell.Cells(Index, 3 + Index2).Value = ""
                    Else
                        CurrentStartCell.Cells(Index, 3 + Index2).Value = 0
                    End If
                Next
            Else
                WaitedValue = WaitedValues(WorkingYear - Provision.FirstYear + 1)
                ' add to receive
                CurrentStartCell.Cells(Index, 3).Value = WaitedValue
                ' add payments
                For Index2 = 1 To NBYears
                    If Index2 < Index Then
                        CurrentStartCell.Cells(Index, 3 + Index2).Value = ""
                    Else
                        WorkingYear2 = FirstYear + Index2 - 1
                        If WorkingYear2 < Provision.FirstYear _
                            Or WorkingYear2 > (Provision.FirstYear + Provision.NBYears - 1) Then
                            CurrentStartCell.Cells(Index, 3 + Index2).Value = 0
                        Else
                            PayedValue = PayedValues(WorkingYear2 - Provision.FirstYear + 1)
                            CurrentStartCell.Cells(Index, 3 + Index2).Value = PayedValue
                        End If
                    End If
                Next
            End If
            ' add to receive
            Specific_Provisions_Theme_Set _
                CurrentStartCell.Cells(Index, 3), _
                True, "LightYellow", False
            ' add payments and provisions
            For Index2 = 1 To NBYears
                If Index2 < Index Then
                    ' add payments
                    Specific_Provisions_Theme_Set _
                        CurrentStartCell.Cells(Index, 3 + Index2), _
                        False, "Grey", False
                    ' add provisions
                    CurrentStartCell.Cells(Index, 4 + NBYears + Index2).Value = ""
                    Specific_Provisions_Theme_Set _
                        CurrentStartCell.Cells(Index, 4 + NBYears + Index2), _
                        False, "Grey", False
                Else
                    ' add payments
                    Specific_Provisions_Theme_Set _
                        CurrentStartCell.Cells(Index, 3 + Index2), _
                        True, "", False
                    ' add provisions
                    CurrentStartCell.Cells(Index, 4 + NBYears + Index2).Formula = "=" _
                        & "0.1*" & CleanAddress( _
                            CurrentStartCell.Cells(Index, 3 + Index2).address(False, False, xlA1, False) _
                        )
                    If Index2 = Index Then
                        CurrentStartCell.Cells(Index, 4 + NBYears + Index2).Formula = _
                            CurrentStartCell.Cells(Index, 4 + NBYears + Index2).Formula _
                            & "+0.25*SUM(" _
                                & CleanAddress( _
                                    CurrentStartCell.Cells(Index, 4 + Index2).address(False, False, xlA1, False) _
                                ) _
                                & ":" _
                                & CleanAddress( _
                                    CurrentStartCell.Cells(Index, 4 + NBYears).address(False, True, xlA1, False) _
                                ) _
                            & ")"
                    End If
                    Specific_Provisions_Theme_Set _
                        CurrentStartCell.Cells(Index, 4 + NBYears + Index2), _
                        True, "LightBlueForTotalForAutoFilledCell", False
                End If
            Next
            ' add waited payments formula
            CurrentStartCell.Cells(Index, 4 + NBYears).Formula = "=MAX(0;" _
                & CleanAddress(CurrentStartCell.Cells(Index, 3).address(False, False, xlA1, False)) _
                & "-SUM(" _
                    & CleanAddress(Range( _
                        CurrentStartCell.Cells(Index, 4), _
                        CurrentStartCell.Cells(Index, 3 + NBYears) _
                    ).address(False, False, xlA1, False)) _
                    & ")" _
            & ")"
            Specific_Provisions_Theme_Set _
                CurrentStartCell.Cells(Index, 4 + NBYears), _
                True, "lightGrey", False
        Next Index
        ' add to receive
        If Not (Provision.RangeForLastYearWaitedValue Is Nothing) Then
            CurrentStartCell.Cells(NBYears, 3).Formula = "=SIERREUR(" _
                    & CleanAddress(Provision.RangeForLastYearWaitedValue.address(True, True, xlA1, True)) _
                    & ";" & CurrentStartCell.Cells(NBYears, 3).Value _
                & ")"
        End If
        ' add to waited payment
        If Not (Provision.RangeForLastYearPayedValue Is Nothing) _
            And (Provision.FirstYear + Provision.NBYears) = (FirstYear + NBYears) Then
            CurrentStartCell.Cells(NBYears, 3 + NBYears).Formula = "=SIERREUR(" _
                    & CleanAddress(Provision.RangeForLastYearPayedValue.address(True, True, xlA1, True)) _
                    & ";" & CurrentStartCell.Cells(NBYears, 3 + NBYears).Value _
                & ")"
        End If
        ' total below waited payments formula
        CurrentStartCell.Cells(NBYears + 1, 4 + NBYears).Value = "Total"
        Specific_Provisions_Theme_Set _
            CurrentStartCell.Cells(NBYears + 1, 4 + NBYears), _
            False, "Blue", True
        ' total below
        For Index = (5 + NBYears) To (5 * NBYears + 13)
            CurrentStartCell.Cells(NBYears + 1, Index).Formula = "=SUM(" _
                & CleanAddress(Range( _
                        CurrentStartCell.Cells(1, Index), _
                        CurrentStartCell.Cells(NBYears, Index) _
                    ).address(False, False, xlA1, False)) _
                & ")"
            Specific_Provisions_Theme_Set _
                CurrentStartCell.Cells(NBYears + 1, Index), _
                True, "LightBlueForTotal", False
        Next Index
        CurrentStartCell.Cells(NBYears + 1, 5 * NBYears + 10).Value = ""
        Specific_Provisions_Theme_Set _
            CurrentStartCell.Cells(NBYears + 1, 5 * NBYears + 10), _
            False, "", False

        ' update next cell
        Set NextCurrentStartCell = CurrentStartCell.Cells(NBYears + 3, 1)

    End If

    Set Provisions_Provision_Add = NextCurrentStartCell
End Function

' search range for forecast of Provisions
' @param Workbook wb
' @param Boolean ForProvisions
' @param Boolean ForForecast
' @return Range Nothing On Error
Public Function Provisions_SearchRange( _
        wb As Workbook, _
        ForProvisions As Boolean, _
        ForForecast As Boolean _
    ) As Range

    Dim Destination As Range
    Dim NBYears As Integer
    Dim ProvisionsSheet As Worksheet

    Set Destination = Nothing

    ' First get the right sheet
    On Error Resume Next
    Set ProvisionsSheet = wb.Worksheets(Nom_Feuille_Provisions)
    On Error GoTo 0
    If Not (ProvisionsSheet Is Nothing) Then
        NBYears = Provisions_Years_getNb(ProvisionsSheet)
        If NBYears > 0 Then
            ' define the right cell
            If ForForecast Then
                If ForProvisions Then
                    Set Destination = ProvisionsSheet.Cells(1, 10 + 5 * NBYears)
                Else
                    Set Destination = ProvisionsSheet.Cells(1, 11 + 5 * NBYears)
                End If
            Else
                If ForProvisions Then
                    Set Destination = ProvisionsSheet.Cells(1, 4 + 2 * NBYears)
                Else
                    Set Destination = ProvisionsSheet.Cells(1, 7 + 4 * NBYears)
                End If
            End If
        End If
    End If

    Set Provisions_SearchRange = Destination
End Function

' search the first years without empty value
' then update years update to current year or maximum year
' @param Worksheet ProvisionsSheet
' @param Data Data
' @return Integer
Public Function Provisions_UpdateNBYears(ProvisionsSheet As Worksheet, Data As Data) As Integer

    Dim CurrentNBYears As Integer
    Dim CurrentIndex As Integer
    Dim CurrentIndexes(1 To 5) As Integer
    Dim CurrentValue As Double
    Dim CurrentValues() As Double
    Dim CurrentYear As Integer
    Dim Index As Integer
    Dim IndexYear As Integer
    Dim Index2 As Integer
    Dim Provision As Provision
    Dim Provisions() As Provision
    Dim NBColsToChange As Integer
    Dim MaximumLastYear As Integer
    Dim MinimumFirstYear As Integer
    Dim WantedFirstYear As Integer
    Dim WantedLastYear As Integer
    Dim WantedNBYears As Integer

    ' default
    WantedLastYear = Provisions_Main_Year_Get(ProvisionsSheet.Parent)
    WantedFirstYear = WantedLastYear - 4

    Provisions = Data.Provisions
    If UBound(Provisions) > 0 Then
        For Index = 1 To UBound(Provisions)
            Provision = Provisions(Index)
            MaximumLastYear = Provision.FirstYear + Provision.NBYears - 1
            MinimumFirstYear = MaximumLastYear

            ' test waited values
            CurrentValues = Provision.WaitedValues
            For IndexYear = 1 To Provision.NBYears
                ' test if empty for this year
                CurrentYear = Provision.FirstYear + IndexYear - 1
                CurrentValue = CurrentValues(IndexYear)
                If CurrentValue <> 0 Then
                    If CurrentYear < MinimumFirstYear Then
                        MinimumFirstYear = CurrentYear
                    End If
                End If
            Next IndexYear

            ' test payments
            CurrentValues = Provision.PayedValues
            CurrentIndex = 1
            For IndexYear = 1 To Provision.NBYears
                ' test if empty for this year
                For Index2 = IndexYear To Provision.NBYears
                    CurrentYear = Provision.FirstYear + Index2 - 1
                    CurrentValue = CurrentValues(CurrentIndex)
                    If CurrentValue <> 0 Then
                        If CurrentYear < MinimumFirstYear Then
                            MinimumFirstYear = CurrentYear
                        End If
                    End If
                    CurrentIndex = CurrentIndex + 1
                Next Index2

            Next IndexYear
            
            ' test RetrievalTenPercent
            CurrentValues = Provision.RetrievalTenPercent
            CurrentIndex = 1
            For IndexYear = 1 To Provision.NBYears
                ' test if empty for this year
                For Index2 = (IndexYear + 1) To Provision.NBYears
                    CurrentYear = Provision.FirstYear + Index2 - 1
                    CurrentValue = CurrentValues(CurrentIndex)
                    If CurrentValue <> 0 Then
                        If CurrentYear < MinimumFirstYear Then
                            MinimumFirstYear = CurrentYear
                        End If
                    End If
                    CurrentIndex = CurrentIndex + 1
                Next Index2

            Next IndexYear

            If MinimumFirstYear < WantedFirstYear Then
                WantedFirstYear = MinimumFirstYear
            End If
            If WantedLastYear < MaximumLastYear Then
                WantedLastYear = MaximumLastYear
            End If
        Next Index
    End If

    WantedNBYears = WantedLastYear - WantedFirstYear + 1

    CurrentNBYears = Provisions_Years_getNb(ProvisionsSheet)
    ' net
    CurrentIndexes(1) = 4 * CurrentNBYears + 9
    ' 10% manual
    CurrentIndexes(2) = 3 * CurrentNBYears + 8
    ' 25% auto
    CurrentIndexes(3) = 2 * CurrentNBYears + 6
    ' provision
    CurrentIndexes(4) = CurrentNBYears + 6
    ' payments
    CurrentIndexes(5) = 5
    If CurrentNBYears > WantedNBYears Then
        ' remove rows
        NBColsToChange = CurrentNBYears - WantedNBYears
        For Index2 = 1 To 5
            For Index = 1 To NBColsToChange
                ProvisionsSheet.Cells(1, CurrentIndexes(Index2)).EntireColumn.Delete Shift:=xlToLeft
            Next Index
        Next Index2
    Else
        If CurrentNBYears < WantedNBYears Then
            ' add rows
            NBColsToChange = WantedNBYears - CurrentNBYears
            For Index2 = 1 To 5
                For Index = 1 To NBColsToChange
                    ProvisionsSheet.Cells(1, CurrentIndexes(Index2)).Select
                    ProvisionsSheet.Cells(1, CurrentIndexes(Index2)).Copy
                    ProvisionsSheet.Cells(1, CurrentIndexes(Index2)).EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                Next Index
            Next Index2
        End If
    End If
    CurrentNBYears = Provisions_Years_getNb(ProvisionsSheet)
    Provisions_UpdateNBYears = CurrentNBYears

    ' update values of years in header
    ' net
    CurrentIndexes(1) = 4 * CurrentNBYears + 8
    ' 10% manual
    CurrentIndexes(2) = 3 * CurrentNBYears + 7
    ' 25% auto
    CurrentIndexes(3) = 2 * CurrentNBYears + 5
    ' provision
    CurrentIndexes(4) = CurrentNBYears + 5
    ' payments
    CurrentIndexes(5) = 4

    ' payments
    For Index = 1 To CurrentNBYears
        ProvisionsSheet.Cells(1, CurrentIndexes(5) + Index - 1).Value = WantedFirstYear + Index - 1
    Next Index
    ' then others
    For Index2 = 1 To 4
        For Index = 1 To CurrentNBYears
            ProvisionsSheet.Cells(1, CurrentIndexes(Index2) + Index - 1).Formula = _
                "=" & CleanAddress( _
                    ProvisionsSheet.Cells(1, CurrentIndexes(5) + Index - 1).address(True, True, xlA1, False) _
                )
        Next Index
    Next Index2

End Function

' search the NB years in Provisions sheet
' @param Worksheet ws
' @return Integer ' return 0 in case of error
Public Function Provisions_Years_getNb(ws As Worksheet) As Integer

    Dim CurrentCounter As Integer
    Dim CurrentRange As Range
    Dim CurrentValue
    Dim Index As Integer
    Dim NBYears As Integer

    CurrentCounter = 0
    NBYears = 0

    Set CurrentRange = ws.Cells(4, 4)
    ' Limit to 20 years
    For Index = 1 To 20
        If NBYears = 0 Then
            CurrentValue = CurrentRange.Cells(1, Index).Value
            If Not (CurrentValue = "" Or CurrentValue = Empty) Then
                If CurrentValue > 2000 _
                    And CurrentValue < 2050 Then
                    CurrentCounter = CurrentCounter + 1
                Else
                    If CurrentValue = Label_Waited_Payments Then
                        NBYears = CurrentCounter
                    End If
                End If
            End If
        End If
    Next Index

    Provisions_Years_getNb = NBYears
End Function



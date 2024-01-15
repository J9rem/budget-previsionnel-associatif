Attribute VB_Name = "Specific"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' constantes
Public Const Nom_de_Fichier_Par_Defaut As String = "InCitu_Budget_Previsionnel_Associatif_Excel"
Public Const BackupDefaultExtension As String = ".xlsx"
Public Const intoOds As Boolean = False


' utils
Public Function choisirNomFicherASauvegarderSansMacro(ByRef FilePath As String) As Boolean
    
    Dim Adresse_dossier_courant As String
    Dim Default_File_Name As String
    Dim Fichier_De_Sauvegarde As String
    
    ' Default FileName
    Default_File_Name = Nom_de_Fichier_Par_Defaut & "_" & Format(Date, "yyyy-mm-dd_") & Format(Time, "hh-nn")
    
    ' Default
    FilePath = ""
    
    ' Changement de dossier
    Adresse_dossier_courant = ThisWorkbook.Path
    ChDir (Adresse_dossier_courant)
    
    ' Fenêtre pour demander le nom du fichier de sauvegarde
    On Error Resume Next
    ' InitialFileName, FileFilter, FiltrerIndex, Title
    Fichier_De_Sauvegarde = Application.GetSaveAsFilename( _
        Default_File_Name, _
        "Excel 2003-2007 (*.xls),*.xls,Excel (*.xlsx),*.xlsx", _
        2, _
        "Choisir le fichier à exporter")
    On Error GoTo 0
    If Fichier_De_Sauvegarde = "" Or Fichier_De_Sauvegarde = Empty Or Fichier_De_Sauvegarde = "Faux" Or Fichier_De_Sauvegarde = "False" Then
        choisirNomFicherASauvegarderSansMacro = False
    Else
        FilePath = Fichier_De_Sauvegarde
        choisirNomFicherASauvegarderSansMacro = True
    End If

End Function

Sub DeleteFile(FileName As String, FolderName As String)
    ChDir (FolderName)
    Kill FileName
End Sub

Public Function saveWorkBookAsCopyNoMacro(FilePath As String, FileName As String, Ext As String, ByRef OpenedWorkbook As Workbook) As Boolean

    Dim DestFolder As String
    Dim TmpFileName As String
    Dim wb As Workbook
    
    saveWorkBookAsCopyNoMacro = False
    
    DestFolder = Left(FilePath, Len(FilePath) - Len(FileName))
    
    TmpFileName = Left(FileName, Len(FileName) - Len(Ext)) & "_" & WorksheetFunction.RandBetween(10000, 45000) & ".xlsm"
    If FileExists(DestFolder & TmpFileName) Then
        saveWorkBookAsCopyNoMacro = False
        Exit Function
    End If
    
    ChDir (DestFolder)
    ThisWorkbook.SaveCopyAs FileName:=TmpFileName
    
    On Error Resume Next
    Workbooks.Open FileName:=DestFolder & TmpFileName, ReadOnly:=False
    On Error GoTo 0
    returnToCurrentPath
    
    If FindOpenedWorkBook(TmpFileName, OpenedWorkbook) Then
        removeShapes OpenedWorkbook
        ' save as new format
        saveWorkbookAs OpenedWorkbook, DestFolder, FileName
        If FileExists(DestFolder & FileName) Then
            saveWorkBookAsCopyNoMacro = True
        End If
    End If
    
    If FileExists(DestFolder & TmpFileName) Then
        If FindOpenedWorkBook(TmpFileName, wb) Then
            closeSilentWorkbook wb
        End If
        deleteSilentWorkbook DestFolder & TmpFileName, TmpFileName
    End If
    
End Function

Public Sub SaveCopyAs(wb As Workbook, FileName As String, FolderName As String)
    ChDir (FolderName)
    wb.SaveCopyAs FileName
End Sub

Public Function removeShapes(wb As Workbook) As Boolean
    Dim Shape As Shape
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        For Each Shape In ws.Shapes
            Shape.Delete
        Next ' Shape
    Next ' Ws
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
    Dim ArrayTmp() As String
    Dim Name As Name
    Dim RangeForType As Range
    Dim Index As Integer
    Dim NBTypesFinancements As Integer
    
    Set RangeForType = Nothing
    For Index = 1 To wb.Names.Count
        If wb.Names(Index).Name = "TYPE_FINANCEUR" Then
            Set RangeForType = wb.Names(Index).RefersToRange
        End If
    Next Index
    
    If RangeForType Is Nothing Then
        TypeFinancementsFromWb = TypeFinancements()
    Else
        NBTypesFinancements = RangeForType.Count
        ReDim ArrayTmp(0 To NBTypesFinancements)
        ArrayTmp(0) = ""
        For Index = 1 To NBTypesFinancements
            ArrayTmp(Index) = RangeForType.Item(Index).value
        Next Index
        
        TypeFinancementsFromWb = ArrayTmp
    End If
    
End Function

Public Sub CleanLineStylesForBudget(BaseCell As Range, HeadCell As Range, IsHeader As Boolean)
    
    Dim TmpVar As Variant
    Dim VarTmp As Variant
    With BaseCell
        If IsHeader Then
            TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
        Else
            If HeadCell.Row = BaseCell.Row - 1 Then
                TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
            Else
                TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal, xlEdgeTop)
                .Cells(0, 1).Borders(xlEdgeBottom).LineStyle = xlNone
            End If
        End If
        For Each VarTmp In TmpVar
            .Borders(VarTmp).LineStyle = xlNone
        Next VarTmp
    End With
End Sub

Public Sub DefineLineStylesForBudget( _
        BaseCell As Range, _
        HeadCell As Range, _
        IsHeader As Boolean _
    )
    Dim TmpVar As Variant
    Dim VarTmp As Variant
    With BaseCell
        If IsHeader Then
            TmpVar = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
            With .Cells(0, 1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlThin
                .TintAndShade = 0
            End With
        Else
            If HeadCell.Row = BaseCell.Row - 1 Then
                TmpVar = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
            Else
                TmpVar = Array(xlEdgeLeft, xlEdgeRight, xlEdgeBottom)
            End If
        End If
        For Each VarTmp In TmpVar
            With .Borders(VarTmp)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlThin
                .TintAndShade = 0
            End With
        Next VarTmp
    End With
End Sub

Public Sub SetFormatForBudget(BaseCell As Range, HeadCell As Range, IsHeader As Boolean)

    Dim IndexBis As Integer
    
    For IndexBis = 1 To 3
        With BaseCell.Cells(1, IndexBis)
            CleanLineStylesForBudget .Cells(1, 1), HeadCell, IsHeader
            DefineLineStylesForBudget .Cells(1, 1), HeadCell, IsHeader
            
            With .Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 8
                .Strikethrough = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                If IsHeader Then
                    .Color = RGB(255, 255, 255)
                Else
                    .ColorIndex = xlAutomatic
                End If
                .Bold = IsHeader
                .Superscript = False
                .Subscript = False
                .ThemeFont = xlThemeFontNone
                .TintAndShade = 0
            End With
            With .Interior
                If IsHeader Then
                    .Pattern = xlSolid
                    .PatternThemeColor = xlThemeColorDark1
                    .Color = 9868950
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                Else
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End If
            End With
            If IndexBis = 1 Or (IndexBis = 3 And IsHeader) Then
                .HorizontalAlignment = xlCenter
            Else
                .HorizontalAlignment = xlLeft
            End If
            .VerticalAlignment = xlTop
            
            If IndexBis = 3 Then
                .NumberFormat = "#,##0.00"" €"""
            Else
                .NumberFormat = "General"
            End If
        End With
    Next IndexBis
End Sub

Public Sub OpenUserForm()

    UserForm1.Show
    
End Sub

Public Sub EnleverBordures(CurrentCell As Range)
    With CurrentCell
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
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

    With CurrentCell
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            If AddTopBorder Then
                .Weight = xlMedium
            Else
                .Weight = xlHairline
            End If
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            If AddBottomBorder Then
                .Weight = xlMedium
            Else
                .Weight = xlHairline
            End If
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Font
            .Name = "Arial"
            If Italic Then
                .FontStyle = "Italic"
            Else
                .FontStyle = "Normal"
            End If
            .Bold = Bold
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            If BlueColor Then
                .Color = RGB(0, 102, 204)
            Else
                .ColorIndex = xlAutomatic
            End If
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        If CurrencyFormat Then
            .NumberFormat = "#,##0.00"" €"""
        Else
            .NumberFormat = "General"
        End If
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
    End With
End Sub

Public Sub CopieLogo(oldWorkbook As Workbook, NewWorkbook As Workbook, Name As String)

    Dim NewChargesSheet As Worksheet
    Dim OldChargesSheet As Worksheet
    Dim CurShape As Shape
    
    On Error Resume Next
    Set OldChargesSheet = oldWorkbook.Worksheets(Name)
    On Error GoTo 0
    If OldChargesSheet Is Nothing Then
        MsgBox "'" & Nom_Feuille_Cout_J_Salaire & "' n'a pas été trouvée dans " & oldWorkbook.Name
        Exit Sub
    End If
    
    On Error Resume Next
    Set NewChargesSheet = NewWorkbook.Worksheets(Name)
    On Error GoTo 0
    If NewChargesSheet Is Nothing Then
        MsgBox "'" & Name & "' n'a pas été trouvée dans " & NewWorkbook.Name
        Exit Sub
    End If
    
    OldChargesSheet.Activate
    OldChargesSheet.Select
    For Each CurShape In OldChargesSheet.Shapes
        ' msoPicture = 13
         If CurShape.Type = msoPicture Then
            CurShape.Copy
            NewChargesSheet.Activate
            NewChargesSheet.Select
            NewChargesSheet.Range(CurShape.TopLeftCell.address).Select
            NewChargesSheet.Paste
            ' Selection.Placement = xlMoveAndSize
            OldChargesSheet.Activate
            OldChargesSheet.Select
        End If
    Next CurShape

End Sub

Public Sub formatChargeCell(CurrentCell As Range, NoBorderOnRightAndLeft As Boolean)
    
    Dim Arr1
    Dim Arr2
    Dim IndexBis
    Dim VarTmp
    
    If NoBorderOnRightAndLeft Then
        Arr1 = Array(xlEdgeLeft, xlEdgeRight, xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
        Arr2 = Array(xlEdgeTop, xlEdgeBottom)
    Else
        Arr1 = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
        Arr2 = Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
    End If

    ' Format cell
    For IndexBis = 1 To 4
        With CurrentCell.Cells(1, IndexBis)
            For Each VarTmp In Arr1
                .Borders(VarTmp).LineStyle = xlNone
            Next VarTmp
            For Each VarTmp In Arr2
                With .Borders(VarTmp)
                    .LineStyle = xlContinuous
                    .ColorIndex = 1
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            Next VarTmp
            With .Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
        End With
    Next IndexBis
End Sub
Public Sub AddBottomBorder(CurrentCell As Range)
    With CurrentCell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub FormatFinancementCells(BaseCell As Range)
    Dim Index As Integer
    Dim TmpVar As Variant
    Dim VarTmp As Variant
    TmpVar = Array(xlEdgeBottom, xlEdgeTop)
    For Index = 2 To 3
        BaseCell.Cells(1, Index).Font.Bold = True
        With BaseCell.Cells(1, Index).Interior
            .Pattern = xlSolid
            .PatternColorIndex = 24
            .Color = 12632256
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        For Each VarTmp In TmpVar
            With BaseCell.Cells(1, Index).Borders(VarTmp)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next VarTmp
    Next Index
End Sub

Public Sub DefinirFormatConditionnelPourLesDossier( _
        SetOfRange As SetOfRange, _
        NBChantiers As Integer _
    )

    Dim CurrentCells As Range
    Dim CurrentFormatCondition As FormatCondition
    Dim FirstCell As Range
    Dim Index As Integer
    Dim ListConditions() As String
    Dim ListColors() As Variant

    ReDim ListConditions(1 To 4)
    ReDim ListColors(1 To 4)

    Set CurrentCells = Range( _
        SetOfRange.HeadCell.Cells(2, 1), _
        SetOfRange.EndCell.Cells(1, 3 + NBChantiers) _
    )
    
    ListConditions(1) = "DOSSIER_OK"
    ListColors(1) = 65280
    ListConditions(2) = "DOSSIER_FAVORABLE_ISSUE_INCERTAINE"
    ListColors(2) = 15773696
    ListConditions(3) = "DOSSIER_INCERTAIN"
    ListColors(3) = 49407
    ListConditions(4) = "DOSSIER_NON_DEPOSE"
    ListColors(4) = 65535
    
    CurrentCells.FormatConditions.Delete
    Set FirstCell = CurrentCells.Cells(1, 1).Cells(2, 1)
    For Index = 1 To 4
        FirstCell.Worksheet.Activate
        CurrentCells.Select
        Set CurrentFormatCondition = CurrentCells.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:= _
                "=SI(" & FirstCell.address( _
                    RowAbsolute:=False, _
                    ColumnAbsolute:=False, _
                    ReferenceStyle:=xlA1 _
                ) & "=" & ListConditions(Index) & ";VRAI();FAUX())" _
            )
        CurrentFormatCondition.StopIfTrue = True
        CurrentFormatCondition.SetFirstPriority
        With CurrentFormatCondition.Interior
            .PatternColorIndex = xlAutomatic
            .Color = ListColors(Index)
            .TintAndShade = 0
        End With
    Next Index
    Set CurrentFormatCondition = CurrentCells.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:= _
            "=MOD(LIGNE(" & FirstCell.address( _
                RowAbsolute:=False, _
                ColumnAbsolute:=False, _
                ReferenceStyle:=xlA1 _
            ) & ");2)" _
        )
    CurrentFormatCondition.StopIfTrue = False
    With CurrentFormatCondition.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(216, 216, 216)
        .TintAndShade = 0
    End With
End Sub

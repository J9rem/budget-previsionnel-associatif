Attribute VB_Name = "Specific"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' constantes
Public Const Nom_de_Fichier_Par_Defaut As String = "InCitu_Budget_Previsionnel_Associatif_v1_11_Excel"
Public Const Nom_de_Fichier_Par_Defaut_xls As String = "InCitu_Budget_Previsionnel_Associatif_v1_11_Excel.xls"
Public Const BackupDefaultExtension As String = ".xlsx"
Public Const intoOds As Boolean = False


' utils
Public Function choisirNomFicherASauvegarderSansMacro(ByRef FilePath As String) As Boolean
    
    Dim Adresse_dossier_courant As String
    Dim Fichier_De_Sauvegarde As String
    
    ' Default
    FilePath = ""
    
    ' Changement de dossier
    Adresse_dossier_courant = ThisWorkbook.Path
    ChDir (Adresse_dossier_courant)
    
    ' Fenêtre pour demander le nom du fichier de sauvegarde
    On Error Resume Next
    ' InitialFileName, FileFilter, FiltrerIndex, Title
    Fichier_De_Sauvegarde = Application.GetSaveAsFilename( _
        Nom_de_Fichier_Par_Defaut, _
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

Public Sub SetFormatForBudget(BaseCell As Range, HeadCell As Range)

    Dim IndexBis As Integer
    Dim TmpVar As Variant
    Dim VarTmp As Variant
    Dim Border
    
    For IndexBis = 1 To 3
        With BaseCell.Cells(1, IndexBis)
            If IndexBis <> 3 Then
                If HeadCell.Row = BaseCell.Row - 1 Then
                    TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal, xlEdgeBottom)
                Else
                    TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal, xlEdgeBottom, xlEdgeTop)
                End If
            Else
                TmpVar = Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
            End If
            For Each VarTmp In TmpVar
                .Borders(VarTmp).LineStyle = xlNone
            Next VarTmp
            
            If IndexBis <> 3 Then
                If HeadCell.Row = BaseCell.Row - 1 Then
                    TmpVar = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop)
                Else
                    TmpVar = Array(xlEdgeLeft, xlEdgeRight)
                End If
            Else
                TmpVar = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
            End If
            For Each VarTmp In TmpVar
                With .Borders(VarTmp)
                    .LineStyle = xlContinuous
                    .ColorIndex = 1
                    .Weight = xlThin
                    .TintAndShade = 0
                End With
            Next VarTmp
            With .Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 8
                .Strikethrough = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .Bold = False
                .Superscript = False
                .Subscript = False
                .ThemeFont = xlThemeFontNone
                .TintAndShade = 0
            End With
            With .Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
    Next IndexBis
End Sub

Public Sub OpenUserForm()

    UserForm1.Show
    
End Sub

Public Sub DefinirBordures(CurrentCell As Range, AddTopBorder As Boolean)

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
                .Weight = xlHairline
            Else
                .Weight = xlMedium
            End If
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub

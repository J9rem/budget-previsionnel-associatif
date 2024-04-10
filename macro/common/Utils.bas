Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la declaration de toutes les variables
Option Explicit

' Types
Type WbRevision
    Majeure As Integer
    Mineure As Integer
    Error As Boolean
End Type

Type TypeCharge
    Nom As String
    Index As Integer
    NomLong As String
End Type

Public Sub NotAvailable()
    MsgBox T_Development_In_Course
End Sub

Public Function FileExists(FilePath As String) As Boolean
    Dim dirResult As String
    
    dirResult = Dir(FilePath)
    If dirResult <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function


Public Function saveWorkbookAs(wb As Workbook, FolderName As String, FileName As String) As Boolean
    Dim previousCalculation
    Dim ShortFileName As String
    previousCalculation = Application.Calculation
    
    ChDir (FolderName)
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    If Right(FileName, 4) = ".xls" Then
        ' Remove Macro
        ShortFileName = Left(FileName, Len(FileName) - 4)
        wb.SaveAs FileName:=ShortFileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
        Workbooks.Open FileName:=FolderName & ShortFileName & ".xlsx"
        If FindOpenedWorkBook(ShortFileName & ".xlsx", wb) Then
            wb.SaveAs FileName:=FileName, FileFormat:=xlExcel8
            wb.Close SaveChanges:=False
        End If
        If FileExists(FolderName & ShortFileName & ".xlsx") Then
            deleteSilentWorkbook FolderName & ShortFileName & ".xlsx", ShortFileName & ".xlsx"
        End If
    Else
        wb.SaveAs FileName:=FileName, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
    End If
        
    Application.DisplayAlerts = True
    Application.Calculation = previousCalculation
    
    returnToCurrentPath
    saveWorkbookAs = True
End Function

Public Function closeSilentWorkbook(ByRef OpenedWorkbook As Workbook) As Boolean

    Dim previousCalculation
    previousCalculation = Application.Calculation
    
    If Not OpenedWorkbook Is Nothing Then

        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
        OpenedWorkbook.Close SaveChanges:=True
        Application.DisplayAlerts = True
        Application.Calculation = previousCalculation
    End If
    
    closeSilentWorkbook = True
End Function

Public Function deleteSilentWorkbook(FilePath As String, FileName As String) As Boolean

    Dim DestFolder As String
    
    DestFolder = Left(FilePath, Len(FilePath) - Len(FileName))
    ChDir (DestFolder)
    
    If FileExists(DestFolder & FileName) Then
        Application.DisplayAlerts = False
        DeleteFile FileName, DestFolder
        Application.DisplayAlerts = True
        If FileExists(DestFolder & FileName) Then
            deleteSilentWorkbook = False
        Else
            deleteSilentWorkbook = True
        End If
    Else
        deleteSilentWorkbook = True
    End If
    
    returnToCurrentPath
End Function

Public Function SaveFileNoMacro(FilePath As String) As Boolean

    Dim Extension As String
    Dim FolderName As String
    Dim MsgBoxResult As Integer
    Dim NewExtension As String
    Dim NewFileName As String
    Dim OpenedWorkbook As Workbook
    Dim SanitizedFilePath As String
    Dim ShortFileName As String
    
    Set OpenedWorkbook = Nothing

    SaveFileNoMacro = False
    If (FilePath = ThisWorkbook.Path) Then
        MsgBox Replace(T_NotPossibleToForceSave, "%n%", Chr(10))
    Else
        
        ' check extension type
        Extension = getExtension(FilePath)
        ShortFileName = getFileNameWithoutExtension(FilePath)
        FolderName = getFolder(FilePath, ShortFileName, Extension)
        If Extension = "xls" Then
            NewExtension = ".xls"
        Else
            NewExtension = BackupDefaultExtension
        End If
        
        NewFileName = ShortFileName & NewExtension
        SanitizedFilePath = FolderName & NewFileName
        
        If FileExists(SanitizedFilePath) Then
            MsgBoxResult = MsgBox( _
                Replace(T_Existing_File_What_To_Do, "%n%", Chr(10)), _
                vbYesNo)
            If MsgBoxResult <> vbYes And MsgBoxResult <> vbOK Then
                Exit Function
            End If
        End If
        
        If FindOpenedWorkBook(NewFileName, OpenedWorkbook) Then
            closeSilentWorkbook OpenedWorkbook
        End If
        
        If FileExists(SanitizedFilePath) Then
            If Not deleteSilentWorkbook(SanitizedFilePath, NewFileName) Then
                MsgBox "Impossible de supprimer " & SanitizedFilePath
                Exit Function
            End If
        End If
        If FileExists(FolderName & ShortFileName & ".xlsx") Then
            If Not deleteSilentWorkbook(FolderName & ShortFileName & ".xlsx", ShortFileName & ".xlsx") Then
                MsgBox "Impossible de supprimer " & FolderName & ShortFileName & ".xlsx"
                Exit Function
            End If
        End If
        
        If saveWorkBookAsCopyNoMacro(SanitizedFilePath, NewFileName, NewExtension, OpenedWorkbook) Then
            SaveFileNoMacro = True
        End If
    End If
End Function

Public Function getExtension(FilePath As String) As String

    Dim SplittedNewFileName As Variant
    
    SplittedNewFileName = Split(FilePath, ".")
    getExtension = SplittedNewFileName(UBound(SplittedNewFileName))
End Function

Public Function getFileNameWithoutExtension(FilePath As String) As String

    Dim tmpName As String
    Dim LastPart As String
    Dim SplittedNewFileName As Variant
    
    SplittedNewFileName = Split(FilePath, "/")
    LastPart = SplittedNewFileName(UBound(SplittedNewFileName))
    
    SplittedNewFileName = Split(LastPart, "\")
    tmpName = SplittedNewFileName(UBound(SplittedNewFileName))
    
    SplittedNewFileName = Split(tmpName, ".")
    LastPart = SplittedNewFileName(UBound(SplittedNewFileName))
    
    getFileNameWithoutExtension = Left(tmpName, Len(tmpName) - Len(LastPart) - 1)
    
End Function

Public Function getFolder(FilePath As String, ShortFileName As String, Ext As String) As String
    getFolder = Left(FilePath, Len(FilePath) - Len(ShortFileName) - 1 - Len(Ext))
End Function

Public Function FindWorkSheet(wb As Workbook, ByRef ws As Worksheet, SheetName As String) As Boolean

    Dim IntWs As Worksheet

    FindWorkSheet = False
    
    For Each IntWs In wb.Worksheets
        If IntWs.Name = SheetName Then
            FindWorkSheet = True
            Set ws = IntWs
        End If
    Next IntWs
    
End Function

Public Function inArray(StrValue As String, Arr As Variant) As Boolean

    Dim Index As Integer
    
    inArray = False
    For Index = LBound(Arr) To UBound(Arr)
        If StrValue = Arr(Index) Then
            inArray = True
        End If
    Next Index

End Function

Public Function inArrayInt(IntValue As Integer, Arr As Variant) As Boolean

    Dim Index As Integer
    
    inArrayInt = False
    For Index = LBound(Arr) To UBound(Arr)
        If IntValue = Arr(Index) Then
            inArrayInt = True
        End If
    Next Index

End Function

' search index of StrValue in Array
' @param String StrValue
' @param Varian Arr
' @return Integer -1 if not found
Public Function indexOfInArrayStr(StrValue As String, Arr As Variant) As Integer

    Dim Index As Integer
    
    indexOfInArrayStr = -1
    For Index = LBound(Arr) To UBound(Arr)
        If indexOfInArrayStr = -1 Then
            If StrValue = Arr(Index) Then
                indexOfInArrayStr = Index
            End If
        End If
    Next Index

End Function

Public Function AddWorksheetAtEnd(wb As Workbook, wsName As String) As Worksheet
    With wb.Worksheets
        .Add after:=.Item(.Count)
        .Item(.Count).Name = wsName
        Set AddWorksheetAtEnd = .Item(.Count)
    End With
End Function

Public Function archiveThisFile() As Boolean

    Dim FileName As String
    Dim Extension As String
    
    archiveThisFile = False
    
    FileName = ThisWorkbook.Name
    
    Extension = getExtension(FileName)
    
    FileName = Left(FileName, Len(FileName) - Len(Extension) - 1) & "-backup-" & Format(Now(), "yyyymdd_hhmmss") & "." & Extension
    
    If FileExists(ThisWorkbook.Path & "\" & FileName) Then
        MsgBox T_NotPossibleToSaveFileBecauseExisting
        Exit Function
    End If
    
    returnToCurrentPath
    SaveCopyAs ThisWorkbook, FileName, ThisWorkbook.Path
    
    If (InStr(ThisWorkbook.Path, "\") And FileExists(ThisWorkbook.Path & "\" & FileName)) Or _
        (InStr(ThisWorkbook.Path, "/") And FileExists(ThisWorkbook.Path & "/" & FileName)) Then
        archiveThisFile = True
    End If
End Function

Public Function DetecteVersion(wb As Workbook) As WbRevision

    Dim PersonnelSheet As Worksheet
    Dim InfoSheet As Worksheet
    Dim Cout_J_Sheet As Worksheet
    Dim Cout_Chantier As Worksheet
    Dim BaseCell As Range
    Dim ExplodedRevisions As Variant
    Dim StrValue As String
    Dim rev As WbRevision
    rev = getDefaultWbRevision()
    
    On Error Resume Next
    Set PersonnelSheet = wb.Worksheets(Nom_Feuille_Personnel)
    Set InfoSheet = wb.Worksheets(Nom_Feuille_Informations)
    Set Cout_J_Sheet = wb.Worksheets(Nom_Feuille_Cout_J_Salaire)
    Set Cout_Chantier = wb.Worksheets(Nom_Feuille_Budget_chantiers)
    On Error GoTo 0
    If Cout_J_Sheet Is Nothing Or Cout_Chantier Is Nothing Then
        rev.Error = True
        rev.Majeure = 0
        rev.Mineure = 0
    Else
        If InfoSheet Is Nothing Or PersonnelSheet Is Nothing Then
            rev.Error = False
            rev.Majeure = 0
            rev.Mineure = 0
        Else
            Set BaseCell = InfoSheet.Range("A:A").Find(Label_Version)
            If BaseCell Is Nothing Then
                rev.Error = True
                rev.Majeure = 1
                rev.Mineure = 0
            Else
                StrValue = BaseCell.Cells(1, 2).Value
                If StrValue = "" Then
                    rev.Error = True
                    rev.Majeure = 1
                    rev.Mineure = 0
                Else
                    ExplodedRevisions = Split(StrValue, ".")
                    rev.Error = False
                    rev.Majeure = ExplodedRevisions(0)
                    rev.Mineure = ExplodedRevisions(1)
                End If
            End If
        End If
    End If
    DetecteVersion = rev

End Function

Public Function FindTypeChargeIndex(StrValue As String) As Integer
    ' return 0 if not found
    Dim TypesCharges() As TypeCharge
    Dim Index As Integer
    Dim IndexFound As Integer
    Dim SimilarIndexFound As Integer
    Dim tmpName As String
    Dim OtherName As String
    Dim typeCh As TypeCharge
    Dim TypChIdx As TypeCharge
    
    TypesCharges = TypesDeCharges().Values
    IndexFound = 0
    For Index = 1 To UBound(TypesCharges)
        typeCh = TypesCharges(Index)
        tmpName = typeCh.NomLong
        OtherName = Replace(tmpName, "É", "E")
        OtherName = Replace(OtherName, "Ê", "E")
        OtherName = Replace(OtherName, "È", "E")
        OtherName = Replace(OtherName, "Ë", "E")
        OtherName = Replace(OtherName, "Ä", "A")
        OtherName = Replace(OtherName, "Â", "A")
        OtherName = Replace(OtherName, "Á", "A")
        OtherName = Replace(OtherName, "À", "A")
        OtherName = Replace(OtherName, "Ò", "O")
        OtherName = Replace(OtherName, "Ó", "O")
        OtherName = Replace(OtherName, "Õ", "O")
        OtherName = Replace(OtherName, "Ö", "O")
        OtherName = Replace(OtherName, "Ô", "O")
        If IndexFound = 0 And Len(tmpName) > 0 _
            And (Left(StrValue, Len(tmpName)) = tmpName _
            Or Left(StrValue, Len(OtherName)) = OtherName) Then
            IndexFound = Index
        End If
    Next Index
    If IndexFound > 0 Then
        SimilarIndexFound = 0
        typeCh = TypesCharges(IndexFound)
        For Index = 1 To UBound(TypesCharges)
            TypChIdx = TypesCharges(Index)
            If SimilarIndexFound = 0 And TypChIdx.Index = typeCh.Index Then
                SimilarIndexFound = Index
            End If
        Next Index
        If SimilarIndexFound > 0 Then
            IndexFound = SimilarIndexFound
        End If
    End If
    
    FindTypeChargeIndex = IndexFound

End Function


Public Function FindTypeFinancementIndex(StrValue As String) As Integer
    ' return 0 if not found
    Dim TypesFinancements() As String
    Dim Index As Integer
    Dim IndexFound As Integer
    Dim IndexAutre As Integer
    
    TypesFinancements = TypeFinancementsFromWb(ThisWorkbook)
    IndexFound = 0
    IndexAutre = 0
    For Index = 1 To UBound(TypesFinancements)
        If IndexFound = 0 And StrValue = TypesFinancements(Index) Then
            IndexFound = Index
        End If
        If IndexAutre = 0 And "Autres" = TypesFinancements(Index) Then
            IndexAutre = Index
        End If
    Next Index
    
    If IndexFound = 0 Then
        IndexFound = IndexAutre
    End If
    
    FindTypeFinancementIndex = IndexFound

End Function
    
Public Function FindTypeChargeIndexFromCode(IntegerValue As Integer) As Integer
    ' return 0 if not found
    Dim TypesCharges() As TypeCharge
    Dim Index As Integer
    Dim IndexFound As Integer
    Dim TypeCharge As TypeCharge
    Dim CurrentIndex As Integer
    
    On Error GoTo ManageLocalError:

    TypesCharges = TypesDeCharges().Values
    IndexFound = 0
    For Index = 1 To UBound(TypesCharges)
        TypeCharge = TypesCharges(Index)
        CurrentIndex = TypeCharge.Index
        If IndexFound = 0 And CurrentIndex = IntegerValue Then
            IndexFound = Index
        End If
    Next Index
ManageLocalError:
    If Err.Number > 0 Then
        IndexFound = 0
    End If
    On Error GoTo 0
    
    FindTypeChargeIndexFromCode = IndexFound

End Function

Public Function openWbSafe(ByRef wb As Workbook, FilePath As String) As Boolean

    Dim DestFolder As String
    Dim Extension As String
    Dim FileName As String
    Dim OpenedWorkbook As Workbook

     openWbSafe = False
     If FileExists(FilePath) Then
        FileName = getFileNameWithoutExtension(FilePath)
        Extension = getExtension(FilePath)
        
        DestFolder = Left(FilePath, Len(FilePath) - Len(FileName) - Len(Extension) - 1)
        ChDir (DestFolder)
    
        If FindOpenedWorkBook(FileName & "." & Extension, OpenedWorkbook) Then
            If OpenedWorkbook.Path & "\" = DestFolder Then
                Set wb = OpenedWorkbook
                openWbSafe = True
            End If
        Else
            Set wb = Workbooks.Open(FileName:=FilePath)
            If FindOpenedWorkBook(FileName & "." & Extension, OpenedWorkbook) Then
                Set wb = OpenedWorkbook
                If wb.Path & "\" = DestFolder Then
                    openWbSafe = True
                End If
            Else
                ' forced for .ods on LibreOffice
                openWbSafe = True
            End If
        End If
     End If

End Function

Public Function removeCrossRef(wb As Workbook, oldWb As Workbook) As Boolean
    
    On Error Resume Next
    wb.ChangeLink oldWb.Path & "\" & oldWb.Name, wb.Path & "\" & wb.Name, xlExcelLinks
    On Error GoTo 0
    removeCrossRef = True
End Function

Public Function returnToCurrentPath() As Boolean
    ChDir (ThisWorkbook.Path)
    returnToCurrentPath = True
End Function
Public Function CleanAddress(address As String) As String
    Dim pos
    Dim Tmp As String
    pos = InStr(1, address, "]")
    If pos Then
        If InStr(1, Left(address, pos), "'") Then
            Tmp = "'"
        Else
            Tmp = ""
        End If
        CleanAddress = Tmp & Mid(address, pos + 1)
    Else
        CleanAddress = address
    End If
End Function

Public Sub SetSilent()
    ' config to be faster
    On Error Resume Next ' pour eviter les erreurs LibreOffice
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = True
    Application.ScreenUpdating = False
    On Error GoTo 0
End Sub

Public Sub SetActive()
    On Error Resume Next ' pour eviter les erreurs LibreOffice
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Function geDefaultJoursChantiers(NBChantiers As Integer)
    Dim newArray() As Double
    Dim idx As Integer
    
    ReDim newArray(1 To NBChantiers)
    For idx = 1 To NBChantiers
        newArray(idx) = 0#
    Next idx
    geDefaultJoursChantiers = newArray
End Function

Public Function geDefaultJoursChantiersStr(NBChantiers As Integer)
    Dim newArray() As String
    Dim idx As Integer
    
    ReDim newArray(1 To NBChantiers)
    For idx = 1 To NBChantiers
        newArray(idx) = 0#
    Next idx
    geDefaultJoursChantiersStr = newArray
End Function

Public Sub Utils_Take_Snapshot()

    Dim CurrentWs As Worksheet
    Dim FoundSheet As Worksheet
    Dim NewName As String
    Dim NewNameSanitized As String
    Dim NewWs As Worksheet
    Dim SuffixIndex As Integer
    Dim wb As Workbook
    Dim wsName As String
    
    Set wb = ThisWorkbook
    Set CurrentWs = wb.ActiveSheet
    wsName = CurrentWs.Name

    NewName = Format(Date, "yyyymmdd-") & wsName
    NewNameSanitized = NewName

    SuffixIndex = 0

    On Error Resume Next
    Set FoundSheet = wb.Worksheets(NewName)
    On Error GoTo 0

    While Not (FoundSheet Is Nothing) And SuffixIndex < 1000
        SuffixIndex = SuffixIndex + 1
        NewNameSanitized = Left(NewName, 31 - Len("-" & SuffixIndex)) & "-" & SuffixIndex
        On Error Resume Next
        Set FoundSheet = Nothing
        Set FoundSheet = wb.Worksheets(NewNameSanitized)
        On Error GoTo 0
    Wend

    If SuffixIndex >= 1000 Then
        Exit Sub
    End If

    SetSilent
    Set NewWs = AddWorksheetAtEnd(wb, NewNameSanitized)

    replaceContentFromWorksheet NewWs, CurrentWs, True

    NewWs.Activate
    NewWs.Cells(1, 1).Select

    SetActive

End Sub


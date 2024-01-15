Attribute VB_Name = "Utils"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit


Public Sub NotAvailable()
    MsgBox "Patience, cette fonction est encore en cours de développement"
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
        MsgBox "Il n'est pas possible d'écraser le fichier courant" & Chr(10) & _
          "Veuillez réessayer avec un autre emplacement ou nom de fichier"
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
            MsgBoxResult = MsgBox("Le fichie cible existe déjà !" & Chr(10) & _
                "Faut-il l'écraser avec le nouveau ?", _
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

Public Function inArray(value As String, Arr As Variant) As Boolean

    Dim Index As Integer
    
    inArray = False
    For Index = LBound(Arr) To UBound(Arr)
        If value = Arr(Index) Then
            inArray = True
        End If
    Next Index

End Function

Public Function AddWorksheetAtEnd(wb As Workbook, WsName As String) As Worksheet
    With wb.Worksheets
        .Add after:=.Item(.Count)
        .Item(.Count).Name = WsName
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
        MsgBox "Impossible de sauvegarder le fichier de sauvegarde car il existe déjà"
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
    Dim value As String
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
                value = BaseCell.Cells(1, 2).value
                If value = "" Then
                    rev.Error = True
                    rev.Majeure = 1
                    rev.Mineure = 0
                Else
                    ExplodedRevisions = Split(value, ".")
                    rev.Error = False
                    rev.Majeure = ExplodedRevisions(0)
                    rev.Mineure = ExplodedRevisions(1)
                End If
            End If
        End If
    End If
    DetecteVersion = rev

End Function

Public Function FindTypeChargeIndex(value As String) As Integer
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
        If IndexFound = 0 And Len(tmpName) > 0 And (Left(value, Len(tmpName)) = tmpName Or Left(value, Len(OtherName)) = OtherName) Then
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


Public Function FindTypeFinancementIndex(value As String) As Integer
    ' return 0 if not found
    Dim TypesFinancements() As String
    Dim Index As Integer
    Dim IndexFound As Integer
    Dim IndexAutre As Integer
    
    TypesFinancements = TypeFinancementsFromWb(ThisWorkbook)
    IndexFound = 0
    IndexAutre = 0
    For Index = 1 To UBound(TypesFinancements)
        If IndexFound = 0 And value = TypesFinancements(Index) Then
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
    
Public Function FindTypeChargeIndexFromCode(value As Integer) As Integer
    ' return 0 if not found
    Dim TypesCharges() As TypeCharge
    Dim Index As Integer
    Dim IndexFound As Integer
    Dim TypeCharge As TypeCharge
    Dim currentIndex As Integer
    
    TypesCharges = TypesDeCharges().Values
    IndexFound = 0
    For Index = 1 To UBound(TypesCharges)
        TypeCharge = TypesCharges(Index)
        currentIndex = TypeCharge.Index
        If IndexFound = 0 And currentIndex = value Then
            IndexFound = Index
        End If
    Next Index
    
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
            On Error Resume Next
            If wb.Path & "\" = DestFolder Then
                openWbSafe = True
            End If
            On Error GoTo 0
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
    On Error Resume Next ' pour éviter les erreurs LibreOffice
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = True
    Application.ScreenUpdating = False
    On Error GoTo 0
End Sub

Public Sub SetActive()
    On Error Resume Next ' pour éviter les erreurs LibreOffice
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


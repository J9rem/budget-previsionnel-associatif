Attribute VB_Name = "Main"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' fonction qui fournit la date de sauvegarde du fichier
' pour pouvoir l'utiliser directement dans les cases

Public Function LastSaveDate() As String
  ' Volatile est présent pour indiquer que c'est une macro qui est recalculée en méme temps que le fichier
  Application.Volatile
  On Error Resume Next
  LastSaveDate = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
  On Error GoTo 0
End Function

Public Sub ExporterSansMacro()
    Dim FilePath As String
    Dim Erreur As Boolean
    
    If throwNotOdsNotInExcel() Then
        Erreur = True
        FilePath = ""
        If choisirNomFicherASauvegarderSansMacro(FilePath) Then
            If SaveFileNoMacro(FilePath) Then
                Erreur = False
            End If
        End If
        
        If Erreur Then
            MsgBox "Fichier non exporté"
        Else
            MsgBox "Fichier sauvegardé"
        End If
    End If
End Sub

Public Sub ImporterDesDonnees()

    Dim Erreur As Boolean
    Dim FilePath As String
    Dim MsgBoxResult As Integer
    Dim continue As Boolean
    
    Erreur = True
    
    If throwNotOdsNotInExcel() Then
        If choisirFichierAImporter(FilePath) Then
            MsgBoxResult = MsgBox( _
                "Faut-il faire une sauvegarde de ce fichier avant l'importation ?" & Chr(10) & _
                "Les données importées remplaceront toutes les données contenues dans le présent fichier.", _
                vbYesNo, _
                "Sauvegarder ce fichier ?" _
                )
            If MsgBoxResult <> vbYes And MsgBoxResult <> vbOK Then
                continue = True
            Else
                continue = archiveThisFile()
            End If
            If continue Then
                If importData(FilePath) Then
                    Erreur = False
                End If
            End If
        End If
        If Erreur Then
            MsgBox "Impossible d'importer le ficher"
        Else
            MsgBox "Fichier importé"
        End If
    End If
End Sub

Public Sub MettreAJourBudgetGlobalForCurrent()
    If throwNotOdsNotInExcel() Then
        MettreAJourBudgetGlobal ThisWorkbook
    End If
End Sub

Public Sub RetirerUnChantier()

    If throwNotOdsNotInExcel() Then
        ChangeUnChantier -1
    End If

End Sub
Public Sub AjouterUnChantier()

    If throwNotOdsNotInExcel() Then
        ChangeUnChantier 1
    End If

End Sub

Public Sub AjoutUnFinancement()
    Dim CurrentNBChantier As Integer
    Dim CurrentWs As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FinancementFantome As FinancementComplet
    
    If Not throwNotOdsNotInExcel() Then
        Exit Sub
    End If
    
    Set wb = ThisWorkbook
    
    SetSilent
    
    ' Current NB
    CurrentNBChantier = GetNbChantiers(wb)
    
    If CurrentNBChantier < 1 Then
        GoTo FinSub
    End If
    
    FinancementFantome.Status = False
    AjoutFinancement wb, CurrentNBChantier, FinancementFantome
    
FinSub:

    Set CurrentWs = wb.ActiveSheet
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Cells(1, 1).Select
    Next 'Ws
    CurrentWs.Activate
    
    SetActive

End Sub

Public Sub RetirerUnSalarie()

    If Not throwNotOdsNotInExcel() Then
        Exit Sub
    End If
    ChangeUnSalarie -1

End Sub
Public Sub AjouterUnSalarie()

    If Not throwNotOdsNotInExcel() Then
        Exit Sub
    End If
    ChangeUnSalarie 1

End Sub

' Macro pour insérer une dépense
Public Sub InsererUneDepense()
    Dim NBChantiers As Integer
    NBChantiers = InsererUneDepenseInternal()
    InsererUneDepenseInternalReal NBChantiers
End Sub


Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

' SPDX-License-Identifier: EUPL-1.2
' Module pour tester si la macro s'exécute en Excel ou non

' Fonction qui vérifie si le fichier est ouvert sur Excel
' sinon affiche un message d'erreur

Public Function throwNotOdsNotInExcel() As Boolean

    On Error GoTo ErreurNestPasExcel
    If xlExcel8 = 56 Then
        throwNotOdsNotInExcel = True
    Else
        throwNotOdsNotInExcel = False
    End If
    GoTo FinErreurExcel
ErreurNestPasExcel:
    throwNotOdsNotInExcel = False
FinErreurExcel:
    On Error GoTo 0
    
    If intoOds Then
        throwNotOdsNotInExcel = True
    End If
    If Not throwNotOdsNotInExcel Then
	    MsgBox "Vous avez ouvert le fichier avec macro avec un autre logiciel qu'Excel" & Chr(10) & _
	      "Le fichier n'est pas fait pour ceci. Rien ne va se passer" & Chr(10) & _
	      "Vous devriez réouvir ce fichier avec Excel pour l'exporter sans macro ou " & Chr(10) & _
	      "l'enregistrer sans macro directement avec votre logiciel"
    End If
End Function



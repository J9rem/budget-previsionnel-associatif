Attribute VB_Name = "Constructors"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la d�claration de toutes les variables
Option Explicit


Public Function TypeFinancements()
    Dim ArrayTmp(0 To 9) As String
    ArrayTmp(0) = ""
    ArrayTmp(1) = "�tat"
    ArrayTmp(2) = "R�gion"
    ArrayTmp(3) = "Communes et intercommunalit�s"
    ArrayTmp(4) = "�tablissements publics"
    ArrayTmp(5) = "Organismes sociaux"
    ArrayTmp(6) = "Fonds europ�ens"
    ArrayTmp(7) = "ASP (emplois aid�s)"
    ArrayTmp(8) = "Fondation"
    ArrayTmp(9) = "Autres"

    TypeFinancements = ArrayTmp
End Function

Public Function TypeStatut()
    Dim ArrayTmp(0 To 4) As String
    ArrayTmp(0) = ""
    ArrayTmp(1) = "Dossier non encore d�pos�, issue et montant incertains"
    ArrayTmp(2) = "Dossier d�pos�, issue et montant incertains"
    ArrayTmp(3) = "Dossier d�pos�, issue favorable et montant incertain"
    ArrayTmp(4) = "Dossier d�pos�, issue et montant certain"

    TypeStatut = ArrayTmp
End Function

Public Function GetNewTypeDeCharge(Title As String, Index As Integer) As TypeCharge
    Dim TmpCharge As TypeCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = Title
    TmpCharge.Index = Index
    If (Title = "") Then
        TmpCharge.NomLong = ""
    Else
        TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    End If

    GetNewTypeDeCharge = TmpCharge
End Function

Public Function TypesDeCharges() As TypesCharges
    Dim ArrayTmp(0 To 10) As TypeCharge
    Dim TmpTypesCharges As TypesCharges
    TmpTypesCharges = getDefaultTypesCharges()
    
    ' Inconnue
    ArrayTmp(0) = GetNewTypeDeCharge("", 0)
    
    ArrayTmp(1) = GetNewTypeDeCharge("Achats", 60)
    ArrayTmp(2) = GetNewTypeDeCharge("Services ext�rieurs", 61)
    ArrayTmp(3) = GetNewTypeDeCharge("Autres services ext�rieurs", 62)
    ArrayTmp(4) = GetNewTypeDeCharge("Imp�ts et taxes", 63)
    ArrayTmp(5) = GetNewTypeDeCharge("Charges de personnel", 64)
    ArrayTmp(6) = GetNewTypeDeCharge("Autres charges de gestion courante", 65)
    ArrayTmp(7) = GetNewTypeDeCharge("Charges financi�res", 66)
    ArrayTmp(8) = GetNewTypeDeCharge("Charges exceptionnelles", 67)
    ArrayTmp(9) = GetNewTypeDeCharge("Dotation aux amortissements", 68)
    ArrayTmp(10) = GetNewTypeDeCharge("Les imp�ts sur les b�n�fices et assimil�s", 69)
    
    TmpTypesCharges.Values = ArrayTmp

    TypesDeCharges = TmpTypesCharges
End Function

Public Function DefaultSheetsNames()
    Dim SheetNames(0 To 11) As String
    SheetNames(0) = Nom_Feuille_Informations
    SheetNames(1) = Nom_Feuille_Personnel
    SheetNames(2) = Nom_Feuille_Cout_J_Salaire
    SheetNames(3) = Nom_Feuille_Charges
    SheetNames(4) = Nom_Feuille_Budget_chantiers
    SheetNames(5) = Nom_Feuille_Budget_global
    SheetNames(6) = Nom_Feuille_Budget_chantiers_realise
    SheetNames(7) = Nom_Feuille_Provisions
    SheetNames(8) = Nom_Feuille_Eupl
    SheetNames(9) = Nom_Feuille_CC_by_SA
    SheetNames(10) = Nom_Feuille_FAQ
    SheetNames(11) = Nom_Feuille_Versions

    DefaultSheetsNames = SheetNames
End Function




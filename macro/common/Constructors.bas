Attribute VB_Name = "Constructors"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit


Public Function TypeFinancements()
    Dim ArrayTmp(0 To 9) As String
    ArrayTmp(0) = ""
    ArrayTmp(1) = "État"
    ArrayTmp(2) = "Région"
    ArrayTmp(3) = "Communes et intercommunalités"
    ArrayTmp(4) = "Établissements publics"
    ArrayTmp(5) = "Organismes sociaux"
    ArrayTmp(6) = "Fonds européens"
    ArrayTmp(7) = "ASP (emplois aidés)"
    ArrayTmp(8) = "Fondation"
    ArrayTmp(9) = "Autres"

    TypeFinancements = ArrayTmp
End Function

Public Function TypeStatut()
    Dim ArrayTmp(0 To 4) As String
    ArrayTmp(0) = ""
    ArrayTmp(1) = "Dossier non encore déposé, issue et montant incertains"
    ArrayTmp(2) = "Dossier déposé, issue et montant incertains"
    ArrayTmp(3) = "Dossier déposé, issue favorable et montant incertain"
    ArrayTmp(4) = "Dossier déposé, issue et montant certain"

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
    ArrayTmp(2) = GetNewTypeDeCharge("Services extérieurs", 61)
    ArrayTmp(3) = GetNewTypeDeCharge("Autres services extérieurs", 62)
    ArrayTmp(4) = GetNewTypeDeCharge("Impôts et taxes", 63)
    ArrayTmp(5) = GetNewTypeDeCharge("Charges de personnel", 64)
    ArrayTmp(6) = GetNewTypeDeCharge("Autres charges de gestion courante", 65)
    ArrayTmp(7) = GetNewTypeDeCharge("Charges financières", 66)
    ArrayTmp(8) = GetNewTypeDeCharge("Charges exceptionnelles", 67)
    ArrayTmp(9) = GetNewTypeDeCharge("Dotations aux amortissements", 68)
    ArrayTmp(10) = GetNewTypeDeCharge("Les impôts sur les bénéfices et assimilés", 69)
    
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




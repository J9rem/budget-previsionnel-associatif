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

Public Function TypesDeCharges() As TypesCharges
    Dim ArrayTmp(0 To 10) As typeCharge
    Dim TmpCharge As typeCharge
    Dim TmpTypesCharges As TypesCharges
    TmpTypesCharges = getDefaultTypesCharges()
    
    ' Inconnue
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = ""
    TmpCharge.Index = 0
    TmpCharge.NomLong = ""
    ArrayTmp(0) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Achats"
    TmpCharge.Index = 60
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(1) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Services extérieurs"
    TmpCharge.Index = 61
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(2) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Autres services extérieurs"
    TmpCharge.Index = 62
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(3) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Impôts et taxes"
    TmpCharge.Index = 63
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(4) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Charges de personnel"
    TmpCharge.Index = 64
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(5) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Autres charges de gestion courante"
    TmpCharge.Index = 65
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(6) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Charges financières"
    TmpCharge.Index = 66
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(7) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Charges exceptionnelles"
    TmpCharge.Index = 67
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(8) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Dotation aux amortissements"
    TmpCharge.Index = 68
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(9) = TmpCharge
    
    TmpCharge = getDefaultTypeCharge()
    TmpCharge.Nom = "Dotations amortissement"
    TmpCharge.Index = 68
    TmpCharge.NomLong = TmpCharge.Index & " - " & UCase(TmpCharge.Nom)
    ArrayTmp(10) = TmpCharge
    
    TmpTypesCharges.Values = ArrayTmp

    TypesDeCharges = TmpTypesCharges
End Function

Public Function DefaultSheetsNames()
    Dim SheetNames(0 To 9) As String
    SheetNames(0) = Nom_Feuille_Informations
    SheetNames(1) = Nom_Feuille_Personnel
    SheetNames(2) = Nom_Feuille_Cout_J_Salaire
    SheetNames(3) = Nom_Feuille_Charges
    SheetNames(4) = Nom_Feuille_Budget_chantiers
    SheetNames(5) = Nom_Feuille_Budget_global
    SheetNames(6) = Nom_Feuille_Eupl
    SheetNames(7) = Nom_Feuille_CC_by_SA
    SheetNames(8) = Nom_Feuille_FAQ
    SheetNames(9) = Nom_Feuille_Versions

    DefaultSheetsNames = SheetNames
End Function




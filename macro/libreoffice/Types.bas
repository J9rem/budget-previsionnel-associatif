﻿Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' Types
Type WbRevision
    Majeure As Integer
    Mineure As Integer
    Error As Boolean
End Type

Type Informations
    Annee As Integer
    AnneeFormula As String
    ConventionCollective As String
    NBConges As Integer
    NBCongesFormula As String
    Pentecote As Boolean
    NBRTT As Integer
    NBRTTFormula As String
    NBJoursSpeciaux As Integer
    NBJoursSpeciauxFormula As String
End Type

Type DepenseChantier
    Nom As String
    Valeur As Double
    BaseCell As Range
    Formula As String
    ValeurReal As Double
    BaseCellReal As Range
    FormulaReal As String
End Type

Type Financement
    Nom As String
    TypeFinancement As Integer ' Index in TypeFinancements
    Valeur As Double
    ValeurReal As Double
    Formula As String
    FormulaReal As String
    Statut As Integer ' 0 = empty
    BaseCell As Range
    BaseCellReal As Range
    IndexInProvisions As Integer ' 0 = not concerned
End Type

Type Chantier
    Nom As String
    Depenses() As DepenseChantier
    Financements() As Financement
    AutoFinancementStructure As Double
    AutoFinancementStructureFormula As String
    AutoFinancementAutres As Double
    AutoFinancementAutresFormula As String
    LivrablesL1 As String
    LivrablesL2 As String
    LivrablesL3 As String
    LivrablesL4 As String
    DetailsL1 As String
    DetailsL2 As String
    DetailsL3 As String
    DetailsL4 As String
    DetailsL5 As String
End Type

Type SetOfChantiers
    Chantiers() As Chantier
End Type

Type FinancementComplet
    Financements() As Financement
    Status As Boolean
End Type

Type DonneesSalarie
    Erreur As Boolean
    Prenom As String
    Nom As String
    TauxDeTempsDeTravail As Double
    TauxDeTempsDeTravailFormula As String
    MasseSalarialeAnnuelle As Double
    MasseSalarialeAnnuelleFormula As String
    TauxOperateur As Double
    TauxOperateurFormula As String
    JoursChantiers() As Double ' Tableau de temps de chantiers même index que le tableau Chantiers
    JoursChantiersFormula() As String
    JoursChantiersReal() As Double
    JoursChantiersFormulaReal() As String
End Type

Type Charge
    Nom As String
    IndexTypeCharge As Integer
    CurrentYearValue As Double
    CurrentRealizedYearValue As Double
    PreviousYearValue As Double
    PreviousN2YearValue As Double
    ChargeCell As Range
    Category As Integer
End Type

Type SetOfCharges
    Charges() As Charge
End Type

Type Provision
    NomDuFinanceur As String
    RangeForTitle As Range ' Nothing if not linked
    NBYears As Integer ' 0 if empty
    FirstYear As Integer ' ex 2020
    WaitedValues() As Double ' for each year
    RangeForLastYearWaitedValue As Range ' Nothing if not linked
    PayedValues() As Double ' one dimension array
      ' NBYears elements for first year
      ' NBYears  - 1 elements for next year
      ' etc up to LastYear
    RangeForLastYearPayedValue As Range ' Nothing if not linked
    RetrievalTenPercent() As Double ' one dimension array
      ' NBYears - 1 elements for first year
      ' NBYears - 2 elements for next year
      ' etc up to LastYear - 1
    RetrievalTenPercentFormula() As String ' one dimension array
      ' NBYears - 1 elements for first year
      ' NBYears - 2 elements for next year
      ' etc up to LastYear - 1
End Type

Type Data
    Salaries() As DonneesSalarie
    Chantiers() As Chantier
    Informations As Informations
    Charges() As Charge
    TitlesForChargesCat() As String
    Provisions() As Provision
End Type

Type TypeCharge
    Nom As String
    Index As Integer
    NomLong As String
End Type

Type TypesCharges
    Values() As TypeCharge
End Type
    
Type NBAndRange
    NB As Integer
    Range As Range
End Type

Type SetOfRange
    EndCell As Range
    EndCellReal As Range
    HeadCell As Range
    HeadCellReal As Range
    ResultCell As Range
    ResultCellReal As Range
    Status As Boolean
    StatusReal As Boolean
    ChantierSheet As Worksheet
    ChantierSheetReal As Worksheet
End Type

Type SetOfCellsCategories
    Cells(60 To 68) As Range
    TotalCell As Range
End Type

Type ListOfCptResult
    Suffix() As String
    Formula() As String
    WithReal() As Boolean
End Type

Public Function getDefaultWbRevision() As WbRevision
    
    Dim wbRevision as new WbRevision
    wbRevision.Majeure = 0
    wbRevision.Mineure = 0
    wbRevision.Error = False
    
    getDefaultWbRevision = wbRevision

End Function

Public Function getDefaultData() As Data
    Dim Data As New Data
    Dim EmptyArrayDonneesSalarie() As DonneesSalarie
    Dim EmptyChantiers() As Chantier
    Dim EmptyCharges() As Charge
    Dim EmptyProvisions() As Provision
    Dim EmptyTitles(1 To 3) As String
    ReDim EmptyArrayDonneesSalarie(0)
    ReDim EmptyChantiers(0)
    ReDim EmptyCharges(0)
    ReDim EmptyProvisions(0)
    Dim Informations As Informations
    
    Data.Salaries = EmptyArrayDonneesSalarie
    Data.Chantiers = EmptyChantiers
    Data.Informations = getDefaultInformations()
    Data.Charges = EmptyCharges
    Data.TitlesForChargesCat = EmptyTitles
    Data.Provisions = EmptyProvisions

    getDefaultData = Data
End Function

Public Function getDefaultDonneesSalarie() As DonneesSalarie

	Dim Donnees As New DonneesSalarie
    Dim EmptyArray() As Double
    Dim EmptyArray2() As Double
    Dim EmptyArrayStr() As String
    Dim EmptyArrayStr2() As String
    
    ReDim EmptyArray(0)
    ReDim EmptyArray2(0)
    ReDim EmptyArrayStr(0)
    ReDim EmptyArrayStr2(0)

    Donnees.Erreur = True
    Donnees.Prenom = ""
    Donnees.Nom = ""
    Donnees.TauxDeTempsDeTravail = 0
    Donnees.TauxDeTempsDeTravailFormula = ""
    Donnees.TauxOperateur = 0
    Donnees.TauxOperateurFormula = ""
    Donnees.MasseSalarialeAnnuelle = 0
    Donnees.MasseSalarialeAnnuelleFormula = ""
    Donnees.JoursChantiers = EmptyArray
    Donnees.JoursChantiersReal = EmptyArray2
    Donnees.JoursChantiersFormula = EmptyArrayStr
    Donnees.JoursChantiersFormulaReal = EmptyArrayStr2
    
    getDefaultDonneesSalarie = Donnees

End Function

Public Function getDefaultChantier(NbDefaultDepenses As Integer) As Chantier

	Dim Chantier As New Chantier
    Dim EmptyFinancements() As Financement
    ReDim EmptyFinancements(0)
    
    Chantier.Nom = ""
    Chantier.Depenses = getDefaultDepenses(NbDefaultDepenses)
    Chantier.Financements = EmptyFinancements
    Chantier.AutoFinancementStructure = 0
    Chantier.AutoFinancementStructureFormula = ""
    Chantier.AutoFinancementAutres = 0
    Chantier.AutoFinancementAutresFormula = ""
    Chantier.LivrablesL1 = ""
    Chantier.LivrablesL2 = ""
    Chantier.LivrablesL3 = ""
    Chantier.LivrablesL4 = ""
    Chantier.DetailsL1 = ""
    Chantier.DetailsL2 = ""
    Chantier.DetailsL3 = ""
    Chantier.DetailsL4 = ""
    Chantier.DetailsL5 = ""
    getDefaultChantier = Chantier
End Function


Public Function getDefaultCharge() As Charge
    Dim ch as new Charge
    getDefaultCharge = ch
End Function

Public Function getDefaultInformations() As Informations
    
    Dim Informations As New Informations
    
    Informations.Annee = Format(Date, "yyyy")
    Informations.AnneeFormula = ""
    Informations.ConventionCollective = ""
    Informations.NBConges = 25
    Informations.NBCongesFormula = ""
    Informations.NBJoursSpeciaux = 0
    Informations.NBJoursSpeciauxFormula = ""
    Informations.Pentecote = True
    Informations.NBRTT = 0
    Informations.NBRTTFormula = ""
    
    getDefaultInformations = Informations

End Function

Public Function getDefaultFinancementComplet() As FinancementComplet
    
    Dim fin As New FinancementComplet
    Dim ArrayTmp() as Financements
    
    fin.Financements = ArrayTmp 
    fin.status = false
    
    getDefaultFinancementComplet = fin

End Function

Public Function getDefaultTypeCharge() As TypeCharge
    Dim ch as new TypeCharge
    getDefaultTypeCharge = ch
End Function

Public Function getDefaultTypesCharges() As TypesCharges
    Dim ch as new TypesCharges
    getDefaultTypesCharges = ch
End Function

Public Function getDefaulNBAndRange() As NBAndRange
	Dim res as New NBAndRange
	getDefaulNBAndRange = res
End Function

Public Function getDefaultDepenses(Nb As Integer)

    Dim ArrayTmp() As DepenseChantier
    Dim IndexChantiers As Integer
    ReDim ArrayTmp(1 To Nb)
    
    For IndexChantiers = 1 To Nb
        ArrayTmp(IndexChantiers) = New DepenseChantier
    Next IndexChantiers
    
    getDefaultDepenses = ArrayTmp
    
End Function

Public Function getDefaultFinancements(Nb As Integer)

    Dim ArrayTmp() As Financement
    Dim Index As Integer
    ReDim ArrayTmp(1 To Nb)
    
    For Index = 1 To Nb
        ArrayTmp(Index) = New Financement
    Next Index
    
    getDefaultFinancements = ArrayTmp
    
End Function

Public Function getDefaultProvision(NBYears As Integer) As Provision

    Dim Provision As New Provision

    Provision = Provisions_Init(Provision, NBYears)
    getDefaultProvision = Provision

End Function


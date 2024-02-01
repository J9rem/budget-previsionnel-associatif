Rem Attribute VBA_ModuleType=VBAModule
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
End Type

Type Financement
    Nom As String
    TypeFinancement As Integer ' Index in TypeFinancements
    Valeur As Double
    Statut As Integer ' 0 = empty
    BaseCell As Range
End Type

Type Chantier
    Nom As String
    Depenses() As DepenseChantier
    Financements() As Financement
    AutoFinancementStructure As Double
    AutoFinancementAutres As Double
    AutoFinancementStructureAnneesPrecedentes As Double
    AutoFinancementAutresAnneesPrecedentes As Double
    CAanneesPrecedentes As Double
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

Type Data
    Salaries() As DonneesSalarie
    Chantiers() As Chantier
    Informations As Informations
    Charges() As Charge
    TitlesForChargesCat() As String
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
    HeadCell As Range
    ResultCell As Range
    Status As Boolean
    ChantierSheet As Worksheet
End Type

Type SetOfCellsCategories
    Cells(60 To 68) As Range
    TotalCell As Range
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
    Dim EmptyTitles(1 To 3) As String
    ReDim EmptyArrayDonneesSalarie(0)
    ReDim EmptyChantiers(0)
    ReDim EmptyCharges(0)
    Dim Informations As Informations
    
    Data.Salaries = EmptyArrayDonneesSalarie
    Data.Chantiers = EmptyChantiers
    Data.Informations = getDefaultInformations()
    Data.Charges = EmptyCharges
    Data.TitlesForChargesCat = EmptyTitles

    getDefaultData = Data
End Function

Public Function getDefaultDonneesSalarie() As DonneesSalarie

	Dim Donnees As New DonneesSalarie
    Dim EmptyArray() As Double
    
    ReDim EmptyArray(0)

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
    Chantier.AutoFinancementStructureAnneesPrecedentes = 0
    Chantier.AutoFinancementAutres = 0
    Chantier.AutoFinancementAutresAnneesPrecedentes = 0
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
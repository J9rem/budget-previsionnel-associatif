Attribute VB_Name = "Types"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la déclaration de toutes les variables
Option Explicit

' Types
Public Type WbRevision
    Majeure As Integer
    Mineure As Integer
    Error As Boolean
End Type

Public Type Informations
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

Public Type DepenseChantier
    Nom As String
    Valeur As Double
    BaseCell As Range
End Type

Public Type Financement
    Nom As String
    TypeFinancement As Integer ' Index in TypeFinancements
    Valeur As Double
    Statut As Integer ' 0 = empty
    BaseCell As Range
End Type

Public Type Chantier
    Nom As String
    Depenses() As DepenseChantier
    Financements() As Financement
    AutoFinancementStructure As Double
    AutoFinancementAutres As Double
    AutoFinancementStructureAnneesPrecedentes As Double
    AutoFinancementAutresAnneesPrecedentes As Double
    CAanneesPrecedentes As Double
End Type

Public Type SetOfChantiers
    Chantiers() As Chantier
End Type

Public Type FinancementComplet
    Financements() As Financement
    Status As Boolean
End Type

Public Type DonneesSalarie
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

Public Type Charge
    Nom As String
    IndexTypeCharge As Integer
    CurrentYearValue As Double
    PreviousYearValue As Double
    PreviousN2YearValue As Double
    ChargeCell As Range
End Type

Public Type SetOfCharges
    Charges() As Charge
End Type

Public Type Data
    Salaries() As DonneesSalarie
    Chantiers() As Chantier
    Informations As Informations
    Charges() As Charge
End Type

Public Type TypeCharge
    Nom As String
    Index As Integer
    NomLong As String
End Type

Public Type TypesCharges
    Values() As TypeCharge
End Type
    
Public Type NBAndRange
    NB As Integer
    Range As Range
End Type

Public Function getDefaultWbRevision() As WbRevision
    
    Dim WbRevision As WbRevision
    WbRevision.Majeure = 0
    WbRevision.Mineure = 0
    WbRevision.Error = False
    
    getDefaultWbRevision = WbRevision

End Function

Public Function getDefaultData(Data As Data) As Data
    Dim EmptyArrayDonneesSalarie() As DonneesSalarie
    Dim EmptyChantiers() As Chantier
    Dim EmptyCharges() As Charge
    ReDim EmptyArrayDonneesSalarie(0)
    ReDim EmptyChantiers(0)
    ReDim EmptyCharges(0)
    Dim Informations As Informations
    
    Data.Salaries = EmptyArrayDonneesSalarie
    Data.Chantiers = EmptyChantiers
    Data.Informations = getDefaultInformations()
    Data.Charges = EmptyCharges

    getDefaultData = Data
End Function

Public Function getDefaultDonneesSalarie() As DonneesSalarie

    Dim Donnees As DonneesSalarie
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

    Dim Chantier As Chantier
    Dim ArrayDepenses() As DepenseChantier
    ReDim ArrayDepenses(1 To NbDefaultDepenses)
    Dim EmptyFinancements() As Financement
    ReDim EmptyFinancements(0)
    
    Chantier.Nom = ""
    Chantier.Depenses = ArrayDepenses
    Chantier.Financements = EmptyFinancements
    Chantier.AutoFinancementStructure = 0
    Chantier.AutoFinancementStructureAnneesPrecedentes = 0
    Chantier.AutoFinancementAutres = 0
    Chantier.AutoFinancementAutresAnneesPrecedentes = 0
    getDefaultChantier = Chantier
End Function

Public Function getDefaultCharge() As Charge
    Dim ch As Charge
    getDefaultCharge = ch
End Function

Public Function getDefaultInformations() As Informations
    
    Dim Informations As Informations
    
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
    
    Dim fin As FinancementComplet
    
    getDefaultFinancementComplet = fin

End Function

Public Function getDefaultTypeCharge() As TypeCharge
    Dim ch As TypeCharge
    getDefaultTypeCharge = ch
End Function

Public Function getDefaultTypesCharges() As TypesCharges
    Dim ch As TypesCharges
    getDefaultTypesCharges = ch
End Function

Public Function getDefaulNBAndRange() As NBAndRange
    Dim res As NBAndRange
    getDefaulNBAndRange = res
End Function

Public Function getDefaultDepenses(NB As Integer) As DepenseChantier()

    Dim ArrayTmp() As DepenseChantier
    Dim IndexChantiers As Integer
    Dim DefaultDepenseChantier As DepenseChantier
    ReDim ArrayTmp(1 To NB)
    
    For IndexChantiers = 1 To NB
        ArrayTmp(IndexChantiers) = DefaultDepenseChantier
    Next IndexChantiers
    
    getDefaultDepenses = ArrayTmp
    
End Function


Public Function getDefaultFinancements(NB As Integer) As Financement()

    Dim ArrayTmp() As Financement
    Dim Index As Integer
    Dim DefaultFinancement As Financement
    ReDim ArrayTmp(1 To NB)
    
    For Index = 1 To NB
        ArrayTmp(Index) = DefaultFinancement
    Next Index
    
    getDefaultFinancements = ArrayTmp
    
End Function

Attribute VB_Name = "Types"
' SPDX-License-Identifier: EUPL-1.2
' Pour forcer la d�claration de toutes les variables
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
    Formula As String
    ValeurReal As Double
    BaseCellReal As Range
    FormulaReal As String
End Type

Public Type Financement
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

Public Type Chantier
    Nom As String
    Depenses() As DepenseChantier
    Financements() As Financement
    AutoFinancementStructure As Double
    AutoFinancementStructureFormula As String
    AutoFinancementAutres As Double
    AutoFinancementAutresFormula As String
    AutoFinancementStructureAnneesPrecedentes As Double
    AutoFinancementStructureAnneesPrecedentesFormula As String
    AutoFinancementAutresAnneesPrecedentes As Double
    AutoFinancementAutresAnneesPrecedentesFormula As String
    CAanneesPrecedentes As Double
    CAanneesPrecedentesFormula As String
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
    JoursChantiers() As Double ' Tableau de temps de chantiers m�me index que le tableau Chantiers
    JoursChantiersFormula() As String
    JoursChantiersReal() As Double
    JoursChantiersFormulaReal() As String
End Type

Public Type Charge
    Nom As String
    IndexTypeCharge As Integer
    CurrentYearValue As Double
    CurrentRealizedYearValue As Double
    PreviousYearValue As Double
    PreviousN2YearValue As Double
    ChargeCell As Range
    Category As Integer
End Type

Public Type SetOfCharges
    Charges() As Charge
End Type

Public Type Provision
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
End Type

Public Type Data
    Salaries() As DonneesSalarie
    Chantiers() As Chantier
    Informations As Informations
    Charges() As Charge
    TitlesForChargesCat() As String
    Provisions() As Provision
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

Public Type SetOfRange
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

Public Type SetOfCellsCategories
    Cells(60 To 68) As Range
    TotalCell As Range
End Type

Public Type ListOfCptResult
    Suffix() As String
    Formula() As String
    WithReal() As Boolean
End Type

Public Function getDefaultWbRevision() As WbRevision
    
    Dim WbRevision As WbRevision
    WbRevision.Majeure = 0
    WbRevision.Mineure = 0
    WbRevision.Error = False
    
    getDefaultWbRevision = WbRevision

End Function

Public Function getDefaultData() As Data
    Dim Data As Data
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

    Dim Donnees As DonneesSalarie
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

    Dim Chantier As Chantier
    Dim ArrayDepenses() As DepenseChantier
    ReDim ArrayDepenses(1 To NbDefaultDepenses)
    Dim EmptyFinancements() As Financement
    ReDim EmptyFinancements(0)
    
    Chantier.Nom = ""
    Chantier.Depenses = ArrayDepenses
    Chantier.Financements = EmptyFinancements
    Chantier.AutoFinancementStructure = 0
    Chantier.AutoFinancementStructureFormula = ""
    Chantier.AutoFinancementStructureAnneesPrecedentes = 0
    Chantier.AutoFinancementStructureAnneesPrecedentesFormula = ""
    Chantier.AutoFinancementAutres = 0
    Chantier.AutoFinancementAutresFormula = ""
    Chantier.AutoFinancementAutresAnneesPrecedentes = 0
    Chantier.AutoFinancementAutresAnneesPrecedentesFormula = ""
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

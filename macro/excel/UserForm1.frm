VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Nouveau financeur"
   ClientHeight    =   2832
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5604
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Validate_Click()
    Dim Nom As String
    Dim TypeFinancement As Integer
    Dim FinancementFantome As FinancementComplet
    Dim TypesFinancements() As String
    
    Dim CurrentNBChantier As Integer
    Dim wb As Workbook
    
    TypesFinancements = TypeFinancementsFromWb(ThisWorkbook)
    
    
    Nom = Me.NomFinancement
    If Me.TypeEtat Then
        TypeFinancement = FindTypeFinancementIndex(Me.TypeEtat.Caption)
    Else
        If Me.TypeRegion Then
            TypeFinancement = FindTypeFinancementIndex(Me.TypeRegion.Caption)
        Else
            If Me.TypeComIntercom Then
                TypeFinancement = FindTypeFinancementIndex(Me.TypeComIntercom.Caption)
            Else
                If Me.TypeEtPubic Then
                    TypeFinancement = FindTypeFinancementIndex(Me.TypeEtPubic.Caption)
                Else
                    If Me.TypeOrgaSociaux Then
                        TypeFinancement = FindTypeFinancementIndex(Me.TypeOrgaSociaux.Caption)
                    Else
                        If Me.TypeFondEuro Then
                            TypeFinancement = FindTypeFinancementIndex(Me.TypeFondEuro.Caption)
                        Else
                            If Me.TypeASP Then
                                TypeFinancement = FindTypeFinancementIndex(Me.TypeASP.Caption)
                            Else
                                If Me.TypeFondation Then
                                    TypeFinancement = FindTypeFinancementIndex(Me.TypeFondation.Caption)
                                Else
                                    If Me.TypeAutre Then
                                        TypeFinancement = FindTypeFinancementIndex(Me.TypeAutre.Caption)
                                    Else
                                        TypeFinancement = 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Unload Me
    If Nom = "" Or Nom = Empty Then
        MsgBox "Le nom ne peut être vide !"
        Exit Sub
    End If
    
    Set wb = ThisWorkbook
    
    ' Current NB
    CurrentNBChantier = GetNbChantiers(wb)
    
    If CurrentNBChantier < 1 Then
        Exit Sub
    End If
    
    FinancementFantome.Status = False
    AjoutFinancement wb, CurrentNBChantier, FinancementFantome, Nom, TypeFinancement
    
End Sub

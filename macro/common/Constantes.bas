Attribute VB_Name = "Constantes"
' SPDX-License-Identifier: EUPL-1.2

' Pour forcer la declaration de toutes les variables
Option Explicit

' constantes
Public Const ColumnOfSecondPartInCharge As Integer = 8
Public Const Label_Annees = "Ann�e courante:"
Public Const Label_Autofinancement_Structure = "Autofinancmt structure"
Public Const Label_Autofinancement_Structure_Previous = "Autofinancmt structure ann�es pr�c�dentes"
Public Const Label_Convention_Collective = "Convention collective: "
Public Const Label_Cout_J_Journalier = "Co�t journalier"
Public Const Label_Cout_J_Salaire_Part_A = "A - Calcul des jours travaillables / an"
Public Const Label_Cout_J_Salaire_Part_B = "B/ Calcul Unit� jours salaires "
Public Const Label_Cout_J_Salaire_Part_D = "Co�t journ�e "
Public Const Label_NBConges = "Nb. Cong�s hors samedi:"
Public Const Label_NBRTT = "Nb. RTT:"
Public Const Label_NB_Jours_speciaux = "Nb. Jours non travaill�s suppl�mentaires:"
Public Const Label_Pentecote = "Pentec�te = f�ri�"
Public Const Label_Total_Financements = "Total des financements"
Public Const Label_Type_Financeur = "Type de financeur"
Public Const Label_Version As String = "Version du tableur:"
Public Const Label_Waited_Payments As String = "En attente de paiement"
Public Const NBExtraCols As Integer = 6
Public Const NBCatOfCharges As Integer = 3
Public Const Nom_Feuille_Budget_chantiers As String = "Budget_chantiers"
Public Const Nom_Feuille_Budget_chantiers_realise As String = "Budget_chantiers_r�alis�"
Public Const Nom_Feuille_Budget_global As String = "Budget_global"
Public Const Nom_Feuille_CC_by_SA As String = "CC BY SA V3.0"
Public Const Nom_Feuille_Charges = "Co�t_j_chgs_indrctes"
Public Const Nom_Feuille_Cout_J_Salaire As String = "Co�t_j_salaires"
Public Const Nom_Feuille_CptResult_prefix As String = "CptResult_"
Public Const Nom_Feuille_CptResult_Real_prefix As String = "CptResultReal_"
Public Const Nom_Feuille_CptResult_suffix As String = "Global"
Public Const Nom_Feuille_Eupl As String = "EUPL v1.2"
Public Const Nom_Feuille_FAQ As String = "FAQ"
Public Const Nom_Feuille_Informations As String = "Informations"
Public Const Nom_Feuille_Personnel As String = "Personnel"
Public Const Nom_Feuille_Provisions As String = "Provisions"
Public Const Nom_Feuille_Versions As String = "Suivi des versions"
Public Const Offset_NB_Cols_For_Percent_In_CptResultReal As Integer = 9
Public Const T_Add_New_Financier_Default As String = "Financement X"
Public Const T_Add_New_Financier_Name As String = "Nom du nouveau financeur ?"
Public Const T_Add_New_Financier_Title As String = "Ajout financement europ�en"
Public Const T_Amout_Salary_of_WorkingPeople As String = "Masse salariale des %n%op�rateurs : "
Public Const T_Charge As String = "la charge"
Public Const T_Choose_File_To_Export As String = "Choisir le fichier � exporter"
Public Const T_Choose_File_To_Import As String = "Choisir le fichier � importer"
Public Const T_Create_CptResult_Formula As String = "Lister les chantiers � prendre en compte ;%n%s�par�s par des virgules ;%n%une s�rie de chantier est indiqu�e par un tiret ;%n%Exemple ""1,4,6-8"" correspond aux chantiers 1, 4, 6, 7 et 8"
Public Const T_Create_CptResult_Formula_Title As String = "Choix des chantiers"
Public Const T_Create_Real_CptResult As String = "Faut-il cr�er un onglet ""Compte de r�sultat r�alis�"" en plus du pr�visionnel ?"
Public Const T_Create_Real_CptResult_Title As String = "Budget r�el ?"
Public Const T_Currency_Format As String = "#,##0.00"" �"""
Public Const T_Development_In_Course As String = "Patience, cette fonction est encore en cours de d�veloppement"
Public Const T_Get_CptResult_Suffix As String = "Quel suffixe pour le nom de l'onglet ?%n%Exemple : le suffixe %suffix% donne l'onglet %onglet%."
Public Const T_Get_CptResult_Suffix_Title As String = "Choix du suffixe"
Public Const T_Data_Were_Replaced As String = "Les donn�es import�es remplaceront toutes les donn�es contenues dans le pr�sent fichier."
Public Const T_Delete_Object_Of_Line As String = "Supprimer de %objectName% de la ligne ?"
Public Const T_Depense As String = "la d�pense"
Public Const T_Error_Bad_Format_For_Charge_Name As String = "Erreur : les deux premiers caract�res du nom doivent �tre compris entre 60 et 68 inclus."
Public Const T_Error_Existing_Tab_for_Suffix As String = "Erreur : le suffixe fourni (%suffix%) correspond d�j� � un onglet existant."
Public Const T_Error_Formula_In_CptResult As String = "Erreur : la formule de la cellule %adr% n'est pas valide. Le calcul ne sera fait que pour le chantier 1."
Public Const T_Error_Incorrect_Formula As String = "Erreur : la formule fournie n'est pas correcte."
Public Const T_Error_Not_Empty_Charge_Name As String = "Erreur : Le nom fourni pour la charge ne peut pas �tre vide"
Public Const T_Error_Not_Possible_To_Associate_Line_To_Type As String = "Erreur : impossible d'associer cette ligne � un type de paiement (entre 60 et 68)"
Public Const T_Error_Not_Possible_To_Found_Type As String = "Erreur : impossible de retrouver les diff�rents types de paiement (60 � 68)"
Public Const T_Existing_File_What_To_Do As String = "Le fichie cible existe d�j� !%n%Faut-il l'�craser avec le nouveau ?"
Public Const T_File_Imported As String = "Fichier import�"
Public Const T_File_Not_Exported As String = "Fichier non export�"
Public Const T_File_Saved As String = "Fichier sauvegard�"
Public Const T_Financement As String = "le financement"
Public Const T_FirstName As String = "Pr�nom"
Public Const T_Formula As String = "Formule de choix des chantiers :"
Public Const T_Given_Line_Is_Not_Line_Of_Object As String = "La ligne entr�e n'est pas la ligne d'%objectName%"
Public Const T_Line_To_Delete_For_Object As String = "Ligne %objectName% � supprimer"
Public Const T_NotFoundPage As String = "'%PageName%' n'a pas �t� trouv�e"
Public Const T_NotFoundFirstName As String = "'Pr�nom' non trouv� dans '%PageName%' !"
Public Const T_NotPossibleToForceSave As String = "Il n'est pas possible d'�craser le fichier courant%n%Veuillez r�essayer avec un autre emplacement ou nom de fichier"
Public Const T_NotPossibleToSaveFileBecauseExisting As String = "Impossible de sauvegarder le fichier de sauvegarde car il existe d�j�"
Public Const T_Provisions_In_CptResult As String = "687 - Provisions pour risques"
Public Const T_Rate_For_Charges As String = "Taux de charges indirectes :"
Public Const T_Retrieval_In_CptResult As String = "787 - Reprises de provisions pour risques"
Public Const T_Salary As String = "R�mun�ration des personnels"
Public Const T_Social_Charges As String = "Charges sociales"
Public Const T_Total_Charges As String = "Total D�penses"
Public Const T_WorkingPeople As String = "Salari�"

Imports System.Threading
Imports System.Globalization
Imports System.Data.OleDb

Public Class VerifMensuelle
    Inherits System.Windows.Forms.Form

    ' Auteur : Thomas FALCHER
    ' 2012

    ' Prérequis aux traitements : les mois de l'année doivent être rangés dans l'ordre. 
    ' Il est conseillé de les nommer avec le formalisme suivant : 01-Janvier.xlsx, 02-Fevrier.xlsx, etc.
    ' Le dossier contenant les mois doit contenir un dossier "result" (en minuscules) ou le fichier excel 
    ' de résultat sera stocké.

    ' Déclaration des variables globales de traitement
    Public RepertoireTravail As String

#Region "Déclaration des variables"
    ' Nombre de ligne du fichier excel courant
    Public nbLignesSource As Integer

    ' Variable de colonne de début d'écriture pour le mois courant
    Public indiceMois As Integer = 4

    ' Variable pour le prix laboratoire
    Public prix As Decimal

    ' Variables stockant la valeur des indicateurs de progression pour les opérations
    Dim progres, progres2 As Integer

    ' Compteur de fichier dans un dossier et indice du fichier actuellement traité (pour affichage progression)
    Dim nbFichiers As Integer
    Dim fichierActuel As String

    ' Label affichage d'écriture
    Dim ecritureEnCours As String

    ' Thread de calcul des opérations
    Public threadCalcul As Thread

    ' Liste de Produits
    Private listeProduits As Hashtable

#End Region
    ' Fonction Permettant l'affichage d'une fenêtre de sélection du répertoire
    Function ChoixDossierFichier(ByVal SelType As Byte) As String

        'Source : http://www.cathyastuce.com/index.php?tg=articles&topics=87
        Dim objShell As Object, objFolder As Object
        Dim Msg As String
        Dim FlagChoix As Long, NbPoint As Integer

        If SelType = 0 Then
            FlagChoix = &H1
            Msg = "Sélectionner un dossier :"
        Else
            FlagChoix = &H4000
            Msg = "Sélectionner un fichier :"
        End If

        objShell = CreateObject("Shell.Application")
        On Error Resume Next
        objFolder = objShell.BrowseForFolder(&H0&, Msg, FlagChoix)
        NbPoint = InStr(objFolder.Title, ":")
        If NbPoint = 0 Then
            RepertoireTravail = objFolder.ParentFolder.ParseName(objFolder.Title).path & ""
        Else
            RepertoireTravail = Mid(objFolder.Title, NbPoint - 1, 2)
        End If
        ChoixDossierFichier = RepertoireTravail
        lblRepertoireCharge.Text = RepertoireTravail
    End Function

    ' Action du clic sur le bouton Parcourir
    Private Sub btnParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParcourir.Click
        ' Traitement des exceptions
        On Error GoTo catchError

        ChoixDossierFichier(0) '0 pour recherche dossier uniquement, 1-255 pour dossier et fichiers
        ' Affichage des mois dans la liste des cases à cocher
        ' DEBUG
        ' RepertoireTravail = "C:\M. CAMPILLO"
        ' Création d'un objet de lecture de fichiers
        Dim fso As Object = CreateObject("Scripting.FileSystemObject")
        nbFichiers = 0
        ' On vide la liste au cas où elle contienne déjà des éléments
        checkedListBoxMois.Items.Clear()
        ' Parcours du dossier RepertoireTravail
        For Each fichier In fso.GetFolder(RepertoireTravail).Files
            ' Controle si le fichier est un fichier excel
            If (Strings.Right(fichier.Name, 4) = ".xls" Or Strings.Right(fichier.Name, 5) = ".xlsx") Then
                checkedListBoxMois.Items.Add(fichier.Name)
            End If
        Next
        checkedListBoxMois.Enabled = True

        Exit Sub

catchError:
        MsgBox("Erreur : " & vbCrLf & Err.Description)
        Err.Clear()
    End Sub

    ' Action du clic sur le bouton Quitter
    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuit.Click
        Me.Dispose()
    End Sub

    ' Action du clic sur le bouton Exécuter : Traitement principal
    Private Sub btnExec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExec.Click

        ' DEBUG
        ' RepertoireTravail = "C:\M. CAMPILLO"
        ' RepertoireTravail = "F:\PERSO\M. CAMPILLO\"

        ' ETAPE : DEFINIR LE NOMBRE DE FICHIERS A TRAITER DANS LE REPERTOIRE
        ' Création d'un objet de lecture de fichiers
        Dim fso As Object = CreateObject("Scripting.FileSystemObject")
        nbFichiers = 0
        ' Parcours du dossier RepertoireTravail
        For Each fichier In fso.GetFolder(RepertoireTravail).Files
            ' Controle si le fichier est un fichier excel
            If (Strings.Right(fichier.Name, 4) = ".xls" Or Strings.Right(fichier.Name, 5) = ".xlsx") Then
                nbFichiers = nbFichiers + 1
            End If
        Next

        ' ETAPE : INSTANCIATION ET LANCEMENT DU THREAD traitement1
        threadCalcul = New Thread(AddressOf Operation) 'Operation est la fonction exécutée par le thread.
        threadCalcul.CurrentCulture = CultureInfo.CreateSpecificCulture("fr-FR")
        threadCalcul.Name = "traitement1"
        threadCalcul.Start() ' Démarrage du thread.

        ' Mise à jour du statut du traitement
        progressBar.Value = progres

    End Sub

    ' Fonction délégué permettant d'accéder au thread d'accéder aux variables
    Public Delegate Sub refreshUI()

    ' Fonction de rafraichissement de l'interface pour la progression des traitements
    Public Sub RefreshProgress()
        'On met la valeur dans le contrôle Windows Forms.
        progressBar.Value = progres
        lblProgress.Text = progres & "%"

        progressBar2.Value = progres2
        lblProgress2.Text = progres2 & "%"

        lblNbFichiers.Text = fichierActuel

        lblEcritureEnCours.Text = ecritureEnCours
    End Sub

    ' Operation est le traitement principal
    Sub Operation()
        Try
            Console.Write("LANCEMENT DU THREAD !!!!" & Environment.NewLine)

            ' Réinitialisation de la ProgressBar
            progressBar.Value = 0

            Dim fso As Object = CreateObject("Scripting.FileSystemObject")
            Dim cptMois As Integer = 0

            Dim feuille As Object(,)

            ' Instanciation Hashtable contenant tous les produits
            listeProduits = New Hashtable

            ' Création d'une Hashtable contenant les fichiers à traiter
            Dim listeFichiers As New Hashtable
            Dim indexFichier As Integer = checkedListBoxMois.CheckedItems.Count

            '-------------------------------------------------------------------------------------------------------------------
            ' NB :
            ' On utilise une décrémentation sur le nombre d'élément coché pour stocker dans listeFichiers les mois dans l'ordre 
            ' inverse de leur lecture dans la checkedListeBox. On stocke ainsi le 1er mois au début de la liste pour qu'il soit
            ' lu en premier dans le parcours de listeFichiers.
            '-------------------------------------------------------------------------------------------------------------------

            ' Parcours de la checkListBox et ajout des fichier à comparer à la liste
            For Each fichier In checkedListBoxMois.CheckedItems
                listeFichiers.Add(indexFichier, fichier)
                indexFichier = indexFichier - 1
            Next

            ' Parcours de la liste des fichiers
            For Each numFichier In listeFichiers.Keys
                ' Attribution des feuilles Excel dans variable
                feuille = Excel_LectureFichier(RepertoireTravail & "\" & listeFichiers.Item(numFichier), "A")

                cptMois = cptMois + 1

                nbLignesSource = feuille.GetLength(0)

                fichierActuel = listeFichiers.Item(numFichier)

                Try
                    Invoke(New refreshUI(AddressOf RefreshProgress))
                Catch ex As Exception
                    Debug.WriteLine(ex.ToString())
                End Try

                For ligne = 0 To nbLignesSource - 1
                    ' Création du produit
                    If (Not listeProduits.ContainsKey(feuille(ligne, 0))) Then
                        Dim unNouveauProduit As New Produit(feuille(ligne, 0))
                        ' Instanciation de l'objet Produit
                        unNouveauProduit.setDesignation(feuille(ligne, 1))
                        unNouveauProduit.setLaboratoire(feuille(ligne, 13))
                        unNouveauProduit.setCodeTVA(Val(feuille(ligne, 2)))

                        listeProduits.Add(feuille(ligne, 0), unNouveauProduit)
                    End If
                    ' Instanciation de l'objet Vente
                    Dim uneNouvelleVente As New VenteMensuelle
                    uneNouvelleVente.setNbBoitesVendues(feuille(ligne, 4))
                    uneNouvelleVente.setPrixHT(feuille(ligne, 7))

                    ' Algorithme de traitement de sélection de valeurs Prix Laboratoire
                    '--------------------------------------------------------------------
                    If Val(feuille(ligne, 11)) = 0 Then
                        If Val(feuille(ligne, 10)) > Val(feuille(ligne, 9)) Then
                            If feuille(ligne, 9) > 0 Then
                                prix = feuille(ligne, 9)
                            End If
                        Else
                            prix = feuille(ligne, 10)
                        End If
                        If Val(feuille(ligne, 10)) = 0 Then
                            prix = feuille(ligne, 9)
                        End If
                        If Val(feuille(ligne, 9)) = 0 And Val(feuille(ligne, 10)) = 0 Then
                            prix = 0
                        End If
                    Else
                        prix = feuille(ligne, 11)
                    End If

                    uneNouvelleVente.setPrixLabo(prix)

                    listeProduits.Item(feuille(ligne, 0)).ajouterVenteMensuelle(cptMois, uneNouvelleVente)

                    progres = ligne * 100 / nbLignesSource

                    Try
                        Invoke(New refreshUI(AddressOf RefreshProgress))
                    Catch ex As Exception
                        Debug.WriteLine(ex.ToString())
                    End Try

                Next ligne
            Next numFichier

            If Dir(RepertoireTravail & "\traitement", vbDirectory) = "" Then
                MkDir(RepertoireTravail & "\traitement")
            End If

            ' Lancement de la procédure d'écriture dans le fichier excel
            EcrireFichierOLEDB(RepertoireTravail & "\traitement\" & txtNomFichierRecap.Text & ".xls", "Traitement_Annuel")
            'EcrireFichierOLEDB("C:\M. CAMPILLO\result\test.xls", "Traitement_Annuel")
            'EcrireFichierOLEDB("F:\PERSO\M. CAMPILLO\result\test.xls", "Traitement_Annuel")

        Catch ex As Exception
            Debug.WriteLine(ex.ToString())
        End Try
    End Sub

#Region "EcritureFichierOLEDB"

    Public Sub EcrireFichierOLEDB(ByVal unFichier As String, ByVal uneFeuille As String)
        ' --------------------------------------------------------------------------------
        ' Création du fichier Excel - source : http://support.microsoft.com/kb/316934/fr
        ' --------------------------------------------------------------------------------
        Debug.Print("Procédure d'écriture dans le fichier excel")
        ecritureEnCours = "écriture dans le fichier en cours..."
        Try
            Invoke(New refreshUI(AddressOf RefreshProgress))
        Catch ex As Exception
            Debug.WriteLine(ex.ToString())
        End Try

        ' Si le fichier (classeur) existe déjà = Proposition de suppression
        Dim answer As MsgBoxResult
        If Dir(unFichier) <> "" Then
            answer = MsgBox("Delete existing workbooks (" & unFichier & ") ?", MsgBoxStyle.YesNo)
            If answer = MsgBoxResult.Yes Then
                If Dir(unFichier) <> "" Then Kill(unFichier)
            Else
                Exit Sub
            End If
        End If

        ' Création du fichier (classeur)
        Dim m_sConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                    "Data Source=" & unFichier & ";" & _
                                    "Extended Properties=""Excel 8.0;HDR=YES"""
        Dim conn As New OleDbConnection()
        conn.ConnectionString = m_sConn
        conn.Open()

        Dim cmdNewTable As New OleDbCommand()
        cmdNewTable.Connection = conn
        cmdNewTable.CommandText = "CREATE TABLE " & uneFeuille & " (CIP char(255), DESIGNATION char(255), LABORATOIRE char(255), TVA int, " & _
                                                                    "NB_BOITES1 int, PRIX_VENTE_UNITAIRE_HT1 double, PRIX_ACHAT_UNITAIRE_HT1 double, " & _
                                                                    "NB_BOITES2 int, PRIX_VENTE_UNITAIRE_HT2 double, PRIX_ACHAT_UNITAIRE_HT2 double, " & _
                                                                    "DELTA_nb_BOITES int, DELTA_PRIX_VENTE double, DELTA_PRIX_ACHAT double, " & _
                                                                    "DELTA_ACHATS double, DELTA_VENTES double" & _
                                                                    ")"
        cmdNewTable.ExecuteNonQuery()

        'version > EXCEL 2007 :
        'MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; data source=" & unFichier & "; Extended Properties='Excel 12.0;HDR=NO;IMEX=1'")

        ' Compteur de progression
        Dim cpt As Integer = 1
        Dim nbProduitsListe As Integer = listeProduits.Count

        ' Parcours de l'objet listeProduits
        For Each produit In listeProduits

            Dim commandData As New OleDbCommand()
            Dim requete As String
            'requete = "INSERT INTO [" & uneFeuille & "$] ([COL 1], [COL 2]) VALUES ('A1', 'A2')"

            ' Remplacement des guillemets pour éviter de casser la requete
            Dim designation As String = Nothing
            designation = Replace(produit.Value.getDesignation(), """", "'")
            Dim laboratoire As String = Nothing
            laboratoire = Replace(produit.Value.getLaboratoire(), """", "'")

            ' Construction de la requete
            requete = "INSERT INTO [" & uneFeuille & "$] VALUES ('" & produit.Value.getCodeCIP() & "', " & _
                                                                    Chr(34) & designation & Chr(34) & ", " & _
                                                                    Chr(34) & laboratoire & Chr(34) & ", " & _
                                                                    produit.Value.getCodeTVA() & ", "
            ' Ajout des ventes
            Dim nombreMois As Integer = produit.Value.getVentes().count()
            Dim comparePrixAch, comparePrixVen As Double

            For Each mois In produit.Value.getVentes.Keys
                ' Si c'est le 1er mois à comparer et qu'il n'y a qu'un mois au total
                'If mois = Val(checkedListBoxMois.CheckedIndices(0)) + 1 And nombreMois = 1 Then
                If mois = 1 And nombreMois = 1 Then
                    requete = requete & produit.Value.getVentes().Item(mois).getNbBoitesVendues() & ", "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixHT() & "', "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixLabo() & "', "
                    requete = requete & "'0', '0', '0', "
                End If
                ' Si c'est le 2ème mois à comparer et qu'il n'y a qu'un mois au total
                If mois = 2 And nombreMois = 1 Then
                    requete = requete & "'0', '0', '0', "
                    requete = requete & produit.Value.getVentes().Item(mois).getNbBoitesVendues() & ", "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixHT() & "', "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixLabo() & "', "
                End If
                ' Si il y a deux mois à comparer
                If nombreMois = 2 Then
                    requete = requete & produit.Value.getVentes().Item(mois).getNbBoitesVendues() & ", "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixHT() & "', "
                    requete = requete & "'" & produit.Value.getVentes().Item(mois).getPrixLabo() & "', "
                End If

            Next mois

            ' Calcul du delta
            Dim deltaBoites As Integer
            Dim deltaPriVen As Decimal
            Dim deltaPriAch As Decimal
            If nombreMois >= 2 Then
                deltaBoites = produit.Value.getVentes().Item(1).getNbBoitesVendues() - produit.Value.getVentes().Item(2).getNbBoitesVendues()
                deltaPriVen = produit.Value.getVentes().Item(1).getPrixHT() - produit.Value.getVentes().Item(2).getPrixHT()
                deltaPriAch = produit.Value.getVentes().Item(1).getPrixLabo() - produit.Value.getVentes().Item(2).getPrixLabo()
                If produit.Value.getVentes().Item(1).getPrixHT() <> 0 Then
                    comparePrixVen = deltaPriVen / (0.01 * produit.Value.getVentes().Item(1).getPrixHT())
                Else
                    comparePrixVen = 0
                End If
                If produit.Value.getVentes().Item(1).getPrixLabo() <> 0 Then
                    comparePrixAch = deltaPriAch / (0.01 * produit.Value.getVentes().Item(1).getPrixLabo())
                Else
                    comparePrixAch = 0
                End If

            Else
                deltaBoites = 0
                deltaPriVen = 0
                deltaPriAch = 0
                comparePrixVen = 0
                comparePrixAch = 0
            End If

            Dim requeteDelta = "'" & deltaBoites & "', '" & deltaPriVen & "', '" & deltaPriAch & "', "
            Dim requeteComparePrix = "'" & comparePrixVen & "', '" & comparePrixAch & "'"

            requete = requete & requeteDelta & requeteComparePrix

            ' Suppression de la dernière virgule et de l'espace
            'requete = Mid(requete, 1, Len(requete) - 2)

            requete = requete & ");"

            'TODO
            '----------------------------------------------------------------------------------------------------------
            'Automatiser le traitement de la requete avec d'autres mois que Janvier (2 dans le code) et Février (1)
            'Attention, dans les objets VenteMensuelle, les mois sont identifiés : janvier = 1, février = 2, etc.
            'Automatiser les traitements avec plus que 2 mois...

            commandData = New OleDbCommand(requete, conn)
            With commandData
                .ExecuteNonQuery()
            End With

            progres2 = cpt * 100 / nbProduitsListe
            cpt = cpt + 1
            Try
                Invoke(New refreshUI(AddressOf RefreshProgress))
            Catch ex As Exception
                Debug.WriteLine(ex.ToString())
            End Try
        Next produit

        conn.Close()

        conn = Nothing

        MsgBox("Traitement Terminé!")
    End Sub

#End Region

#Region "LectureFichierOLEDB"
    Public Function Excel_LectureFichier(ByVal pChemin As String, ByVal pNomFeuille As String) As Object(,)

        Dim DS As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim lRetour(0, 0) As Object

        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & "data source=" & pChemin & "; " & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'")

        ' Select the data from Sheet1 of the workbook.
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & pNomFeuille & "$]", MyConnection)
        DS = New System.Data.DataSet
        MyCommand.Fill(DS)
        Dim lFeuilleExcel As DataTable
        lFeuilleExcel = DS.Tables(0)
        Dim lLigneExcel As DataRow
        Dim lStr As String
        Dim lCompteur As Long = 0
        Dim lNB_Ligne As Long
        Dim lNB_Colonne As Long
        'Console.WriteLine(lFeuilleExcel.Rows.Count)
        ' Détermination du nombre de ligne excel
        If lFeuilleExcel.Rows(lFeuilleExcel.Rows.Count - 2).Item(0).ToString = "Total produit" Then
            lNB_Ligne = lFeuilleExcel.Rows.Count - 2
        Else
            lNB_Ligne = lFeuilleExcel.Rows.Count
        End If
        'lNB_Ligne = lFeuilleExcel.Rows.Count
        lNB_Colonne = lFeuilleExcel.Rows(0).ItemArray.GetUpperBound(0)
        ReDim lRetour(lNB_Ligne - 1, lNB_Colonne)
        For Each lLigneExcel In lFeuilleExcel.Rows
            ' Si on est arrivé à la fin du fichier excel, on sort de la fonction
            If lLigneExcel.Item(0).ToString = "Total produit" Then
                GoTo NextElement
            End If
            Dim lArray() As Object
            lArray = lLigneExcel.ItemArray
            For i As Long = 0 To lArray.GetUpperBound(0)
                lStr = ""
                If IsDBNull(lArray(i)) = True Then
                    lStr = "vide"
                Else
                    lStr = lArray(i)
                End If
                lRetour(lCompteur, i) = lStr
            Next
            lCompteur += 1
        Next
NextElement:
        Console.ReadLine()
        MyConnection.Close()
        Return lRetour
    End Function

#End Region

    Private Sub checkedListBoxMois_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles checkedListBoxMois.SelectedIndexChanged
        ' Règles de Gestion :
        '------------------------------------------------------------------------------------------------------------------------
        ' Cette version de l'application permet de comparer 2 mois. Lorsque 2 mois sont sélectionnés, les autres sont grisés.
        '------------------------------------------------------------------------------------------------------------------------

        If checkedListBoxMois.GetItemCheckState(checkedListBoxMois.Items.IndexOf(checkedListBoxMois.SelectedItem)).ToString() = "Checked" Then
            checkedListBoxMois.SetItemChecked(checkedListBoxMois.SelectedIndex, False)
        Else
            If checkedListBoxMois.CheckedItems.Count = 2 Then
                MsgBox("Cette version de l'application ne permet pas de comparer plus de 2 mois")
                checkedListBoxMois.SetItemChecked(checkedListBoxMois.SelectedIndex, False)
            Else
                checkedListBoxMois.SetItemChecked(checkedListBoxMois.SelectedIndex, True)
            End If
        End If

        deverouilleBtnComparer()

        Debug.Print("click effectué. Nb éléments cochés : " & checkedListBoxMois.CheckedItems.Count)
    End Sub

    Public Sub deverouilleBtnComparer()
        ' Déverouillage du bouton Execution si 2 cases sont cochées
        If checkedListBoxMois.CheckedItems.Count = 2 And txtNomFichierRecap.Text <> "" Then
            btnExec.Enabled = True
        Else
            btnExec.Enabled = False
        End If
    End Sub

    Private Sub txtNomFichierRecap_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNomFichierRecap.TextChanged
        deverouilleBtnComparer()
    End Sub
End Class

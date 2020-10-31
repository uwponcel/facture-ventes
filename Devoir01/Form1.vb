'********************************************************************
'* Devoir 01 - William Poncelet
'********************************************************************

Public Class Form1

    '********************************************************************
    '* Variables globale
    '********************************************************************
    Private Structure Vente
        Public noProduit As String
        Public marque As String
        Public modele As String
        Public annee As String
        Public couleur As String
        Public categorie As String
        Public noSerie As String
        Public prix As String
        Public dateVente As String

        Public nomClient As String
        Public prenomClient As String
        Public noClient As String
        Public adresse As String
        Public ville As String
        Public province As String
        Public codePostal As String
        Public courriel As String
        Public telephone As String

    End Structure

    Private VenteClient As New SortedList(Of String, Vente)
    Private enregistrement As String

    '********************************************************************
    '* Chargement du formulaire, fermeture et initialization.
    '********************************************************************
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initialize()
    End Sub
    Private Sub Initialize()
        Me.mtxtNoProduit.Text = Nothing ' ""
        Me.txtMarque.Text = Nothing
        Me.txtModele.Text = Nothing
        Me.mtxtAnnee.Text = "2020"
        Me.txtCouleur.Text = Nothing
        Me.txtCategorie.Text = Nothing
        Me.txtNoSerie.Text = Nothing
        Me.mtxtPrix.Text = Nothing
        Me.dtpickerDateVente.Value = Date.Now
        Me.mtxtNoClient.Text = Nothing
        Me.txtNom.Text = Nothing
        Me.txtPrenom.Text = Nothing
        Me.txtAdresse.Text = Nothing
        Me.txtVille.Text = Nothing
        Me.CboProvince.SelectedIndex = CboProvince.FindStringExact("ON")
        Me.CboProvince.BackColor = Color.Empty
        Me.mtxtCodePostal.Text = Nothing
        Me.txtCourriel.Text = Nothing
        Me.mtxtNoTel.Text = Nothing
        Me.btnAfficherVente.Enabled = False
        Me.btnAjouterCollection.Enabled = False
        Me.btnExtraireCollection.Enabled = False
        Me.btnSupprimerCollection.Enabled = False
        Me.btnSauvegarderCollection.Enabled = False
        Me.btnFermer.Enabled = False



    End Sub
    Private Sub btnInitialiser_Click(sender As Object, e As EventArgs) Handles btnInitialiser.Click
        Initialize()
    End Sub
    Private Sub btnFermer_Click(sender As Object, e As EventArgs) Handles btnFermer.Click
        Me.Close()
    End Sub

    '********************************************************************
    '* Validation du formulaire
    '********************************************************************
    Private Sub btnValider_Click(sender As Object, e As EventArgs) Handles btnValider.Click
        Dim marques() As String = {"Dell", "Asus", "Acer", "HP", "MSI"}
        Dim modeles() As String = {"x1", "x2", "x3", "x4", "x5", "x6"}
        Dim couleurs() As String = {"rouge", "jaune", "vert", "bleu", "mauve"}
        Dim categories() As String = {"Cuisine", "Salon", "Bureau", "Loisir", "Professionel", "Autres"}
        Dim telephone() As String = {"514", "450", "819", "418", "613"}

        Dim FicheValide As Boolean = True
        Dim vente As Vente


        'No de produit check
        If mtxtNoProduit.Text.Length < 6 Then
            mtxtNoProduit.Text = "N00000"
            mtxtNoProduit.ForeColor = Color.Red
            FicheValide = False
        End If

        'Marque check
        If LenghtCheck(txtMarque, "Inscrire marque") Or Not ContainsWord(marques, txtMarque.Text) Then
            txtMarque.ForeColor = Color.Red
            FicheValide = False
        End If

        'Modèle check
        If LenghtCheck(txtModele, "Inscrire modele") Or Not ContainsWord(modeles, txtModele.Text) Then
            txtModele.ForeColor = Color.Red
            FicheValide = False
        End If

        'Année check
        Dim annee As String
        If mtxtAnnee.Text.Length <> 0 Then
            annee = Convert.ToInt32(mtxtAnnee.Text)
            If (annee <= 2015 Or annee >= 2019) Then
                mtxtAnnee.ForeColor = Color.Red
                FicheValide = False
            End If
        Else
            mtxtAnnee.Text = "2017"
            mtxtAnnee.ForeColor = Color.Red
            FicheValide = False
        End If


        'Couleur check
        If LenghtCheck(txtCouleur, "Inscrire couleur") Or Not ContainsWord(couleurs, txtCouleur.Text) Then
            txtCouleur.ForeColor = Color.Red
            FicheValide = False
        End If

        'Catégorie check
        If LenghtCheck(txtCategorie, "Inscrire catégorie") Or Not ContainsWord(categories, txtCategorie.Text) Then
            txtCategorie.ForeColor = Color.Red
            FicheValide = False
        End If

        'No de série check
        If LenghtCheck(txtNoSerie, "Inscrire no de série") Or txtNoSerie.Text.Length < 8 Then
            txtNoSerie.ForeColor = Color.Red
            FicheValide = False
        End If


        'Prix check
        If mtxtPrix.Text = "$   ,   ." Then
            mtxtPrix.Text = "000 000.000"
            mtxtPrix.ForeColor = Color.Red
            FicheValide = False
        End If

        '--------------- Info sur client ---------------

        'No de client check
        If mtxtNoClient.Text.Length < 5 Then
            mtxtNoClient.Text = "A0000"
            mtxtNoClient.ForeColor = Color.Red
            FicheValide = False
        End If

        'Nom de client check
        If LenghtCheck(txtNom, "Inscrire nom") Then
            FicheValide = False
        End If

        'Prénom de client check
        If LenghtCheck(txtPrenom, "Inscrire prénom") Then
            FicheValide = False
        End If

        'Adresse check
        If LenghtCheck(txtAdresse, "Inscrire adresse") Then
            FicheValide = False
        End If


        'Ville check
        If LenghtCheck(txtVille, "Inscrire ville") Then
            FicheValide = False
        End If

        'Province
        If String.IsNullOrEmpty(CboProvince.Text) Then
            CboProvince.BackColor = Color.Red
            FicheValide = False
        End If

        'Code postal check
        If mtxtCodePostal.Text.Length < 6 Then
            mtxtCodePostal.Text = "A0A0A0"
            mtxtCodePostal.ForeColor = Color.Red
            FicheValide = False
        End If

        'Courriel check
        If txtCourriel.Text.Length > 0 Then
            If Not (txtCourriel.Text.Contains("@")) Then
                txtCourriel.ForeColor = Color.Red
                FicheValide = False
            End If
        End If

        'Téléphone check
        Dim indicatif As String = mtxtNoTel.Text.Substring(1, 3)

        If mtxtNoTel.Text.Length < 13 Or Not ContainsWord(telephone, indicatif) Then
            mtxtNoTel.Text = "(000)-000-0000"
            mtxtNoTel.ForeColor = Color.Red
            FicheValide = False
        End If

        If FicheValide = True Then
            btnAfficherVente.Enabled = True
            btnAjouterCollection.Enabled = True
        Else
            btnAfficherVente.Enabled = False
            btnAjouterCollection.Enabled = False
        End If

        Console.WriteLine(FicheValide.ToString)
    End Sub
    Private Function LenghtCheck(p_text As TextBox, p_errorValue As String) As Boolean
        If p_text.Text.Length = 0 Then
            p_text.Text = p_errorValue
            p_text.ForeColor = Color.Red
            Return True
        Else
            Return False
        End If
    End Function

    '* Vérifie si un string est contenu dans un array (case insensitive)
    Private Function ContainsWord(p_wordsArray As String(), p_word As String) As Boolean
        Return Array.Exists(p_wordsArray, Function(s As String) s.Equals(p_word.Trim, StringComparison.CurrentCultureIgnoreCase))
    End Function

    '* Change la couleur du texte à noir si le texte est overwritten
    Private Sub TextAsChanged(sender As Object, e As EventArgs) _
        Handles mtxtNoProduit.TextChanged, txtMarque.TextChanged,
        txtModele.TextChanged, mtxtAnnee.TextChanged, txtCouleur.TextChanged, txtCategorie.TextChanged,
        txtNoSerie.TextChanged, mtxtPrix.TextChanged, mtxtNoClient.TextChanged, txtNom.TextChanged,
        txtPrenom.TextChanged, txtAdresse.TextChanged, txtVille.TextChanged, mtxtCodePostal.TextChanged,
        txtCourriel.TextChanged, mtxtNoTel.TextChanged
        sender.ForeColor = Color.Black
    End Sub

    '* Vérifie si la selection du comboBox à changé
    Private Sub SelectionAsChanged(sender As Object, e As EventArgs) _
        Handles CboProvince.DropDown
        sender.Backcolor = Color.Empty
    End Sub


    '********************************************************************
    '* Fonctions pour les bouttons post validation
    '********************************************************************
    Private Sub btnAfficherVente_Click(sender As Object, e As EventArgs) Handles btnAfficherVente.Click
        Dim Facture As String
        Facture =
            "Client:  " & vbTab & vbTab & txtNom.Text & ", " & txtPrenom.Text & vbCrLf &
            "No de client: " & vbTab & mtxtNoClient.Text & vbCrLf &
            "Adresse: " & vbTab & vbTab & txtAdresse.Text & vbCrLf &
            "Ville: " & vbTab & vbTab & txtVille.Text & vbCrLf &
            "Province: " & vbTab & CboProvince.SelectedItem.ToString & vbCrLf &
            "Code postal: " & vbTab & mtxtCodePostal.Text & vbCrLf &
            "Courriel: " & vbTab & vbTab & txtCourriel.Text & vbCrLf &
            "Téléphone: " & vbTab & mtxtNoTel.Text & vbCrLf & vbCrLf &
            "==============================================" & vbCrLf & vbCrLf &
            "No de produit: " & vbTab & mtxtNoProduit.Text & vbCrLf &
            "Marque: " & vbTab & vbTab & txtMarque.Text & vbCrLf &
            "Modèle: " & vbTab & vbTab & txtModele.Text & vbCrLf &
            "Année: " & vbTab & vbTab & mtxtAnnee.Text & vbCrLf &
            "Couleur: " & vbTab & vbTab & txtCouleur.Text & vbCrLf &
            "Catégorie: " & vbTab & txtCategorie.Text & vbCrLf &
            "No de série: " & vbTab & txtNoSerie.Text & vbCrLf &
            "Prix: " & vbTab & vbTab & mtxtPrix.Text & vbCrLf &
            "Date de vente: " & vbTab & dtpickerDateVente.Value.ToShortDateString & vbCrLf

        txtFacturation.Text = Facture
    End Sub
    Private Sub btnAjouterCollection_Click(sender As Object, e As EventArgs) Handles btnAjouterCollection.Click
        Dim cle As String
        Dim Vente As New Vente

        cle = Me.mtxtNoClient.Text & "/" & Me.mtxtNoProduit.Text

        Vente.noProduit = Me.mtxtNoProduit.Text
        Vente.marque = Me.txtMarque.Text
        Vente.modele = Me.txtModele.Text
        Vente.annee = Me.mtxtAnnee.Text
        Vente.couleur = Me.txtCouleur.Text
        Vente.categorie = Me.txtCategorie.Text
        Vente.noSerie = Me.txtNoSerie.Text
        Vente.prix = Me.mtxtPrix.Text
        Vente.dateVente = Me.dtpickerDateVente.Value.ToShortDateString
        Vente.noClient = Me.mtxtNoClient.Text
        Vente.nomClient = Me.txtNom.Text
        Vente.prenomClient = Me.txtPrenom.Text
        Vente.adresse = Me.txtAdresse.Text
        Vente.ville = Me.txtVille.Text
        Vente.province = Me.CboProvince.SelectedItem.ToString
        Vente.codePostal = Me.mtxtCodePostal.Text
        Vente.courriel = Me.txtCourriel.Text
        Vente.telephone = Me.mtxtNoTel.Text


        If VenteClient.ContainsKey(cle) Then
            MsgBox("ERREUR : La clé (no de client) ==> " & cle & " <== existe déjà.")
        Else
            VenteClient.Add(cle, Vente)
            MsgBox("La fiche client ==> " & cle & " <== a été sauvegardée.")
            btnExtraireCollection.Enabled = True
            btnSupprimerCollection.Enabled = True
            btnSauvegarderCollection.Enabled = True
        End If
        Me.cboClientProduit.Items.Clear()
        For Each element In VenteClient.Keys
            Me.cboClientProduit.Items.Add(element)
        Next element
        Me.cboClientProduit.Sorted = True
    End Sub
    Private Sub btnExtraireCollection_Click(sender As Object, e As EventArgs) Handles btnExtraireCollection.Click
        Dim cle As String
        Dim Vente As New Vente
        If Me.cboClientProduit.Text <> "" Then
            cle = Me.cboClientProduit.Text
            If Not VenteClient.ContainsKey(cle) Then
                MsgBox("ERREUR : La clé (no de client) ==> " & CStr(cle) & " <== n'existe pas.")
            Else
                Vente = VenteClient.Item(cle)
            End If
            Me.mtxtNoProduit.Text = Vente.noProduit
            Me.txtMarque.Text = Vente.marque
            Me.txtModele.Text = Vente.modele
            Me.mtxtAnnee.Text = Vente.annee
            Me.txtCouleur.Text = Vente.couleur
            Me.txtCategorie.Text = Vente.categorie
            Me.txtNoSerie.Text = Vente.noSerie
            Me.mtxtPrix.Text = Vente.prix
            Me.mtxtNoProduit.Text = Vente.noProduit
            Me.dtpickerDateVente.Value = Vente.dateVente
            Me.mtxtNoClient.Text = Vente.noClient
            Me.txtNom.Text = Vente.nomClient
            Me.txtPrenom.Text = Vente.prenomClient
            Me.mtxtNoClient.Text = Vente.noClient
            Me.txtAdresse.Text = Vente.adresse
            Me.txtVille.Text = Vente.ville
            Me.CboProvince.SelectedIndex = CboProvince.FindStringExact(Vente.province)
            Me.mtxtCodePostal.Text = Vente.codePostal
            Me.txtCourriel.Text = Vente.courriel
            Me.mtxtNoTel.Text = Vente.telephone
        Else
            MsgBox("Sélectionner une valeur de la liste.")
        End If
    End Sub
    Private Sub btnSupprimerCollection_Click(sender As Object, e As EventArgs) Handles btnSupprimerCollection.Click
        If Me.cboClientProduit.Text = "" Or IsNothing(Me.cboClientProduit.Text) Then
            MsgBox("ERREUR : Sélectionner une  clé (no de client).")
            Exit Sub
        End If
        If Not VenteClient.ContainsKey(Me.cboClientProduit.Text) Then
            MsgBox("ERREUR : La clé (no de client) ==> " & Me.cboClientProduit.Text & " <== n'existe pas.")
        Else
            VenteClient.Remove(Me.cboClientProduit.Text)
            MsgBox("La fiche client ==> " & Me.cboClientProduit.Text & " <== a été supprimée.")

            '*If Me.mtxtNoClient.Text & "/" & Me.mtxtNoProduit.Text = Me.cboClientProduit.Text Then
            'Initialize()
            ' End If
            Me.cboClientProduit.Items.Remove(Me.cboClientProduit.Text)
            Me.cboClientProduit.Text = ""
        End If
        If VenteClient.Count = 0 Then
            btnExtraireCollection.Enabled = False
            btnSupprimerCollection.Enabled = False
            btnSauvegarderCollection.Enabled = False
        End If
    End Sub
    Private Sub btnSauvegarderCollection_Click(sender As Object, e As EventArgs) Handles btnSauvegarderCollection.Click
        Dim cle As String
        Dim Vente As Vente
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\Exercice009.txt")
        For Each cle In VenteClient.Keys
            Vente = VenteClient.Item(cle)
            enregistrement = Vente.noProduit & "|" &
                Vente.marque & "|" &
                Vente.modele & "|" &
                Vente.annee & "|" &
                Vente.couleur & "|" &
                Vente.categorie & "|" &
                Vente.noSerie & "|" &
                Vente.prix & "|" &
                Vente.dateVente & "|" &
                Vente.noClient & "|" &
                Vente.nomClient & "|" &
                Vente.prenomClient & "|" &
                Vente.adresse & "|" &
                Vente.ville & "|" &
                Vente.province & "|" &
                Vente.codePostal & "|" &
                Vente.courriel & "|" &
                Vente.telephone & vbCrLf
            If fileExists Then
                My.Computer.FileSystem.WriteAllText("..\..\..\Devoir01.txt",
                                                                        enregistrement & vbCrLf, True)
            Else
                My.Computer.FileSystem.WriteAllText("..\..\..\Devoir01.txt",
                                                   enregistrement & vbCrLf, False)
                fileExists = True
            End If
        Next
        If VenteClient.Count >= 5 Then
            btnFermer.Enabled = True
        End If
    End Sub


End Class

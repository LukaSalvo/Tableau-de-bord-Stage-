Private Sub btnExportTableStockage_Click()
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim db As DAO.Database
    Dim sql As String
    Dim mois As Integer
    Dim annee As Integer
    Dim dataFound As Boolean

    ' Récupérer les valeurs du mois et de l'année depuis le formulaire
    mois = Val(Nz(Forms!Formulaire!txtMois, 0)) ' Utiliser Val pour convertir en entier, 0 si vide
    annee = Val(Nz(Forms!Formulaire!txtAnnee, 0)) ' Utiliser Val pour convertir en entier, 0 si vide

    ' Vérifier si l'année est valide
    If annee = 0 Then
        MsgBox "Veuillez entrer une année valide.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Vérifier si le mois est valide
    If mois < 1 Or mois > 12 Then
        MsgBox "Le mois doit être une valeur entre 1 et 12.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Crée une nouvelle instance d'Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    ' Crée un nouveau classeur
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    ' Connexion à la base de données
    Set db = CurrentDb

    ' Définir les en-têtes
    xlSheet.Cells(1, 1).value = "Libellé"
    xlSheet.Cells(1, 2).value = "CASC"
    xlSheet.Cells(1, 3).value = "VILLE"
    xlSheet.Cells(1, 4).value = "ECOLE"
    xlSheet.Cells(1, 5).value = "COMMUN"

    ' Définir les données
    Dim data(1 To 10, 1 To 5) As Variant
    data(1, 1) = "Nombre de demandes ouvertes"
    data(2, 1) = "Nombre d'incidents ouverts"
    data(3, 1) = "Nombre de tickets ouverts"
    data(4, 1) = "Nombre de tickets fermés"
    data(5, 1) = "Pourcentage de demandes"
    data(6, 1) = "Pourcentage d'incidents"
    data(7, 1) = "Pourcentage de Tickets clos en moins de 4h"
    data(8, 1) = "Pourcentage de Tickets clos en moins de 24h"
    data(9, 1) = "Pourcentage de Tickets clos en moins de 7j"
    data(10, 1) = "Pourcentage de Tickets clos en plus de 7j"

    ' Initialiser le drapeau pour vérifier si des données sont trouvées
    dataFound = False

    ' Récupérer les données pour CASC avec le mois et l'année
    data(1, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 25 AND Mois = " & mois & " AND Annee = " & annee)
    data(2, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 29 AND Mois = " & mois & " AND Annee = " & annee)
    data(3, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 37 AND Mois = " & mois & " AND Annee = " & annee)
    data(4, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 33 AND Mois = " & mois & " AND Annee = " & annee)
    data(5, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 17 AND Mois = " & mois & " AND Annee = " & annee)
    data(6, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 21 AND Mois = " & mois & " AND Annee = " & annee)
    data(7, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 1 AND Mois = " & mois & " AND Annee = " & annee)
    data(8, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 5 AND Mois = " & mois & " AND Annee = " & annee)
    data(9, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 9 AND Mois = " & mois & " AND Annee = " & annee)
    data(10, 2) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 13 AND Mois = " & mois & " AND Annee = " & annee)

    ' Récupérer les données pour VILLE
    data(1, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 26 AND Mois = " & mois & " AND Annee = " & annee)
    data(2, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 30 AND Mois = " & mois & " AND Annee = " & annee)
    data(3, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 38 AND Mois = " & mois & " AND Annee = " & annee)
    data(4, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 34 AND Mois = " & mois & " AND Annee = " & annee)
    data(5, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 18 AND Mois = " & mois & " AND Annee = " & annee)
    data(6, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 22 AND Mois = " & mois & " AND Annee = " & annee)
    data(7, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 2 AND Mois = " & mois & " AND Annee = " & annee)
    data(8, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 6 AND Mois = " & mois & " AND Annee = " & annee)
    data(9, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 10 AND Mois = " & mois & " AND Annee = " & annee)
    data(10, 3) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 14 AND Mois = " & mois & " AND Annee = " & annee)

    ' Récupérer les données pour ECOLE
    data(1, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 27 AND Mois = " & mois & " AND Annee = " & annee)
    data(2, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 31 AND Mois = " & mois & " AND Annee = " & annee)
    data(3, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 39 AND Mois = " & mois & " AND Annee = " & annee)
    data(4, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 35 AND Mois = " & mois & " AND Annee = " & annee)
    data(5, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 19 AND Mois = " & mois & " AND Annee = " & annee)
    data(6, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 23 AND Mois = " & mois & " AND Annee = " & annee)
    data(7, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 3 AND Mois = " & mois & " AND Annee = " & annee)
    data(8, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 7 AND Mois = " & mois & " AND Annee = " & annee)
    data(9, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 11 AND Mois = " & mois & " AND Annee = " & annee)
    data(10, 4) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 15 AND Mois = " & mois & " AND Annee = " & annee)

    ' Récupérer les données pour COMMUN
    data(1, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 28 AND Mois = " & mois & " AND Annee = " & annee)
    data(2, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 32 AND Mois = " & mois & " AND Annee = " & annee)
    data(3, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 40 AND Mois = " & mois & " AND Annee = " & annee)
    data(4, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 36 AND Mois = " & mois & " AND Annee = " & annee)
    data(5, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 20 AND Mois = " & mois & " AND Annee = " & annee)
    data(6, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 24 AND Mois = " & mois & " AND Annee = " & annee)
    data(7, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 4 AND Mois = " & mois & " AND Annee = " & annee)
    data(8, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 8 AND Mois = " & mois & " AND Annee = " & annee)
    data(9, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 12 AND Mois = " & mois & " AND Annee = " & annee)
    data(10, 5) = GetSQLValue(db, "SELECT Resultats FROM TableStockageResultats WHERE id_requete = 16 AND Mois = " & mois & " AND Annee = " & annee)

    ' Vérifier si des données ont été trouvées (par exemple, pour CASC)
    Dim i As Integer, j As Integer
    For i = 1 To 10
        For j = 2 To 5
            If Not IsEmpty(data(i, j)) Then
                dataFound = True
                Exit For
            End If
        Next j
        If dataFound Then Exit For
    Next i

    ' Si aucune donnée n'est trouvée, afficher un message et arrêter
    If Not dataFound Then
        MsgBox "Aucune donnée trouvée dans TableStockageResultats pour le mois " & mois & " de l'année " & annee & ".", vbExclamation, "Erreur"
        xlBook.Close False
        xlApp.Quit
        Set db = Nothing
        Exit Sub
    End If

    ' Remplir les données dans la feuille Excel
    For i = 1 To 10
        For j = 1 To 5
            xlSheet.Cells(i + 1, j).value = data(i, j)
        Next j
    Next i

    ' Mise en forme des en-têtes
    With xlSheet.Range("A1:E1")
        .Interior.Color = vbLightBlue ' Utiliser une couleur grise (peut être ajustée avec un code RGB si besoin)
        .Font.Bold = True ' Mettre les en-têtes en gras
        .HorizontalAlignment = -4108 ' Centrer le texte (xlCenter)
    End With

    ' Ajuster la largeur des colonnes
    xlSheet.Columns("A:E").AutoFit

    ' Ajouter des bordures
    With xlSheet.Range("A1:E11").Borders
        .LineStyle = 1 ' xlContinuous
        .Weight = 2 ' xlThin
    End With

    ' Ajuster l'alignement du texte pour les données
    xlSheet.Range("A2:E11").HorizontalAlignment = -4108 ' Centrer le texte (xlCenter)

    ' Enregistrer le fichier avec le mois et l'année dans le nom
    filePath = Environ("USERPROFILE") & "\Desktop\Resultats_" & annee & "_" & Format(mois, "00") & ".xlsx"
    xlBook.SaveAs filePath
    xlBook.Close
    xlApp.Quit

    MsgBox "Les données ont été exportées avec succès vers " & filePath, vbInformation, "Succès"
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'exportation des données : " & Err.Description
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set db = Nothing
End Sub
Function GetSQLValue(db As DAO.Database, sqlQuery As String) As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sqlQuery, dbOpenSnapshot)
    If Not rs.EOF Then
        GetSQLValue = rs.fields(0).value
    Else
        GetSQLValue = 0
    End If
    rs.Close
    Set rs = Nothing
End Function



Private Sub btnExportGeneral_Click()
On Error GoTo ErrorHandler

    Dim filePath As String
    
    filePath = Environ("USERPROFILE") & "\Desktop\DataWarehouse_Tickets.xlsx"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "DataWarehouse_Tickets", filePath, True
    MsgBox "Les données ont été exportées avec succès vers " & filePath
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'exportation des données : " & Err.Description
End Sub
Private Sub btnImportCSVManualWithTypes_Click()
    
    Dim boite As FileDialog: Dim chemin As String
    Dim instanceE As Excel.Application
    Dim base As Database: Dim requete As String
    Dim ligne As Long: Dim nbLignes As Long
    Dim nom As String: Dim act As String: Dim de As String
    
    Set boite = Application.FileDialog(msoFileDialogFilePicker)
    If boite.Show Then chemin = boite.SelectedItems(1)
    
    If chemin <> "" Then
        Set instanceE = CreateObject("Excel.application")
        instanceE.Visible = False
        instanceE.Workbooks.Open chemin
        
        nbLignes = instanceE.Sheets(1).Cells.SpecialCells(xlCellTypeLastCell).row
        Set base = CurrentDb()
        ligne = 2
        
        While instanceE.Sheets(1).Cells(ligne, 2).value <> ""
            With instanceE.Sheets(1)
                Installation_Time = .Cells(ligne, 2).value
                Computer = .Cells(ligne, 3).value
                Installation_status = .Cells(ligne, 4).value
                Installation_code = .Cells(ligne, 5).value
                cisco_Systems_Inc = .Cells(ligne, 6).value
                Software = .Cells(ligne, 7).value
                Installed_Version = .Cells(ligne, 8).value
                Previously_Installed = .Cells(ligne, 9).value
                Company_Name = .Cells(ligne, 10).value
                
                Severity = .Cells(ligne, 11).value
                Category = .Cells(ligne, 12).value
                CVE_ID = .Cells(ligne, 13).value
                KB_ID = .Cells(ligne, 14).value
    
                requete = "INSERT INTO installation_history VALUES ('" & Installation_Time & "' , '" & Computer & "' , '" & Installation_status & "' , '" & Installation_code & "' , '" & cisco_Systems_Inc & "', '" & Software & "', '" & Installed_Version & "', '" & Previously_Installed & "', '" & Company_Name & "', '" & Severity & "', '" & Category & "', '" & CVE_ID & "', '" & KB_ID & "' )"
                base.Execute requete
                
                evolution.value = Int((ligne * 100) / nbLignes)
                ligne = ligne + 1
                End With
        Wend
        instanceE.Quit
        Set instanceE = Nothing
        base.Close
        Set base = Nothing
    End If
    
    MsgBox "L'importation est réussie"
    Set boite = Nothing
    
    
    End Sub
Private Sub btnSelectionMultiple_Click()
    Dim db As DAO.Database
    Dim sql As String
    Dim selectedIDs As Collection
    Dim rs As DAO.Recordset
    Dim selectedID As Long
    Dim i As Integer
    Dim selItem As Variant

    Application.Echo False
    DoCmd.Echo False
    DoCmd.SetWarnings False

    On Error GoTo ErrorHandler
    Set db = CurrentDb
    Set selectedIDs = New Collection

    Debug.Print "Vérification des enregistrements dans le sous-formulaire..."
    If Me!SousFormulaire.Form.Recordset.EOF And Me!SousFormulaire.Form.Recordset.BOF Then
        MsgBox "Aucune ligne sélectionnée dans le sous-formulaire.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If

    ' Vérifiez le nombre de lignes sélectionnées
    Debug.Print "Nombre de lignes sélectionnées : " & Me!SousFormulaire.Form.SelCount
    If Me!SousFormulaire.Form.SelCount = 0 Then
        MsgBox "Aucune ligne sélectionnée dans le sous-formulaire.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If

    ' Parcourir les lignes sélectionnées
    Debug.Print "Parcours des lignes sélectionnées..."
    For Each selItem In Me!SousFormulaire.Form.SelItems
        Debug.Print "Vérification de la ligne sélectionnée : " & selItem
        If Me!SousFormulaire.Form.Recordset.AbsolutePosition = selItem Then
            selectedIDs.Add Me!SousFormulaire.Form!ID.value
            Debug.Print "ID sélectionné : " & Me!SousFormulaire.Form!ID.value
        End If
    Next selItem

    If selectedIDs.Count = 0 Then
        MsgBox "Aucune ligne valide sélectionnée ou ID non défini.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If

    On Error Resume Next
    db.Execute "DROP TABLE Temp_SingleFiltered_Tickets", dbFailOnError
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    sql = "SELECT id, entities_id, nom_entities, titre_tickets, status_tickets, description, type, nom_type, id_categorie, nom_categorie, demandeurs, techniciens, cout " & _
          "INTO Temp_MultiFiltered_Tickets " & _
          "FROM Temp_Filtered_Tickets " & _
          "WHERE id IN (" & Join(selectedIDs, ",") & ")"

    Debug.Print "Exécution de la requête SQL : " & sql
    db.Execute sql, dbFailOnError

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM Temp_MultiFiltered_Tickets")
    If rs!Total = 0 Then
        MsgBox "Aucune donnée trouvée pour les IDs sélectionnés.", vbInformation, "Information"
        rs.Close
        GoTo Cleanup
    End If
    rs.Close

    On Error Resume Next
    Me!SousFormulaire.Form.RecordSource = "Temp_MultiFiltered_Tickets"
    If Err.Number <> 0 Then
        Debug.Print "Erreur lors de la mise à jour du RecordSource : " & Err.Description
        GoTo Cleanup
    End If

    Me!SousFormulaire.Requery
    If Err.Number <> 0 Then
        Debug.Print "Erreur lors du Requery : " & Err.Description
        GoTo Cleanup
    End If

    MsgBox "Affichage limité aux lignes sélectionnées pour les IDs " & Join(selectedIDs, ", ") & " ! Vous pouvez maintenant modifier les champs via les combo boxes.", vbInformation, "Succès"

Cleanup:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    DoCmd.SetWarnings True
    DoCmd.Echo True
    Application.Echo True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.Description & vbCrLf & "Requête : " & sql, vbCritical, "Erreur"
    GoTo Cleanup
End Sub



    


Private Sub CommandeTest_Click()
    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field
    Dim rsTickets As DAO.Recordset
    Dim rsEntities As DAO.Recordset
    Dim rsCost As DAO.Recordset
    Dim rsCategorie As DAO.Recordset
    Dim rsDemandeurs As DAO.Recordset
    Dim rsTechniciens As DAO.Recordset
    Dim rsDW As DAO.Recordset
    Dim sql As String
    Dim demandeursList As String
    Dim techniciensList As String

    Dim mois As Variant
    Dim annee As Integer

    On Error GoTo ErrorHandler
    Set db = CurrentDb

    mois = Nz(Forms!Formulaire!txtMois, "")
    annee = Val(Nz(Forms!Formulaire!txtAnnee, 0))
    If annee = 0 Then
        MsgBox "Veuillez entrer une année valide.", vbExclamation, "Erreur"
        Exit Sub
    End If

    Dim moisInt As Integer
    If mois = "" Then
        moisInt = 0
    Else
        moisInt = Val(mois)
        If moisInt < 1 Or moisInt > 12 Then
            MsgBox "Le mois doit être une valeur entre 1 et 12.", vbExclamation, "Erreur"
            Exit Sub
        End If
    End If

    On Error Resume Next
    db.TableDefs.Delete "DataWarehouse_Tickets"
    On Error GoTo ErrorHandler

    ' Créer la table DataWarehouse_Tickets
    Set tblDef = db.CreateTableDef("DataWarehouse_Tickets")

    ' Ajouter les champs à la table
    With tblDef
        ' id (clé primaire)
        Set fld = .CreateField("id", dbLong)
        .fields.Append fld

        ' id_entities
        .fields.Append .CreateField("entities_id", dbLong)

        ' nom_entities
        .fields.Append .CreateField("nom_entities", dbText, 50)

        ' titre_tickets
        .fields.Append .CreateField("titre_tickets", dbMemo)

        ' mois
        .fields.Append .CreateField("mois", dbLong)

        ' annee
        .fields.Append .CreateField("annee", dbLong)

        ' date_ouverture
        .fields.Append .CreateField("date_ouverture", dbDate)

        ' date_fermeture
        .fields.Append .CreateField("date_fermeture", dbDate)

        ' date_resolution
        .fields.Append .CreateField("date_resolution", dbDate)

        ' status_tickets
        .fields.Append .CreateField("status_tickets", dbLong)

        ' description
        .fields.Append .CreateField("description", dbMemo)

        ' type
        .fields.Append .CreateField("type", dbLong)

        ' nom_type
        .fields.Append .CreateField("nom_type", dbText, 50)

        ' id_categorie
        .fields.Append .CreateField("id_categorie", dbLong)

        ' nom_categorie
        .fields.Append .CreateField("nom_categorie", dbText, 255)

        ' demandeurs (concaténation des noms et prénoms)
        .fields.Append .CreateField("demandeurs", dbText, 255)

        ' techniciens (concaténation des noms et prénoms)
        .fields.Append .CreateField("techniciens", dbText, 255)

        ' cout
        .fields.Append .CreateField("cout", dbDouble)
    End With

    db.TableDefs.Append tblDef

    ' Vérifier si des données existent déjà pour le mois et l'année spécifiés
    If moisInt > 0 Then
        sql = "SELECT COUNT(*) AS Total FROM DataWarehouse_Tickets WHERE Mois = " & moisInt & " AND Annee = " & annee
        Set rsDW = db.OpenRecordset(sql, dbOpenSnapshot)
        If rsDW!Total > 0 Then
            MsgBox "Les données pour le mois " & moisInt & " de l'année " & annee & " sont déjà insérées dans DataWarehouse_Tickets.", vbExclamation, "Avertissement"
            rsDW.Close
            GoTo Cleanup
        End If
        rsDW.Close
    End If

    sql = "SELECT t.id, t.entities_id, t.name, t.date, t.closedate, t.solvedate, t.status, t.content, t.type, t.itilcategories_id " & _
          "FROM glpi_tickets AS t " & _
          "WHERE Year(t.date) = " & annee
    If moisInt > 0 Then
        sql = sql & " AND Month(t.date) = " & moisInt
    End If
    Set rsTickets = db.OpenRecordset(sql, dbOpenDynaset)

    If rsTickets.EOF Then
        MsgBox "Aucun ticket trouvé pour l'année " & annee & IIf(moisInt > 0, " et le mois " & moisInt, "") & ".", vbInformation, "Information"
        rsTickets.Close
        GoTo Cleanup
    End If

    sql = "SELECT e.id, e.name FROM glpi_entities AS e"
    Set rsEntities = db.OpenRecordset(sql, dbOpenDynaset)

    Set rsDW = db.OpenRecordset("DataWarehouse_Tickets", dbOpenDynaset)

    Do While Not rsTickets.EOF
        rsDW.AddNew

        rsDW!ID = rsTickets!ID
        rsDW!entities_id = rsTickets!entities_id

        If Nz(rsTickets!Name, "") = "" Then
            rsDW!titre_tickets = "Titre non défini"
        Else
            rsDW!titre_tickets = rsTickets!Name
        End If

        If Not IsNull(rsTickets!Date) Then
            rsDW!mois = month(rsTickets!Date)
            rsDW!annee = year(rsTickets!Date)
        Else
            rsDW!mois = 0
            rsDW!annee = 0
        End If

        rsDW!date_ouverture = rsTickets!Date
        rsDW!date_fermeture = Nz(rsTickets!closedate, Null)
        rsDW!date_resolution = Nz(rsTickets!solvedate, Null)

        rsDW!Description = CleanText(Nz(rsTickets!content, "Aucune description"))

        rsDW!status_tickets = rsTickets!status
        rsDW!Type = rsTickets!Type

        If rsTickets!Type = 1 Then
            rsDW!nom_type = "Incident"
        ElseIf rsTickets!Type = 2 Then
            rsDW!nom_type = "Demande"
        Else
            rsDW!nom_type = "Inconnu"
        End If

        rsEntities.FindFirst "id = " & rsTickets!entities_id
        If Not rsEntities.NoMatch Then
            rsDW!nom_entities = rsEntities!Name
        Else
            rsDW!nom_entities = "Entité inconnue"
        End If

        sql = "SELECT id AS id_categorie ,completename AS nom_categorie  FROM glpi_itilcategories WHERE id = " & Nz(rsTickets!itilcategories_id, 0)
        Set rsCategorie = db.OpenRecordset(sql, dbOpenDynaset)

        If Not rsCategorie.EOF And Not IsNull(rsCategorie!nom_categorie) Then
            rsDW!id_categorie = rsCategorie!id_categorie
            rsDW!nom_categorie = rsCategorie!nom_categorie
        Else
            rsDW!nom_categorie = " "
        End If
        rsCategorie.Close

        demandeursList = ""
        sql = "SELECT u.firstname, u.realname " & _
              "FROM glpi_tickets_users tu " & _
              "LEFT JOIN glpi_users u ON u.id = tu.users_id " & _
              "WHERE tu.tickets_id = " & rsTickets!ID & " AND tu.type = 1"
        Set rsDemandeurs = db.OpenRecordset(sql, dbOpenDynaset)
        Do While Not rsDemandeurs.EOF
            If Not IsNull(rsDemandeurs!realname) And Not IsNull(rsDemandeurs!firstname) Then
                If demandeursList = "" Then
                    demandeursList = rsDemandeurs!realname & " " & rsDemandeurs!firstname
                Else
                    demandeursList = demandeursList & ", " & rsDemandeurs!realname & " " & rsDemandeurs!firstname
                End If
            End If
            rsDemandeurs.MoveNext
        Loop
        rsDemandeurs.Close
        rsDW!demandeurs = IIf(demandeursList = "", " ", demandeursList)

        techniciensList = ""
        sql = "SELECT u.firstname, u.realname " & _
              "FROM glpi_tickets_users tu " & _
              "LEFT JOIN glpi_users u ON u.id = tu.users_id " & _
              "WHERE tu.tickets_id = " & rsTickets!ID & " AND tu.type = 2"
        Set rsTechniciens = db.OpenRecordset(sql, dbOpenDynaset)
        Do While Not rsTechniciens.EOF
            If Not IsNull(rsTechniciens!realname) And Not IsNull(rsTechniciens!firstname) Then
                If techniciensList = "" Then
                    techniciensList = rsTechniciens!realname & " " & rsTechniciens!firstname
                Else
                    techniciensList = techniciensList & ", " & rsTechniciens!realname & " " & rsTechniciens!firstname
                End If
            End If
            rsTechniciens.MoveNext
        Loop
        rsTechniciens.Close
        rsDW!techniciens = IIf(techniciensList = "", " ", techniciensList)

        sql = "SELECT SUM(actiontime) / 60 AS cout FROM glpi_ticketcosts WHERE tickets_id = " & rsTickets!ID
        Set rsCost = db.OpenRecordset(sql, dbOpenDynaset)
        If Not rsCost.EOF And Not IsNull(rsCost!cout) Then
            rsDW!cout = rsCost!cout
        Else
            rsDW!cout = 0
        End If
        rsCost.Close

        Debug.Print "ID: " & rsTickets!ID & ", Titre: " & Nz(rsTickets!Name, "NULL") & _
                   ", Description: " & Nz(rsTickets!content, "NULL") & _
                   ", Catégorie: " & Nz(rsDW!nom_categorie, "NULL") & _
                   ", Demandeurs: " & rsDW!demandeurs & _
                   ", Techniciens: " & rsDW!techniciens

        rsDW.Update

        rsTickets.MoveNext
    Loop

Cleanup:

    rsTickets.Close
    rsEntities.Close
    rsDW.Close
    Set rsTickets = Nothing
    Set rsEntities = Nothing
    Set rsCost = Nothing
    Set rsCategorie = Nothing
    Set rsDemandeurs = Nothing
    Set rsTechniciens = Nothing
    Set rsDW = Nothing
    Set tblDef = Nothing
    Set db = Nothing

    MsgBox "Table créée et données insérées avec succès dans le data warehouse !", vbInformation, "Succès"
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur"
    End If
    Resume Cleanup
End Sub


Private Sub insert_Click()
    On Error GoTo ErrorHandler

    ' Désactiver les avertissements
    DoCmd.SetWarnings False

    ' Déclarer les variables pour les paramètres
    Dim mois As Integer
    Dim annee As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Integer
    Dim moisInsere As Boolean
    Dim moisExistants As String
    Dim moisManquants As String

    ' Récupérer les valeurs des contrôles du formulaire
    mois = Val(Nz(Forms!Formulaire!txtMois, 0)) ' Utiliser Val pour convertir en entier, 0 si vide
    annee = Val(Nz(Forms!Formulaire!txtAnnee, 0)) ' Utiliser Val pour convertir en entier, 0 si vide

    ' Vérifier si l'année est valide
    If annee = 0 Then
        MsgBox "Veuillez entrer une année valide.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Vérifier si le mois est valide (si spécifié)
    If mois <> 0 Then
        If mois < 1 Or mois > 12 Then
            MsgBox "Le mois doit être une valeur entre 1 et 12.", vbExclamation, "Erreur"
            Exit Sub
        End If
    End If

    Set db = CurrentDb

    ' Si l'utilisateur veut insérer pour une année entière, vérifier que DataWarehouse_Tickets contient des données
    If mois = 0 Then
        moisManquants = ""
        For i = 1 To 12
            sql = "SELECT COUNT(*) AS Total FROM DataWarehouse_Tickets WHERE Mois = " & i & " AND Annee = " & annee
            Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
            If rs!Total = 0 Then
                If moisManquants = "" Then
                    moisManquants = CStr(i)
                Else
                    moisManquants = moisManquants & ", " & CStr(i)
                End If
            End If
            rs.Close
            Set rs = Nothing
        Next i

        If moisManquants <> "" Then
            MsgBox "Aucune donnée trouvée dans DataWarehouse_Tickets pour les mois " & moisManquants & " de l'année " & annee & ". Veuillez d'abord générer les données via le bouton 'Générer Data Warehouse'.", vbExclamation, "Erreur"
            Set db = Nothing
            DoCmd.SetWarnings True
            Exit Sub
        End If
    End If


    Call ExecuteQueryForCreate(db, "creation_table_stockage_resultats")

    If mois <> 0 Then
        sql = "SELECT COUNT(*) AS Total FROM DataWarehouse_Tickets WHERE Mois = " & mois & " AND Annee = " & annee
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        If rs!Total = 0 Then
            MsgBox "Aucune donnée trouvée dans DataWarehouse_Tickets pour le mois " & mois & " de l'année " & annee & ". Veuillez d'abord générer les données via le bouton 'Générer Data Warehouse'.", vbExclamation, "Erreur"
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            DoCmd.SetWarnings True
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        sql = "SELECT COUNT(*) AS Total FROM TableStockageResultats WHERE Mois = " & mois & " AND Annee = " & annee
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        If rs!Total > 0 Then
            MsgBox "Les données pour le mois " & mois & " de l'année " & annee & " sont déjà insérées dans TableStockageResultats.", vbExclamation, "Avertissement"
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            DoCmd.SetWarnings True
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_COMMUN", mois, annee)
        
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_CASC", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_VILLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_ECOLE", mois, annee)
        Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_COMMUN", mois, annee)


        sql = "UPDATE TableStockageResultats SET Resultats = Round([Resultats], 2) WHERE Resultats IS NOT NULL AND Mois = " & mois & " AND Annee = " & annee
        db.Execute sql, dbFailOnError

        MsgBox "Données insérées avec succès pour le mois " & mois & " de l'année " & annee & " !", vbInformation, "Succès"
    Else

        moisInsere = False
        moisExistants = ""

        For i = 1 To 12

            sql = "SELECT COUNT(*) AS Total FROM TableStockageResultats WHERE Mois = " & i & " AND Annee = " & annee
            Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
            If rs!Total > 0 Then

                If moisExistants = "" Then
                    moisExistants = CStr(i)
                Else
                    moisExistants = moisExistants & ", " & CStr(i)
                End If
            Else

                Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_demandes_ouvertes_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_incidents_ouverts_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_fermés_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "nb_tickets_ouverts_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_demandes_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_incidents_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_4h_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_24h_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_7j_COMMUN", i, annee)
                
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_CASC", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_VILLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_ECOLE", i, annee)
                Call ExecuteSavedQueryWithParameters(db, "pourcentage_clos_+7j_COMMUN", i, annee)


                sql = "UPDATE TableStockageResultats SET Resultats = Round([Resultats], 2) WHERE Resultats IS NOT NULL AND Mois = " & i & " AND Annee = " & annee
                db.Execute sql, dbFailOnError

                moisInsere = True
            End If
            rs.Close
            Set rs = Nothing
        Next i

        If moisInsere Then
            Dim msg As String
            msg = "Données insérées avec succès pour l'année " & annee & " !"
            If moisExistants <> "" Then
                msg = msg & vbCrLf & "Note : Les données pour les mois " & moisExistants & " étaient déjà insérées et n'ont pas été modifiées."
            End If
            MsgBox msg, vbInformation, "Succès"
        Else
            MsgBox "Aucune nouvelle donnée n'a été insérée pour l'année " & annee & " car tous les mois sont déjà présents dans TableStockageResultats.", vbInformation, "Information"
        End If
    End If


    Set db = Nothing
    DoCmd.SetWarnings True
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'exécution des requêtes : " & Err.Description
    DoCmd.SetWarnings True
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Sub ExecuteSavedQueryWithParameters(db As DAO.Database, queryName As String, mois As Integer, annee As Integer)

    On Error GoTo QueryErrorHandler
    
    Dim qdef As DAO.QueryDef
    Set qdef = db.QueryDefs(queryName)
    qdef.Parameters("Mois").value = mois
    qdef.Parameters("Annee").value = annee
    qdef.Execute dbFailOnError
    
    Exit Sub

QueryErrorHandler:
    MsgBox "Erreur lors de l'exécution de la requête '" & queryName & "': " & Err.Description
End Sub
Private Sub ExecuteQueryForCreate(db As DAO.Database, queryName As String)

    On Error GoTo QueryErrorHandler
    
    Dim qdef As DAO.QueryDef
    Set qdef = db.QueryDefs(queryName)
    qdef.Execute dbFailOnError
    Exit Sub

QueryErrorHandler:
    MsgBox "Erreur lors de l'exécution de la requête '" & queryName & "': " & Err.Description
End Sub



Private Function tableExists(tableName As String) As Boolean
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    
    Set db = CurrentDb
    For Each tbl In db.TableDefs
        If tbl.Name = tableName Then
            tableExists = True
            Exit Function
        End If
    Next tbl
    tableExists = False
End Function


Private Function CleanText(Text As String) As String

    If Text = "" Or Text = "Aucune description" Then
        CleanText = Text
        Exit Function
    End If
    
    Text = Replace(Text, "&#60;p&#62;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;/p&#62;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;p class=", " ", , , vbTextCompare)
    Text = Replace(Text, "&#62;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#38;nb", " ", , , vbTextCompare)
    Text = Replace(Text, "sp;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;br", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;strong", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;/strong", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;span style=", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;/span", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;span", " ", , , vbTextCompare)
    Text = Replace(Text, "mso-ligatures: none; mso-fareast-language: FR;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#38;lt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#38;gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "mso-outline-level: 1;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#38;gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;", " ", , , vbTextCompare)
    Text = Replace(Text, "style=", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;div class", " ", , , vbTextCompare)
    Text = Replace(Text, "&#60;h1 class=", " ", , , vbTextCompare)
    Text = Replace(Text, "class=", " ", , , vbTextCompare)
    Text = Replace(Text, "MsoNormal", " ", , , vbTextCompare)
    Text = Replace(Text, """", " ", , , vbTextCompare)
    Text = Replace(Text, "mso-fareast-font-family: 'Times New Roman';", " ", , , vbTextCompare)
    Text = Replace(Text, "href=", " ", , , vbTextCompare)
    Text = Replace(Text, "margin-bottom: 12.", " ", , , vbTextCompare)
    Text = Replace(Text, "0pt;", " ", , , vbTextCompare)
    Text = Replace(Text, "href=", " ", , , vbTextCompare)
    Text = Replace(Text, "margin-bottom:", " ", , , vbTextCompare)
    Text = Replace(Text, "pt;", " ", , , vbTextCompare)
    Text = Replace(Text, "mso-fareast-language: FR;", " ", , , vbTextCompare)
    Text = Replace(Text, "&lt;p&gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&lt;/p&gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&lt;strong&gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&lt;/strong&gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "&lt;", " ", , , vbTextCompare)
    Text = Replace(Text, "/&gt;", " ", , , vbTextCompare)
    Text = Replace(Text, "br", " ", , , vbTextCompare)
    Text = Replace(Text, "&#039;", " ", , , vbTextCompare)
    Text = Replace(Text, "&gt;", " ", , , vbTextCompare)
    
    Text = Replace(Text, vbCrLf, " ", , , vbTextCompare)
    Text = Replace(Text, vbTab, " ", , , vbTextCompare)
    Text = Trim(Text)
    
    If Text = "" Then
        CleanText = "Aucune description (texte nettoyé vide)"
    Else
    
        CleanText = Text
    End If
    
End Function

Private Sub CommandeFilter_Click()
    Dim db As DAO.Database
    Dim sql As String
    Dim whereClause As String
    Dim entityFilter As String
    Dim typeFilter As String
    Dim categoryFilter As String
    Dim keywordFilter As String
    Dim rsCheck As DAO.Recordset
    
    Application.Echo False
    DoCmd.Echo False
    DoCmd.SetWarnings False
    
    On Error GoTo ErrorHandler
    Set db = CurrentDb
    

    whereClause = ""

    entityFilter = Nz(Forms!Formulaire!liste_entities, "")
    If entityFilter <> "" Then
        whereClause = whereClause & "nom_entities = '" & entityFilter & "'"
    End If

    typeFilter = Nz(Forms!Formulaire!liste_type, "")
    If typeFilter <> "" Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        whereClause = whereClause & "nom_type = '" & typeFilter & "'"
    End If
    
  
    categoryFilter = Nz(Forms!Formulaire!liste_categorie, "")
    If categoryFilter <> "" Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        If categoryFilter = "Sans catégorie" Then
            whereClause = whereClause & "nom_categorie IS NULL OR nom_categorie = ' '"
        Else
            whereClause = whereClause & "nom_categorie = '" & categoryFilter & "'"
        End If
    End If
    
    
    keywordFilter = Nz(Forms!Formulaire!MotAFiltrer, "")
    If keywordFilter <> "" Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        whereClause = whereClause & "(UCASE(titre_tickets) LIKE '*" & UCase(keywordFilter) & "*' OR UCASE(description) LIKE '*" & UCase(keywordFilter) & "*')"
    End If
    
    If Forms!Formulaire!ToggleSansCategorie.value = True Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        whereClause = whereClause & "(nom_categorie IS NULL OR nom_categorie = ' ')"
    End If
    

    On Error Resume Next
    db.Execute "DELETE FROM Temp_Filtered_Tickets", dbFailOnError
    If Err.Number = 0 Then
    Else
        sql = "SELECT id, entities_id, nom_entities, titre_tickets, status_tickets, description, type, nom_type, id_categorie, nom_categorie, demandeurs, techniciens, cout INTO Temp_Filtered_Tickets FROM DataWarehouse_Tickets WHERE 1=0"
        db.Execute sql, dbFailOnError
    End If
    On Error GoTo ErrorHandler
    
    
    sql = "INSERT INTO Temp_Filtered_Tickets (id, entities_id, nom_entities, titre_tickets, status_tickets, description, type, nom_type, id_categorie, nom_categorie, demandeurs, techniciens, cout) " & _
          "SELECT id, entities_id, nom_entities, titre_tickets, status_tickets, description, type, nom_type, id_categorie, nom_categorie, demandeurs, techniciens, cout FROM DataWarehouse_Tickets"
    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If
   
    db.Execute sql, dbFailOnError
  
    On Error Resume Next
    Set rsCheck = db.OpenRecordset("SELECT COUNT(*) AS Total FROM Temp_Filtered_Tickets")
    If Err.Number <> 0 Then
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    If rsCheck!Total = 0 Then
        MsgBox "Aucun ticket trouvé avec les filtres sélectionnés.", vbInformation, "Information"
        GoTo Cleanup
    End If
    rsCheck.Close
    
    On Error Resume Next
    Me!SousFormulaire.Form.RecordSource = "Temp_Filtered_Tickets"
    If Err.Number <> 0 Then
        GoTo Cleanup
    End If
    Debug.Print "RecordSource mis à jour avec Temp_Filtered_Tickets"
    
    Me!SousFormulaire.Requery
    If Err.Number <> 0 Then
        GoTo Cleanup
    End If
   
    MsgBox "Résultats filtrés affichés dans une table éditable !", vbInformation, "Succès"
    
Cleanup:
    If Not rsCheck Is Nothing Then rsCheck.Close
    Set rsCheck = Nothing
    Set db = Nothing
    DoCmd.SetWarnings True
    DoCmd.Echo True
    Application.Echo True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.Description & vbCrLf & "Requête : " & sql, vbCritical, "Erreur"
    Debug.Print "Erreur capturée : " & Err.Description
    GoTo Cleanup
End Sub


Private Sub btnSeulFiltre_Click()
    Dim db As DAO.Database
    Dim sql As String
    Dim selectedID As Long
    Dim rs As DAO.Recordset
    
    Application.Echo False
    DoCmd.Echo False
    DoCmd.SetWarnings False
    
    On Error GoTo ErrorHandler
    Set db = CurrentDb
    
  
    If Me!SousFormulaire.Form.Recordset.EOF And Me!SousFormulaire.Form.Recordset.BOF Then
        MsgBox "Aucune ligne sélectionnée dans le sous-formulaire.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If
    
 
    selectedID = Me!SousFormulaire.Form!ID.value
    If selectedID = 0 Then
        MsgBox "Aucune ligne valide sélectionnée ou ID non défini.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If
    
    
    On Error Resume Next
    db.Execute "DROP TABLE Temp_SingleFiltered_Tickets", dbFailOnError
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
   
    sql = "SELECT id, entities_id, nom_entities, titre_tickets, status_tickets, description, type, nom_type, id_categorie, nom_categorie, demandeurs, techniciens, cout " & _
          "INTO Temp_SingleFiltered_Tickets " & _
          "FROM Temp_Filtered_Tickets " & _
          "WHERE id = " & selectedID

    db.Execute sql, dbFailOnError

   
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM Temp_SingleFiltered_Tickets")
    If rs!Total = 0 Then
        MsgBox "Aucune donnée trouvée pour l'ID sélectionné.", vbInformation, "Information"
        rs.Close
        GoTo Cleanup
    End If
    rs.Close
    

    On Error Resume Next
    Me!SousFormulaire.Form.RecordSource = "Temp_SingleFiltered_Tickets"
    If Err.Number <> 0 Then
        Debug.Print "Erreur lors de la mise à jour du RecordSource : " & Err.Description
        GoTo Cleanup
    End If

    Me!SousFormulaire.Requery
    If Err.Number <> 0 Then
        Debug.Print "Erreur lors du Requery : " & Err.Description
        GoTo Cleanup
    End If

    MsgBox "Affichage limité à la ligne sélectionnée pour l'ID " & selectedID & " ! Vous pouvez maintenant modifier les champs via les combo boxes.", vbInformation, "Succès"
    
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    DoCmd.SetWarnings True
    DoCmd.Echo True
    Application.Echo True
    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.Description & vbCrLf & "Requête : " & sql, vbCritical, "Erreur"
    GoTo Cleanup
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandler
    

    Me!comboNomEntities.RowSource = "SELECT id, name FROM [glpi_entities] WHERE id IN (0, 1, 2, 3, 4, 6) ORDER BY name;"
    Me!comboNomEntities.ColumnCount = 2
    Me!comboNomEntities.ColumnWidths = "0cm;3cm"
    Me!comboNomEntities.BoundColumn = 1

    Me!comboNomType.RowSource = "SELECT DISTINCT type, nom_type FROM DataWarehouse_Tickets WHERE nom_type IN ('Incident', 'Demande') ORDER BY nom_type;"
    Me!comboNomType.ColumnCount = 2
    Me!comboNomType.ColumnWidths = "0cm;3cm"
    Me!comboNomType.BoundColumn = 1
    If Me!comboNomType.ListCount = 0 Then
        MsgBox "Aucune valeur de type ('Incident' ou 'Demande') trouvée dans DataWarehouse_Tickets.", vbExclamation, "Erreur"
    End If
   
    Me!comboNomCategorie.RowSource = "SELECT id, completename FROM glpi_itilcategories ORDER BY completename;"
    Me!comboNomCategorie.ColumnCount = 2
    Me!comboNomCategorie.ColumnWidths = "0cm;3cm"
    Me!comboNomCategorie.BoundColumn = 1

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tableExists As Boolean
    Dim sqlCreateTable As String
    
    Set db = CurrentDb
    tableExists = False
    
    For Each tdf In db.TableDefs
        If tdf.Name = "Temp_Filtered_Tickets" Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
          sqlCreateTable = "CREATE TABLE Temp_Filtered_Tickets (" & _
                         "id AUTOINCREMENT PRIMARY KEY, " & _
                         "nom_entities TEXT(255), " & _
                         "nom_type TEXT(50), " & _
                         "nom_categorie TEXT(255), " & _
                         "date_ouverture DATETIME, " & _
                         "date_fermeture DATETIME, " & _
                         "entities_id INTEGER, " & _
                         "type INTEGER, " & _
                         "id_categorie INTEGER);"
        db.Execute sqlCreateTable, dbFailOnError
    End If
      Me!SousFormulaire.Form.RecordSource = "SELECT * FROM Temp_Filtered_Tickets;"
    
    If DCount("*", "Temp_Filtered_Tickets") = 0 Then
            db.Execute "INSERT INTO Temp_Filtered_Tickets " & _
            "SELECT * FROM DataWarehouse_Tickets;", dbFailOnError
          End If
    
    Me!SousFormulaire.Form.Requery

    
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Erreur dans Form_Load"
    
    
    
    Debug.Print "Erreur dans Form_Load : " & Err.Number & " - " & Err.Description
    Set db = Nothing
End Sub
Private Sub btnSauvegarderModifications_Click()
    MsgBox "Bouton Sauvegarder cliqué !", vbInformation, "Test"
    
    Dim db As DAO.Database
    Dim sql As String
    Dim selectedID As Long
    Dim rs As DAO.Recordset
    Dim updateFields As String
    Dim hasUpdates As Boolean
    Dim entities_id As Long
    Dim ticketType As Long
    Dim id_categorie As Long
    Dim nom_entities As String
    Dim nom_type As String
    Dim nom_categorie As String
    Dim current_entities_id As Long
    Dim current_ticketType As Long
    Dim current_id_categorie As Long
    
    On Error GoTo ErrorHandler
    Debug.Print "=== Début de btnSauvegarderModifications_Click ==="
    Set db = CurrentDb


    If Not tableExists("Temp_SingleFiltered_Tickets") Then
        MsgBox "La table Temp_SingleFiltered_Tickets n'existe pas. Veuillez d'abord sélectionner une ligne en cliquant sur 'Sélection'.", vbExclamation, "Erreur"
        GoTo Cleanup
    End If
   
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM Temp_SingleFiltered_Tickets")
    If rs!Total = 0 Then
        MsgBox "Aucune ligne sélectionnée pour modification. Veuillez d'abord sélectionner une ligne en cliquant sur 'Sélection'.", vbExclamation, "Erreur"
        rs.Close
        Set rs = Nothing
        GoTo Cleanup
    End If
    rs.Close
    Set rs = Nothing

    
   
    Set rs = db.OpenRecordset("SELECT id, entities_id, type, id_categorie FROM Temp_SingleFiltered_Tickets")
    If rs.EOF Then
        MsgBox "Erreur : Aucune donnée dans Temp_SingleFiltered_Tickets.", vbCritical, "Erreur"
        rs.Close
        Set rs = Nothing
        GoTo Cleanup
    End If
    selectedID = rs!ID
    current_entities_id = Nz(rs!entities_id, 0)
    current_ticketType = Nz(rs!Type, 0)
    current_id_categorie = Nz(rs!id_categorie, 0)
    rs.Close
    Set rs = Nothing

    If Not tableExists("DataWarehouse_Tickets") Then
        MsgBox "La table DataWarehouse_Tickets n'existe pas ou n'est plus accessible.", vbCritical, "Erreur"
        GoTo Cleanup
    End If
   
    updateFields = ""
    hasUpdates = False
 
    
   
    If Not IsNull(Me!comboNomEntities) Then
        If IsNumeric(Nz(Me!comboNomEntities, "")) Then
            entities_id = CLng(Me!comboNomEntities)
            If entities_id <> current_entities_id Then
                Set rs = db.OpenRecordset("SELECT name FROM [glpi_entities] WHERE id = " & entities_id)
                If Not rs.EOF Then
                    nom_entities = Nz(rs!Name, "")
                    If updateFields <> "" Then updateFields = updateFields & ", "
                    updateFields = updateFields & "nom_entities = '" & Replace(nom_entities, "'", "''") & "', entities_id = " & entities_id
                    hasUpdates = True
                Else
                    MsgBox "Entité non trouvée pour l'ID : " & entities_id, vbExclamation, "Erreur"
                    rs.Close
                    Set rs = Nothing
                    GoTo Cleanup
                End If
                rs.Close
                Set rs = Nothing
            End If
        Else
            MsgBox "La valeur sélectionnée dans la combo box 'Entités' n'est pas un ID valide.", vbExclamation, "Erreur"
            GoTo Cleanup
        End If
    End If
    

    If Not IsNull(Me!comboNomType) Then
        If IsNumeric(Nz(Me!comboNomType, "")) Then
            ticketType = CLng(Me!comboNomType)
            If ticketType <> current_ticketType Then
                Select Case ticketType
                    Case 1
                        nom_type = "Incident"
                    Case 2
                        nom_type = "Demande"
                    Case Else
                        MsgBox "Type non reconnu pour l'ID : " & ticketType, vbExclamation, "Erreur"
                        GoTo Cleanup
                End Select
                If nom_type <> "" Then
                    If updateFields <> "" Then updateFields = updateFields & ", "
                    updateFields = updateFields & "nom_type = '" & Replace(nom_type, "'", "''") & "', type = " & ticketType
                    hasUpdates = True
                End If

            End If
        Else
            MsgBox "La valeur sélectionnée dans la combo box 'Type' n'est pas un ID valide. Veuillez sélectionner 'Incident' ou 'Demande'.", vbExclamation, "Erreur"
            GoTo Cleanup
        End If
    End If
    
   
    If Not IsNull(Me!comboNomCategorie) Then
        If IsNumeric(Nz(Me!comboNomCategorie, "")) Then
            id_categorie = CLng(Me!comboNomCategorie)
            If id_categorie <> current_id_categorie Then
                Set rs = db.OpenRecordset("SELECT completename FROM glpi_itilcategories WHERE id = " & id_categorie)
                If Not rs.EOF Then
                    nom_categorie = Nz(rs!completename, "")
                    If updateFields <> "" Then updateFields = updateFields & ", "
                    updateFields = updateFields & "nom_categorie = '" & Replace(nom_categorie, "'", "''") & "', id_categorie = " & id_categorie
                    hasUpdates = True
                Else
                    MsgBox "Catégorie non trouvée pour l'ID : " & id_categorie, vbExclamation, "Erreur"
                    rs.Close
                    Set rs = Nothing
                    GoTo Cleanup
                End If
                rs.Close
                Set rs = Nothing
            End If
        Else
          
            MsgBox "La valeur sélectionnée dans la combo box 'Catégorie' n'est pas un ID valide.", vbExclamation, "Erreur"
            GoTo Cleanup
        End If
    End If

    If Not hasUpdates Then
        MsgBox "Aucune modification à sauvegarder. Veuillez sélectionner au moins une valeur dans les combo box.", vbInformation, "Information"
        GoTo Cleanup
    End If
    sql = "UPDATE Temp_SingleFiltered_Tickets SET " & updateFields & " WHERE id = " & selectedID
    db.Execute sql, dbFailOnError

    
    Set db = Nothing
    Set db = CurrentDb
       If Not tableExists("DataWarehouse_Tickets") Then
         MsgBox "La table DataWarehouse_Tickets n'existe pas ou n'est plus accessible après réinitialisation.", vbCritical, "Erreur"
        GoTo Cleanup
    End If

    sql = "UPDATE DataWarehouse_Tickets SET " & updateFields & " WHERE id = " & selectedID
    db.Execute sql, dbFailOnError
    
   
    Me!SousFormulaire.Requery
    
    Me!comboNomEntities = Null
    Me!comboNomType = Null
    Me!comboNomCategorie = Null
    MsgBox "Modifications sauvegardées avec succès pour l'ID " & selectedID & " !", vbInformation, "Succès"
    
Cleanup:
      If Not rs Is Nothing Then
        On Error Resume Next 
        rs.Close
        If Err.Number = 0 Then
            Debug.Print "Recordset fermé avec succès."
        Else
            Debug.Print "Erreur lors de la fermeture du recordset : " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de la sauvegarde des modifications : " & Err.Description & vbCrLf & "Requête : " & sql, vbCritical, "Erreur"
    GoTo Cleanup
End Sub

Private Sub btnOuvrirGraphiques_Click()
    On Error GoTo ErrorHandler
    

    DoCmd.Close acForm, "Formulaire", acSaveNo
    

    DoCmd.OpenForm "Graphiques"
    
    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de l'ouverture du formulaire Graphiques : " & Err.Description, vbCritical, "Erreur"
End Sub




Private Sub btnExporterResultatsExcel_Click()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdef As DAO.QueryDef
    Dim sql As String
    Dim filePath As String
    Dim mois As Integer
    Dim annee As Integer
    Dim moisFilter As String


    mois = Val(Nz(Forms!Formulaire!txtMois, 0))
    annee = Val(Nz(Forms!Formulaire!txtAnnee, 0))

    If annee = 0 Then
        MsgBox "Veuillez entrer une année valide.", vbExclamation, "Erreur"
        Exit Sub
    End If


    filePath = Environ("USERPROFILE") & "\Desktop\Resultats_GLPI_" & annee & ".xlsx"

    Set db = CurrentDb


    sql = "TRANSFORM Sum(Resultats) AS SommeDeResultats " & _
          "SELECT Libelle, Entite " & _
          "FROM TableStockageResultats " & _
          "WHERE Annee = " & annee


    If mois > 0 Then
        If mois < 1 Or mois > 12 Then
            MsgBox "Le mois doit être une valeur entre 1 et 12.", vbExclamation, "Erreur"
            Exit Sub
        End If
        sql = sql & " AND Mois <= " & mois
    End If

    sql = sql & " GROUP BY Libelle, Entite " & _
                "PIVOT Choose([Mois], 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre') " & _
                "IN ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')"


    On Error Resume Next
    db.QueryDefs.Delete "Temp_RequeteTableauCroise"
    On Error GoTo ErrorHandler

    Set qdef = db.CreateQueryDef("Temp_RequeteTableauCroise", sql)


    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Temp_RequeteTableauCroise", filePath, True


    db.QueryDefs.Delete "Temp_RequeteTableauCroise"


    Call FormatExcelFile(filePath)

    MsgBox "Les résultats ont été exportés avec succès vers " & filePath, vbInformation, "Succès"
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors de l'exportation des résultats : " & Err.Description, vbCritical, "Erreur"
End Sub

Private Sub FormatExcelFile(filePath As String)
    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object


    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(filePath)
    Set xlWorksheet = xlWorkbook.Sheets(1)

    xlWorksheet.Rows(1).insert
    xlWorksheet.Cells(1, 1).value = "Ticket GLPI " & Forms!Formulaire!txtAnnee
    xlWorksheet.Cells(1, 1).Font.Bold = True
    xlWorksheet.Cells(1, 1).Font.Size = 14


    With xlWorksheet.UsedRange
        .Borders.LineStyle = 1 
        .Borders.Weight = 2 
    End With


    xlWorksheet.Rows(2).Font.Bold = True


    xlWorksheet.Columns.AutoFit


    xlWorkbook.Save
    xlWorkbook.Close
    xlApp.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite lors du formatage du fichier Excel : " & Err.Description, vbCritical, "Erreur"
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
End Sub

Private Sub btnSynchronize_Click()
    Dim db As DAO.Database
    Dim rsDW As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    Set db = CurrentDb
    
    sql = "SELECT id, id_categorie FROM DataWarehouse_Tickets WHERE id_categorie IS NOT NULL"
    Set rsDW = db.OpenRecordset(sql, dbOpenDynaset)
    
    Do While Not rsDW.EOF
        sql = "UPDATE glpi_tickets SET itilcategories_id = " & rsDW!id_categorie & " WHERE id = " & rsDW!ID
        db.Execute sql, dbFailOnError
        rsDW.MoveNext
    Loop
    
    
    sql = "SELECT id, type FROM DataWarehouse_Tickets WHERE type IS NOT NULL"
    Set rsDW = db.OpenRecordset(sql, dbOpenDynaset)
    
    Do While Not rsDW.EOF
        sql = "UPDATE glpi_tickets SET type = " & rsDW!Type & " WHERE id = " & rsDW!ID
        db.Execute sql, dbFailOnError
        rsDW.MoveNext
    Loop
    
    MsgBox "Les catégories ont été synchronisées avec succès vers glpi_tickets.", vbInformation, "Succès"
    
Cleanup:
    rsDW.Close
    Set rsDW = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de la synchronisation : " & Err.Description, vbCritical, "Erreur"
    Resume Cleanup
End Sub

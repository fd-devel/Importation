VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EtapeS 
   Caption         =   "Traitement de fichiers d'�critures"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840.001
   OleObjectBlob   =   "EtapeS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EtapeS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    
    Dim ModeDev As Boolean
    ModeDev = Worksheets("Listes").Range("D1").Value
    
    ' Variables Appli
    Dim LecteurAppli As String
    Dim LecteurOrdi As String
    If ModeDev Then LecteurAppli = "C" Else LecteurAppli = "I"
    If ModeDev Then LecteurOrdi = "C" Else LecteurOrdi = "V"
    
    
    'Valeurs par d�faut
    TerminerBT.Enabled = True
    BT1.Enabled = True
    BT2.Enabled = False
    DosNum = Cells(1, 2).Value
    DosName = Cells(2, 2).Value
    Me.Caption = DosNum & " - " & DosName & " | Traitement de fichier d'�critures"
    NomClasseurExcel = ActiveWorkbook.Name
    
    AppliPath = LecteurAppli & ":\Importations\XLX\"
    DepotPath = LecteurAppli & ":\DEPOT_Importations\"
    TMPPath = AppliPath & "TMP\"
    FichierAImporterDansCegid = DepotPath & DosNum & " - import.csv" 'Nom du fichier qui sera import� par Cegid
    FichierIntermediaire = TMPPath & DosNum & "\FichierATraiter.txt" 'Nom du fichier qui sera trait� sur le serveur
    FichierIntermediaire2 = TMPPath & DosNum & "\FichierATraiter2.txt" 'Nom du fichier qui sera trait� sur le serveur quand le fichier interm�diaire a besoin d'un traitement pr�alable (Isagri par exemple)
    TypeImport = Sheets("Dossier").Cells(4, 2)
    ModelesPath = AppliPath & "MODELES\" 'R�pertoire o� sont stock�s les mod�les de page import
    
    FichierSourcePC = ""
    PathPC = LecteurOrdi & ":\IMPORT" 'R�pertoire sur le PC du comptable o� sont stock�s les fichiers re�us des clients
    PathBackup = PathPC & "\Transf�r�s" 'R�pertoire backup sur le PC du comptable
    DrivePC = LecteurOrdi & ":\" 'Lecteur sur poste du comptable
        
        
    '######## V�rifications ########
    Etape = "1 / 4"
    Titre = "V�rifications"
    Titre.ForeColor = vbBlack
    Infos = "Patientez un instant"
    
    'V�rifier que les param�tres soient renseign�s
    If Len(DosNum) = 0 Or Len(DosName) = 0 Or Len(TypeImport) = 0 Then GoTo errorParametres
    
    'V�rifier qu'un fichier ne soit pas en attente d'importation par Cegid
    If Len(Dir(FichierAImporterDansCegid)) <> 0 Then GoTo errorFichierPresent

    'V�rifier que l'on peut �crire sur le disque du comptable
    On Error GoTo errorAccesEcritureDisquePC
    ChDrive DrivePC
 
    testfic = FreeFile
    Open PathPC & "\test.test" For Output As testfic
    Close testfic
    
    
    'Cr�ation des r�pertoires sur le poste du comptable si ils n'existent pas
    On Error GoTo errorCreationRepertoires
    If Len(Dir(PathPC, vbDirectory)) = 0 Then
        MkDir PathPC
    End If
    If Len(Dir(PathBackup, vbDirectory)) = 0 Then
        MkDir PathBackup
    End If
    
    'Cr�ation du r�pertoire temp si il n'existe pas
    On Error GoTo errorCreationRepertoireServeur
    If Len(Dir(TMPPath & DosNum, vbDirectory)) = 0 Then
        MkDir TMPPath & DosNum
    End If
    
    
    'Si tout va bien...
    Titre = "Traitement de fichier"
    Infos = Chr(10) & "Importer : Choix du fichier � importer et lancement du traitement" _
            & Chr(10) & Chr(10) & "Param�tres : Acc�s aux tables de correspondance et aux param�tres de l'outil" _
            & Chr(10) & Chr(10) & "Quitter : Quitter l'outil d'import"
    BT2.Enabled = True
    GoTo theEnd
    
    '######## Messages d'erreur ########

errorParametres:
    Titre = "Param�tres manquants"
    Titre.ForeColor = vbRed
    Titre.Font.Size = 18
    Infos = "Certains param�tres d'importation doivent �tre renseign�s." & Chr(10) & Chr(10) _
        & "Param�tres obligatoires:" & Chr(10) & Chr(10) _
        & "- N� de Dossier Cegid" & Chr(10) _
        & "- Nom du client" & Chr(10) & Chr(10) _
        & " Cliquez sur le bouton ''Param�tres'' et renseignez les cases."
    Etape = "Anomalie #V1"
    Me.Repaint
    csvLink.Visible = True
    GoTo theEnd

errorFichierPresent:
    Titre = "Fichier d�j� pr�sent"
    Titre.ForeColor = vbRed
    Titre.Font.Size = 18
    Infos = "Un fichier est d�j� en attente d'int�gration dans Cegid." & Chr(10) & Chr(10) _
        & "Pour rappel, vous devez :" & Chr(10) & Chr(10) _
        & "1 - Ouvrir votre dossier comptable Cegid" & Chr(10) _
        & "2 - Module 'Traitements annexes'" & Chr(10) _
        & "3 - Section 'R�cup�ration de donn�es'" & Chr(10) _
        & "4 - Lancer 'Grand-livres - Journaux'"
    Etape = "Anomalie #V2"
    Me.Repaint
    csvLink.Visible = True
    GoTo theEnd
        
errorAccesEcritureDisquePC:
    Titre = "Droits d'acc�s en �criture au disque dur du poste"
    Titre.ForeColor = vbRed
    Titre.Font.Size = 18
    Infos = "Vous n'avez pas acc�s en �criture au poste sur lequel vous travaillez." _
            & Chr(10) & "Merci de quitter l'environnement Cegid et de le relancer en veillant � donner les droits complets."
    Etape = "Anomalie #V3"
    Me.Repaint
    GoTo theEnd

errorCreationRepertoires:
    Titre = "Erreur lors de la cr�ation des r�pertoires sur le poste"
    Titre.ForeColor = vbRed
    Titre.Font.Size = 18
    Infos = "Il est impossible de cr�er des r�pertoires sur votre poste." & Chr(10) & "Merci de quitter l'environnement Cegid et de le relancer en veillant � donner les droits complets."
    Etape = "Anomalie #V4"
    Me.Repaint
    GoTo theEnd

errorCreationRepertoireServeur:
    Titre = "Erreur lors de la cr�ation du r�pertoire temporaire"
    Titre.ForeColor = vbRed
    Titre.Font.Size = 18
    Infos = "R�pertoire temporaire � cr�er : " & TMPPath & DosNum
    Etape = "Anomalie #V5"
    Me.Repaint
    GoTo theEnd


'On active le bouton d'importation si tous les tests ont �t� pass�s
BT2.Enabled = True

'Sortie pr�ventive en cas de test �chou�
theEnd:

End Sub

Private Sub UserForm_Terminate()
    
    'On r�affiche la fen�tre Excel si l'UserForm se ferme
    Application.Visible = True

End Sub

Private Sub BT2_Click()

    
    '=============== Etape 2 =============== Copie du fichier sur le serveur
    
    'Copier le fichier sur le serveur
    Titre = "R�cup�ration du fichier du client"
    Titre.ForeColor = vbBlack
    Infos = "Transfert du fichier depuis le poste du comptable sur le serveur Cegid"
    Etape = "Etape 2 / 4"
    Me.Repaint
    BT2.Enabled = False


    ' ######
    ' ###### Transfert du fichier sur le serveur ########
    On Error GoTo errorTransfertFichiers
    'Choisir le fichier � transf�rer sur le serveur
    ChDir PathPC
    FichierSourcePC = Application.GetOpenFilename
    'V�rifier qu'un fichier a �t� choisi
    If FichierSourcePC = False Then GoTo errorChoixFichier
    'Copie du fichier sur le serveur
    FileCopy FichierSourcePC, FichierIntermediaire
    
    'D�place et renomme le fichier dans le r�pertoire backup
    On Error GoTo errorBackup
    Name FichierSourcePC As PathBackup & "\Dossier " & DosNum & " - Date " & Format(Date, "yyyy.mm.dd") & " Heure " & Format(Time, "hh.mm.ss") & ".bak"

    On Error GoTo 0
    '=============== Etape 3 =============== Traitement des fichiers
    Titre = "Traitement du fichier avant Import"
    Titre.ForeColor = vbBlack
    Infos = "Application des tables de correspondance" & Chr(10) & "Journaux et comptes"
    Etape = "Etape 3 / 4"
    Me.Repaint


    ' ######
    ' ###### Supprimer les feuilles d'import / export si elles existent
    If IsError(Evaluate("='Import'!A1")) <> True Then
        Application.DisplayAlerts = False
        Sheets("Import").Delete
        Application.DisplayAlerts = True
    End If
    If IsError(Evaluate("='ExportCegid'!A1")) <> True Then
        Application.DisplayAlerts = False
        Sheets("ExportCegid").Delete
        Application.DisplayAlerts = True
    End If
    If IsError(Evaluate("='Export'!A1")) <> True Then
        Application.DisplayAlerts = False
        Sheets("Export").Delete
        Application.DisplayAlerts = True
    End If
    
    
    ' ######
    ' ###### Cr�ation des feuilles d'import
    On Error GoTo errorMoulinette
    
    Dim ShImport As Worksheet
    Dim ShExport As Worksheet
    Set ShImport = Sheets.Add(After:=Sheets(Sheets.Count))
    ShImport.Name = "Import"
    Set ShExport = Sheets.Add(After:=Sheets(Sheets.Count))
    ShExport.Name = "Export"
    
    Dim StringCol, Champ As String      ' nom de la colonne
    Dim Col_X As Integer                ' pour incrementation des cellules Col1, Col2, ... de la page Param
    Dim Col_Sheet_Import As Integer     ' pour incr�mentation des colonnes de la page Import (seulement celles r�cup�r�es)
    Dim Col_File_Import As Integer      ' pour compter le Nbre de colonne du fichier a importer ( pour colonne a taille fixe)
    Dim TitresImports As Variant        ' Liste des Titres de colonne d'Import - sur page import
    Dim TitresExports As Variant        ' Liste des Titres de colonne d'Export - sur page import
    Dim TitresChamps(30) As String      ' Pour lister les tires des champs utilis�s
    Dim vZone() As Integer                ' Array pour : TextFileColumnDataTypes
    Dim indexZone As Integer            ' la dimension du array pour TextFileColumnDataTypes
    Dim Montant As String, Sens As String
    Dim Type_import As Integer          ' Importation ecritures ou balance
    
    
    ' ######
    ' ###### Initialise la valeur de TextFileColumnDataTypes
    ' pour toutes les colonnes � " Ignorer " = 9
    ReDim vZone(30)
    For i = 0 To 30
        vZone(i) = 9
    Next i

    indexZone = 0
    
    ' ######
    ' ###### Quelle importation? Ecr ou BAL
    Type_import = Worksheets("Listes").Range("K1").Value
    
    
    ' ######
    ' ###### Initialisation titres de colonnes sur feuille import
    Col_X = 1
    Col_Sheet_Import = 1
    
    TitresImports = Array("", "I_JRNL", "I_DATE", "I_DATE_FACT", "I_CPT", "I_AUX", "I_LBL", "I_TIERS", "I_NATURE", "I_REFER", "I_DBT", "I_CRDT", "I_MONT", "I_SENS", "I_LETT", "I_QUT1", "", "", "", "", "")
    TitresExports = Array("", "E_JRNL", "E_DATE", "E_DATE_FACT", "E_CPT", "E_AUX", "E_LBL", "E_TIERS", "E_NATURE", "E_REFER", "E_DBT", "E_CRDT", "E_MONT", "E_SENS", "E_LETT", "E_QUT1", "", "", "", "", "")
    
    
    
        
    ' ######
    ' ###### Format des dates
    
    ' Format des charact�res de colonnes du fichier importer
    
    '   1 = STANDARD
    '   2 = TEXTE
    '   3 = MJA     - date
    '   4 = JMA     - date
    '   5 = AMJ     - date
    '   6 = MAJ     - date
    '   7 = JAM     - date
    '   8 = AJM     - date
    '   9 = IGNORER

    If Type_import = 1 Then         '####  Ne concerne que l'import d'�critures
        ' pour convertir la colonne,
        ' a r�cup�rer sur page Param pour valeur
        ' et sur page Listes (Masqu�e!) pour Valeurs possibles
        Dim FDate As Integer
        Dim WhatIsTheDate As String
        WhatIsTheDate = Worksheets("Param").Range("I3").Value
        Select Case WhatIsTheDate
            Case Worksheets("Listes").Range("H5").Value
                FDate = 3
            Case Worksheets("Listes").Range("H6").Value
                FDate = 4
            Case Worksheets("Listes").Range("H7").Value
                FDate = 5
            Case Worksheets("Listes").Range("H8").Value
                FDate = 6
            Case Worksheets("Listes").Range("H9").Value
                FDate = 7
            Case Worksheets("Listes").Range("H10").Value
                FDate = 8
            Case Else
                FDate = 4
        End Select
    End If
    
    
    
    ' ######
    ' ###### Quelle est la Premi�re ligne a importer
    Dim FirstLine As Variant
    FirstLine = Worksheets("Param").Range("I1").Value
    
    
    
    
    ' ######
    ' ###### Creation ent�tes des colonnes "Zone Import"
    Sheets("Param").Select
    
    ' Sur la page Param on passe une par une les cellules A6, B6...
    ' Tant que la cellule commence par 'Col'
    ' Comme ca on peut en ajouter si besoin et pas la peine de changer cette partie de code
    ' ... Tant que col ...
    ' la cellule de dessous contient le param�tre du champ
    ' On r�cup�re ce param�tre (dans StringCol) (s'il n'est pas a Ignorer) et on note sur la page import a partir de A1
    ' Le premier "n'existe pas" stope la boucle
    
    StringCol = Cells(6, Col_X).Value
    Do While Left(StringCol, 3) = "Col"
    
'        On Error GoTo Suivant
        ' On v�rifie si la colonne existe dans le param�trage
        ' et si on doit l'ignorer ou en tenir compte
        ' note :
        '   IGNORER = existe mais on en tient pas compte
        '   N'EXISTE pas = La colonne n'existe pas
        ' Attention a la derni�re colonne, apr�s le dernier d�limiteur!
        
        
        If Cells(7, Col_X).Value <> "IGNORER" Then Champ = Cells(7, Col_X).Value Else GoTo Suivant
        
        Sheets("Import").Select
        Select Case Champ
        
            Case "Journal"
                Cells(1, Col_Sheet_Import).Value = TitresImports(1) ' On ecrit le nom de la colonne
                TitresChamps(Col_Sheet_Import) = TitresExports(1)   ' On r�serve le nom de la colonne - pour zone export
                vZone(Col_X - 1) = 1 ' type de donn�e du fichier pour TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante


            Case "Date"
                Cells(1, Col_Sheet_Import).Value = TitresImports(2)
                TitresChamps(Col_Sheet_Import) = TitresExports(2)
                vZone(Col_X - 1) = FDate ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
                
                
            Case "Date Fact"
                Cells(1, Col_Sheet_Import).Value = TitresImports(3)
                TitresChamps(Col_Sheet_Import) = TitresExports(3)
                vZone(Col_X - 1) = FDate ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
 
                
            Case "Compte"
                Cells(1, Col_Sheet_Import).Value = TitresImports(4)
                TitresChamps(Col_Sheet_Import) = TitresExports(4)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
                
            
            Case "Cpte Auxilaire"
                Cells(1, Col_Sheet_Import).Value = TitresImports(5)
                TitresChamps(Col_Sheet_Import) = TitresExports(5)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante

            
            Case "Libell�"
                Cells(1, Col_Sheet_Import).Value = TitresImports(6)
                TitresChamps(Col_Sheet_Import) = TitresExports(6)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante

            
            Case "Tiers"
                Cells(1, Col_Sheet_Import).Value = TitresImports(7)
                TitresChamps(Col_Sheet_Import) = TitresExports(7)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Nature"
                Cells(1, Col_Sheet_Import).Value = TitresImports(8)
                TitresChamps(Col_Sheet_Import) = TitresExports(8)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Pi�ce"
                Cells(1, Col_Sheet_Import).Value = TitresImports(9)
                TitresChamps(Col_Sheet_Import) = TitresExports(9)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "D�bit"
                Cells(1, Col_Sheet_Import).Value = TitresImports(10)
                TitresChamps(Col_Sheet_Import) = TitresExports(10)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Cr�dit"
                Cells(1, Col_Sheet_Import).Value = TitresImports(11)
                TitresChamps(Col_Sheet_Import) = TitresExports(11)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Montant"
                Cells(1, Col_Sheet_Import).Value = TitresImports(12)
                TitresChamps(Col_Sheet_Import) = TitresExports(12)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                'On stocke la lettre de la colonne pour plus tard
                Montant = Split(Columns(Col_Sheet_Import).Address(ColumnAbsolute:=False), ":")(1)
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Sens"
                Cells(1, Col_Sheet_Import).Value = TitresImports(13)
                TitresChamps(Col_Sheet_Import) = TitresExports(13)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                ' On stocke la lettre de la colonne pour plus tard
                Sens = Split(Columns(Col_Sheet_Import).Address(ColumnAbsolute:=False), ":")(1)
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Lettrage"
                Cells(1, Col_Sheet_Import).Value = TitresImports(14)
                TitresChamps(Col_Sheet_Import) = TitresExports(14)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
            
            
            Case "Quantit�1"
                Cells(1, Col_Sheet_Import).Value = TitresImports(15)
                TitresChamps(Col_Sheet_Import) = TitresExports(15)
                vZone(Col_X - 1) = 1 ' TextFileColumnDataTypes
                
                Col_Sheet_Import = Col_Sheet_Import + 1 ' Colonne page import suivante
                
            Case "N'EXISTE PAS"
                Exit Do
            Case Else
                Exit Do
            
        End Select
    
        ' fin de boucle on incr�mente l'indice de la colonne
Suivant:
        Col_File_Import = Col_File_Import + 1
        Sheets("Param").Select      ' on revient sur la page Param
        Col_X = Col_X + 1           ' on passe a la colonne 'Col X' suivante
        StringCol = Cells(6, Col_X)
        indexZone = indexZone + 1   ' compte le nombre de colonnes du fichier ( a traiter + a ignorer. Sans n'existe pas)

    Loop
    
    
    
    
    
    ' ######
    ' ###### Creation des ent�tes de colonnes "Zone Export" - page import
    Sheets("Import").Select
    Col_Sheet_Import = Col_Sheet_Import ' -1 = DernierCol_Nun?
    Dim DernierCol_Nun As Integer
    DernierCol_Nun = Range("A1").End(xlToRight).Column
    

    Dim k As Integer: k = 1     ' pour incr�menter le nombre de colonnes de la zone export
                                ' Diff�rent de j qui incr�mente le colonnes zone import

                                Dim Aux_GO As Boolean       ' pour savoir si on cr�e la colonne cpte auxiliaire
    Aux_GO = True       ' On cr�e? Oui?    - ! - On cr�e ici parce qu'on s'en re-sert pour les formules ligne 2
    
    For j = 1 To Col_Sheet_Import
        ' On inscrit le nom de la colonne
        Cells(1, k + DernierCol_Nun).Value = TitresChamps(j)
        
        ' Quand on passe sur "CPT" on teste l'existance de "AUX"
        ' Si existe : rien a faire
        ' Sinon cr�er la colonne E_AUX
        
        If TitresChamps(j) = "E_CPT" Then
            For i = 0 To UBound(TitresChamps)
                If TitresChamps(i) = "E_AUX" Then
                    Aux_GO = False  ' Eh non! pas besoin la colonne existe.
                    Exit For
                End If
            Next i
            
            If Aux_GO Then
                k = k + 1
                Cells(1, k + DernierCol_Nun).Value = "E_AUX"
            End If
        End If
        
        k = k + 1
    Next j
    
    
    
    ' ######
    ' ###### Cr�ation des formules de conversion : ligne 2, zone export
        
    Dim Lettre As String        ' Lettre de la colone
    k = 1                  ' on s'en re-sert
    

    
    For j = 1 To DernierCol_Nun
        ' Lettre de la colonne
        Lettre = Split(Columns(j).Address(ColumnAbsolute:=False), ":")(1)
        
        Select Case TitresChamps(j)
            Case TitresExports(1)                       ' E_JRNL
                Macro = "=CorresJRNL(" & Lettre & "2)"
                Cells(2, k + DernierCol_Nun).Select
                Selection.FormulaLocal = Macro
                
            Case TitresExports(2)                       ' E_DATE
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                ' On en profite pour convertir la colonne au format Date fran�ais
                col_Dt = Split(Columns(k + DernierCol_Nun).Address(ColumnAbsolute:=False), ":")(1)
                Range(col_Dt & ":" & col_Dt).Select
                Selection.NumberFormat = "m/d/yyyy" ' WhatIsTheDate
                
            Case TitresExports(3)                       'E_DATE_FACT
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                ' On en profite pour convertir la colonne au format Date
                col_Dt = Split(Columns(k + DernierCol_Nun).Address(ColumnAbsolute:=False), ":")(1)
                Range(col_Dt & ":" & col_Dt).Select
                Selection.NumberFormat = "m/d/yyyy" ' WhatIsTheDate
                
            Case TitresExports(4)                       'E_CPT
                Macro = "=CorresCPT(" & Lettre & "2;""CPT"")"
                Cells(2, k + DernierCol_Nun).Select

                Selection.FormulaLocal = Macro
                
                If Aux_GO Then
                    k = k + 1   ' on incremente la colonne
                    Macro = "=CorresCPT(" & Lettre & "2;""AUX"")"
                    Cells(2, k + DernierCol_Nun).Select
                    Selection.FormulaLocal = Macro
                    
                End If
                
            Case TitresExports(5)                       'E_AUX
                Macro = "=CorresAUX(" & Lettre & "2)"
                Cells(2, k + DernierCol_Nun).Select
                Selection.FormulaLocal = Macro
                
            Case TitresExports(6)                       'E_LBL
                Cells(2, k + DernierCol_Nun).Value = "=Replac_Spe(" & Lettre & "2)"
                
            Case TitresExports(7)                       'E_TIERS
                Cells(2, k + DernierCol_Nun).Value = "=Replac_Spe(" & Lettre & "2)"
                
            Case TitresExports(8)                       'E_NATURE
                Cells(2, k + DernierCol_Nun).Value = "=Replac_Spe(" & Lettre & "2)"
                
            Case TitresExports(9)                       'E_REFER
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(10)                      'E_DBT
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(11)                      'E_CRDT
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(12)                      'E_MONT --> E_DBT
                Macro = "=si(" & Sens & "2 = ""D"" ; " & Montant & "2 ; """" )"
                Cells(2, k + DernierCol_Nun).Select
                Selection.FormulaLocal = Macro
                
            Case TitresExports(13)                      'E_SENS --> E_CRDT
                Macro = "=si(" & Sens & "2 = ""C"" ; " & Montant & "2 ; """" )"
                Cells(2, k + DernierCol_Nun).Select
                Selection.FormulaLocal = Macro
                
            Case TitresExports(14)                      'E_LETT
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(15)                      'E_QUT1
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(16)  '
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(17)  '
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
            Case TitresExports(18)  '
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
                
            Case Else  '
                Cells(2, k + DernierCol_Nun).Value = "=" & Lettre & "2"
                
        End Select
        k = k + 1
    Next j
    
    
        '# Range a exporter ?      ######  ENCORE UTILE????   #####
    Dim Range_Db, Range_Fin As Variant
    Dim Aux_Oui As Integer
    If Aux_GO Then Aux_Oui = 1 Else Aux_Oui = 0
    
    Range_Db = Split(Columns(DernierCol_Nun + 1).Address(ColumnAbsolute:=False), ":")(1)
    Range_Fin = Split(Columns(DernierCol_Nun + DernierCol_Nun + Aux_Oui).Address(ColumnAbsolute:=False), ":")(1)
    
    ExportRange = Range_Db & ":" & Range_Fin    'Colonnes � copier
    Cells(1, DernierCol_Nun + DernierCol_Nun + Aux_Oui + 1).Value = "ExportRange = " & ExportRange   ' pour m�moire
    
    
    
    
    ' ######
    ' ######    SPECIAL ISAGRI
    Dim IndexFichierIntermediaire As Integer
    IndexFichierIntermediaire = 1 ' Fichier normal <> isagri
    
    If Worksheets("Dossier").Range("B4").Value = "ISAGRI" And Type_import = 1 Then    ' isa et �critures (<> balance)
        IndexFichierIntermediaire = 2
        
    ' D�clarations
        Dim isaFic, cegFic As Integer
        Dim isaLigne, isaECR As String
          
        isaFic = FreeFile
        Open FichierIntermediaire For Input As isaFic
            
        cegFic = FreeFile
        Open FichierIntermediaire2 For Output As cegFic
            
        While Not EOF(isaFic)
            Line Input #isaFic, isaLigne
            If Left(isaLigne, 3) = "ECR" Then isaECR = Left(isaLigne, 246)
            If Left(isaLigne, 3) = "MVT" And Mid(isaECR, 120, 4) = "0EUR" Then Print #cegFic, isaECR & isaLigne
        Wend
        Close isaFic
        Close cegFic
    
    End If
    

    ' ######
    ' ###### D�limiteur
    Dim Process As Variant  ' Suivant le d�limiteur initialisation de QuertTable est diff�rent
    Dim Delimiteur As Variant
    Dim ColumnWidths() As Variant
        
    Select Case Worksheets("Param").Range("D3").Value
        Case Worksheets("Listes").Range("E5").Value      ' point virgule
            Process = 1
            Delimiteur = ";"
        Case Worksheets("Listes").Range("E6").Value      ' virgule
            Process = 1
            Delimiteur = ","
        Case Worksheets("Listes").Range("E7").Value      ' Demi colonne
            Process = 1
            Delimiteur = "|"
        Case Worksheets("Listes").Range("E8").Value      ' Di�se
            Process = 1
            Delimiteur = "#"
        Case Worksheets("Listes").Range("E9").Value      ' Tabulation
            Process = 2
            
        Case Worksheets("Listes").Range("E10").Value      ' Champs fixes
            Process = 3
            ' On r�cup�re le nbre de caract�res de chaques colonnes
            For j = 1 To Col_File_Import
                ReDim Preserve ColumnWidths(j - 1)
                ColumnWidths(j - 1) = Worksheets("Param").Cells(5, j).Value
            Next j
        
    End Select


    ' ######
    ' ######    TextFileColumnDataTypes
    ReDim Preserve vZone(indexZone)
    
    

    ' ######
    ' ###### Montage de l'objet QueryTables
    Range("A2").Select
    Set sht = ActiveSheet
    
    If IndexFichierIntermediaire = 2 Then     ' Pour isagri
    
        Set QrResults = sht.QueryTables.Add( _
                Connection:="TEXT;" & FichierIntermediaire2, _
                Destination:=Range("$A$2"))
    Else
     
        Set QrResults = sht.QueryTables.Add( _
                Connection:="TEXT;" & FichierIntermediaire, _
                Destination:=Range("$A$2"))
    End If

                
    With QrResults
        .Name = "FichierATraiter_1"
        .FieldNames = False
        .RowNumbers = False
        .FillAdjacentFormulas = True
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = FirstLine
        .TextFileTrailingMinusNumbers = True
        .TextFileDecimalSeparator = "."
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileColumnDataTypes = vZone
    End With
                        
    Select Case Process
        Case 1
            With QrResults
                .TextFileParseType = xlDelimited
                .TextFileConsecutiveDelimiter = False
                .TextFileOtherDelimiter = Delimiteur
            End With
        Case 2
            With QrResults
                .TextFileParseType = xlDelimited
                .TextFileTabDelimiter = True
            End With
        Case 3
            With QrResults
                .TextFileParseType = xlFixedWidth
                .TextFileFixedColumnWidths = ColumnWidths
            End With
    End Select
            
            With QrResults
                .Refresh BackgroundQuery:=False
            End With
    
    GoTo Lasuite
    
    
    ' ######
    ' ###### Traitement des donn�es externes
      
    
Lasuite:
    'Actualiser l'import ---- N�cessaire � cause d'un Bug ? Apparemment quand on cr�e l'import il ne duplique pas les formules � cot� de l'import... donc bon...
    Worksheets("Import").UsedRange.Rows("2:2").Calculate 'n�cessaire pour la premi�re ligne
    Sheets("Import").QueryTables(1).Refresh
    
    
    
    
    ' ######
    ' ######    FEUILLE EXPORT
    
    '## On remet les colonnes dans le m�me ordre quelque soit le fichier import�
    '##
    '## Pour qu'elles correspondent au Script
    '## d'import dans C�gid => 1 seul script
    '##
    '## Script C�gid :  _IMPORT_CSV    pour �critures
    '##
    '## Script C�gid :  _IMPORT_BAL_CSV    pour BALANCE
    
    Dim Script_import As String
    If Type_import = 1 Then
        Script_import = "_IMPORT_CSV"
    Else
        Script_import = "_IMPORT_BAL_CSV"
    End If
    
    
    
    ' Ordre choisi : (subit...)
        
    ' Journal                        E_JNL
    ' Date                           E_DATE
    ' Date facture                   E_DATE_FACT
    ' Num�ro de pi�ce                E_REFER
    ' Quantit�                       E_QUT1
    ' Lettrage                       E_LETT
    ' Compte gen�ral                 E_CPT
    ' Compte Auxiliaire              E_AUX
    ' Code Analytique
    ' Libell�
    ' Tiers                          E_TIERS
    ' Nature                         E_NATURE
    ' D�bit                          E_DBT
    ' Cr�dit                         E_CRDT
    
    
    Sheets("Export").Select
    Dim ColumnsSheetExport As Variant
    If Type_import = 1 Then
        ColumnsSheetExport = Array("E_JRNL", "E_DATE", "E_DATE_FACT", "E_REFER", "E_QUT1", "E_LETT", "E_CPT", "E_AUX", "E_TIERS", "E_NATURE", "E_DBT", "E_CRDT")
    Else
        ColumnsSheetExport = Array("E_CPT", "E_AUX", "E_LBL", "E_QUT1", "E_DBT", "E_CRDT")
    End If
    Dim ColFind As Boolean
    ColFind = False
   
    j = 1   ' pour chaque colonne de la feuille Export
    For Each ColField In ColumnsSheetExport
        i = 1   ' on cherche la colonne dans la feuille Import
        Sheets("Import").Select
        
        Do While Cells(1, i).Value <> ""
            If Cells(1, i).Value = "E_MONT" Then Cells(1, i).Value = "E_DBT"
            If Cells(1, i).Value = "E_SENS" Then Cells(1, i).Value = "E_CRDT"
            
            ' Pour les balances : si Tiers est choisi au lieu de Lbl
            If Type_import = 2 Then
                If Cells(1, i).Value = "E_TIERS" Then Cells(1, i).Value = "E_LBL"
            End If
            
            ' quand on trouve on copie
            If ColField = Cells(1, i).Value Then
                Columns(i).Select
                Selection.Copy
                Sheets("Export").Select
                Columns(j).Select
                Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Selection.Columns.AutoFit
                
                ColFind = True
                j = j + 1
                Sheets("Import").Select
            End If
            i = i + 1
        Loop
        
        
        ' Si la colone n'a pas �t� trouv�e dans import on la cr�e quand m�me dans Export.
        ' C'est pour respecter le nombre de colones dans le script C�gid
        If ColFind = False Then
            Sheets("Export").Select
            Cells(1, j).Value = ColField
            j = j + 1
        Else
            ColFind = False
            
        End If

    Next
       
    
    
    ' ######
    ' ###### G�n�ration du CSV � destination de Cegid
    
    Sheets("Export").Select
    If Type_import = 1 Then
        Columns("A:M").Select
    Else
        Columns("A:F").Select
    End If
    Selection.Copy
    Workbooks.Add (1)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveWindow.DisplayZeros = False
    ActiveWorkbook.SaveAs Filename:= _
        FichierAImporterDansCegid, _
        FileFormat:=xlCSV, CreateBackup:=False, local:=True
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    

    ' ######
    ' ###### On fait le m�nage apr�s
    Sheets("Dossier").Select
    Application.DisplayAlerts = False
    Sheets("Import").Delete
    Sheets("Export").Delete
    Application.DisplayAlerts = True
    
    'On met � jour la date de dernier import
    Worksheets("Dossier").Range("B11").Value = Date
    Worksheets("Dossier").Range("B12").Value = Time
    
    'On sauvegarde le fichier
    ActiveWorkbook.Save

    '=============== Etape 4 =============== Fini !!!!
    
    Titre = "Traitement termin�"
    Titre.ForeColor = RGB(0, 160, 64)
    Titre.Font.Size = 18
    Infos = "Le fichier � importer dans Cegid a �t� cr��. " & Chr(10) & Chr(10) _
            & "Pour rappel, vous devez maintenant :" & Chr(10) _
            & "1 - Ouvrir votre dossier comptable Cegid" & Chr(10) _
            & "2 - Module 'Traitements annexes'" & Chr(10) _
            & "3 - Section 'R�cup�ration de donn�es'" & Chr(10) _
            & "4 - Lancer 'Grand-livres - Journaux'" & Chr(10) _
            & "5 - Script d'import:'" & Script_import
    Etape = "Etape 4 / 4"
    csvLink.Visible = True
    Me.Repaint
    
'Si tout va bien
GoTo theEnd


'######## Messages d'erreur ########
 
errorSuite:
    Titre = "Test Echou�"
    Titre.ForeColor = vbRed
    Infos = Chr(10) & "Aller essaye encore."
    Etape = "Anomalie Bizarre"
    BT2.Enabled = True
    Me.Repaint
    GoTo theEnd
    
errorChoixFichier:
    Titre = "Aucun fichier choisi"
    Titre.ForeColor = vbRed
    Infos = Chr(10) & "Merci de cliquer � nouveau sur importer puis de choisir le fichier � importer sur votre poste."
    Etape = "Anomalie #C1"
    BT2.Enabled = True
    Me.Repaint
    GoTo theEnd

errorTransfertFichiers:
    Titre = "Erreur lors du transfert de fichier"
    Titre.ForeColor = vbRed
    Infos = ""
    Etape = "Anomalie #C2"
    Me.Repaint
    GoTo theEnd

errorBackup:
    Titre = "Erreur lors de la sauvegarde du fichier transf�r�"
    Titre.ForeColor = vbRed
    Infos = ""
    Etape = "Anomalie #C3"
    Me.Repaint
    GoTo theEnd

errorModele:
    Titre = "Mod�le d'import inconnu"
    Titre.ForeColor = vbRed
    Infos = "V�rifiez le param�trage sur la feuille 'Param'. "
    Etape = "Anomalie #T1"
    Me.Repaint
    GoTo theEnd
    
errorMoulinette:
    Titre = "Erreur! Un param�tre n'est pas g�r�! "
    Titre.ForeColor = vbRed
    Infos = "V�rifiez le param�trage sur la feuille 'Param'. "
    Etape = "Anomalie #F1"
    Me.Repaint
    GoTo theEnd


theEnd:

End Sub

Private Sub BT1_Click()

    'On r�affiche la fen�tre Excel avant de fermer l'UserForm
    Application.Visible = True
    Unload Me

End Sub


Private Sub TerminerBT_Click()

    'On quitte le document et Excel si c'�tait le seul document ouvert
    Unload Me
    
    If Workbooks.Count = 1 Then
        SendKeys "%{F4}" 'Envoi des touches Alt-F4 si un seul document ouvert
    Else
        ActiveWorkbook.Close
    End If

End Sub

Private Sub csvLink_MouseDown(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 1 Then 'Ouverture du .csv sur clic gauche
        Workbooks.Open Filename:=FichierAImporterDansCegid, local:=True
    End If
    
    If Button = 2 Then 'Ouverture du r�pertoite contenant le .csv sur clic droit
        ThisWorkbook.FollowHyperlink DepotPath
    End If

End Sub











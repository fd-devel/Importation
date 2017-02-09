Attribute VB_Name = "Module1"
Sub Launcher()

    MiseAJour
    Etape1

End Sub


Sub Etape1() 'Etape de préparation / vérification

    'On affiche l'UserForm
    EtapeS.Show 0
    
    'On tri les tables de correspondance pour éviter les soucis de recherches
    Sheets("Journaux").Select
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Journaux").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Journaux").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Journaux").Sort
        .SetRange Range("A1:B2")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
    Sheets("Comptes").Select
    Columns("A:C").Select
    ActiveWorkbook.Worksheets("Comptes").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Comptes").Sort.SortFields.Add Key:=Range("A2:A13") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Comptes").Sort
        .SetRange Range("A1:C13")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
    Sheets("Dossier").Activate
    Range("B1").Select
    
    'On évite de demander à sauvegarder si on a touché à rien avant de quitter
    ActiveWorkbook.Save

End Sub


Sub MiseAJour() 'Mise à jour du fichier XLS client (libellés, listes de choix, etc)
    
    Sheets("Dossier").Activate
    Range("B4").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ISAGRI, POMO, Cote Ouest, CFC Caisse, CFC Fact, AUTRE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Choisir le modèle d'écritures"
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    Range("B1").Select
    Cells(1, 1).Value = "N° Dossier Cegid"
    Cells(2, 1).Value = "Nom du client"
    Cells(3, 1).Value = "Comptable"
    Cells(4, 1).Value = "Logiciel / Ecritures du client"
    
    'On évite de demander à sauvegarder si on a touché à rien avant de quitter
    ActiveWorkbook.Saved = True

End Sub

Sub LancerProcedure() 'Clic sur le bouton

    ActiveWorkbook.Save 'On sauvegarde car on est rentré dans les paramètres
    Etape1

End Sub

Function CorresJRNL(Journal As Variant) As Variant 'Gestion des correspondances de codes journaux

    On Error GoTo ErreurJRNL:

    CorresJRNL = WorksheetFunction.VLookup(Journal, Range("JNX"), 2, False)
    
    Exit Function
    
ErreurJRNL:
    CorresJRNL = Journal

End Function

Function CorresCPT(ClientCompte As Variant, Info As String) As Variant 'Gestion des correspondances de codes journaux

    On Error GoTo ErreurCPT:
    Dim CegidCPT, CegidAUX As Variant
    Dim Regroupement As Boolean
    Regroupement = Worksheets("Comptes").Range("E1").Value
    
    If Len(ClientCompte) >= 8 Then ClientCompte = Left(ClientCompte, 7)
    
    Select Case Info
        Case "CPT"
            CorresCPT = WorksheetFunction.VLookup(ClientCompte, Range("CPTS"), 2, False)
        Case "AUX"
            CorresCPT = WorksheetFunction.VLookup(ClientCompte, Range("CPTS"), 3, False)
            
            
    End Select
    
    If IsNull(CorresCPT) Or Len(CorresCPT) = 0 Then CorresCPT = ""
    
    Exit Function
    
ErreurCPT:
    
    Select Case Info
        Case "CPT"
            CorresCPT = ClientCompte
            If Left(ClientCompte, 3) = 421 Then CorresCPT = 421
            If Left(ClientCompte, 3) = 411 Then CorresCPT = 411
            If Left(ClientCompte, 3) = 401 Then CorresCPT = 401
        Case "AUX"
            CorresCPT = ""
            
            If Left(ClientCompte, 3) = 421 Then
                If Regroupement = "Vrai" Then
                   CorresCPT = "S0000000"
                Else
                    CorresCPT = "S" & Right(ClientCompte, Len(ClientCompte) - 3)
                End If
            End If
            If Left(ClientCompte, 3) = 411 Then
                If Regroupement = "Vrai" Then
                   CorresCPT = "C0000000"
                Else
                    CorresCPT = "C" & Right(ClientCompte, Len(ClientCompte) - 3)
                End If
            End If
            If Left(ClientCompte, 3) = 401 Then
                If Regroupement = "Vrai" Then
                   CorresCPT = "F0000000"
                Else
                    CorresCPT = "F" & Right(ClientCompte, Len(ClientCompte) - 3)
                End If
            End If
        Case Else
            CorresCPT = Null
    End Select
    
    If IsNull(CorresCPT) Or Len(CorresCPT) = 0 Then CorresCPT = ""

End Function

Function CorresAUX(Auxiliaire As Variant) As Variant 'Gestion des correspondances de codes auxiliaires

    On Error GoTo ErreurAUX:
    
    If (Len(Auxiliaire) >= 8) Then Auxiliaire = Left(auxilaire, 7)

    If (Len(Auxiliaire) <> 0) Then
        CorresAUX = "C" & Auxiliaire
    Else
        CorresAUX = auxilaire
    End If
        
    
    Exit Function
    
ErreurAUX:
    CorresAUX = Auxiliaire

End Function

Function Replac_Spe(Texte As Variant) As Variant
    On Error GoTo ErreurRpl:
    
    Replac_Spe = Replace(Texte, ";", "-")
'    Texte = Replace(Texte, "", " ")
    
    Exit Function
    
ErreurRpl:
    Replac_Spe = Texte
End Function

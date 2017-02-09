Attribute VB_Name = "Module2"
Option Explicit

Sub Go_Vide()
'
' Go_Vide Macro
' Réinitialise les paramétrages par défaut
'

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A40:AD40").Select
    Selection.Copy
    Sheets("Listes").Visible = False
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 9).Value = 1
    Cells(3, 4).Value = "Point virgule ( ; )"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "A DEFINIR"
End Sub

Sub Go_POMO()
'
' Go_POMO Macro
' Pré-paramétrage pour importation POMO
'

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A43:AD43").Select
    Selection.Copy
    Sheets("Listes").Visible = False
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 9).Value = 1
    Cells(3, 4).Value = "Demi colonne ( | )"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "POMO"
End Sub

Sub Go_Cote_Ouest()

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A46:AD46").Select
    Selection.Copy
    Sheets("Listes").Visible = False
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 9).Value = 1
    Cells(3, 4).Value = "Point virgule ( ; )"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "Cote Ouest"
End Sub

Sub Go_ISAGRI()

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A50:AD50").Select
    Selection.Copy
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Listes").Select
    Range("A49:AD49").Select
    Selection.Copy
    Sheets("Param").Select
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Listes").Visible = False
    Cells(1, 9).Value = "1"
    Cells(3, 4).Value = "Champ fixe"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "ISAGRI"
End Sub

Sub Go_CFC_Caisse()

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A54:AD54").Select
    Selection.Copy
    Sheets("Listes").Visible = False
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 9).Value = 2
    Cells(3, 4).Value = "Point virgule ( ; )"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "CFC Caisse"
End Sub

Sub Go_CFC_Fact()

    Sheets("Listes").Visible = True
    Sheets("Listes").Select
    Cells(1, 11).Value = 1
    Range("A57:AD57").Select
    Selection.Copy
    Sheets("Listes").Visible = False
    Sheets("Param").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 9).Value = 2
    Cells(3, 4).Value = "Point virgule ( ; )"
    Cells(3, 9).Value = "jj/mm/aaaa"
    Range("A1").Select
    Sheets("Dossier").Select
    Cells(4, 2).Value = "CFC Fact"
End Sub


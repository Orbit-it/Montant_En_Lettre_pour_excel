Option Explicit
Option Base 1

Function NombreEnLettre(ByVal Nombre As Double) As String
    ' Fonction principale qui convertit un nombre en lettres / arrondire la Partie decimale à 2 chiffre après la virgule
    ' Version compatible avec Excel 2016 ou plus !
    ' Auteur: Serigne Mansour Diop --> github: Orbit-it (sdiop)
    
    Dim Entier As Double
    Dim PartieDecimale As Integer
    Dim Resultat As String
    
    ' Gestion du zéro
    If Nombre = 0 Then
        NombreEnLettre = "Zéro Franc CFA"
        Exit Function
    End If
    
    ' Gestion du signe négatif
    If Nombre < 0 Then
        NombreEnLettre = "Moins " & NombreEnLettre(Abs(Nombre))
        Exit Function
    End If
    
    ' Séparation de la partie entière et décimale
    Entier = Int(Nombre)
    PartieDecimale = Round((Nombre - Entier) * 100)
    If PartieDecimale >= 100 Then
        PartieDecimale = 0
        Entier = Entier + 1
    End If
    
    ' Conversion de la partie entière
    If Entier > 0 Then
        Resultat = ConvertirPartieEntiere(Entier)
    End If
    
    ' Ajout de la devise avec gestion du pluriel
    If Entier > 1 Then
        Resultat = Resultat & " Francs"
    ElseIf Entier = 1 Then
        Resultat = Resultat & " Franc"
    End If
    
    ' Ajout des centimes si nécessaire
    If PartieDecimale > 0 Then
        If Entier > 0 Then
            Resultat = Resultat & " et"
        End If
        Resultat = Resultat & " " & ConvertirCentimes(PartieDecimale)
        If PartieDecimale > 1 Then
            Resultat = Resultat & " Centimes"
        Else
            Resultat = Resultat & " Centime"
        End If
    End If
    
    Resultat = Resultat & " CFA"
    NombreEnLettre = Application.Trim(Resultat)
    
End Function

Private Function ConvertirPartieEntiere(ByVal Nombre As Double) As String
    ' Convertit la partie entière en lettres
    
    Dim Milliards As Long
    Dim Millions As Long
    Dim Milliers As Long
    Dim Unites As Integer
    Dim Resultat As String
    Dim Temp As String
    Dim Reste As Double
    
    ' Initialisation
    Resultat = ""
    Reste = Nombre
    
    ' Extraction des milliards (max 999)
    If Reste >= 1000000000# Then
        Milliards = Int(Reste / 1000000000#)
        Reste = Reste - Milliards * 1000000000#
        
        Temp = ConvertirGroupe(Milliards)
        If Temp = "un" Then
            Resultat = Resultat & "un milliard "
        Else
            Resultat = Resultat & Temp & " milliards "
        End If
    End If
    
    ' Extraction des millions
    If Reste >= 1000000 Then
        Millions = Int(Reste / 1000000)
        Reste = Reste - Millions * 1000000
        
        Temp = ConvertirGroupe(Millions)
        If Temp = "un" Then
            Resultat = Resultat & "un million "
        Else
            Resultat = Resultat & Temp & " millions "
        End If
    End If
    
    ' Extraction des milliers
    If Reste >= 1000 Then
        Milliers = Int(Reste / 1000)
        Reste = Reste - Milliers * 1000
        
        Temp = ConvertirGroupe(Milliers)
        If Temp = "un" Then
            Resultat = Resultat & "mille "
        Else
            Resultat = Resultat & Temp & " mille "
        End If
    End If
    
    ' Traitement des unités (ce qui reste < 1000)
    If Reste > 0 Then
        Resultat = Resultat & ConvertirGroupe(Reste)
    ElseIf Nombre = 0 Then
        Resultat = "zéro"
    End If
    
    ' Nettoyage des espaces superflus
    ConvertirPartieEntiere = Trim(Resultat)
    
End Function

Private Function ConvertirGroupe(ByVal Nombre As Integer) As String
    ' Convertit un nombre de 1 à 999 en lettres
    
    Dim Centaines As Integer
    Dim Dizaines As Integer
    Dim Unites As Integer
    Dim Resultat As String
    
    ' Tableaux de conversion
    Dim UnitesLettres As Variant
    Dim DizainesLettres As Variant
    Dim DixSeizeLettres As Variant
    
    ' Initialisation des tableaux
    UnitesLettres = Array("un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf")
    DizainesLettres = Array("dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    DixSeizeLettres = Array("onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf")
    
    ' Cas particulier du zéro
    If Nombre = 0 Then
        ConvertirGroupe = ""
        Exit Function
    End If
    
    ' Extraction des composants
    Centaines = Int(Nombre / 100)
    Dizaines = Int((Nombre Mod 100) / 10)
    Unites = Nombre Mod 10
    
    ' TRAITEMENT DES CENTAINES
    If Centaines > 0 Then
        If Centaines = 1 Then
            Resultat = "cent"
        Else
            Resultat = UnitesLettres(Centaines) & " cent"
        End If
        
        ' Ajout du "s" à cent si pas de suite et centaines > 1
        If Dizaines = 0 And Unites = 0 And Centaines > 1 Then
            Resultat = Resultat & "s"
        End If
        Resultat = Resultat & " "
    End If
    
    ' TRAITEMENT DES DIZAINES ET UNITES
    If Dizaines = 0 Then
        ' Pas de dizaine, juste les unités
        If Unites > 0 Then
            Resultat = Resultat & UnitesLettres(Unites) & " "
        End If
    ElseIf Dizaines = 1 Then
        ' Cas des nombres de 10 à 19
        Resultat = Resultat & DixSeizeLettres(Unites + 1) & " "
    ElseIf Dizaines = 7 Or Dizaines = 9 Then
        ' Cas particuliers : 70-79 et 90-99 (soixante-dix, quatre-vingt-dix)
        If Unites = 0 Then
            ' Dizaine ronde (70 ou 90)
            If Dizaines = 7 Then
                Resultat = Resultat & "soixante-dix "
            Else
                Resultat = Resultat & "quatre-vingt-dix "
            End If
        ElseIf Unites = 1 And Dizaines = 7 Then
            Resultat = Resultat & "soixante et onze "
        ElseIf Unites = 1 And Dizaines = 9 Then
            Resultat = Resultat & "quatre-vingt-onze "
        Else
            Resultat = Resultat & DizainesLettres(Dizaines) & "-" & DixSeizeLettres(Unites + 1) & " "
        End If
    Else
        ' Cas général (20-60, 80)
        If Unites = 1 And Dizaines <> 8 Then
            ' Exception pour 21, 31, 41, 51, 61 (sauf 81)
            Resultat = Resultat & DizainesLettres(Dizaines) & " et un "
        ElseIf Unites = 0 Then
            ' Dizaine ronde
            If Dizaines = 8 Then
                Resultat = Resultat & "quatre-vingts "
            Else
                Resultat = Resultat & DizainesLettres(Dizaines) & " "
            End If
        Else
            ' Dizaine + unité normale
            Resultat = Resultat & DizainesLettres(Dizaines) & "-" & UnitesLettres(Unites) & " "
        End If
    End If
    
    ' Nettoyage et retour
    ConvertirGroupe = Trim(Resultat)
    
End Function

Private Function ConvertirCentimes(ByVal Centimes As Integer) As String
    ' Convertit les centimes (0-99) en lettres
    
    If Centimes = 0 Then
        ConvertirCentimes = ""
    ElseIf Centimes < 10 Then
        ConvertirCentimes = UnitesLettres(Centimes)
    Else
        ConvertirCentimes = ConvertirGroupe(Centimes)
    End If
    
End Function

' Fonctions auxiliaires pour les tableaux
Private Function UnitesLettres(Index As Integer) As String
    Dim Tbl As Variant
    Tbl = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf")
    If Index >= 0 And Index <= 9 Then
        UnitesLettres = Tbl(Index)
    Else
        UnitesLettres = ""
    End If
End Function

Private Function DizainesLettres(Index As Integer) As String
    Dim Tbl As Variant
    Tbl = Array("", "dix", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    If Index >= 0 And Index <= 9 Then
        DizainesLettres = Tbl(Index)
    Else
        DizainesLettres = ""
    End If
End Function

Private Function DixSeizeLettres(Index As Integer) As String
    Dim Tbl As Variant
    Tbl = Array("dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf")
    If Index >= 1 And Index <= 10 Then
        DixSeizeLettres = Tbl(Index - 1)
    Else
        DixSeizeLettres = ""
    End If
End Function

' Fonction simplifiée pour utilisation dans Excel
Function NbreEnLettre(ByVal Cellule As Range) As String
    ' Utilisation : =NbreEnLettre(F28)
    On Error GoTo Erreur
    
    If Cellule.Cells.Count > 1 Then
        NbreEnLettre = "Erreur : sélectionnez une seule cellule"
        Exit Function
    End If
    
    If IsNumeric(Cellule.Value) Then
        NbreEnLettre = NombreEnLettre(CDbl(Cellule.Value))
    Else
        NbreEnLettre = "Erreur : cellule non numérique"
    End If
    
    Exit Function
    
Erreur:
    NbreEnLettre = "Erreur de calcul"
End Function

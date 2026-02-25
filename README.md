# Montant_En_Lettre_pour_excel
Une solution VBA complète et robuste pour convertir un nombre en lettres (en français) avec gestion des millions, milliers, centaines, etc.

# Comment installer le code VBA :
Ouvrez l'éditeur VBA : Alt + F11

Insérez un module : Clic droit sur "VBAProject" > Insérer > Module

Copier le code du fichier "VBA" et Collez le dans le module

Fermez l'éditeur et enregistrez le fichier au format .xlsm (classeur macro-enabled)

# Comment utiliser ces fonctions dans Excel :
Appeler la fonction sur n'importe quelle cellule: =NombreEnLettre(Cellule) avec Cellule le nom de la cellule contenant le montant en chiffre à convertir.

Exemples de résultats :
Nombre	Résultat
0	Zéro Franc CFA
1	un Franc CFA
42	quarante-deux Francs CFA
100	cent Francs CFA
101	cent un Francs CFA
1000	mille Francs CFA
1500	mille cinq cents Francs CFA
1999	mille neuf cent quatre-vingt-dix-neuf Francs CFA
1000000	un million Francs CFA
1500000	un million cinq cent mille Francs CFA
1234.56	mille deux cent trente-quatre Francs cinquante-six Centimes CFA


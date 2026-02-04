\# Générateur de devis “plus” – Documentation technique (VBA / Excel)



Cette version “plus” étend le générateur de devis avec de nouvelles fonctionnalités d’automatisation et une macro de filtrage/masquage de lignes.\[file:3]\[file:4]  

L’objectif est de simplifier la lecture et l’impression du devis en ne montrant que les lignes réellement utilisées.



\---



\## 1. Objectifs fonctionnels supplémentaires



Par rapport à la version 2010 :



\- Améliorer l’ergonomie de la feuille de devis.  

\- Cacher automatiquement les lignes non utilisées (quantité vide, ou cellule vide dans la colonne de contrôle).  

\- Préparer un devis propre pour impression / export (sans “lignes vides”).\[file:3]\[file:4]  



\---



\## 2. Structure du classeur



La structure de base reste similaire à la version 2010 : catalogue + feuille de saisie du devis.\[file:4]  



\### 2.1. Feuille DEVIS (exemple)



\- Colonnes typiques :  

&#x20; - Code de Référence de l'Article  

&#x20; - Libellé de l'Article  

&#x20; - Unité de Mesure  

&#x20; - Quantité  

&#x20; - Prix unitaire HT  

&#x20; - Montant total  

&#x20; - Colonne “contrôle / activation” utilisée pour savoir si la ligne doit être affichée.\[file:4]  



\- Ligne 4 (par exemple en F4) sert de point d’entrée pour le filtrage : la macro sélectionne à partir de cette cellule.\[file:3]  



\---



\## 3. Macro de masquage des lignes



\### 3.1. Code VBA



```vba

Private Sub CommandButton1\_Click()



&#x20;   Range("F4").Select

&#x20;   For Each o In Selection

&#x20;       If o.Value = "" Then

&#x20;           o.EntireRow.Hidden = True

&#x20;       End If

&#x20;   Next



End Sub




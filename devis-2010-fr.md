\# Générateur de devis 2010 – Documentation technique (VBA / Excel)



Ce fichier Excel constitue un générateur de devis basé sur un bordereau de prix détaillé pour les travaux d’aménagement paysager et de génie civil (Ville de Rennes, etc.).\[file:2]  

C’est l’un des premiers projets d’\*\*automatisation\*\* de l’auteur : l’objectif est de transformer un catalogue de prix complexe en outil interactif de construction de devis.



\---



\## 1. Objectifs fonctionnels



\- Centraliser le bordereau de prix dans un classeur unique.  

\- Permettre la sélection rapide d’articles (codes, libellés, unités, prix unitaires).  

\- Calculer automatiquement les montants en fonction des quantités saisies.  

\- Structurer le devis par familles (A, B, C, D… : terrassement, plantations, réseaux, etc.).\[file:2]  



\---



\## 2. Structure du classeur



> Les noms d’onglets exacts peuvent varier, mais la structure logique est la suivante.



\### 2.1. Feuille catalogue / bordereau



\- Colonnes principales :  

&#x20; - `Code Catalogue` (ex. JAE)  

&#x20; - `Code de Référence de l'Article` (ex. A, A1, A10001…)  

&#x20; - `Type d'Article (1 ou 2)`  

&#x20; - `Libellé de l'Article`  

&#x20; - `Code Unité de Mesure` (F, u, m², m³, ml, etc.)  

&#x20; - `Code TVA`  

&#x20; - `Prix Unitaire Hors Taxe`  

&#x20; - `Texte Descriptif de l'Article` (long, type CCTP).\[file:2]  



\- Hiérarchie :  

&#x20; - Niveau 2 = familles / sous-familles (A, A1, B, B1…).  

&#x20; - Niveau 1 = articles tarifés (code complet type A10001, B10101…).\[file:2]  



\### 2.2. Feuille de saisie du devis



\- Zone de saisie :  

&#x20; - Code article (ou liste déroulante)  

&#x20; - Libellé  

&#x20; - Unité  

&#x20; - Prix unitaire HT  

&#x20; - Quantité  

&#x20; - Montant total (Qté × PU).\[file:2]  



\- Totaux :  

&#x20; - Sous-totaux par chapitre (A, B, C…).  

&#x20; - Total général HT, TVA, TTC (si implémenté).



\### 2.3. Feuilles auxiliaires (optionnel)



\- Paramètres (TVA, taux divers).  

\- Impressions / mise en forme du devis client.



\---



\## 3. Logique d’automatisation (sans code détaillé)



\### 3.1. Alimentation des lignes de devis



\- Recherche de l’article par code dans la feuille catalogue (via RECHERCHEV / INDEX+EQUIV ou macro).  

\- Récupération automatique : libellé, unité, prix unitaire, texte descriptif si nécessaire.\[file:2]  



\### 3.2. Calculs automatisés



\- Formule de montant : `=Quantité \* Prix\_Unitaire\_HT`.  

\- Agrégation des montants par famille (somme conditionnelle sur le code A, B, C…).\[file:2]  



\### 3.3. Structuration par chapitres



\- Utilisation des codes (A, B, C…) pour afficher les titres de chapitre dans le devis.  

\- Possibilité d’inclure ou non certains chapitres en fonction du projet.



\---



\## 4. Macros VBA (principes)



> Le fichier 2010 contient principalement la logique de catalogue et de calcul.  

> La macro de filtrage/masquage documentée dans la version suivante (plus) est un prolongement logique.



Exemples de responsabilités possibles des macros :



\- Initialisation d’un nouveau devis (effacer les anciennes quantités / montants).  

\- Copie d’une sélection d’articles depuis le catalogue vers la feuille devis.  

\- Mise à jour des montants et totaux.



\---



\## 5. Vision “projet d’automatisation”



\- À l’époque, le vocabulaire “RPA / automation” n’était pas utilisé, mais le besoin était déjà :  

&#x20; - réduire le temps de préparation des devis,  

&#x20; - fiabiliser les prix et libellés (une seule source de vérité),  

&#x20; - structurer un processus métier récurrent (réponse à appel d’offres, marchés à bons de commande, etc.).\[file:2]  



\- Le fichier Excel + VBA représente un précurseur de ce que l’on ferait aujourd’hui avec :  

&#x20; - une appli web,  

&#x20; - une API de catalogue,  

&#x20; - ou un workflow no-code.



\### 5.1. Gain de temps estimé



Avant la mise en place du générateur de devis, la préparation d’un devis complet à partir du bordereau prenait typiquement entre 1 h 30 et 2 h (recherche manuelle des lignes, recopies, contrôles).  

Avec le classeur automatisé (recherche d’articles, calculs et structure prêts à l’emploi), le même devis peut être préparé en 30 à 45 minutes, soit un gain d’environ 1 h par devis.  



Sur une base prudente de 40 à 60 devis par an, cela représente un ordre de grandeur de 40 à 60 heures de travail économisées chaque année, uniquement sur la phase de préparation des devis.  

Ce gain se traduit en pratique par plus de disponibilité pour l’analyse technique des projets et les échanges avec la maîtrise d’ouvrage, plutôt que par un calcul d’économies financières (hors périmètre dans un contexte de fonction publique).\[file:2]\[file:4]  



\---



\## 6. Pistes d’évolution modernes



\- Export du devis en PDF avec mise en page soignée.  

\- Connexion à une base de données (SQLite, PostgreSQL, etc.) au lieu d’Excel.  

\- Interface web (Django / FastAPI + front) reprenant la même logique de bordereau.  

\- Intégration avec des outils d’automatisation (n8n, Make, Zapier) pour enchaîner : demande → chiffrage → envoi du PDF.




# Calculateur de Primes d'Assurance en VBA

Ce projet est un outil simplifié de calcul de primes d'assurance développé en VBA pour Microsoft Excel.  
L'outil permet à l'utilisateur de saisir son prénom, son nom, son âge et le nombre de sinistres,  
puis de calculer et d'afficher la prime d'assurance correspondante à l'aide d'un UserForm.

## Accès au code source

Le code source du projet VBA est accessible de deux manières :

1. **Fichier Excel (.xlsm)**  
   Vous pouvez télécharger le fichier `CalculateurPrimeAssurance.xlsm` et l’ouvrir dans Microsoft Excel. Pour accéder au code, appuyez sur `Alt + F11` pour ouvrir l'éditeur VBA. N'oubliez pas d'activer les macros si nécessaire.

2. **Modules exportés**  
   Le dossier `src/` contient les modules et UserForms exportés au format texte (.bas, .frm, .frx). Vous pouvez ouvrir ces fichiers avec un éditeur de texte pour consulter le code et suivre l’évolution du projet.


## Fonctionnalités

- **Saisie des informations personnelles :**  
  L'utilisateur saisit son **prénom** et son **nom** dans des zones de texte dédiées.

- **Saisie des données d'assurance :**  
  - **Âge du conducteur**  
  - **Nombre de sinistres** (le nombre de sinistres est pris en compte pour ajuster le montant de la prime)

- **Calcul de la prime d'assurance :**  
  La prime de base est fixée à 100 €. Selon l'âge du conducteur, la prime est ajustée :  
  - Si l'âge est inférieur à 25 ans, la prime est multipliée par 1,5.  
  - Si l'âge est compris entre 25 et 65 ans, la prime est multipliée par 1,2.  
  - Si l'âge est supérieur à 65 ans, la prime est multipliée par 1,3.  
  Ensuite, pour chaque sinistre, on ajoute 50 € à la prime.

- **Affichage du résultat :**  
  Le résultat s'affiche dans un label du UserForm sous la forme :  
  `Bonjour [Prénom] [Nom], la prime d'assurance est de [Montant] €`

- **Validation des entrées :**  
  - Le programme vérifie que l'âge et le nombre de sinistres sont des valeurs numériques.  
  - Si l'âge est inférieur à 18 ans, un message d'erreur s'affiche.

## Calcul des Primes

Le calcul de la prime se fait en deux étapes principales :

1. **Ajustement en fonction de l'âge :**

   - **Prime de base :** 100 €
   - **Multiplicateur d'âge :**
     - Si `age < 25` : prime = 100 * 1.5  
     - Si `25 <= age <= 65` : prime = 100 * 1.2  
     - Si `age > 65` : prime = 100 * 1.3

2. **Ajout en fonction du nombre de sinistres :**

   Pour chaque sinistre, on ajoute 50 €.  
   Par exemple, pour 2 sinistres, on ajoute 2 * 50 = 100 €.

**Formule finale :**

prime = (Prime de base ajustée selon l'âge) + (Nombre de sinistres * 50)

# Rapport des ventes en excel

Ce programme Go génère un fichier Excel contenant un rapport de ventes mensuelles. Il utilise la bibliothèque `excelize` pour créer un fichier Excel avec des données d'exemple, des styles et des formules.

## Fonctionnalités

- Création d'un nouveau fichier Excel
- Ajout de données de ventes pour différents produits
- Calcul automatique des totaux avec des formules Excel
- Application de styles aux en-têtes (fond bleu, texte blanc, bordures)
- Largeur fixe (20) pour les colonnes A à D
- Sauvegarde du fichier sous le nom `rapport_ventes.xlsx`

## Dépendances

- Go 1.24 ou supérieur (version minimale exigée par `excelize` v2.10)
- Bibliothèque `github.com/xuri/excelize/v2`

## Installation

1. Assurez-vous que Go est installé sur votre système.
2. [Clonez ou téléchargez ce projet](https://github.com/lemire/goexcel/archive/refs/heads/main.zip).
3. Installez les dépendances :

   ```bash
   go mod tidy
   ```

## Utilisation

Exécutez le programme :

```bash
go run main.go
```

Le fichier `rapport_ventes.xlsx` sera généré dans le répertoire courant. Vous pouvez l'ouvrir avec Excel ou tout autre lecteur de fichiers Excel. Vous pouvez ouvrir le document.


## Activité

Examinez le contenu du programme (`main.go`). Pouvez-vous modifier les données et générer un nouveau chiffrier excel?

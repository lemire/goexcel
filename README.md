# Rapport de Ventes Excel

Ce programme Go génère un fichier Excel contenant un rapport de ventes mensuelles. Il utilise la bibliothèque `excelize` pour créer un fichier Excel avec des données d'exemple, des styles et des formules.

## Fonctionnalités

- Création d'un nouveau fichier Excel
- Ajout de données de ventes pour différents produits
- Calcul automatique des totaux avec des formules Excel
- Application de styles aux en-têtes (fond bleu, texte blanc, bordures)
- Ajustement automatique de la largeur des colonnes
- Sauvegarde du fichier sous le nom `sales_report.xlsx`

## Dépendances

- Go 1.16 ou supérieur
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

Le fichier `sales_report.xlsx` sera généré dans le répertoire courant. Vous pouvez l'ouvrir avec Excel ou tout autre lecteur de fichiers Excel. Vous pouvez ouvrir le document.

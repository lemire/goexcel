package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Créer un nouveau fichier Excel
	f := excelize.NewFile()

	// Renommer la feuille par défaut (optionnel)
	f.SetSheetName("Sheet1", "Rapport de Ventes")

	// Définir quelques valeurs de cellules avec différents styles
	f.SetCellValue("Rapport de Ventes", "A1", "Rapport de Ventes Mensuelles")
	f.SetCellValue("Rapport de Ventes", "A2", "Produit")
	f.SetCellValue("Rapport de Ventes", "B2", "Quantité")
	f.SetCellValue("Rapport de Ventes", "C2", "Prix")
	f.SetCellValue("Rapport de Ventes", "D2", "Total")

	// Données d'exemple
	products := []struct {
		Name     string
		Quantity int
		Price    float64
	}{
		{"Ordinateur portable", 15, 999.99},
		{"Souris", 120, 25.50},
		{"Clavier", 80, 75.00},
		{"Moniteur", 35, 299.99},
	}

	for i, p := range products {
		row := i + 3 // commencer à la ligne 3
		f.SetCellValue("Rapport de Ventes", fmt.Sprintf("A%d", row), p.Name)
		f.SetCellValue("Rapport de Ventes", fmt.Sprintf("B%d", row), p.Quantity)
		f.SetCellValue("Rapport de Ventes", fmt.Sprintf("C%d", row), p.Price)
		// Total = Quantité × Prix
		f.SetCellFormula("Rapport de Ventes", fmt.Sprintf("D%d", row), fmt.Sprintf("B%d*C%d", row, row))
	}

	// Styler la ligne d'en-tête
	styleHeader, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#4472C4"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		log.Fatal(err)
	}
	f.SetCellStyle("Rapport de Ventes", "A2", "D2", styleHeader)

	// Ajouter une ligne de total en bas
	totalRow := len(products) + 3
	f.SetCellValue("Rapport de Ventes", "C"+fmt.Sprint(totalRow), "Total Général")
	f.SetCellFormula("Rapport de Ventes", "D"+fmt.Sprint(totalRow), fmt.Sprintf("SUM(D3:D%d)", totalRow-1))

	// Ajuster automatiquement la largeur des colonnes
	f.SetColWidth("Rapport de Ventes", "A", "D", 20)

	// Sauvegarder le fichier
	filename := "rapport_ventes.xlsx"
	if err := f.SaveAs(filename); err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Fichier Excel '%s' généré avec succès !\n", filename)
}

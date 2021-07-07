package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"log"
	"os"
)

// Referência para colunas do Excel
var colunas = []string{"A", "B", "C", "D", "E", "F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"}

// Nome da coluna a ser removida
var coluna string = "cartão de crédito"

func main() {

	// Diretório onde os arquivos serão colocados e manipulados
	dirname := "arquivos"

	// Abre o diretório onde os arquivos foram colocados
	f, err := os.Open(dirname)
	if err != nil {
		log.Fatal(err)
	}

	// Coloca os arquivos em um array
	files, err := f.Readdir(-1)
	f.Close()
	if err != nil {
		log.Fatal(err)
	}

	// Para cada arquivo encontrado, chama-se a rotina para remover a coluna
	for _, file := range files {
		fmt.Println(file.Name())
		removeColuna(file.Name())
	}

}

// Remove uma coluna do Excel quando é encontrada
func removeColuna(arquivo string) {

	// Abre o arquivo Excel
	f, err := excelize.OpenFile("arquivos/" + arquivo)
	if err != nil {
		log.Fatalln(err)
	}

	// Abre a primeira pasta (Sheet)
	Sheet := f.WorkBook.Sheets.Sheet[0].Name

	// Carrega as colunas
	cols, _ := f.GetCols(Sheet)

	// Procura a coluna e remove se for encontrada
	for i, col := range cols {
		if col[0] == coluna {
			f.RemoveCol(Sheet,colunas[i])
		}
	}

	// Salva a planilha Excel
	f.Save()

}

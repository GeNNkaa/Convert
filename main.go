package main

import (
	"encoding/csv"
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
	"os"
	"sort"
	"strings"
)

func main() {
	var startRows []int

	myApp := app.New()
	myWindow := myApp.NewWindow("Выбор файла")

	// Переменная для хранения имени выбранного файла
	var selectedFileNime string

	selectButton := widget.NewButton("Выбрать файл", func() {
		dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err == nil && reader != nil {
				defer reader.Close()
				selectedFileNime = reader.URI().Path()
			}
		}, myWindow)
	})

	//Создаем контейнер для размещения виджетов
	content := container.NewVBox(
		selectButton,
	)

	myWindow.SetContent(content)
	myWindow.ShowAndRun()

	//Открываем необходимый файл
	sprav, err := excelize.OpenFile(selectedFileNime)
	if err != nil {
		fmt.Println(err)
		return
	}

	//Присваиваем значение первого листа переменной
	osn := "Основной справочник"

	//Находим объединенные ячейки
	mergedCells, err := sprav.GetMergeCells(osn)
	if err != nil {
		fmt.Println(err)
		return
	}

	//Из полученных ячеек находим координаты
	//И присваиваем значения координат переменной
	for _, cell := range mergedCells {
		startCell := cell.GetStartAxis()
		_, startRow, _ := excelize.CellNameToCoordinates(startCell)
		startRows = append(startRows, startRow)
	}

	//Сортируем значения от последнего к первому
	sort.Slice(startRows, func(i, j int) bool {
		return startRows[i] > startRows[j]
	})

	// Удаляем объединенные ячейки
	for _, v := range startRows {
		sprav.RemoveRow(osn, v)
	}

	// Присваиваем переменной office значение ячейки А1
	office, err := sprav.GetCellValue(osn, "A1")
	if err != nil {
		fmt.Println(err)
		return
	}

	//Если ячейка содержит в себе "ОФИС" удаляем строку
	if office == "ОФИС " {
		sprav.RemoveRow(osn, 1)
	}

	// Сохраняем файл
	err = sprav.SaveAs("Справочник.xlsx")
	if err != nil {
		fmt.Println(err)
	}

	//Создаем файл csv
	csvFile, err := os.Create("Contacts.csv")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer csvFile.Close()

	//Создаем писатель
	writer := csv.NewWriter(charmap.Windows1251.NewEncoder().Writer(csvFile))
	defer writer.Flush()

	sort := []string{"Имя", "Должность", "Отдел", "Рабочий телефон", "Телефон раб. 2", "Адрес эл. почты", "\r\n"}
	writer.Write(sort)

	rows, err := sprav.GetRows(osn)
	if err != nil {
		fmt.Println(err)
		return
	}

	//Выбираем столбцы которые необходимо считать
	columnIndices := []int{1, 2, 3, 4, 6, 8}

	// Считываем данные из необходимых столбцов и меняем казахские символы на кирилицу
	for _, row := range rows {
		exportedColumns := make([]string, 0)

		for _, columnIndex := range columnIndices {
			if columnIndex < len(row) {
				row[columnIndex] = strings.Replace(row[columnIndex], "ә", "а", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "ғ", "г", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "ң", "н", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "ө", "о", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "қ", "к", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "ұ", "у", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "ү", "у", -1)

				row[columnIndex] = strings.Replace(row[columnIndex], "Ә", "А", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "Ғ", "Г", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "Ң", "Н", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "Ө", "О", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "Қ", "К", -1)
				row[columnIndex] = strings.Replace(row[columnIndex], "Ұ", "У", -1)

				exportedColumns = append(exportedColumns, row[columnIndex])
			} else {
				exportedColumns = append(exportedColumns, "")
			}

		}
		//Меняем кодировку из LF в CRLF
		//exportedColumns = append(exportedColumns, "\r\n")

		// Запись в SCV файл
		err := writer.Write(exportedColumns)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
}

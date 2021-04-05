// ea_unpacking
package main

import (
	// "fmt"
	"io"
	"io/ioutil"
	"log"
	"math"
	"os"
	"path/filepath"
	"strconv"

	"os/exec"

	"time"

	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type abonentStruct struct {
	fatherid            string  `json:"fatherid"`
	ratio_count         uint    `json:"ratio_count"`
	inside_losses_ratio float64 `json:"inside_losses_ratio"`
	inside_losses       float32 `json:"inside_losses"`
	gp_losses           float32 `json:"gp_losses"`
	desc                string  `json:"desc"`
	costs               float32 `json:"costs"`
}

type configStruct struct {
	registryName   string
	decryptionPath string
	profilesPath   string
	unpackingName  string
	logFileName    string
	currentPeriod  time.Time
}

var (
	config      configStruct
	abonentBase map[string]abonentStruct
)

func main() {
	setConfig()
	logFile, err := os.OpenFile(config.logFileName, os.O_WRONLY|os.O_CREATE, 0644)
	check(err)
	defer logFile.Close() // don't forget to close it

	mw := io.MultiWriter(os.Stdout, logFile)
	log.SetOutput(mw)
	log.Println("!!! Начата работа !!!")

	abonentBase := mainReadRegistry() // чтение реестра абонентов и субабонентов
	mainReadDecryption(abonentBase)   // чтение расшифровок абонентов из файлов ГП
	mainReadProfiles(abonentBase)     // чтение профилей
	mainCleanBase(abonentBase)        // очистка базы от незаполненных абонентов
	mainCreateUnpacking(abonentBase)  // создание и вывод таблицы распаковки

	log.Println("!!! Работа завершена !!! Вся информация записана в файл", config.unpackingName)
	err = exec.Command("cmd", "/C", "start", "chrome.exe", config.logFileName).Run()
	check(err)
}

func setConfig() {
	config.registryName = "реестр_распаковки.xlsx"
	config.decryptionPath = "descriptions/"
	config.profilesPath = "profiles/"
	config.unpackingName = "итоговый с распаковкой.xlsx"
	config.logFileName = "ea_unpacking.log"
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}

func xlsOpenFile(filename string) *excelize.File {
	f, err := excelize.OpenFile(filename)
	check(err)
	return f
}

func string2uint(str string) uint {
	i, err := strconv.Atoi(str)
	check(err)
	return uint(i)
}

func string2float32(str string) float32 {
	out, err := strconv.ParseFloat(str, 32)
	check(err)
	return float32(out)
}

func string2float64(str string) float64 {
	out, err := strconv.ParseFloat(str, 64)
	check(err)
	return out
}

func int2string(number int) string {
	return strconv.Itoa(number)
}

// если коэф-т потерь от абонента пустой, то возвращаем 0 текстом
func inside_losses_check(num string) string {
	if num == "" {
		return "0"
	} else {
		return num
	}
}

// чтение реестра
func mainReadRegistry() map[string]abonentStruct {
	localAbonentBase := make(map[string]abonentStruct)

	xlsRegistry := xlsOpenFile(config.registryName)
	xlsRegistrySheet := xlsRegistry.GetSheetName(1)

	rows := xlsRegistry.GetRows(xlsRegistrySheet)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if row[3] == "#DIV/0!" {
			log.Println("Абонент ", row[0], ", не указан коэф-т потерь. Пропускаем.")
			continue
		}
		var abonentRecord abonentStruct
		abonentRecord.fatherid = row[1]
		abonentRecord.ratio_count = string2uint(row[2])
		abonentRecord.inside_losses_ratio = string2float64(inside_losses_check(row[3]))
		abonentRecord.desc = row[4]

		localAbonentBase[row[0]] = abonentRecord
	}
	return localAbonentBase
}

// чтение файла с расшифровками, добавляем потери и расход основных абонентов
func oneReadDecryption(filename string, abonentBase map[string]abonentStruct) {
	xlsFile := xlsOpenFile(filename)
	xlsFileSheet := xlsFile.GetSheetName(1)

	rows := xlsFile.GetRows(xlsFileSheet)
	for _, row := range rows {
		if val, ok := abonentBase[row[2]]; ok {
			var abonentRecord abonentStruct
			abonentRecord.fatherid = abonentBase[row[2]].fatherid
			abonentRecord.ratio_count = abonentBase[row[2]].ratio_count
			abonentRecord.inside_losses_ratio = abonentBase[row[2]].inside_losses_ratio
			abonentRecord.desc = abonentBase[row[2]].desc

			if row[10] != "" { // Основной расход от ГП
				abonentRecord.costs = string2float32(row[10])
			} else if val.costs != 0 {
				abonentRecord.costs = abonentBase[row[2]].costs
			}

			if row[13] != "" { // Потери от ГП
				abonentRecord.gp_losses = string2float32(row[13])
			} else if val.gp_losses != 0 {
				abonentRecord.gp_losses = abonentBase[row[2]].gp_losses
			}

			abonentBase[row[2]] = abonentRecord
		}
	}

}

// перебор всех файлов с расшифровками
func mainReadDecryption(abonentBase map[string]abonentStruct) {
	files, err := ioutil.ReadDir(config.decryptionPath)
	check(err)

	for _, f := range files {
		if filepath.Ext(f.Name()) == ".xlsx" && f.Name()[0:1] != "~" {
			log.Println("[расшифровка] найден файл \"" + f.Name())
			oneReadDecryption(config.decryptionPath+f.Name(), abonentBase)
		}
	}
}

// перебор всех файлов с профилями
func mainReadProfiles(abonentBase map[string]abonentStruct) {
	files, err := ioutil.ReadDir(config.profilesPath)
	check(err)

	for _, f := range files {
		if filepath.Ext(f.Name()) == ".xlsx" && f.Name()[0:1] != "~" {
			log.Println("[профиль] найден файл \"" + f.Name())
			oneReadProfiles(config.profilesPath+f.Name(), abonentBase)
		}
	}
}

// чтение данных из файла профиля
func oneReadProfiles(filename string, abonentBase map[string]abonentStruct) {
	xlsFile := xlsOpenFile(filename)
	xlsFileSheet := xlsFile.GetSheetName(1)

	setCurrentPeriod(xlsFile.GetCellValue(xlsFileSheet, "A2"))

	rows := xlsFile.GetRows(xlsFileSheet)
	for _, row := range rows {
		rowsAbonent := row[0]             // строка с номером абонента из эксель файла
		for key, _ := range abonentBase { // перебираем всех абонентов из базы в памяти
			if strings.Contains(rowsAbonent, key) { // проверяем есть ли наши абоненты из базы в строке из экселя

				var abonentRecord abonentStruct
				abonentRecord.fatherid = abonentBase[key].fatherid
				abonentRecord.ratio_count = abonentBase[key].ratio_count
				abonentRecord.inside_losses_ratio = abonentBase[key].inside_losses_ratio
				abonentRecord.desc = abonentBase[key].desc

				if abonentBase[key].costs == 0 && abonentBase[key].fatherid == "" {
					log.Println("Абонент ", key, ", значение взято из профиля, т.к. отсутствует у ГП.")
					abonentRecord.costs = string2float32(row[7])
				} else if abonentBase[key].costs == 0 && abonentBase[key].fatherid != "" {
					abonentRecord.costs = string2float32(row[7])
				} else {
					abonentRecord.costs = abonentBase[key].costs
				}

				abonentRecord.gp_losses = abonentBase[key].gp_losses

				abonentBase[key] = abonentRecord

			}
		}
	}

}

// очистка базы от абонентов и субабонентов с недостатком данных
func mainCleanBase(abonentBase map[string]abonentStruct) {
	var abonents2Delete []string // перечень абонентов для удаления, т.к. по ним нет показаний. Удаляются как субабоненты, так и их родители
	log.Println("Проверяем базу на полноту заполнения.")

	// делаем выборку абонентов, у которых отсутствуют показания
	for key, value := range abonentBase {
		if value.costs == 0 {
			log.Println("Внимание! Профиль по абоненту ", key, " отсутствует. Расчеты по ним производиться не будут. Исправьте реестр или добавьте профили или расшифровки по ним.")
			abonents2Delete = append(abonents2Delete, key)
			if value.fatherid != "" {
				abonents2Delete = append(abonents2Delete, value.fatherid)
			}
		}
	}

	// делаем выборку абонентов, у которых отсутствуют субабоненты
	for key, value := range abonentBase {
		if value.fatherid == "" {
			if len(getSubAbonents(key, abonentBase)) == 0 {
				abonents2Delete = append(abonents2Delete, key)
			}
		}
	}

	// сначала удаляем записи, у которых родитель из удаляемых
	for _, v := range abonents2Delete {
		for key, value := range abonentBase {
			if value.fatherid == v {
				delete(abonentBase, key)
				log.Println("Субабонент ", key, " исключен из расчета, т.к. абонент "+v+" не имеет показаний.")
			}
		}
	}

	// удаляем записи, которые в числе удаляемых
	for _, v := range abonents2Delete {
		if _, ok := abonentBase[v]; ok {
			delete(abonentBase, v)
			log.Println("Абонент ", v, " исключен из расчета, т.к. не имеет показаний.")
		}
	}

}

// распаковка
func mainCreateUnpacking(abonentBase map[string]abonentStruct) {
	xlsUnpacking := excelize.NewFile()
	// Создать новый лист
	xlsUnpackingSheetName := "Sheet1"
	xlsUnpackingSheet := xlsUnpacking.NewSheet(xlsUnpackingSheetName)

	rowNumber := 3 // номер строки для начала заголовка
	// перебор главных абонентов и работа по ним
	for abonent, abonentItem := range abonentBase {
		if abonentItem.fatherid == "" { // отбираем только главных абонентов
			subAbonents := getSubAbonents(abonent, abonentBase)

			subCreateUnpackingSetHeading(rowNumber, xlsUnpackingSheetName, xlsUnpacking)                                   // создаем заголовок
			subCreateUnpackingSetValues(rowNumber, xlsUnpackingSheetName, xlsUnpacking, abonent, subAbonents, abonentBase) // создаем содержательную часть
			rowNumber = rowNumber + 5 + len(subAbonents)
		}
	}

	// Установить активный лист рабочей книги
	xlsUnpacking.SetActiveSheet(xlsUnpackingSheet)
	// Сохранить файл xlsx по данному пути
	err := xlsUnpacking.SaveAs(config.unpackingName)
	check(err)

}

// по номеру абонента возвращаем slice с номерами его субабонентов
func getSubAbonents(abonent string, abonentBase map[string]abonentStruct) []string {
	var subAbonents []string
	for key, value := range abonentBase {
		if value.fatherid == abonent {
			subAbonents = append(subAbonents, key)
		}
	}
	return subAbonents
}

// создаем заголовок таблицы в распаковке
func subCreateUnpackingSetHeading(rowNumber int, xlsUnpackingSheetName string, xlsUnpacking *excelize.File) {
	styleHead, err := xlsUnpacking.NewStyle(`{
		"fill":
		{
			"type":"pattern",
			"color":["#E7E6E6"],
			"pattern":1
		},
		"alignment":
	    {
	        "horizontal": "center",
			"vertical": "center",
	        "wrap_text": true
	    },
		"font":
	    {
	        "bold": true
	    }
	}`)
	check(err)

	styleHeadBoldRight, err := xlsUnpacking.NewStyle(`{
		"alignment":
	    {
	        "horizontal": "right"
	    },
		"font":
	    {
	        "bold": true
	    }
	}`)
	check(err)

	styleHeadBoldDate, err := xlsUnpacking.NewStyle(`{"font":{"bold": true},"custom_number_format": "mmmm yyyy"}`)
	check(err)

	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "B1", "Отчетный период:")
	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "B1", "B1", styleHeadBoldRight)

	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "C1", 25569+(config.currentPeriod.Unix()/86400))
	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "C1", "C1", styleHeadBoldDate)

	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "A"+int2string(rowNumber), "Счетчик")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "B"+int2string(rowNumber), "Описание ТУ")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "C"+int2string(rowNumber), "Объем")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "D"+int2string(rowNumber), "% в общем потреблении")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "E"+int2string(rowNumber), "потери ГП")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "F"+int2string(rowNumber), "потери субабонента")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "G"+int2string(rowNumber), "К выставлению")
	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "A"+int2string(rowNumber), "G"+int2string(rowNumber), styleHead)

	xlsUnpacking.SetColWidth(xlsUnpackingSheetName, "A", "A", 25)
	xlsUnpacking.SetColWidth(xlsUnpackingSheetName, "B", "B", 51.71)
	xlsUnpacking.SetColWidth(xlsUnpackingSheetName, "C", "F", 15)
	xlsUnpacking.SetColWidth(xlsUnpackingSheetName, "D", "D", 14)
	xlsUnpacking.SetColWidth(xlsUnpackingSheetName, "G", "G", 14.43)
}

// создаем содержательную часть
func subCreateUnpackingSetValues(rowNumber int, xlsUnpackingSheetName string, xlsUnpacking *excelize.File, mainAbonent string, subAbonents []string, abonentBase map[string]abonentStruct) {
	// стили
	styleNumber, err := xlsUnpacking.NewStyle(`{"number_format": 4}`)
	check(err)
	styleDesc, err := xlsUnpacking.NewStyle(`{
		"alignment":
	    {
	        "wrap_text": true
	    }
	}`)
	check(err)
	stylePercent, err := xlsUnpacking.NewStyle(`{"number_format": 9}`)
	check(err)

	hoursInCurrentPeriod := hoursInCurrentPeriod() // расчитываем количество часов в отчетном периоде

	firstRow := rowNumber // запоминаем строку с которой начали
	// строка с главным абонентом
	rowNumber++
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "A"+int2string(rowNumber), mainAbonent+" (главный)")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "B"+int2string(rowNumber), abonentBase[mainAbonent].desc)
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "C"+int2string(rowNumber), abonentBase[mainAbonent].costs)
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "D"+int2string(rowNumber), 1)
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "E"+int2string(rowNumber), abonentBase[mainAbonent].gp_losses)
	xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "G"+int2string(rowNumber), "C"+int2string(rowNumber)+"+E"+int2string(rowNumber))

	// строка с субабонентами
	rowNumber++
	for key, value := range subAbonents {
		xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "A"+int2string(rowNumber+key), value)
		xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "B"+int2string(rowNumber+key), abonentBase[value].desc)
		xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "C"+int2string(rowNumber+key), abonentBase[value].costs)
		xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "D"+int2string(rowNumber+key), "C"+int2string(rowNumber+key)+"/$C$"+int2string(rowNumber-1))
		xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "E"+int2string(rowNumber+key), "E"+int2string(firstRow+1)+"*D"+int2string(rowNumber))
		xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "F"+int2string(rowNumber+key), abonentBase[value].inside_losses_ratio*math.Pow(float64(abonentBase[value].costs), 2)/float64(hoursInCurrentPeriod))
		xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "G"+int2string(rowNumber+key), "C"+int2string(rowNumber+key)+"+E"+int2string(rowNumber+key)+"+F"+int2string(rowNumber+key))
	}

	// строка с остатком главного абонента
	rowNumber = rowNumber + len(subAbonents)
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "A"+int2string(rowNumber), mainAbonent+" (остаток)")
	xlsUnpacking.SetCellValue(xlsUnpackingSheetName, "B"+int2string(rowNumber), abonentBase[mainAbonent].desc)
	xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "C"+int2string(rowNumber), "C"+int2string(firstRow+1)+"-SUM(C"+int2string(firstRow+2)+":C"+int2string(rowNumber-1)+")")
	xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "D"+int2string(rowNumber), "C"+int2string(rowNumber)+"/$C$"+int2string(firstRow+1))
	xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "E"+int2string(rowNumber), "E"+int2string(firstRow+1)+"-SUM(E"+int2string(firstRow+2)+":E"+int2string(rowNumber-1)+")")
	xlsUnpacking.SetCellFormula(xlsUnpackingSheetName, "G"+int2string(rowNumber), "C"+int2string(rowNumber)+"+E"+int2string(rowNumber)+"-"+"SUM(F"+int2string(firstRow+2)+":F"+int2string(rowNumber-1)+")")

	// проставляем стили

	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "B"+int2string(firstRow+1), "B"+int2string(rowNumber), styleDesc)
	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "C"+int2string(firstRow+1), "G"+int2string(rowNumber), styleNumber)
	xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "D"+int2string(firstRow+1), "D"+int2string(rowNumber), stylePercent)
	// xlsUnpacking.SetCellStyle(xlsUnpackingSheetName, "E"+int2string(firstRow+1), "G"+int2string(rowNumber), styleNumber)

}

// устанавливаем начало текущего периода, взятого из профиля
func setCurrentPeriod(dateFromFile string) {
	trimedStr := strings.TrimSpace(dateFromFile)
	// 01.03.2021
	time, err := time.Parse("02.01.2006", trimedStr[len(trimedStr)-10:])
	check(err)
	config.currentPeriod = time.AddDate(0, -1, 0)
	log.Println("Установлен период: " + config.currentPeriod.Format("2006.01.02"))
}

// возвращает количество дней в текущем периоде
func hoursInCurrentPeriod() int {
	end := config.currentPeriod.AddDate(0, +1, 0)
	difference := end.Sub(config.currentPeriod).Hours()
	return int(difference)
}

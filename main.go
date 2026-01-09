package main

import (
	"errors"
	"flag"
	"fmt"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Row struct {
	xVal float64
	yVal float64
}

type Config struct {
	fileName        string
	sheetName       string
	newSheetName    string
	categoryColName string
	xColName        string
	yColName        string
	categoryColIdx  int
	xColIdx         int
	yColIdx         int
	categoryColumns map[string][]Row
}

func main() {
	fileNamePtr := flag.String("fileName", "Book2.xlsx", "the name of the file to work on")
	sheetNamePtr := flag.String("sheetName", "dummy data", "the name of the sheet to work on")
	newSheetNamePtr := flag.String("newSheetName", "mias_chart", "the name of the new sheet to put the data/chart")
	categoryColNamePtr := flag.String("categoryColName", "Measurement Name", "the name column that has the different categories")
	xColNamePtr := flag.String("xColName", "Days on Study", "the name of the column with the x axis values")
	yColNamePtr := flag.String("yColName", "Times Upper Reference Value", "the name of the column with the y axis values")
	flag.Parse()

	c := NewConfig(
		*fileNamePtr,
		*sheetNamePtr,
		*newSheetNamePtr,
		*categoryColNamePtr,
		*xColNamePtr,
		*yColNamePtr,
	)

	f, err := excelize.OpenFile(c.fileName)
	if err != nil {
		log.Fatal(err)
	}

	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	rows, err := f.GetRows(c.sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = c.GetColumnIndexes(rows[0])
	if err != nil {
		fmt.Println("failed to GetColumnIndexes")
		log.Fatal(err)
	}

	err = c.MapColumnData(rows[1:])
	if err != nil {
		fmt.Println("failed to MapColumnData")
		log.Fatal(err)
	}

	_, err = f.NewSheet(c.newSheetName)
	if err != nil {
		log.Fatal(err)
	}

	chartSeries := []excelize.ChartSeries{}
	chartSeriesCategories := ""

	colIdx := 1
	for key, value := range c.categoryColumns {
		if colIdx == 1 {
			// set first col
			values := []any{c.xColName}
			for _, item := range value {
				values = append(values, item.xVal)
			}

			cord, err := GetCordinate(colIdx, 1)
			if err != nil {
				log.Fatal(err)
			}

			err = f.SetSheetCol(c.newSheetName, cord, &values)
			if err != nil {
				fmt.Println("failed to add x column data")
				log.Fatal(err)
			}

			cordRange, err := GetCordinateRange(c.newSheetName, colIdx, 2, len(value)+1)
			if err != nil {
				fmt.Println("failed to category range")
				log.Fatal(err)
			}
			chartSeriesCategories = cordRange
		}

		colIdx++

		catColValues := []any{key}
		for _, item := range value {
			catColValues = append(catColValues, item.yVal)
		}

		cord, err := GetCordinate(colIdx, 1)
		if err != nil {
			log.Fatal(err)
		}

		err = f.SetSheetCol(c.newSheetName, cord, &catColValues)
		if err != nil {
			fmt.Printf("failed to add y column data. key: %s\n", key)
			log.Fatal(err)
		}

		nameCord, err := GetStaticSheetCordinate(c.newSheetName, colIdx, 1)
		if err != nil {
			fmt.Println("failed to get nameCord")
			log.Fatal(err)
		}

		chartSeriesValues, err := GetCordinateRange(c.newSheetName, colIdx, 2, len(value)+1)
		if err != nil {
			fmt.Println("failed to get chartSeriesValues")
			log.Fatal(err)
		}

		chartSeries = append(chartSeries, excelize.ChartSeries{
			Name:       nameCord,
			Categories: chartSeriesCategories,
			Values:     chartSeriesValues,
			Line: excelize.ChartLine{
				Smooth: true,
			},
		})
	}

	err = CreateLineChart(f, c.newSheetName, chartSeries)
	if err != nil {
		fmt.Println("failed to add chart")
		log.Fatal(err)
	}

	if err = f.SaveAs(c.fileName); err != nil {
		fmt.Println(err)
	}
}

func NewConfig(fileName string, sheetName string, newSheetName string, categoryColName string, xColName string, yColName string) Config {
	return Config{
		fileName:        fileName,
		sheetName:       sheetName,
		newSheetName:    newSheetName,
		categoryColName: categoryColName,
		xColName:        xColName,
		yColName:        yColName,
		categoryColIdx:  -1,
		xColIdx:         -1,
		yColIdx:         -1,
		categoryColumns: map[string][]Row{},
	}
}

func (c *Config) GetColumnIndexes(rows []string) error {
	for idx, row := range rows {
		switch row {
		case c.categoryColName:
			c.categoryColIdx = idx
		case c.xColName:
			c.xColIdx = idx
		case c.yColName:
			c.yColIdx = idx
		}
	}

	if c.categoryColIdx == -1 || c.xColIdx == -1 || c.yColIdx == -1 {
		return errors.New("missing index")
	}

	return nil
}

func (c *Config) MapColumnData(rows [][]string) error {
	for _, row := range rows {
		xIntVal, err := strconv.ParseFloat(row[c.xColIdx], 64)
		if err != nil {
			return fmt.Errorf("failed to parse cell to float: %s", row[c.xColIdx])
		}

		yIntVal, err := strconv.ParseFloat(row[c.yColIdx], 64)
		if err != nil {
			return fmt.Errorf("failed to parse cell to float: %s", row[c.yColIdx])
		}

		rowData := Row{
			xVal: xIntVal,
			yVal: yIntVal,
		}

		cat := row[c.categoryColIdx]
		_, ok := c.categoryColumns[cat]
		if !ok {
			c.categoryColumns[cat] = []Row{rowData}
		} else {
			c.categoryColumns[cat] = append(c.categoryColumns[cat], rowData)
		}
	}

	return nil
}

var cordinateMap = map[int]string{
	1: "A",
	2: "B",
	3: "C",
	4: "D",
	5: "E",
	6: "F",
	7: "G",
	8: "H",
	9: "I",
}

func GetCordinate(col int, row int) (string, error) {
	letter, ok := cordinateMap[col]
	if !ok {
		return "", fmt.Errorf("no letter found for index %d", col)
	}

	return fmt.Sprintf("%s%d", letter, row), nil
}

func GetStaticSheetCordinate(sheetName string, col int, row int) (string, error) {
	letter, ok := cordinateMap[col]
	if !ok {
		return "", fmt.Errorf("no letter found for index %d", col)
	}

	return fmt.Sprintf("%s!$%s$%d", sheetName, letter, row), nil
}

func GetCordinateRange(sheetName string, col int, start int, end int) (string, error) {
	letter, ok := cordinateMap[col]
	if !ok {
		return "", fmt.Errorf("no letter found for index %d", col)
	}

	return fmt.Sprintf("%s!$%s$%d:$%s$%d", sheetName, letter, start, letter, end), nil
}

func CreateLineChart(f *excelize.File, sheetName string, series []excelize.ChartSeries) error {
	font := excelize.Font{
		Color:  "000000",
		Size:   12,
		Family: "Arial",
	}

	return f.AddChart(sheetName, "E1", &excelize.Chart{
		Type:   excelize.Line,
		Series: series,
		Format: excelize.GraphicOptions{
			OffsetX: 15,
			OffsetY: 10,
			AutoFit: true,
		},
		Legend: excelize.ChartLegend{
			Position: "top",
			Font:     &font,
		},
		XAxis: excelize.ChartAxis{
			Title: []excelize.RichTextRun{
				{
					Text: "xULN",
					Font: &font,
				},
			},
			Font: font,
		},
		YAxis: excelize.ChartAxis{
			Title: []excelize.RichTextRun{
				{
					Text: "Days On Study",
					Font: &font,
				},
			},
			Font:           font,
			MajorGridLines: true,
		},
		// Title: []excelize.RichTextRun{
		// 	{
		// 		Text: "Mia's Line Chart",
		// 		Font: &font,
		// 	},
		// },
		PlotArea: excelize.ChartPlotArea{
			// ShowCatName:     false,
			// ShowLeaderLines: true,
			// ShowPercent:     true,
			// ShowSerName:     true,
			// ShowVal:         true,
		},
	})
}

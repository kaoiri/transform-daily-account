package main

import (
	"encoding/csv"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"github.com/BurntSushi/toml"
)

type Config struct {
	Filename struct {
		Zitem    string `toml:"zitem"`
		Ztime    string `toml:"ztime"`
		Ztotal   string `toml:"ztotal"`
		Ttotal   string `toml:"ttotal"`
		Template string `toml:"template"`
		SheetName string `toml:"sheetname"`
	}
	Exclusion struct {
		Tables               []int16  `toml:"tables"`
		KeywordsFromEatin    []string `toml:"keywords_from_eatin"`
		KeywordsFromCatering []string `toml:"keywords_from_catering"`
	}
}

func main() {
	exe, err := os.Executable()
	if err != nil {
		panic(err)
	}
	path := filepath.Dir(exe)
	/*
	path, err := os.Getwd()
	if err != nil {
		panic(err)
	}
	*/

	var config Config
	_, err = toml.DecodeFile(filepath.Join(path, "config.toml"), &config)
	if err != nil {
		panic(err)
	}

	var ttotals Ttotals
	ttotals.CollectFromCSV(filepath.Join(path, config.Filename.Ttotal))

	var ztotals Ztotals
	ztotals.CollectFromCSV(filepath.Join(path, config.Filename.Ztotal))

	var ztimes Ztimes
	ztimes.CollectFromCSV(filepath.Join(path, config.Filename.Ztime))

	var zitems Zitems
	zitems.CollectFromCSV(filepath.Join(path, config.Filename.Zitem))

	output, err := excelize.OpenFile(filepath.Join(path, config.Filename.Template))
	if err != nil {
		panic(err)
	}

	date, err := getDate(filepath.Join(path, config.Filename.Ttotal))
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "E1", date.Display)

	zt, err := ztotals.Get(1)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "A3", zt.GuestCount)

	orders := CountLunchOrder(ttotals, config.Exclusion.Tables)
	output.SetCellValue(config.Filename.SheetName, "A8", orders)
	output.SetCellValue(config.Filename.SheetName, "K2", orders)
	orders = CountDinnerOrder(ttotals, config.Exclusion.Tables)
	output.SetCellValue(config.Filename.SheetName, "A9", orders)
	output.SetCellValue(config.Filename.SheetName, "M2", orders)

	zt, err = ztotals.Get(112)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D18", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F18", zt.Total)

	zt, err = ztotals.Get(113)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D19", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F19", zt.Total)

	zt, err = ztotals.Get(114)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D20", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F20", zt.Total)

	zt, err = ztotals.Get(116)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D21", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F21", zt.Total)

	zt, err = ztotals.Get(48)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D22", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "X2", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F22", zt.Total)
	output.SetCellValue(config.Filename.SheetName, "W2", zt.Total)
	check := 0 - zt.Total

	zt, err = ztotals.Get(99)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D23", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F23", zt.Total)

	zt, err = ztotals.Get(84)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D24", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F24", zt.Total)

	zt, err = ztotals.Get(134)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "D25", zt.OrderCount)
	output.SetCellValue(config.Filename.SheetName, "F25", zt.Total)

	zt, err = ztotals.Get(151)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F26", zt.Total)

	zt, err = ztotals.Get(78)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F27", zt.Total)

	zt, err = ztotals.Get(300)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F28", zt.Total)

	zt, err = ztotals.Get(301)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F29", zt.Total)

	zt, err = ztotals.Get(302)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F30", zt.Total)

	zt, err = ztotals.Get(303)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F31", zt.Total)

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F32", zt.Total)

	zt, err = ztotals.Get(305)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F33", zt.Total)

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F34", zt.Total)

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F35", zt.Total)

	zt, err = ztotals.Get(46)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F3", zt.Total)
	output.SetCellValue(config.Filename.SheetName, "H2", zt.Total)

	zt, err = ztotals.Get(190)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F4", zt.Total)
	output.SetCellValue(config.Filename.SheetName, "I2", zt.Total)

	zt, err = ztotals.Get(79)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F6", zt.Total)
	check += zt.Total

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
	temp := TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = int32(math.Floor(float64(temp) / 13.5))
	temp = zt.Total - temp
	output.SetCellValue(config.Filename.SheetName, "F7", temp)
	f7 := temp
	check += temp

	temp = TotalLunchEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	output.SetCellValue(config.Filename.SheetName, "F8", temp)
	output.SetCellValue(config.Filename.SheetName, "J2", temp)
	check += temp

	temp = TotalDinnerEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	output.SetCellValue(config.Filename.SheetName, "F9", temp)
	output.SetCellValue(config.Filename.SheetName, "L2", temp)
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"宴会"})) / 1.1))
	output.SetCellValue(config.Filename.SheetName, "F10", temp)
	f10 := temp
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"T テイク宴会", "Ｔ テイク宴会", "Tテイク宴会", "Ｔテイク宴会"})) / 1.08))
	output.SetCellValue(config.Filename.SheetName, "F11", temp)
	output.SetCellValue(config.Filename.SheetName, "N2", f10 + temp)
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"法事"})) / 1.1))
	output.SetCellValue(config.Filename.SheetName, "F12", temp)
	f12 := temp
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"T テイク法事", "Ｔ テイク法事", "Tテイク法事", "Ｔテイク法事"})) / 1.08))
	output.SetCellValue(config.Filename.SheetName, "F13", temp)
	output.SetCellValue(config.Filename.SheetName, "P2", f12 + temp)
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"葬儀"})) / 1.1))
	output.SetCellValue(config.Filename.SheetName, "F14", temp)
	f14 := temp
	check += temp

	temp = int32(math.Floor(float64(TotalWithKeywords(zitems, []string{"T テイク葬儀", "Ｔ テイク葬儀", "Tテイク葬儀", "Ｔテイク葬儀"})) / 1.08))
	output.SetCellValue(config.Filename.SheetName, "F15", temp)
	output.SetCellValue(config.Filename.SheetName, "R2", f14 + temp)
	check += temp

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
	output.SetCellValue(config.Filename.SheetName, "F16", zt.Total)
	f16 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "T2", zt.Total)
	check += zt.Total

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
	temp = TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = zt.Total - temp - f7
	output.SetCellValue(config.Filename.SheetName, "F17", temp)
	output.SetCellValue(config.Filename.SheetName, "U2", temp)
	output.SetCellValue(config.Filename.SheetName, "V2", f16+temp)
	check += temp

	output.SetCellValue(config.Filename.SheetName, "F5", check)

	storeName, err := output.GetCellValue(config.Filename.SheetName, "C1")
	if err != nil {
		panic(err)
	}

	output.UpdateLinkedValue();

	if err := output.SaveAs(filepath.Join(path, fmt.Sprintf("%s %s 日計.xlsx", date.Raw, storeName))); err != nil {
		panic(err)
	}
}

type Date struct {
	Raw     string
	Display string
}

func getDate(path string) (Date, error) {
	var date Date

	file, err := os.Open(path)
	if err != nil {
		return date, err
	}
	defer file.Close()

	reader := csv.NewReader(file)

	var line []string
	line, err = reader.Read()
	if err != nil {
		return date, err
	}

	date.Raw = strings.TrimSpace(line[3])
	res, err := formatDate(date.Raw)
	if err != nil {
		return date, err
	}
	date.Display = res

	return date, nil
}

func formatDate(s string) (string, error) {
	wdays := []string{"日", "月", "火", "水", "木", "金", "土"}
	time, err := time.Parse("20060102", s)
	if err != nil {
		return "", err
	}

	return fmt.Sprintf("%s(%s)", time.Format("2006年01月02日"), wdays[time.Weekday()]), nil
}

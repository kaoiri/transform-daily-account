package main

import (
	"encoding/csv"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/BurntSushi/toml"
	"github.com/xuri/excelize/v2"
)

type Config struct {
	Filename struct {
		Zitem     string `toml:"zitem"`
		Ztime     string `toml:"ztime"`
		Ztotal    string `toml:"ztotal"`
		Ttotal    string `toml:"ttotal"`
		Template  string `toml:"template"`
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
    a3 := zt.GuestCount;
	output.SetCellValue(config.Filename.SheetName, "A3", a3)

	orders := CountLunchOrder(ttotals, config.Exclusion.Tables)
    a8 := orders
	output.SetCellValue(config.Filename.SheetName, "A8", a8)
	output.SetCellValue(config.Filename.SheetName, "K2", a8)
	orders = CountDinnerOrder(ttotals, config.Exclusion.Tables)
    a9 := orders
	output.SetCellValue(config.Filename.SheetName, "A9", a9)
	output.SetCellValue(config.Filename.SheetName, "M2", a9)

	zt, err = ztotals.Get(112)
	if err != nil {
		panic(err)
	}
    d18 := zt.OrderCount
    f18 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D18", d18)
	output.SetCellValue(config.Filename.SheetName, "F18", f18)

	zt, err = ztotals.Get(113)
	if err != nil {
		panic(err)
	}
    d19 := zt.OrderCount
    f19 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D19", d19)
	output.SetCellValue(config.Filename.SheetName, "F19", f19)

	zt, err = ztotals.Get(114)
	if err != nil {
		panic(err)
	}
    d20 := zt.OrderCount
    f20 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D20", d20)
	output.SetCellValue(config.Filename.SheetName, "F20", f20)

	zt, err = ztotals.Get(116)
	if err != nil {
		panic(err)
	}
    d21 := zt.OrderCount
    f21 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D21", d21)
	output.SetCellValue(config.Filename.SheetName, "F21", f21)

	zt, err = ztotals.Get(48)
	if err != nil {
		panic(err)
	}
    d22 := zt.OrderCount
    f22 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D22", d22)
	output.SetCellValue(config.Filename.SheetName, "F22", f22)

	zt, err = ztotals.Get(99)
	if err != nil {
		panic(err)
	}
    d23 := zt.OrderCount
    f23 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D23", d23)
	output.SetCellValue(config.Filename.SheetName, "F23", f23)

	zt, err = ztotals.Get(84)
	if err != nil {
		panic(err)
	}
    d24 := zt.OrderCount
    f24 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D24", d24)
	output.SetCellValue(config.Filename.SheetName, "F24", f24)

	zt, err = ztotals.Get(134)
	if err != nil {
		panic(err)
	}
    d25 := zt.OrderCount
    f25 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "D25", d25)
	output.SetCellValue(config.Filename.SheetName, "F25", f25)

	zt, err = ztotals.Get(151)
	if err != nil {
		panic(err)
	}
    f26 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F26", f26)

	zt, err = ztotals.Get(78)
	if err != nil {
		panic(err)
	}
    f27 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F27", f27)

	zt, err = ztotals.Get(300)
	if err != nil {
		panic(err)
	}
    f28 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F28", f28)

	zt, err = ztotals.Get(301)
	if err != nil {
		panic(err)
	}
    f29 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F29", f29)

	zt, err = ztotals.Get(302)
	if err != nil {
		panic(err)
	}
    f30 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F30", f30)

	zt, err = ztotals.Get(303)
	if err != nil {
		panic(err)
	}
    f31 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F31", f31)

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
    f32 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F32", f32)

	zt, err = ztotals.Get(305)
	if err != nil {
		panic(err)
	}
    f33 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F33", f33)

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
    f34 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F34", f34)

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
    f35 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F35", f35)

	zt, err = ztotals.Get(46)
	if err != nil {
		panic(err)
	}
    f3 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F3", f3)

	zt, err = ztotals.Get(190)
	if err != nil {
		panic(err)
	}
    f4 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F4", f4)

	zt, err = ztotals.Get(79)
	if err != nil {
		panic(err)
	}
    f6 := zt.Total
	output.SetCellValue(config.Filename.SheetName, "F6", f6)

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
	temp := TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = int32(math.Round(float64(temp) / 13.5))
    f7 := zt.Total - temp
	output.SetCellValue(config.Filename.SheetName, "F7", f7)

	temp = TotalLunchEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	temp = int32(math.Round(float64(temp) * 1.1))
    f8 := temp
	output.SetCellValue(config.Filename.SheetName, "F8", f8)

	temp = TotalDinnerEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	temp = int32(math.Round(float64(temp) * 1.1))
    // f9 := temp
    f9 := f28 + f29 - f8
	output.SetCellValue(config.Filename.SheetName, "F9", f9)

    f10 := TotalWithKeywords(zitems, []string{"宴会"})
	output.SetCellValue(config.Filename.SheetName, "F10", f10)

    f11 := TotalWithKeywords(zitems, []string{"T テイク宴会", "Ｔ テイク宴会", "Tテイク宴会", "Ｔテイク宴会"})
	output.SetCellValue(config.Filename.SheetName, "F11", f11)

    f12 := TotalWithKeywords(zitems, []string{"法事"})
	output.SetCellValue(config.Filename.SheetName, "F12", f12)

    f13 := TotalWithKeywords(zitems, []string{"T テイク法事", "Ｔ テイク法事", "Tテイク法事", "Ｔテイク法事"})
	output.SetCellValue(config.Filename.SheetName, "F13", f13)

    f14 := TotalWithKeywords(zitems, []string{"葬儀"})
	output.SetCellValue(config.Filename.SheetName, "F14", f14)

    f15 := TotalWithKeywords(zitems, []string{"T テイク葬儀", "Ｔ テイク葬儀", "Tテイク葬儀", "Ｔテイク葬儀"})
	output.SetCellValue(config.Filename.SheetName, "F15", f15)

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
    // f16 := int32(math.Round(float64(zt.Total) * 1.08))
    f16 := f32 + f33
	output.SetCellValue(config.Filename.SheetName, "F16", f16)

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
	temp = TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = zt.Total - temp - f7
	temp = int32(math.Round(float64(temp) * 1.08))
    f17 := temp
	output.SetCellValue(config.Filename.SheetName, "F17", f17)

    check := 0 - f22 + f8 + f9 + f10 + f11 + f12 + f13 + f14 + f15 + f16 + f17

	output.SetCellValue(config.Filename.SheetName, "F5", check)

	storeName, err := output.GetCellValue(config.Filename.SheetName, "C1")
	if err != nil {
		panic(err)
	}

	output.UpdateLinkedValue()

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

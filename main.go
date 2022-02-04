package main

import (
	"encoding/csv"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/BurntSushi/toml"
)

type Config struct {
	Filename struct {
		Zitem    string `toml:"zitem"`
		Ztime    string `toml:"ztime"`
		Ztotal   string `toml:"ztotal"`
		Ttotal   string `toml:"ttotal"`
		Template string `toml:"template"`
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
	output.SetCellValue("Sheet1", "E1", date.Display)

	zt, err := ztotals.Get(1)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "A3", zt.GuestCount)

	orders := CountLunchOrder(ttotals, config.Exclusion.Tables)
	output.SetCellValue("Sheet1", "A8", orders)
	output.SetCellValue("Sheet1", "K2", orders)
	orders = CountDinnerOrder(ttotals, config.Exclusion.Tables)
	output.SetCellValue("Sheet1", "A9", orders)
	output.SetCellValue("Sheet1", "M2", orders)

	zt, err = ztotals.Get(112)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D18", zt.OrderCount)
	output.SetCellValue("Sheet1", "F18", zt.Total)

	zt, err = ztotals.Get(113)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D19", zt.OrderCount)
	output.SetCellValue("Sheet1", "F19", zt.Total)

	zt, err = ztotals.Get(114)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D20", zt.OrderCount)
	output.SetCellValue("Sheet1", "F20", zt.Total)

	zt, err = ztotals.Get(116)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D21", zt.OrderCount)
	output.SetCellValue("Sheet1", "F21", zt.Total)

	zt, err = ztotals.Get(48)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D22", zt.OrderCount)
	output.SetCellValue("Sheet1", "X2", zt.OrderCount)
	output.SetCellValue("Sheet1", "F22", zt.Total)
	output.SetCellValue("Sheet1", "W2", zt.Total)
	check := 0 - zt.Total

	zt, err = ztotals.Get(99)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D23", zt.OrderCount)
	output.SetCellValue("Sheet1", "F23", zt.Total)

	zt, err = ztotals.Get(84)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D24", zt.OrderCount)
	output.SetCellValue("Sheet1", "F24", zt.Total)

	zt, err = ztotals.Get(134)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "D25", zt.OrderCount)
	output.SetCellValue("Sheet1", "F25", zt.Total)

	zt, err = ztotals.Get(151)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F26", zt.Total)

	zt, err = ztotals.Get(78)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F27", zt.Total)

	zt, err = ztotals.Get(300)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F28", zt.Total)

	zt, err = ztotals.Get(301)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F29", zt.Total)

	zt, err = ztotals.Get(302)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F30", zt.Total)

	zt, err = ztotals.Get(303)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F31", zt.Total)

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F32", zt.Total)

	zt, err = ztotals.Get(305)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F33", zt.Total)

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F34", zt.Total)

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F35", zt.Total)

	zt, err = ztotals.Get(46)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F3", zt.Total)
	output.SetCellValue("Sheet1", "H2", zt.Total)

	zt, err = ztotals.Get(190)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F4", zt.Total)
	output.SetCellValue("Sheet1", "I2", zt.Total)

	zt, err = ztotals.Get(79)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F6", zt.Total)
	check += zt.Total

	zt, err = ztotals.Get(307)
	if err != nil {
		panic(err)
	}
	temp := TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = int32(math.Trunc(float64(temp) / 13.5))
	fmt.Print(temp)
	temp = zt.Total - temp
	fmt.Print(temp)
	output.SetCellValue("Sheet1", "F7", temp)
	f7 := temp
	check += temp

	temp = TotalLunchEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	output.SetCellValue("Sheet1", "F8", temp)
	output.SetCellValue("Sheet1", "J2", temp)
	check += temp

	temp = TotalDinnerEatIn(zitems, config.Exclusion.KeywordsFromEatin)
	output.SetCellValue("Sheet1", "F9", temp)
	output.SetCellValue("Sheet1", "L2", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"宴会"})
	output.SetCellValue("Sheet1", "F10", temp)
	output.SetCellValue("Sheet1", "N2", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク宴会", "Ｔ テイク宴会", "Tテイク宴会", "Ｔテイク宴会"})
	output.SetCellValue("Sheet1", "F11", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"法事"})
	output.SetCellValue("Sheet1", "F12", temp)
	output.SetCellValue("Sheet1", "P2", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク法事", "Ｔ テイク法事", "Tテイク法事", "Ｔテイク法事"})
	output.SetCellValue("Sheet1", "F13", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"葬儀"})
	output.SetCellValue("Sheet1", "F14", temp)
	output.SetCellValue("Sheet1", "R2", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク葬儀", "Ｔ テイク葬儀", "Tテイク葬儀", "Ｔテイク葬儀"})
	output.SetCellValue("Sheet1", "F15", temp)
	check += temp

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F16", zt.Total)
	f16 := zt.Total
	output.SetCellValue("Sheet1", "T2", zt.Total)
	check += zt.Total

	zt, err = ztotals.Get(306)
	if err != nil {
		panic(err)
	}
	temp = TotalWithKeywords(zitems, config.Exclusion.KeywordsFromCatering)
	temp = zt.Total - temp - f7
	output.SetCellValue("Sheet1", "F17", temp)
	output.SetCellValue("Sheet1", "U2", temp)
	output.SetCellValue("Sheet1", "V2", f16+temp)
	check += temp

	output.SetCellValue("Sheet1", "F5", check)

	storeName := output.GetCellValue("Sheet1", "C1")
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

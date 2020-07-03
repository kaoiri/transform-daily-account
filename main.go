package main

import (
	"encoding/csv"
	"fmt"
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
		Output   string `toml:"output"`
	}
	Exclusion struct {
		Tables   []int16  `toml:"tables"`
		Keywords []string `toml:"keywords"`
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

	check := int32(0)

	date, err := getDate(filepath.Join(path, config.Filename.Ttotal))
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "E1", date)

	zt, err := ztotals.Get(1)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "A3", zt.GuestCount)

	output.SetCellValue("Sheet1", "A8", CountLunchOrder(ttotals, config.Exclusion.Tables))
	output.SetCellValue("Sheet1", "A9", CountDinnerOrder(ttotals, config.Exclusion.Tables))

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
	output.SetCellValue("Sheet1", "F22", zt.Total)
	check -= zt.Total

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

	zt, err = ztotals.Get(190)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F4", zt.Total)

	zt, err = ztotals.Get(79)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F6", zt.Total)
	check += zt.Total

	zt, err = ztotals.Get(80)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F7", zt.Total)

	temp := TotalLunchEatIn(zitems, config.Exclusion.Keywords)
	output.SetCellValue("Sheet1", "F8", temp)
	check += temp

	temp = TotalDinnerEatIn(zitems, config.Exclusion.Keywords)
	output.SetCellValue("Sheet1", "F9", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"宴会"})
	output.SetCellValue("Sheet1", "F10", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク宴会", "Ｔ テイク宴会", "Tテイク宴会", "Ｔテイク宴会"})
	output.SetCellValue("Sheet1", "F11", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"法事"})
	output.SetCellValue("Sheet1", "F12", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク法事", "Ｔ テイク法事", "Tテイク法事", "Ｔテイク法事"})
	output.SetCellValue("Sheet1", "F13", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"葬儀"})
	output.SetCellValue("Sheet1", "F14", temp)
	check += temp

	temp = TotalWithKeywords(zitems, []string{"T テイク葬儀", "Ｔ テイク葬儀", "Tテイク葬儀", "Ｔテイク葬儀"})
	output.SetCellValue("Sheet1", "F15", temp)
	check += temp

	zt, err = ztotals.Get(304)
	if err != nil {
		panic(err)
	}
	output.SetCellValue("Sheet1", "F16", zt.Total)
	check += zt.Total

	output.SetCellValue("Sheet1", "F5", check)

	if err := output.SaveAs(filepath.Join(path, config.Filename.Output)); err != nil {
		panic(err)
	}
}

func getDate(path string) (string, error) {
	file, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer file.Close()

	reader := csv.NewReader(file)

	var line []string
	line, err = reader.Read()
	if err != nil {
		return "", err
	}

	res, err := formatDate(strings.TrimSpace(line[3]))
	if err != nil {
		return "", err
	}

	return res, nil
}

func formatDate(s string) (string, error) {
	wdays := []string{"日", "月", "火", "水", "木", "金", "土"}
	time, err := time.Parse("20060102", s)
	if err != nil {
		return "", err
	}

	return fmt.Sprintf("%s(%s)", time.Format("2006年01月02日"), wdays[time.Weekday()]), nil
}

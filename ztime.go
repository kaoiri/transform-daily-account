package main

import (
	"encoding/csv"
	"errors"
	"os"
	"strconv"
	"strings"

	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

type Ztime struct {
	ID          int16
	Start       int16
	End         int16
	EatIn       SubTotals
	TakeOut     SubTotals
	Restaurant  SubTotals
	Floor       SubTotals
	SalesAmount int32
}

type SubTotals struct {
	PriceInternal int32
	CountInternal int16
	PriceForeign  int32
	CountForeign  int16
}

func (z *Ztime) FromCSV(s []string) error {
	if len(s) < 32 {
		return errors.New("Invalid format.")
	}

	// A = 0
	i, err := strconv.ParseInt(strings.TrimSpace(s[0]), 10, 16)
	if err != nil {
		return err
	}
	z.ID = int16(i)

	// B = 1
	i, err = strconv.ParseInt(strings.TrimSpace(s[1]), 10, 16)
	if err != nil {
		return err
	}
	z.Start = int16(i)

	// C = 2
	i, err = strconv.ParseInt(strings.TrimSpace(s[2]), 10, 16)
	if err != nil {
		return err
	}
	z.End = int16(i)

	// D = 3
	i, err = strconv.ParseInt(strings.TrimSpace(s[3]), 10, 32)
	if err != nil {
		return err
	}

	// E = 4
	j, err := strconv.ParseInt(strings.TrimSpace(s[4]), 10, 16)
	if err != nil {
		return err
	}

	// F = 5
	k, err := strconv.ParseInt(strings.TrimSpace(s[5]), 10, 32)
	if err != nil {
		return err
	}

	// G = 6
	l, err := strconv.ParseInt(strings.TrimSpace(s[6]), 10, 16)
	if err != nil {
		return err
	}

	z.EatIn = SubTotals{int32(i), int16(j), int32(k), int16(l)}

	// H = 7
	i, err = strconv.ParseInt(strings.TrimSpace(s[7]), 10, 32)
	if err != nil {
		return err
	}

	// I = 8
	j, err = strconv.ParseInt(strings.TrimSpace(s[8]), 10, 16)
	if err != nil {
		return err
	}

	// J = 9
	k, err = strconv.ParseInt(strings.TrimSpace(s[9]), 10, 32)
	if err != nil {
		return err
	}

	// K = 10
	l, err = strconv.ParseInt(strings.TrimSpace(s[10]), 10, 16)
	if err != nil {
		return err
	}

	z.TakeOut = SubTotals{int32(i), int16(j), int32(k), int16(l)}

	// L = 11
	i, err = strconv.ParseInt(s[11], 10, 32)
	if err != nil {
		return err
	}

	// M = 12
	j, err = strconv.ParseInt(strings.TrimSpace(s[12]), 10, 16)
	if err != nil {
		return err
	}

	// N = 13
	k, err = strconv.ParseInt(strings.TrimSpace(s[13]), 10, 32)
	if err != nil {
		return err
	}

	// O = 14
	l, err = strconv.ParseInt(strings.TrimSpace(s[14]), 10, 16)
	if err != nil {
		return err
	}

	z.Restaurant = SubTotals{int32(i), int16(j), int32(k), int16(l)}

	// P = 15
	i, err = strconv.ParseInt(strings.TrimSpace(s[15]), 10, 32)
	if err != nil {
		return err
	}

	// Q = 16
	j, err = strconv.ParseInt(strings.TrimSpace(s[16]), 10, 16)
	if err != nil {
		return err
	}

	// R = 17
	k, err = strconv.ParseInt(strings.TrimSpace(s[17]), 10, 32)
	if err != nil {
		return err
	}

	// S = 18
	l, err = strconv.ParseInt(strings.TrimSpace(s[18]), 10, 16)
	if err != nil {
		return err
	}

	z.Floor = SubTotals{int32(i), int16(j), int32(k), int16(l)}

	// AF = 31
	i, err = strconv.ParseInt(strings.TrimSpace(s[31]), 10, 16)
	if err != nil {
		return err
	}
	z.SalesAmount = int32(i)

	return nil
}

type Ztimes []Ztime

func (ts *Ztimes) CollectFromCSV(path string) error {
	file, err := os.Open(path)
	if err != nil {
		return err
	}
	defer file.Close()

	utf8 := transform.NewReader(file, japanese.ShiftJIS.NewDecoder())
	reader := csv.NewReader(utf8)
	reader.FieldsPerRecord = -1
	var line []string
	for {
		line, err = reader.Read()
		if err != nil {
			break
		}

		t := Ztime{}
		err := t.FromCSV(line)
		if err != nil {
			continue
		}

		*ts = append(*ts, t)
	}

	return nil
}

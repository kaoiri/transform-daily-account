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

type Zitem struct {
	ID          int16
	Name        string
	Section     Entry
	Parent      Entry
	Unit        int32
	TakeOutType Entry
	GP          Entry
	DP          Entry
	AMCount     int16
	PMCount     int16
}

type Entry struct {
	ID   int16
	Name string
}

func (z *Zitem) FromCSV(s []string) error {
	if len(s) < 27 {
		return errors.New("Invalid format.")
	}

	// A = 0
	i, err := strconv.ParseInt(strings.TrimSpace(s[0]), 10, 16)
	if err != nil {
		return err
	}
	z.ID = int16(i)

	// C = 2
	z.Name = strings.TrimSpace(s[2])

	// D = 3
	i, err = strconv.ParseInt(strings.TrimSpace(s[3]), 10, 16)
	if err != nil {
		return err
	}

	// E = 4
	z.Section = Entry{int16(i), strings.TrimSpace(s[4])}

	// F = 5
	i, err = strconv.ParseInt(strings.TrimSpace(s[5]), 10, 16)
	if err != nil {
		return err
	}

	// G = 6
	z.Parent = Entry{int16(i), strings.TrimSpace(s[6])}

	// H = 7
	i, err = strconv.ParseInt(strings.TrimSpace(s[7]), 10, 32)
	if err != nil {
		return err
	}
	z.Unit = int32(i)

	// O = 14
	i, err = strconv.ParseInt(strings.TrimSpace(s[14]), 10, 16)
	if err != nil {
		return err
	}

	// P = 15
	z.TakeOutType = Entry{int16(i), strings.TrimSpace(s[15])}

	// R = 17
	i, err = strconv.ParseInt(strings.TrimSpace(s[17]), 10, 16)
	if err != nil {
		return err
	}

	// S = 18
	z.GP = Entry{int16(i), strings.TrimSpace(s[18])}

	// T = 19
	i, err = strconv.ParseInt(strings.TrimSpace(s[19]), 10, 16)
	if err != nil {
		return err
	}

	// U = 20
	z.DP = Entry{int16(i), strings.TrimSpace(s[20])}

	// X = 23
	i, err = strconv.ParseInt(strings.TrimSpace(s[23]), 10, 16)
	if err != nil {
		return err
	}
	z.AMCount = int16(i)

	// AA = 26
	i, err = strconv.ParseInt(strings.TrimSpace(s[26]), 10, 16)
	if err != nil {
		return err
	}
	z.PMCount = int16(i)

	return nil
}

type Zitems []Zitem

func (ts *Zitems) CollectFromCSV(path string) error {
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

		t := Zitem{}
		err := t.FromCSV(line)
		if err != nil {
			continue
		}

		*ts = append(*ts, t)
	}

	return nil
}

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

type Ztotal struct {
	ID         int16
	Name       string
	Total      int32
	ItemCount  int16
	OrderCount int16
	SheetCount int16
	GroupCount int16
	GuestCount int16
}

func (z *Ztotal) FromCSV(s []string) error {
	if len(s) < 14 {
		return errors.New("Invalid format.")
	}

	// A = 0
	i, err := strconv.ParseInt(strings.TrimSpace(s[0]), 10, 16)
	if err != nil {
		return err
	}
	z.ID = int16(i)

	// B = 1
	z.Name = strings.TrimSpace(s[1])

	// D = 3
	i, err = strconv.ParseInt(strings.TrimSpace(s[3]), 10, 32)
	if err != nil {
		i = 0
	}
	z.Total = int32(i)

	// F = 5
	i, err = strconv.ParseInt(strings.TrimSpace(s[5]), 10, 16)
	if err != nil {
		i = 0
	}
	z.ItemCount = int16(i)

	// H = 7
	i, err = strconv.ParseInt(strings.TrimSpace(s[7]), 10, 16)
	if err != nil {
		i = 0
	}
	z.OrderCount = int16(i)

	// J = 9
	i, err = strconv.ParseInt(strings.TrimSpace(s[9]), 10, 16)
	if err != nil {
		i = 0
	}
	z.SheetCount = int16(i)

	// L = 11
	i, err = strconv.ParseInt(strings.TrimSpace(s[11]), 10, 16)
	if err != nil {
		i = 0
	}
	z.GroupCount = int16(i)

	// N = 13
	i, err = strconv.ParseInt(strings.TrimSpace(s[13]), 10, 16)
	if err != nil {
		i = 0
	}
	z.GuestCount = int16(i)

	return nil
}

type Ztotals []Ztotal

func (ts *Ztotals) CollectFromCSV(path string) error {
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

		t := Ztotal{}
		err := t.FromCSV(line)
		if err != nil {
			continue
		}

		*ts = append(*ts, t)
	}

	return nil
}

func (ts Ztotals) Get(id int16) (*Ztotal, error) {
	for _, t := range ts {
		if t.ID == id {
			return &t, nil
		}
	}
	return nil, errors.New("Failed to get Ztotal by ID.")
}

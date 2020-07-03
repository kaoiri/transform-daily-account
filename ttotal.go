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

type Ttotal struct {
	ID    int16
	No    int16
	Time  int32
	Count int16
}

func (t *Ttotal) FromCSV(s []string) error {
	if len(s) < 19 {
		return errors.New("Invalid format.")
	}

	// A = 0
	i, err := strconv.ParseInt(strings.TrimSpace(s[0]), 10, 16)
	if err != nil {
		return err
	}
	t.ID = int16(i)

	// J = 9
	i, err = strconv.ParseInt(strings.TrimSpace(s[9]), 10, 16)
	if err != nil {
		i = 0
	}
	t.No = int16(i)

	// M = 12
	timeRaw := strings.TrimSpace(s[12])
	i = 0
	if len(timeRaw) > 9 {
		i, err = strconv.ParseInt(timeRaw[8:], 10, 32)
	}
	if err != nil {
		i = 0
	}
	t.Time = int32(i)

	// S = 18
	i, err = strconv.ParseInt(strings.TrimSpace(s[18]), 10, 16)
	if err != nil {
		i = 0
	}
	t.Count = int16(i)

	return nil
}

type Ttotals []Ttotal

func (ts *Ttotals) CollectFromCSV(path string) error {
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

		t := Ttotal{}
		err := t.FromCSV(line)
		if err != nil {
			continue
		}

		*ts = append(*ts, t)
	}

	return nil
}

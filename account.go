package main

import (
	"strings"
)

func contains(list []int16, target int16) bool {
	for _, v := range list {
		if v == target {
			return true
		}
	}
	return false
}

func wordIncluded(list []string, target string) bool {
	for _, s := range list {
		if strings.Contains(target, s) {
			return true
		}
	}
	return false
}

func wordMatches(list []string, target string) bool {
	for _, s := range list {
		if s == target {
			return true
		}
	}
	return false
}

func CountLunchOrder(ts Ttotals, tables []int16) int16 {
	var res int16
	res = 0
	for _, t := range ts {
		if !contains(tables, t.No) {
			continue
		}

		if t.Time > 150000 {
			continue
		}

		res += t.Count
	}
	return res
}

func CountDinnerOrder(ts Ttotals, tables []int16) int16 {
	var res int16
	res = 0
	for _, t := range ts {
		if !contains(tables, t.No) {
			continue
		}

		if t.Time <= 150000 {
			continue
		}

		res += t.Count
	}
	return res
}

func TotalLunchEatIn(zs Zitems, words []string) int32 {
	var res int32
	res = 0
	for _, z := range zs {
		if wordIncluded(words, z.Name) {
			continue
		}

		if wordIncluded(words, z.GP.Name) {
			continue
		}

		if wordIncluded(words, z.DP.Name) {
			continue
		}

		res += z.Unit * int32(z.AMCount)
	}
	return res
}

func TotalDinnerEatIn(zs Zitems, words []string) int32 {
	var res int32
	res = 0
	for _, z := range zs {
		if wordIncluded(words, z.Name) {
			continue
		}

		if wordIncluded(words, z.GP.Name) {
			continue
		}

		if wordIncluded(words, z.DP.Name) {
			continue
		}

		res += z.Unit * int32(z.PMCount)
	}
	return res
}

func TotalWithKeywords(zs Zitems, keywords []string) int32 {
	var res int32
	res = 0
	for _, z := range zs {
		if wordMatches(keywords, z.Name) {
			res += z.Unit * int32(z.AMCount+z.PMCount)
		}
	}
	return res
}

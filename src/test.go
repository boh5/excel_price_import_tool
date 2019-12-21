package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	rFPath := "excels/自控材料价格明细表191015.xlsx"
	wFPath := "excels/龙泉项目SNCR.xlsx"
	_, err := excelize.OpenFile(rFPath)
	if err != nil {
		fmt.Println(err)
		return
	}
	wF, err := excelize.OpenFile(wFPath)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows := wF.GetRows("Sheet1")
	fmt.Println(len(rows))
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
	fmt.Println("rune A: ",rune("A"[0]))
	fmt.Println(fmt.Sprint("A", 1))
}

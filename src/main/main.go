package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

const BaseColumnRune rune = 'A'

func main() {
	rFPath := "excels/自控材料价格明细表191015.xlsx"
	wFPath := "excels/龙泉项目SNCR.xlsx"
	rF, err := excelize.OpenFile(rFPath)
	if err != nil {
		fmt.Println(err)
		return
	}
	wF, err := excelize.OpenFile(wFPath)
	if err != nil {
		fmt.Println(err)
		return
	}
	do(rF, wF, "C", "G", "C", "G")
	wF.SaveAs("excels/龙泉项目SNCR--test.xlsx")
}

func do(rF, wF *excelize.File, rFRColumn, rFPColumn, wFRColumn, wFPColumn string) {
	totalCount := 0

	emptyCellStyle, _ := wF.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"fill":{"type":"pattern","color":["#FF5733"],"pattern":1}}`)
	priceCellStyle, _ := wF.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"number_format":2}`)

	for _, wFSheetName := range wF.GetSheetMap() {
		wFSheetMaxRow := len(wF.GetRows(wFSheetName))
		for i := 1; i <= wFSheetMaxRow; i++ {
			searched := false
			wFNowPrice := wF.GetCellValue(wFSheetName, fmt.Sprint(wFPColumn, i))
			if wFNowPrice != "" { // 如果当前价格cell不为空，跳过该行
				continue
			}
		rFLoop:
			for _, rFSheetName := range rF.GetSheetMap() {
				rFSheetRows, err := rF.Rows(rFSheetName)
				if err != nil {
					fmt.Println(err)
				}
				for rFSheetRows.Next() {
					rFSheetRow := rFSheetRows.Columns()
					wFNowRef := wF.GetCellValue(wFSheetName, fmt.Sprint(wFRColumn, i))
					rfNowRef := rFSheetRow[rune(rFRColumn[0])-BaseColumnRune]
					defer func() {
						if r := recover(); r != nil {
							fmt.Println("rFSheetName:", rFSheetName)
							fmt.Println("rFSheetRow:", rFSheetRow)
							fmt.Println(r)
						}
					}()
					if wFNowRef == rfNowRef {
						rFSheetPrice, _ := strconv.ParseFloat(rFSheetRow[rune(rFPColumn[0])-BaseColumnRune], 64)
						wF.SetCellValue(wFSheetName, fmt.Sprint(wFPColumn, i), rFSheetPrice)
						wF.SetCellStyle(wFSheetName, fmt.Sprint(wFPColumn, i), fmt.Sprint(wFPColumn, i), priceCellStyle)
						totalCount += 1
						searched = true
						break rFLoop
					}
				}
			}
			if searched == false {
				wF.SetCellStyle(wFSheetName, fmt.Sprint(wFPColumn, i), fmt.Sprint(wFPColumn, i), emptyCellStyle)
			}
		}
	}
}

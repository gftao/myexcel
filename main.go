package main

import (
	"fmt"
	"time"
	"github.com/aswjh/excel"
)

func main() {
	option := excel.Option{"Visible": true, "DisplayAlerts": true}
	xl, _ := excel.New(option)      //xl, _ := excel.Open("test_excel.xls", option)
	defer xl.Quit()

	sheet, _ := xl.Sheet(1)         //xl.Sheet("sheet1")
	defer sheet.Release()
	sheet.Cells(1, 1, "hello")
	sheet.PutCell(1, 2, 2006)
	sheet.MustCells(1, 3, 3.14159)
	sheet.MustCells(1, 4, "商户号")

	cell := sheet.MustCell(5, 6)
	defer cell.Release()
	cell.Put("go")
	cell.Put("font", map[string]interface{}{"name": "Arial", "size": 26, "bold": true})
	cell.Put("interior", "colorindex", 6)

	sheet.PutRange("a3:c3", []string {"@1", "@2", "@3"})
	rg := sheet.Range("d3:f3")
	defer rg.Release()
	rg.Put([]string {"~4", "~5", "~6"})

	urc := sheet.MustGet("UsedRange", "Rows", "Count").(int32)
	println("str:"+sheet.MustCells(1, 2), sheet.MustGetCell(1, 2).(float64), cell.MustGet().(string), urc)

	cnt := 0
	sheet.ReadRow("A", 1, "F", 9, func(row []interface{}) (rc int) {    //"A", 1 or 1, 9 or 1 or nothing
		cnt ++
		fmt.Println(cnt, row)
		return                                                                   //-1: break
	})

	time.Sleep(3000000000)
	xl.SaveAs("E:/go/src/myexcel/test_excel.xls")    //xl.SaveAs("test_excel", "html")


}

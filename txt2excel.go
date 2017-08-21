package main

import (
	"flag"
	"io/ioutil"
	"fmt"
	"golib/gerror"
	"github.com/aswjh/excel"
	"os"
	"bufio"
	"io"
	"strings"
	"path/filepath"
	"os/exec"
)

var FilePath = flag.String("fp","./file/","txt file path")

func main()  {
	flag.Parse()
	option := excel.Option{"Visible": false, "DisplayAlerts": false}
	xl, _ := excel.New(option)      //xl, _ := excel.Open("test_excel.xls", option)
	defer xl.Quit()

	af, err := GetPathAllFile(*FilePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(af)

	sheet, _ := xl.Sheet(1)
	defer sheet.Release()
	curP ,err := GetCurrentPath()
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(curP, curP[len(curP)-1:])
	if (*FilePath)[len(*FilePath) -1:] != curP[len(curP)-1:] {
		*FilePath = *FilePath + curP[len(curP)-1:]
	}

	for i, _ := range af {
		nm := *FilePath + af[i]
		fmt.Println("->",nm)
		Txt2excl(nm, &sheet)
		//savenm := "E:\\go\\src\\myexcel\\" + af[i][:len(af[i])-4] + ".xlsx"
		savenm := curP + af[i][:len(af[i])-4] + ".xlsx"
		xl.SaveAs(savenm, "xlsx")
	}
	//sheet.MustCells(1, 4, "商户号")

	//xl.SaveAs("E:\\go\\src\\myexcel\\test_excel.xls", "xls")
}

func Txt2excl(fn string, sheet *excel.Sheet) gerror.IError{
	f, err := os.Open(fn)
	if err != nil {
		return gerror.NewR(101,err,"")
	}
	defer f.Close()
	bf := bufio.NewReader(f)
	row := 1
	for  {
		l, _, e := bf.ReadLine()
		if e == io.EOF {
			break
		}
		ls := string(l)
		//fmt.Println(ls)
		for i, cell := range strings.Split(ls,","){
			//fmt.Println(cell)
			sheet.Cells(row, i+1, cell)
		}
		row ++
	}

	return nil
}

func GetPathAllFile(p string) ([]string, gerror.IError) {
	fi, err := ioutil.ReadDir(p)
	//af := []string{}
	af := make([]string,0)
	if err != nil {
		fmt.Println("read dir error")
		return nil, gerror.NewR(100,err,"read dir error")
	}
	for _, v := range fi {
		//fmt.Println(i, "=",  v.Name())
		fn := v.Name()
		af = append(af, fn)
	}

	return af, nil
}

func GetCurrentPath() (string, gerror.IError) {
	file, err := exec.LookPath(os.Args[0])
	if err != nil {
		return "",   gerror.NewR(106,err,"")
	}
	path, err := filepath.Abs(file)
	if err != nil {
		return "",  gerror.NewR(105,err,"")
	}
	i := strings.LastIndex(path, "/")
	if i < 0 {
		i = strings.LastIndex(path, "\\")
	}
	if i < 0 {
		return "", gerror.NewR(104,err,`error: Can't find "/" or "\".`)
	}
	return string(path[0 : i+1]), nil
}
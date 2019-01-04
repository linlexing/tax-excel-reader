package main

import (
	"encoding/csv"
	"flag"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	var rootPath = flag.String("path", "", "path the directory")
	flag.Parse()
	if *rootPath == "" {
		println("must set path param")
		return
	}
	println("search for", *rootPath)
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	exPath := filepath.Dir(ex)
	fs, err := os.Create(filepath.Join(exPath, "output.csv"))
	if err != nil {
		panic(err)
	}
	defer fs.Close()
	csvW := csv.NewWriter(fs)
	csvW.UseCRLF = true
	if err := csvW.Write([]string{"工号", "姓名", "所属部门", "身份证号", "专项附加扣除分类", "开始时间", "结束时间", "月扣除金额"}); err != nil {
		panic(err)
	}

	if err := filepath.Walk(*rootPath, func(path string, info os.FileInfo, err error) error {
		if filepath.Ext(path) != ".xlsx" {
			return nil
		}
		xlFile, err := xlsx.OpenFile(path)
		if err != nil {
			return err
		}
		println(path)
		for rowIndex := 1; rowIndex < 16; rowIndex++ {
			line := []string{}
			for colIndex := 0; colIndex < 8; colIndex++ {
				val := xlFile.Sheet["分项汇总"].Rows[rowIndex].Cells[colIndex].String()
				//日期要转换干净
				if colIndex == 5 || colIndex == 6 {

					if val == `1899\-12` {
						val = ""
					}
					val = strings.Replace(val, `\-`, "-", -1)
				}
				line = append(line, val)
			}
			// fmt.Printf("%v\n", line)
			csvW.Write(line)
		}
		return nil
	}); err != nil {
		panic(err)
	}
	csvW.Flush()
}

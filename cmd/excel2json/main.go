package main

import (
	"errors"
	"fmt"
	"io/fs"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/spf13/cast"

	"github.com/suhanyujie/excel2json/pkg/utils/jsonx"

	"github.com/xuri/excelize/v2"

	"github.com/urfave/cli/v2"
)

const (
	TargetDir = "./example/output"
)

func main() {
	app := &cli.App{
		Name:   "toJson",
		Usage:  "将特定的 xlsx 文件转换为 json 文件",
		Action: DoConvert,
	}

	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}

func DoConvert(ctx *cli.Context) error {
	// 1. 读取配置文件
	// 2. 读取 excel 文件
	// 3. 生成 json 文件
	fmt.Println("ok!")
	params := ctx.Args().Slice()
	if len(params) == 0 {
		// 转换当前路径下的所有 xlsx 文件 todo
		ConvertByDir("./example/data")
	} else if len(params) == 1 {
		// 转换当前路径下的所有 xlsx 文件 todo
	} else {
		fmt.Printf("参数错误")
		return nil
	}
	// fmt.Printf("%s", jsonx.ToJsonIgnoreErr(params))
	return nil
}

func ConvertByDir(dir string) []string {
	filepath.WalkDir(dir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			log.Printf("[readDir] err: %v", err)
			return err
		}
		if d.IsDir() {
			return nil
		}
		if filepath.Ext(path) != ".xlsx" || strings.HasPrefix(d.Name(), "~") {
			return nil
		}
		ConvertOneFile(path)
		return nil
	})
	return []string{}
}

func ConvertOneFile(fileFullPath string) error {
	dataList := make([]map[string]any, 0)
	fs, err := excelize.OpenFile(fileFullPath)
	if err != nil {
		log.Printf("[ConvertOneFile] err: %v", err)
		return err
	}
	defer fs.Close()
	sheetList := fs.GetSheetList()
	if len(sheetList) < 1 {
		return errors.New("sheetList is empty")
	}
	rows, err := fs.GetRows(sheetList[0])
	if err != nil {
		log.Printf("[ConvertOneFile] err: %v", err)
		return err
	}
	colKeys := make([]string, 0)
	colNameArr := make([]string, 0)
	typeArr := make([]string, 0)
	for rowIdx, row := range rows {
		if rowIdx >= 2 {
			break
		}
		switch rowIdx {
		case 0:
			colNameArr = append(colNameArr, row...)
		case 1:
			colKeys = append(colKeys, row...)
		case 2:
			typeArr = append(typeArr, row...)
		}
	}
	var colKey string
	for rowIdx, row := range rows {
		if rowIdx <= 2 {
			continue
		}
		rowMap := make(map[string]any, 0)
		for i, cell := range row {
			if i > len(colKeys) || len(colKeys[i]) < 1 {
				continue
			}
			colKey = colKeys[i]
			var cellValue any
			var typeName string
			if i >= len(typeArr) {
				typeName = "string"
			} else {
				typeName = typeArr[i]
			}
			switch typeName {
			case "int":
				cellValue = cast.ToInt64(cell)
			case "list":
				tmpList := make([]any, 0)
				jsonx.FromJson(cell, &tmpList)
				cellValue = tmpList
			default:
				cellValue = cell
			}
			rowMap[colKey] = cellValue
		}
		dataList = append(dataList, rowMap)

		// 写入文件
		fileName := filepath.Base(fileFullPath)
		fileSuffix := path.Ext(fileFullPath)
		fileNamePrefix := fileName[0 : len(fileName)-len(fileSuffix)]
		targetFilePath := path.Join(TargetDir, fileNamePrefix+".json")
		tmpFs, err := os.OpenFile(targetFilePath, os.O_CREATE|os.O_RDWR, 0666)
		if err != nil {
			log.Printf("[ConvertOneFile] OpenFile err: %v", err)
			return err
		}
		tmpFs.WriteString(jsonx.ToJsonIgnoreErr(dataList))
		tmpFs.Close()
	}
	return nil
}

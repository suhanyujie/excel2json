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
	"github.com/urfave/cli/v2"
	"github.com/xuri/excelize/v2"
)

var (
	// default value
	input  = "./"
	output = "./output"
)

func main() {
	app := &cli.App{
		Name:   "toJson",
		Usage:  "将特定的 xlsx 文件转换为 json 文件",
		Action: DoConvert,
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "input",
				Value:       "./",
				Aliases:     []string{"i"},
				Usage:       "要转换的文件所在路径",
				Destination: &input,
			},
			&cli.StringFlag{
				Name:        "output",
				Value:       "./output",
				Aliases:     []string{"o"},
				Usage:       "转换成 json 存放的路径",
				Destination: &output,
			},
		},
	}

	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}

func DoConvert(ctx *cli.Context) error {
	// 1. 读取配置文件
	// 2. 读取 excel 文件
	// 3. 生成 json 文件
	params := ctx.Args().Slice()
	fileNum := 0
	if len(params) == 0 {
		// 没有参数时，意味着，转换当前目录下的 xlsx 文件，并将其输出到当前文件夹的 output 文件夹下
		// 转换当前路径下的所有 xlsx 文件
		fileNum = ConvertByDir(input)
	} else if len(params) == 1 {
		// 有一个参数，有两种情况：
		// 1.参数是输入文件夹
		// 2.参数是输出文件夹
		// 暂不支持一个参数
	} else if len(params) == 2 {
		if input != ctx.Args().Get(0) {
			input = ctx.Args().Get(0)
		}
		if output != ctx.Args().Get(1) {
			output = ctx.Args().Get(1)
		}
		// 2 个参数，一个是输入目录，一个是输出目录
		fileNum = ConvertByDir(input)
	} else {
		// 其他参数，暂不支持
		fmt.Printf("[err] 参数错误")
		return nil
	}
	fmt.Printf("[ok] 转换完成，转换了 %d 个文件...\n", fileNum)
	return nil
}

func handleForInputParam(p1, p2 string) (inputDir string) {
	// 对于只有输入目录的情况下，用户可能输入：`i=./`，也可能只输入 `./`
	inputDir = p2
	if p1 == "" {
		inputDir = p1
	}
	return inputDir
}

func ConvertByDir(inputDir string) int {
	cnt := 0
	filepath.WalkDir(inputDir, func(path string, d fs.DirEntry, err error) error {
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
		cnt++
		return nil
	})
	return cnt
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
		targetFilePath := path.Join(output, fileNamePrefix+".json")
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

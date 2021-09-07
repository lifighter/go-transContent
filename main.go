package main

import (
	"bufio"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	fmt.Println("本软件遵循Apache License 2.0开源协议，可以对其进行修改、商用、分发；" +
		"\nhttp://www.apache.org/licenses/LICENSE-2.0" +
		"\n程序用于将txt文本按照既定格式转换为Excel，详情请参见：" +
		"\nhttps://github.com/")

	for {
		fmt.Println("请输入需要转换的txt文本文件：（输入q退出程序）")
		r := bufio.NewReader(os.Stdin)
		i, _, _ := r.ReadLine()
		txtPath := string(i)
		if txtPath == "q" {
			return
		}
		txtFile, err := os.Open(txtPath)
		if err != nil {
			fmt.Println("打开文件出错！", err)
		}
		defer txtFile.Close()

		xlsxFile := xlsx.NewFile()
		sheet, err := xlsxFile.AddSheet("统计表")
		row := sheet.AddRow()
		if err != nil {
			fmt.Println("创建文件出错", err)
		}
		titles := [...]string{"车号", "姓名", "吨位", "手机号", "驾驶证号", "货型", "其它"}
		itemsMap := make(map[string]string)

		for i := range titles {
			cell := row.AddCell()
			cell.Value = titles[i]
		}

		txtBr := bufio.NewReader(txtFile)
		for {
			a, _, c := txtBr.ReadLine()
			if c == io.EOF {
				fmt.Println("文件已全部读完.")
				break
			}
			tmpLine := string(a)
			// tmpLine := "车号：宁E61081"
			tmpLine = strings.ReplaceAll(tmpLine, ":", "：")
			tmpItems := strings.Split(tmpLine, "：")
			if len(tmpItems) != 2 {
				fmt.Println("文本【" + tmpLine + "】格式错误，必须以：分隔.")
				break
			}

			itemsMap[tmpItems[0]] = tmpItems[1]

		}

		for i := range titles {
			cell := sheet.Cell(1, i)
			cell.Value = itemsMap[titles[i]]
			delete(itemsMap, titles[i])
		}
		if len(itemsMap) > 0 {
			var otherStr string
			for k, v := range itemsMap {
				otherStr = k + ":" + v + ","
			}
			cell := sheet.Cell(1, len(titles)-1)
			cell.Value = otherStr
		}

		err = xlsxFile.Save("test.xlsx")
		if err != nil {
			fmt.Println("文件保存出错", err)
		}
	}
}

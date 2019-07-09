/*
* @Author: apple
* @Date:   2019-07-10 02:01:02
* @Last Modified by:   scottxiong
* @Last Modified time: 2019-07-10 05:32:58
 */
package xlsx

import (
	"github.com/tealeg/xlsx"
)

type Point struct {
	row   int
	col   int
	value string
}

var (
	points = []Point{}
)

func Write(xlsxFile string, sheetName string) {

	newFile := xlsxFile

	file, error := xlsx.OpenFile(newFile)

	if error != nil {
		panic(error)
	}
	defer file.Save(newFile)
	for _, point := range points {
		file.Sheet[sheetName].Rows[point.row].Cells[point.col].SetValue(point.value)
	}
}

func NewPoint(row, col int, value string) []Point {
	points = append(points, Point{
		row:   row,
		col:   col,
		value: value,
	})
	return points
}

/*
* @Author: scottxiong
* @Date:   2019-07-10 02:06:51
* @Last Modified by:   scottxiong
* @Last Modified time: 2019-07-10 05:33:57
 */
package main

import (
	"github.com/scott-x/utils/xlsx"
)

func main() {
	xlsx.NewPoint(3, 1, "some notes")
	xlsx.NewPoint(3, 0, "12/12/2012")
	xlsx.Do("./template.xlsx", "Sheet1")
}

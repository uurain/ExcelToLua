package main

import (
	"errors"
	"fmt"
	"io/ioutil"

	"os"
	"path"
	"strconv"
	"strings"

	"encoding/json"

	"github.com/excelize"
	//	"LuaAdapter"
)

type ExcelData struct {
	Name       string
	TypeName   string
	SplitStr   string
	Annotation string
}

func main() {
	var argCount = len(os.Args)
	if argCount < 3 {
		fmt.Println("input or output path error!")
		return
	}
	var srcFilePath = os.Args[1]
	var targetPath = os.Args[2]

	//	var srcFilePath = "./test.xlsx"
	//	var targetPath = "./"

	Generate(srcFilePath, targetPath)
}

func Generate(filePath string, tPath string) {
	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(filePath + " OpenExcelError:" + err.Error())
		return
	}
	fileName := path.Base(filePath)
	fileName = strings.Replace(fileName, ".xlsx", "", -1)
	fmt.Println("begin:" + filePath)
	//	var sheetMap = xlsx.GetSheetMap()
	//	for i := range sheetMap {
	//		fmt.Println(i)
	//		fmt.Println(sheetMap[i])
	//	}

	var sheetName = xlsx.GetSheetName(1)
	if sheetName == "" {
		fmt.Println("Sheet为空")
		return
	}
	rows := xlsx.GetRows(sheetName)
	rowCount := len(rows)
	if rowCount < 4 {
		fmt.Println("行数太少")
		return
	}

	colCellTypeMap := make(map[int]*ExcelData)
	var lineData [][]interface{}

	for rowIndex, row := range rows {
		var objValAry []interface{}
		var errStr = ""
		for colIndex, colCell := range row {
			colCellType, ok := colCellTypeMap[colIndex]
			if !ok {
				colCellType = new(ExcelData)
				colCellTypeMap[colIndex] = colCellType
			}
			if rowIndex == 0 {
				colCellType.Name = colCell
			} else if rowIndex == 1 {
				// 注释
				colCellType.Annotation = colCell
			} else if rowIndex == 2 {
				// 分割符
				colCellType.SplitStr = colCell
				if colCellType.SplitStr == "" {
					colCellType.SplitStr = ";"
				}
			} else if rowIndex == 3 {
				// 类型
				colCellType.TypeName = colCell
			} else {
				objVal, err := Transit(colCellType, colCell)
				if err != nil {
					errStr = fmt.Sprintf("%drow,%dcol,error:%s", rowIndex, colIndex, err.Error())
					break
				}
				objValAry = append(objValAry, objVal)
			}
		}
		if rowIndex > 3 {
			if errStr != "" {
				fmt.Println(fileName + ":" + errStr)
				return
			}
			lineData = append(lineData, objValAry)
		}
	}

	jsonPath := fmt.Sprintf("%v%v.json", tPath, fileName)
	CreateToJson(jsonPath, colCellTypeMap, lineData)

	luaPath := fmt.Sprintf("%v%v.lua", tPath, fileName)
	CreateToLua(luaPath, fileName, colCellTypeMap, lineData)

	fmt.Println(filePath + " Complete!")
}

func Transit(typeData *ExcelData, val string) (interface{}, error) {
	if typeData.TypeName == "int" {
		return strconv.Atoi(val)
	} else if typeData.TypeName == "int64" {
		return strconv.ParseInt(val, 10, 64)
	} else if typeData.TypeName == "float" {
		return strconv.ParseFloat(val, 32)
	} else if typeData.TypeName == "string" {
		return val, nil
	} else if typeData.TypeName == "bool" {
		if val == "true" {
			return true, nil
		} else if val == "false" {
			return false, nil
		} else {
			return false, errors.New("bool值错误")
		}
	} else if typeData.TypeName == "ints" {
		var valStrAry = strings.Split(val, typeData.SplitStr)
		var valAry []int
		for i := range valStrAry {
			tempVal, e := strconv.Atoi(valStrAry[i])
			if e != nil {
				return nil, e
			}
			valAry = append(valAry, tempVal)
		}
		return valAry, nil
	} else if typeData.TypeName == "int64s" {
		var valStrAry = strings.Split(val, typeData.SplitStr)
		var valAry []int64
		for i := range valStrAry {
			tempVal, e := strconv.ParseInt(valStrAry[i], 10, 64)
			if e != nil {
				return nil, e
			}
			valAry = append(valAry, tempVal)
		}
		return valAry, nil
	} else if typeData.TypeName == "floats" {
		var valStrAry = strings.Split(val, typeData.SplitStr)
		var valAry []float64
		for i := range valStrAry {
			tempVal, e := strconv.ParseFloat(valStrAry[i], 32)
			if e != nil {
				return nil, e
			}
			valAry = append(valAry, tempVal)
		}
		return valAry, nil
	} else if typeData.TypeName == "strings" {
		var valStrAry = strings.Split(val, typeData.SplitStr)
		return valStrAry, nil
	} else if typeData.TypeName == "bools" {
		var valStrAry = strings.Split(val, typeData.SplitStr)
		var valAry []bool
		for i := range valStrAry {
			var tempVal = false
			if valStrAry[i] == "true" {
				tempVal = true
			} else if valStrAry[i] == "false" {
				tempVal = false
			} else {
				return nil, errors.New("bool值错误")
			}
			valAry = append(valAry, tempVal)
		}
		return valAry, nil
	}
	return nil, nil
}

func CreateToLua(luaPath string, fileName string, colCellTypeMap map[int]*ExcelData, lineData [][]interface{}) {
	var strAll = fileName + " = {\n"
	var lineCount = len(lineData)

	for m := range lineData {
		var objValAry = lineData[m]
		var strLine = "{ "
		var objValCount = len(objValAry)
		for i := range objValAry {
			typeData, ok := colCellTypeMap[i]
			if !ok {
				fmt.Println(string(i) + " UnKnowlineError")
				return
			}
			var strVal = TransitTypeVal(typeData.Name, objValAry[i])
			strLine += strVal
			if i < objValCount-1 {
				strLine += ", "
			}
		}
		strLine += "}"
		strAll += "    " + strLine
		if m < lineCount-1 {
			strAll += ",\n"
		}
	}
	strAll += "\n}"

	data := []byte(strAll)
	wErr := ioutil.WriteFile(luaPath, data, 0644)
	if wErr != nil {
		fmt.Sprintf("%v write lua file Error:%v", luaPath, wErr.Error())
		return
	}
}

func CreateToJson(jsonPath string, colCellTypeMap map[int]*ExcelData, lineData [][]interface{}) {
	var newLineData []interface{}
	for m := range lineData {
		var objValAry = lineData[m]
		var objValMap = make(map[string]interface{})
		for i := range objValAry {
			typeData, ok := colCellTypeMap[i]
			if !ok {
				fmt.Println(string(i) + " UnKnowlineError")
				return
			}
			objValMap[typeData.Name] = objValAry[i]
		}
		newLineData = append(newLineData, objValMap)
	}

	jsonStr, jErr := json.Marshal(newLineData)
	if jErr != nil {
		fmt.Printf("%v json marshal err:", jErr)
	} else {
		wErr := ioutil.WriteFile(jsonPath, jsonStr, 0644)
		if wErr != nil {
			fmt.Sprintf("%v write json file Error:%v", jsonPath, wErr.Error())
			return
		}
	}

}

func TransitTypeVal(keyName string, val interface{}) string {
	var strVal = ""
	switch val.(type) {
	case int:
		strVal = fmt.Sprintf("%s = %d", keyName, val.(int))
		break
	case int64:
		strVal = fmt.Sprintf("%s = %d", keyName, val.(int64))
		break
	case float64:
		strVal = fmt.Sprintf("%s = %f", keyName, val.(float64))
		break
	case string:
		strVal = fmt.Sprintf("%s = \"%s\"", keyName, val.(string))
		break
	case bool:
		strVal = fmt.Sprintf("%s = %v", keyName, val.(bool))
		break
	case []int:
		var valAry = val.([]int)
		var tempStr = ""
		var valCount = len(valAry)
		for i := range valAry {
			tempStr += fmt.Sprintf("%v", valAry[i])
			if i < valCount-1 {
				tempStr += ", "
			}
		}
		strVal = fmt.Sprintf("%v = {%v}", keyName, tempStr)
		break
	case []int64:
		var valAry = val.([]int64)
		var tempStr = ""
		var valCount = len(valAry)
		for i := range valAry {
			tempStr += fmt.Sprintf("%v", valAry[i])
			if i < valCount-1 {
				tempStr += ", "
			}
		}
		strVal = fmt.Sprintf("%v = {%v}", keyName, tempStr)
		break
	case []float64:
		var valAry = val.([]float64)
		var tempStr = ""
		var valCount = len(valAry)
		for i := range valAry {
			tempStr += fmt.Sprintf("%v", valAry[i])
			if i < valCount-1 {
				tempStr += ", "
			}
		}
		strVal = fmt.Sprintf("%v = {%v}", keyName, tempStr)
		break
	case []bool:
		var valAry = val.([]bool)
		var tempStr = ""
		var valCount = len(valAry)
		for i := range valAry {
			tempStr += fmt.Sprintf("%v", valAry[i])
			if i < valCount-1 {
				tempStr += ", "
			}
		}
		strVal = fmt.Sprintf("%v = {%v}", keyName, tempStr)
		break
	case []string:
		var valAry = val.([]string)
		var tempStr = ""
		var valCount = len(valAry)
		for i := range valAry {
			tempStr += "\"" + valAry[i] + "\""
			if i < valCount-1 {
				tempStr += ", "
			}
		}
		strVal = fmt.Sprintf("%v = {%v}", keyName, tempStr)
		break
	}
	return strVal
}

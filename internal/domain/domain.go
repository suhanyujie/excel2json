package domain

import "strings"

// 分离命令行参数
// `i=./`  => `{key:"i", value:"./"}`
func SeparateCliParam(paramStr string) (key, value string) {
	vArr := strings.Split(paramStr, "=")
	return vArr[0], vArr[1]
}

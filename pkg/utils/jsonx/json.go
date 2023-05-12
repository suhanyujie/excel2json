package jsonx

import (
	"bytes"

	"github.com/goccy/go-json"
)

func ToJsonIgnoreErr(v interface{}) string {
	bArr, _ := json.Marshal(v)
	return string(bArr)
}

func ToJson(v interface{}) (string, error) {
	bArr, err := json.Marshal(v)
	if err != nil {
		return "", err
	}
	return string(bArr), nil
}

func FromJson(jsonStr string, o interface{}) error {
	err := json.Unmarshal([]byte(jsonStr), &o)
	if err != nil {
		return err
	}
	return nil
}

func JsonStrFormat(jsonCont string) string {
	src := []byte(jsonCont)
	dstBuf := bytes.NewBuffer(make([]byte, 0, len(src)))
	json.Indent(dstBuf, src, "", "    ")
	return dstBuf.String()
}

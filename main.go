package main

import (
	"flag"
	"fmt"
	"os"
	"sort"
	"strconv"

	jsoniter "github.com/json-iterator/go"
	"github.com/pkg/errors"
	"github.com/samber/lo"
	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/maps"
)

var json = jsoniter.ConfigCompatibleWithStandardLibrary

func main() {
	var srcFile = flag.String("if", "src.json", "源 JSON 文件")
	var dstFile = flag.String("of", "dst.xlsx", "结果 Excel 文件")
	var runMode = flag.String("mode", "product", "工作模式（product/debug）")
	flag.Parse()

	err := run(*srcFile, *dstFile)
	if err != nil {
		if *runMode == "debug" {
			fmt.Printf("%+v", err)
		} else {
			fmt.Println(err.Error())
		}
	}
}

type HeadIndex struct {
	count      int
	head2index map[string]int
	index2head map[int]string
}

func newHeadIndex() *HeadIndex {
	return &HeadIndex{
		head2index: make(map[string]int),
		index2head: make(map[int]string),
	}
}

func (h *HeadIndex) addHead(head string) {
	_, ok := h.head2index[head]
	if ok {
		return
	}

	index := h.count
	h.head2index[head] = index
	h.index2head[index] = head

	h.count++
}

func (h *HeadIndex) buildRow(data map[string]interface{}) (row []interface{}) {
	row = make([]interface{}, 0)

	for i := 0; i < h.count; i++ {
		head := h.index2head[i]
		value := data[head]
		row = append(row, value)
	}

	return
}

func (h *HeadIndex) getHeads() (heads []string) {
	heads = maps.Keys(h.head2index)
	sort.Slice(heads, func(i, j int) bool {
		return h.head2index[heads[i]] < h.head2index[heads[j]]
	})
	return
}

func run(srcFile, dstFile string) (err error) {
	// read json
	buf, err := os.ReadFile(srcFile)
	if err != nil {
		err = errors.WithStack(err)
		return
	}

	arr := make([]map[string]interface{}, 0)
	err = json.Unmarshal(buf, &arr)
	if err != nil {
		err = errors.WithStack(err)
		return
	}

	// collect heads
	headIndex := newHeadIndex()
	for _, row := range arr {
		newHeads := maps.Keys(row)
		collectHeads(headIndex, newHeads)
	}

	// create excel
	ef := excelize.NewFile()
	defer ef.Close()
	sw, err := ef.NewStreamWriter("Sheet1")
	if err != nil {
		err = errors.WithStack(err)
		return
	}

	// write table head
	heads := headIndex.getHeads()
	err = sw.SetRow("A1", lo.Map(heads, func(s string, _ int) interface{} {
		return s
	}))
	if err != nil {
		err = errors.WithStack(err)
		return
	}

	// write data
	for i, item := range arr {
		row := headIndex.buildRow(item)
		err = sw.SetRow("A"+strconv.Itoa(i+2), row)
		if err != nil {
			err = errors.WithStack(err)
			return
		}
	}

	// save file
	err = sw.Flush()
	if err != nil {
		err = errors.WithStack(err)
		return
	}
	err = ef.SaveAs(dstFile)
	if err != nil {
		err = errors.WithStack(err)
		return
	}

	return
}

func collectHeads(headIndex *HeadIndex, newHeads []string) {
	for _, head := range newHeads {
		headIndex.addHead(head)
	}
}

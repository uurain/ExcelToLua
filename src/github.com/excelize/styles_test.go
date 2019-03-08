package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestStyleFill(t *testing.T) {
	cases := []struct {
		label      string
		format     string
		expectFill bool
	}{{
		label:      "no_fill",
		format:     `{"alignment":{"wrap_text":true}}`,
		expectFill: false,
	}, {
		label:      "fill",
		format:     `{"fill":{"type":"pattern","pattern":1,"color":["#000000"]}}`,
		expectFill: true,
	}}

	for _, testCase := range cases {
		xl := NewFile()
		const sheet = "Sheet1"

		styleID, err := xl.NewStyle(testCase.format)
		if err != nil {
			t.Fatalf("%v", err)
		}

		styles := xl.stylesReader()
		style := styles.CellXfs.Xf[styleID]
		if testCase.expectFill {
			assert.NotEqual(t, style.FillID, 0, testCase.label)
		} else {
			assert.Equal(t, style.FillID, 0, testCase.label)
		}
	}
}

func TestSetConditionalFormat(t *testing.T) {
	cases := []struct {
		label  string
		format string
		rules  []*xlsxCfRule
	}{{
		label: "3_color_scale",
		format: `[{
			"type":"3_color_scale",
			"criteria":"=",
			"min_type":"num",
			"mid_type":"num",
			"max_type":"num",
			"min_value": "-10",
			"mid_value": "0",
			"max_value": "10",
			"min_color":"ff0000",
			"mid_color":"00ff00",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "-10",
				}, {
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "10",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF00FF00",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}, {
		label: "3_color_scale default min/mid/max",
		format: `[{
			"type":"3_color_scale",
			"criteria":"=",
			"min_type":"num",
			"mid_type":"num",
			"max_type":"num",
			"min_color":"ff0000",
			"mid_color":"00ff00",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "50",
				}, {
					Type: "num",
					Val:  "0",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF00FF00",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}, {
		label: "2_color_scale default min/max",
		format: `[{
			"type":"2_color_scale",
			"criteria":"=",
			"min_type":"num",
			"max_type":"num",
			"min_color":"ff0000",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "0",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}}

	for _, testCase := range cases {
		xl := NewFile()
		const sheet = "Sheet1"
		const cellRange = "A1:A1"

		err := xl.SetConditionalFormat(sheet, cellRange, testCase.format)
		if err != nil {
			t.Fatalf("%s", err)
		}

		xlsx := xl.workSheetReader(sheet)
		cf := xlsx.ConditionalFormatting
		assert.Len(t, cf, 1, testCase.label)
		assert.Len(t, cf[0].CfRule, 1, testCase.label)
		assert.Equal(t, cellRange, cf[0].SQRef, testCase.label)
		assert.EqualValues(t, testCase.rules, cf[0].CfRule, testCase.label)
	}
}

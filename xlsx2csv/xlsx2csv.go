package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/jessevdk/go-flags"
)

var (
	opts struct {
		All bool `short:"a" long:"all"
			description:"export all sheets"`
		Sheet int `short:"s" long:"sheet"
			description:"sheet number to convert"`
		SheetName string `short:"n" long:"sheetname"
			description:"sheet name to convert"`
		Delimiter string `short:"d" long:"delimiter"
			description:"delimiter - columns delimiter in csv, 'tab' or 'x09' for a tab (default: comma ',')"`
		DateFormat string `short:"f" long:"dateformat"
			description:"override date/time format (ex. %Y/%m/%d)"`
		IgnoreEmpty bool `short:"i" long:"ignoreempty"
			description:"skip empty lines"`
		Escape bool `short:"e" long:"escape"
			description:"Escape \r\n\t characters"`
		SheetDelimiter string `short:"p" long:"sheetdelimiter"
			description:"sheet delimiter used to separate sheets, pass '' if you do not need delimiter (default: '--------')"`
		Hyperlinks bool `long:"hyperlinks"
			description:"include hyperlinks" desctiption:"include hyperlinks"`
		IncludeSheetPattern []string `short:"I" long:"include_sheet_pattern"
			description:"only include sheets named matching given pattern, only effects when -a option is enabled."`
		ExcludeSheetPattern []string `short:"E" long:"exclude_sheet_pattern"
			description:"exclude sheets named matching given pattern, only effects when -a option is enabled."`
		MergeCells bool `short:"m" long:"merge-cells"
			description:"merge cells"`
	}

	FORMATS = map[string]string{
		`general`:                  `float`,
		`0`:                        `float`,
		`0.00`:                     `float`,
		`#,##0`:                    `float`,
		`#,##0.00`:                 `float`,
		`0%`:                       `percentage`,
		`0.00%`:                    `percentage`,
		`0.00e+00`:                 `float`,
		`mm-dd-yy`:                 `date`,
		`d-mmm-yy`:                 `date`,
		`d-mmm`:                    `date`,
		`mmm-yy`:                   `date`,
		`h:mm am/pm`:               `date`,
		`h:mm:ss am/pm`:            `date`,
		`h:mm`:                     `time`,
		`h:mm:ss`:                  `time`,
		`m/d/yy h:mm`:              `date`,
		`#,##0 ;(#,##0)`:           `float`,
		`#,##0 ;[red](#,##0)`:      `float`,
		`#,##0.00;(#,##0.00)`:      `float`,
		`#,##0.00;[red](#,##0.00)`: `float`,
		`mm:ss`:                    `time`,
		`[h]:mm:ss`:                `time`,
		`mmss.0`:                   `time`,
		`##0.0e+0`:                 `float`,
		`@`:                        `float`,
		`yyyy\\-mm\\-dd`:           `date`,
		`dd/mm/yy`:                 `date`,
		`hh:mm:ss`:                 `time`,
		"dd/mm/yy\\ hh:mm":         `date`,
		`dd/mm/yyyy hh:mm:ss`:      `date`,
		`yy-mm-dd`:                 `date`,
		`d-mmm-yyyy`:               `date`,
		`m/d/yy`:                   `date`,
		`m/d/yyyy`:                 `date`,
		`dd-mmm-yyyy`:              `date`,
		`dd/mm/yyyy`:               `date`,
		`mm/dd/yy hh:mm am/pm`:     `date`,
		`mm/dd/yyyy hh:mm:ss`:      `date`,
		`yyyy-mm-dd hh:mm:ss`:      `date`,
	}
	STANDARD_FORMATS = map[int]string{
		0:  `general`,
		1:  `0`,
		2:  `0.00`,
		3:  `#,##0`,
		4:  `#,##0.00`,
		9:  `0%`,
		10: `0.00%`,
		11: `0.00e+00`,
		12: `# ?/?`,
		13: `# ??/??`,
		14: `mm-dd-yy`,
		15: `d-mmm-yy`,
		16: `d-mmm`,
		17: `mmm-yy`,
		18: `h:mm am/pm`,
		19: `h:mm:ss am/pm`,
		20: `h:mm`,
		21: `h:mm:ss`,
		22: `m/d/yy h:mm`,
		37: `#,##0 ;(#,##0)`,
		38: `#,##0 ;[red](#,##0)`,
		39: `#,##0.00;(#,##0.00)`,
		40: `#,##0.00;[red](#,##0.00)`,
		45: `mm:ss`,
		46: `[h]:mm:ss`,
		47: `mmss.0`,
		48: `##0.0e+0`,
		49: `@`,
	}
)

type (
	ErrInvalidXlsxFile struct {
		Name string
	}
	ErrSheetNotFound struct {
		SheetId int
	}
	OutFileAlreadyExistsException struct {
		Name string
	}
)

type (
	Workbook struct {
		Sheets []*Sheet
		Date1904 bool
	}
	Sheet struct {

	}
)

func (Workbook *w) Parse(r *io.Reader) {
	workbookDoc := nil

	if workbookDoc.firstChild
        if workbookDoc.firstChild.namespaceURI:
            fileVersion = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "fileVersion")
        else:
            fileVersion = workbookDoc.firstChild.getElementsByTagName("fileVersion")
        if len(fileVersion) == 0:
            self.appName = 'unknown'
        else:
            try:
                if workbookDoc.firstChild.namespaceURI:
                    self.appName = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "fileVersion")[0]._attrs['appName'].value
                else:
                    self.appName = workbookDoc.firstChild.getElementsByTagName("fileVersion")[0]._attrs['appName'].value
            except KeyError:
                # no app name
                self.appName = 'unknown'

        try:
            if workbookDoc.firstChild.namespaceURI:
                self.date1904 = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "workbookPr")[0]._attrs['date1904'].value.lower().strip() != "false"
            else:
                self.date1904 = workbookDoc.firstChild.getElementsByTagName("workbookPr")[0]._attrs['date1904'].value.lower().strip() != "false"
        except:
            pass

        if workbookDoc.firstChild.namespaceURI:
            sheets = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "sheets")[0]
        else:
            sheets = workbookDoc.firstChild.getElementsByTagName("sheets")[0]
        if workbookDoc.firstChild.namespaceURI:
            sheetNodes = sheets.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "sheet")
        else:
            sheetNodes = sheets.getElementsByTagName("sheet")
        for sheetNode in sheetNodes:
            attrs = sheetNode._attrs
            name = attrs["name"].value
            if self.appName == 'xl' and len(attrs["r:id"].value) > 2:
                if 'r:id' in attrs: id = int(attrs["r:id"].value[3:])
                else: id = int(attrs['sheetId'].value)
            else:
                if 'sheetId' in attrs: id = int(attrs["sheetId"].value)
                else: id = int(attrs['r:id'].value[3:])
            self.sheets.append({'name': name, 'id': id})

func main() {
	args, err := flags.ParseArgs(&opts, os.Args)
	if err != nil {
		os.Exit(2)
	}

	b, _ := json.Marshal(&opts)
	fmt.Println(string(b))

	b, _ = json.Marshal(&args)
	fmt.Println(string(b))
}

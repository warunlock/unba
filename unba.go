package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"strconv"
	"strings"

	"github.com/xuri/excelize"
)

type data struct {
	Items []struct {
		ID         string `json:"id"`
		Surname    string `json:"surname"`
		Firstname  string `json:"firstname"`
		Middlename string `json:"middlename"`
		Racalc     string `json:"racalc"`
		Certnum    string `json:"certnum"`
		Certat     string `json:"certat"`
		Certcalc   string `json:"certcalc"`
	} `json:"items"`
	Total string `json:"total"`
}

var mdata = data{}
var rada string = "27"
var offset int = 0
var wline int = 2

func main() {

	flag.StringVar(&rada, "r", rada, "Выбор рады 1-28")
	flag.Parse()

	f := excelize.NewFile()
	fmt.Println("ПІБ;Обліковується у;№ Свідоцтва;Дата видачі;Орган, що видав свідоцтво;Адреса;Область;Місто;Телефон;email;Форми діяльності;Адреса;Область;Місто;Телефон;link;Чи актичний;Атестації;")
	f.SetCellValue("Sheet1", "A1", "ПІБ")
	f.SetCellValue("Sheet1", "B1", "Обліковується у")
	f.SetCellValue("Sheet1", "C1", "№ Свідоцтва")
	f.SetCellValue("Sheet1", "D1", "Дата видачі")
	f.SetCellValue("Sheet1", "E1", "Орган, що видав свідоцтво")
	f.SetCellValue("Sheet1", "F1", "Адреса")

	f.SetCellValue("Sheet1", "G1", "Область")
	f.SetCellValue("Sheet1", "H1", "Місто")

	f.SetCellValue("Sheet1", "I1", "Телефон")
	f.SetCellValue("Sheet1", "J1", "Email")
	f.SetCellValue("Sheet1", "K1", "Форми діяльності")
	f.SetCellValue("Sheet1", "L1", "Адреса")

	f.SetCellValue("Sheet1", "M1", "Область")
	f.SetCellValue("Sheet1", "N1", "Місто")

	f.SetCellValue("Sheet1", "O1", "Телефон")
	f.SetCellValue("Sheet1", "P1", "link")
	f.SetCellValue("Sheet1", "Q1", "Чи актичний")
	f.SetCellValue("Sheet1", "R1", "Атестації")

	for {
		//	resp, err := http.Get("https://erau.unba.org.ua/search?limit=10&offset=61000&order%5Bsurname%5D=ASC&addation%5Bprobono%5D=0&foreigner=0")
		resp, err := http.Get("https://erau.unba.org.ua/search?limit=8&offset=" + strconv.Itoa(offset) + "&order%5Bsurname%5D=ASC&raid=" + rada + "&addation%5Bprobono%5D=0&foreigner=0")
		//Request URL: https://erau.unba.org.ua/search?limit=8&offset=0&order%5Bsurname%5D=ASC&raid=28&addation%5Bprobono%5D=0&foreigner=0

		if err != nil {
			fmt.Println(err)
			return
		}
		defer resp.Body.Close()

		gdata, err := ioutil.ReadAll(resp.Body)

		err = json.Unmarshal(gdata, &mdata)

		//fmt.Println(mdata)
		if len(mdata.Items) == 0 {
			break
		}
		for _, obj := range mdata.Items {
			//fmt.Println(obj)
			fmt.Print(obj.Surname + " " + obj.Firstname + " " + obj.Middlename + ";" + obj.Racalc + ";" + obj.Certnum + ";" + obj.Certat + ";" + obj.Certcalc + ";")
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(wline), obj.Surname+" "+obj.Firstname+" "+obj.Middlename)
			f.SetCellValue("Sheet1", "B"+strconv.Itoa(wline), obj.Racalc)
			f.SetCellValue("Sheet1", "C"+strconv.Itoa(wline), obj.Certnum)
			f.SetCellValue("Sheet1", "D"+strconv.Itoa(wline), obj.Certat)
			f.SetCellValue("Sheet1", "E"+strconv.Itoa(wline), obj.Certcalc)

			//gat data from personal card
			req, err := http.NewRequest("GET", "https://erau.unba.org.ua/profile/"+obj.ID, nil)
			if err != nil {
				// handle err
			}
			req.Header.Set("Connection", "keep-alive")
			req.Header.Set("Cache-Control", "max-age=0")
			req.Header.Set("Sec-Ch-Ua", "^^Chromium^^;v=^^92^^, ^^")

			resp1, err := http.DefaultClient.Do(req)
			if err != nil {
				// handle err
			}
			defer resp1.Body.Close()
			body, _ := ioutil.ReadAll(resp1.Body)
			sbody := string(body)
			wstring := ""
			windex := strings.Index(sbody, "<h2 style=\"font-size: 13px;margin: 0;padding-bottom: 5px;\">")
			if windex > -1 {
				wstring = sbody[windex+59:]
				windex = strings.Index(wstring, "</h2>")
				wadress := wstring[:windex]
				fmt.Print(wadress + ";")
				f.SetCellValue("Sheet1", "F"+strconv.Itoa(wline), wadress)
				split := strings.Split(wadress, ", ")
				if len(split) > 4 {
					f.SetCellValue("Sheet1", "G"+strconv.Itoa(wline), split[1])
					if split[1] == "Київ" || split[1] == "Севастополь" {
						f.SetCellValue("Sheet1", "H"+strconv.Itoa(wline), split[2])
					} else {
						f.SetCellValue("Sheet1", "H"+strconv.Itoa(wline), split[3])
					}

				} else {
					f.SetCellValue("Sheet1", "G"+strconv.Itoa(wline), "-")
					f.SetCellValue("Sheet1", "H"+strconv.Itoa(wline), "-")
				}

			} else {
				fmt.Print("-;-;-;")
				f.SetCellValue("Sheet1", "F"+strconv.Itoa(wline), "-")
				f.SetCellValue("Sheet1", "G"+strconv.Itoa(wline), "-")
				f.SetCellValue("Sheet1", "H"+strconv.Itoa(wline), "-")
			}
			//fmt.Println(wstring)
			windex = strings.Index(wstring, "<a href=\"tel:")
			if windex > -1 {
				wstring = wstring[windex+13:]
				windex = strings.Index(wstring, "\">")
				tel := wstring[:windex]
				fmt.Print(tel + ";")
				f.SetCellValue("Sheet1", "I"+strconv.Itoa(wline), tel)
			} else {
				fmt.Print("-" + ";")
				f.SetCellValue("Sheet1", "I"+strconv.Itoa(wline), "-")
			}

			windex = strings.Index(wstring, "<a href=\"mailto:")
			if windex > -1 {
				wstring = wstring[windex+16:]
				windex = strings.Index(wstring, "\">")
				email := wstring[:windex]
				fmt.Print(email + ";")
				f.SetCellValue("Sheet1", "J"+strconv.Itoa(wline), email)
			} else {
				fmt.Print("-" + ";")
				f.SetCellValue("Sheet1", "J"+strconv.Itoa(wline), "-")
			}

			windex = strings.Index(wstring, "Форми адвокатської діяльності.")
			if windex > -1 {
				wstring = wstring[windex:]
				windex = strings.Index(wstring, "<div class=\"column-right__header col-md-12\">")
				wstring = wstring[windex+44:]
				windex = strings.Index(wstring, "</div>")
				forms := strings.TrimSpace(wstring[:windex])
				fmt.Print(forms + ";")
				f.SetCellValue("Sheet1", "K"+strconv.Itoa(wline), forms)
			} else {
				fmt.Print("-" + ";")
				f.SetCellValue("Sheet1", "K"+strconv.Itoa(wline), "-")
			}

			windex = strings.Index(wstring, "Адреса:")
			if windex > -1 {
				wstring = wstring[windex:]
				windex = strings.Index(wstring, "<div class=\"text-info col-md-9\">")
				wstring = wstring[windex+32:]
				windex = strings.Index(wstring, "</div>")
				wadress := wstring[:windex]
				fmt.Print(wadress + ";")
				f.SetCellValue("Sheet1", "L"+strconv.Itoa(wline), wadress)
				split := strings.Split(wadress, ", ")
				if len(split) > 4 {
					f.SetCellValue("Sheet1", "M"+strconv.Itoa(wline), split[1])
					if split[1] == "Київ" || split[1] == "Севастополь" {
						f.SetCellValue("Sheet1", "N"+strconv.Itoa(wline), split[2])
					} else {
						f.SetCellValue("Sheet1", "N"+strconv.Itoa(wline), split[3])
					}
				} else {
					f.SetCellValue("Sheet1", "M"+strconv.Itoa(wline), "-")
					f.SetCellValue("Sheet1", "N"+strconv.Itoa(wline), "-")
				}
			} else {
				fmt.Print("-;-;-;")
				f.SetCellValue("Sheet1", "L"+strconv.Itoa(wline), "-")
				f.SetCellValue("Sheet1", "M"+strconv.Itoa(wline), "-")
				f.SetCellValue("Sheet1", "N"+strconv.Itoa(wline), "-")
			}

			windex = strings.Index(wstring, "<a href=\"tel:")
			if windex > -1 {
				wstring = wstring[windex+13:]
				windex = strings.Index(wstring, "\">")
				tel := wstring[:windex]
				fmt.Print(tel + ";")
				f.SetCellValue("Sheet1", "O"+strconv.Itoa(wline), tel)
			} else {
				fmt.Print("-" + ";")
				f.SetCellValue("Sheet1", "O"+strconv.Itoa(wline), "-")
			}

			fmt.Print("https://erau.unba.org.ua/profile/" + obj.ID + ";")
			f.SetCellValue("Sheet1", "P"+strconv.Itoa(wline), "https://erau.unba.org.ua/profile/"+obj.ID)

			errindex := strings.Index(sbody, "<div class=\"error-alert\">")
			if errindex > -1 {
				fmt.Print("not active" + ";")
				f.SetCellValue("Sheet1", "Q"+strconv.Itoa(wline), "not active")
			} else {
				fmt.Print("active" + ";")
				f.SetCellValue("Sheet1", "Q"+strconv.Itoa(wline), "active")
			}

			windex = strings.Index(wstring, "Підвищення кваліфікації")
			if windex > -1 {
				wstring = wstring[windex:]
				PQ := ""
				for strings.Index(wstring, "<div class=\"type-info col-md-3\">") > -1 {
					windex = strings.Index(wstring, "<div class=\"type-info col-md-3\">")
					wstring = wstring[windex+32:]
					windex = strings.Index(wstring, "</div>")
					PQ = PQ + wstring[:windex] + " "
					wstring = wstring[windex+6:]
					windex = strings.Index(wstring, ">")
					wstring = wstring[windex+1:]
					windex = strings.Index(wstring, "</div>")
					PQ = PQ + strings.TrimSpace(wstring[:windex]) + " "
					wstring = wstring[windex+6:]
				}
				fmt.Print(PQ + ";")
				f.SetCellValue("Sheet1", "R"+strconv.Itoa(wline), PQ)
			} else {
				fmt.Print("-" + ";")
				f.SetCellValue("Sheet1", "R"+strconv.Itoa(wline), "-")
			}

			//

			fmt.Println(" ")
			wline++
			if err := f.SaveAs(rada + ".xlsx"); err != nil {
				fmt.Println(err)
			}
		}
		offset = offset + 8

		/*		if offset > 40 {
				break
			}*/

	}
	if err := f.SaveAs(rada + ".xlsx"); err != nil {
		log.Fatal(err)
	}
	log.Println("End")
}

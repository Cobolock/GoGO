package main

import (
	"bytes"
	"crypto/tls"
	"database/sql"
	"encoding/base64"
	"encoding/json"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"github.com/kardianos/service"
	"github.com/tealeg/xlsx"
	"golang.org/x/sys/windows/registry"
	"io/ioutil"
	"log"
	"net/smtp"
	"os"
	"strconv"
	"time"
)

var logger service.Logger
var Path string

type program struct{}

type EmailUser struct {
	Username    string
	Password    string
	EmailServer string
	Port        int
}

type LogString struct {
	LogTime  string
	ClientIp string
	DName    string
	UName    string
	Text     string
	OpName   string
}

type Config struct {
	Settings struct {
		MailFrom   string        `json:"mail_from,omitempty"`
		MailPass   string        `json:"mail_pass,omitempty"`
		MailServer string        `json:"mail_server,omitempty"`
		MailTo     string        `json:"mail_to,omitempty"`
		ServerIp   string        `json:"server_ip,omitempty"`
		ServerPass string        `json:"server_pass,omitempty"`
		ServerUser string        `json:"server_user,omitempty"`
		Day1       int           `json:"day_1,omitempty"`
		Day2       int           `json:"day_2,omitempty"`
		Day3       int           `json:"day_3,omitempty"`
		Day4       int           `json:"day_4,omitempty"`
		Day5       int           `json:"day_5,omitempty"`
		Day6       int           `json:"day_6,omitempty"`
		Day7       int           `json:"day_7,omitempty"`
		Delay      time.Duration `json:"delay",omitempty"`
		DoorsFlag  int           `json:"doors_flag,omitempty"`
		EcD        int           `json:"ecD,omitempty"`
		EcIP       int           `json:"ecIP,omitempty"`
		EcO        int           `json:"ecO,omitempty"`
		EcU        int           `json:"ecU,omitempty"`
		StartHour  int           `json:"start_hour,omitempty"`
		StartMin   int           `json:"start_min,omitempty"`
		UsersFlag  int           `json:"users_flag,omitempty"`
		MailPort   int           `json:"mail_port,omitempty"`
		ServerPort int           `json:"server_Port,omitempty"`
	} `json:"settings"`
	Filters struct {
		Doors map[string]int
		Users map[string]int
	} `json:"filters"`
}

func checkErr(err error) {
	if err != nil {
		t := fmt.Sprintf("%v\n", err)
		ioutil.WriteFile(Path+"log.txt", []byte(t), 0644)
		panic(err)
	}
}

func SendMail(from, to, server, port, message string) {
	tlc := &tls.Config{
		InsecureSkipVerify: true,
		ServerName:         server,
	}

	c, err := smtp.Dial(server + ":" + port)
	checkErr(err)

	err = c.StartTLS(tlc)
	checkErr(err)

	err = c.Mail(from)
	checkErr(err)

	err = c.Rcpt(to)
	checkErr(err)

	w, err := c.Data()
	checkErr(err)

	_, err = w.Write([]byte(message))
	checkErr(err)

	err = w.Close()
	checkErr(err)

	c.Quit()
}

func getJSON(fileName string) Config {

	var C Config

	C.Settings.MailPort = 25
	C.Settings.ServerPort = 3305
	C.Settings.ServerUser = "root"
	C.Settings.Delay = 30

	jsFile, err := ioutil.ReadFile(fileName)
	checkErr(err)

	err = json.Unmarshal(jsFile, &C)
	checkErr(err)

	return C
}

func getLastId(db *sql.DB, id *int) {
	q := "SELECT id FROM `tc-db-main`.userlog ORDER BY id DESC LIMIT 1"
	err := db.QueryRow(q).Scan(*&id)
	checkErr(err)
}

func decodePass(code string) string {
	p := ""
	for l := 0; l < len(code); l += 2 {
		p += string(code[l+1]) + string(code[l])
	}
	b, _ := base64.StdEncoding.DecodeString(string(p))
	return string(b)
}

func (p *program) run() {

	var id int
	var q string
	var daySend int = 0

	conf, err := os.Stat(Path + "config.json")

	config := getJSON(Path + "config.json")

	cS := config.Settings
	cF := config.Filters

	days := []int{cS.Day1, cS.Day2, cS.Day3, cS.Day4, cS.Day5, cS.Day6, cS.Day7}

	emailUser := &EmailUser{
		cS.MailFrom,
		decodePass(cS.MailPass),
		cS.MailServer,
		cS.MailPort,
	}

	usersList := ""
	lim := ""

	for k, _ := range cF.Users {
		a := cF.Users[k]
		if a == 1 {
			usersList += lim + k
		}
		lim = ", "
	}

	doorsList := ""
	lim = ""

	for k, _ := range cF.Doors {
		a, _ := cF.Doors[k]
		if a == 1 {
			doorsList += lim + k
		}
		lim = ", "
	}

	summaryFilters := ""
	if usersList != "" && cS.UsersFlag == 1 {
		summaryFilters += " AND l.USERID IN (" + usersList + ")"
	}
	if doorsList != "" && cS.DoorsFlag == 1 {
		summaryFilters += " AND l.APID IN (" + doorsList + ")"
	}

	db, err := sql.Open("mysql",
		cS.ServerUser+
			":"+
			decodePass(cS.ServerPass)+
			"@tcp("+
			cS.ServerIp+
			":"+
			strconv.Itoa(cS.ServerPort)+
			")/")
	checkErr(err)

	getLastId(db, &id)

	q = `SELECT l.LOGTIME, l.CLIENTIP,
        CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
        CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
        l.TEXT, p.NAME as OPNAME
        FROM ` + "`tc-db-main`" + `.userlog AS l
        LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
        LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
        LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as p ON l.USERID=p.ID
        WHERE l.ID > ? ` + summaryFilters + ` ORDER BY l.LOGTIME`

	for true {
		today := int(time.Time.Weekday(time.Now()))
		nowHour := int(time.Time.Hour(time.Now()))
		nowMin := int(time.Time.Minute(time.Now()))

		check, err := os.Stat(Path + "config.json")
		if conf.ModTime() != check.ModTime() {
			go p.run()
			return
		}

		if today == 0 {
			today = 7
		}

		for _, v := range days {
			if v == today && cS.StartHour == nowHour && cS.StartMin == nowMin && daySend != today {
				weekAgo := time.Time.AddDate(time.Now(), 0, 0, -7)
				now := time.Now()
				qxls := ""

				qxls = `SELECT l.LOGTIME, l.CLIENTIP,
                    CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
                    CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
                    l.TEXT, p.NAME as OPNAME
                    FROM ` + "`tc-db-main`" + `.userlog AS l
                    LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
                    LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
                    LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as p ON l.USERID=p.ID
                    WHERE l.LOGTIME BETWEEN '` + weekAgo.Format(time.RFC3339) + `' AND '` + now.Format(time.RFC3339) + `' ` + summaryFilters + ` ORDER BY l.LOGTIME`

				rowsXls, err := db.Query(qxls)
				checkErr(err)

				var logStringsXls []LogString

				for rowsXls.Next() {
					LSX := LogString{}
					err := rowsXls.Scan(&LSX.LogTime,
						&LSX.ClientIp,
						&LSX.UName,
						&LSX.DName,
						&LSX.Text,
						&LSX.OpName)
					checkErr(err)
					logStringsXls = append(logStringsXls, LSX)
				}

				var file *xlsx.File
				var sheet *xlsx.Sheet
				var row *xlsx.Row
				var cell *xlsx.Cell
				var buf bytes.Buffer

				file = xlsx.NewFile()
				sheet = file.AddSheet("Отчёт")
				row = sheet.AddRow()
				cell = row.AddCell()
				cell.Value = fmt.Sprintf("Отчёт за %02d.%02d - %02d.%02d", int(weekAgo.Month()), weekAgo.Day(), int(now.Month()), now.Day())

				row = sheet.AddRow()

				for _, v := range logStringsXls {
					row = sheet.AddRow()
					cell = row.AddCell()
					cell.Value = v.LogTime

					if t := cS.EcU; t == 1 {
						cell = row.AddCell()
						cell.Value = v.OpName
					}
					if t := cS.EcIP; t == 1 {
						cell = row.AddCell()
						cell.Value = v.ClientIp
					}
					if t := cS.EcD; t == 1 {
						cell = row.AddCell()
						cell.Value = v.DName
					}
					if t := cS.EcO; t == 1 {
						cell = row.AddCell()
						cell.Value = v.UName
					}
					cell = row.AddCell()
					cell.Value = v.Text
				}

				err = file.Save(Path + "spnx.xlsx")
				checkErr(err)

				xlsxFile, _ := ioutil.ReadFile(Path + "spnx.xlsx")
				checkErr(err)

				encoded := base64.StdEncoding.EncodeToString(xlsxFile)
				lineMaxLength := 500
				nbrLines := len(encoded) / lineMaxLength

				for i := 0; i < nbrLines; i++ {
					buf.WriteString(encoded[i*lineMaxLength:(i+1)*lineMaxLength] + "\n")
				}

				buf.WriteString(encoded[nbrLines*lineMaxLength:])

				header := make(map[string]string)
				header["From"] = emailUser.Username
				header["To"] = cS.MailTo
				header["Subject"] = "Отчёт"
				header["MIME-Version"] = "1.0"
				header["Content-Type"] = "application/csv; name=\"spnx.xlsx\""
				header["Content-Transfer-Encoding"] = "base64"
				header["Content-Disposition"] = "attachment; filename=\"spnx.xlsx\""

				message := ""
				for k, v := range header {
					message += fmt.Sprintf("%s: %s\r\n", k, v)
				}
				message += "\r\n" + fmt.Sprintf("%s\r\n", buf.String())

				SendMail(emailUser.Username,
					cS.MailTo,
					emailUser.EmailServer,
					strconv.Itoa(emailUser.Port),
					message,
				)

				daySend = today
			}
		}

		lastId := id

		rows, err := db.Query(q, id)
		checkErr(err)

		getLastId(db, &id)

		var logStrings []LogString

		for rows.Next() {
			LS := LogString{}
			err := rows.Scan(&LS.LogTime,
				&LS.ClientIp,
				&LS.UName,
				&LS.DName,
				&LS.Text,
				&LS.OpName)
			checkErr(err)
			logStrings = append(logStrings, LS)
		}

		if lastId != id && lastId != 0 {

			text := ""
			for _, v := range logStrings {
				text += fmt.Sprintf("%s", v.LogTime)
				if t := cS.EcU; t == 1 {
					text += fmt.Sprintf("\tПользователь: %s", v.OpName)
				}
				if t := cS.EcIP; t == 1 {
					text += fmt.Sprintf(" (%s)", v.ClientIp)
				}
				if t := cS.EcD; t == 1 {
					text += fmt.Sprintf("\tТочка прохода: %s", v.DName)
				}
				if t := cS.EcO; t == 1 {
					text += fmt.Sprintf("\tОбъект: %s", v.UName)
				}
				text += fmt.Sprintf("\t%s\n", v.Text)
			}

			header := make(map[string]string)
			header["From"] = emailUser.Username
			header["To"] = cS.MailTo
			header["Subject"] = "Отчёт"

			message := ""
			for k, v := range header {
				message += fmt.Sprintf("%s: %s\r\n", k, v)
			}
			message += "\r\n" + text

			SendMail(emailUser.Username,
				cS.MailTo,
				emailUser.EmailServer,
				strconv.Itoa(emailUser.Port),
				message,
			)

		}

		time.Sleep(cS.Delay * time.Second)
	}
}

func (p *program) Stop(s service.Service) error {
	// Stop should not block. Return with a few seconds.
	return nil
}

func (p *program) Start(s service.Service) error {
	// Start should not block. Do the actual work async.
	go p.run()
	return nil
}

func main() {

	k, err := registry.OpenKey(registry.LOCAL_MACHINE, "SOFTWARE\\INT", registry.QUERY_VALUE)
	checkErr(err)
	defer k.Close()

	Path, _, err = k.GetStringValue("WorkingDirectory")
	checkErr(err)

	fmt.Printf("%s\n", Path)

	svcConfig := &service.Config{
		Name:        "SphinxMailer",
		DisplayName: "SphinxMailer",
		Description: "Отправляет письма при обнаружении событий в СКУД Сфинкс.",
	}

	prg := &program{}
	s, err := service.New(prg, svcConfig)
	if err != nil {
		log.Fatal(err)
	}
	logger, err = s.Logger(nil)
	if err != nil {
		log.Fatal(err)
	}
	err = s.Run()
	if err != nil {
		logger.Error(err)
	}
}

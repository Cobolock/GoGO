package main

import (
	"bytes"
	"database/sql"
	"encoding/base64"
	"fmt"
	simplejson "github.com/bitly/go-simplejson"
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
	"tls"
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
	// LogTime     time.Time
	LogTime  string
	ClientIp string
	DName    string
	UName    string
	Text     string
	OpName   string
}

func checkErr(err error) {
	if err != nil {
		t := fmt.Sprintf("%v\n", err)
		ioutil.WriteFile(Path+"log.txt", []byte(t), 0644)
		panic(err)
	}
}

func getLastId(db *sql.DB, id *int) {
	q := "SELECT id FROM `tc-db-main`.userlog ORDER BY id DESC LIMIT 1"
	err := db.QueryRow(q).Scan(*&id)
	checkErr(err)
}

func getDBVErsion(db *sql.DB, dbv *int) {
	q := "SELECT PARAMVALUE FROM `tc-db-main`.parami WHERE NAME='DBVER'"
	err := db.QueryRow(q).Scan(*&dbv)
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

	var id, dbv int
	var q string
	var daySend int = 0

	conf, err := os.Stat(Path + "config.json")

	jsFile, err := ioutil.ReadFile(Path + "config.json")
	checkErr(err)

	jsData, err := simplejson.NewJson(jsFile)
	checkErr(err)

	jsSettings, _ := jsData.Get("settings").Map()
	jsMailPort, _ := jsData.Get("settings").Get("mail_port").Int()
	usersFlag, _ := jsData.Get("settings").Get("users_flag").Int()
	doorsFlag, _ := jsData.Get("settings").Get("doors_flag").Int()
	sendHour, _ := jsData.Get("settings").Get("start_hour").Int()
	sendMin, _ := jsData.Get("settings").Get("start_min").Int()

	days := []int{}
	for i := 1; i <= 7; i++ {
		t, _ := jsData.Get("settings").Get("day_" + strconv.Itoa(i)).Int()
		if t == 1 {
			days = append(days, i)
		}
	}

	emailUser := &EmailUser{
		jsSettings["mail_from"].(string),
		decodePass(jsSettings["mail_pass"].(string)),
		jsSettings["mail_server"].(string),
		jsMailPort,
	}

	auth := smtp.PlainAuth("",
		emailUser.Username,
		emailUser.Password,
		emailUser.EmailServer,
	)

	usersList := ""
	lim := ""

	usersMap, _ := jsData.Get("filters").Get("users").Map()
	for k, _ := range usersMap {
		a, _ := jsData.Get("filters").Get("users").Get(k).Int()
		if a == 1 {
			usersList += lim + k
		}
		lim = ", "
	}

	doorsList := ""
	lim = ""

	doorsMap, _ := jsData.Get("filters").Get("doors").Map()
	for k, _ := range doorsMap {
		a, _ := jsData.Get("filters").Get("doors").Get(k).Int()
		if a == 1 {
			doorsList += lim + k
		}
		lim = ", "
	}

	summaryFilters := ""
	if usersList != "" && usersFlag == 1 {
		summaryFilters += " AND l.USERID IN (" + usersList + ")"
	}
	if doorsList != "" && doorsFlag == 1 {
		summaryFilters += " AND l.APID IN (" + doorsList + ")"
	}

	db, err := sql.Open("mysql",
		jsSettings["server_user"].(string)+
			":"+
			decodePass(jsSettings["server_pass"].(string))+
			"@tcp("+
			jsSettings["server_ip"].(string)+
			":"+
			jsSettings["server_port"].(string)+
			")/")
	checkErr(err)

	getLastId(db, &id)
	getDBVErsion(db, &dbv)

	if dbv <= 161 {
		q = `SELECT l.LOGTIME, l.CLIENTIP,
            CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
            CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
            p.USERNAME as OPNAME, l.TEXT
            FROM ` + "`tc-db-main`" + `.userlog AS l
            LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
            LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
            LEFT OUTER JOIN ` + "`tc-common`" + `.profiles as p ON l.USERID=p.ID
            WHERE l.ID > ? ` + summaryFilters + ` ORDER BY l.LOGTIME`
	} else {
		q = `SELECT l.LOGTIME, l.CLIENTIP,
            CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
            CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
            l.TEXT, p.NAME as OPNAME
            FROM ` + "`tc-db-main`" + `.userlog AS l
            LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
            LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
            LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as p ON l.USERID=p.ID
            WHERE l.ID > ? ` + summaryFilters + ` ORDER BY l.LOGTIME`
	}

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
			if v == today && sendHour == nowHour && sendMin == nowMin && daySend != today {
				weekAgo := time.Time.AddDate(time.Now(), 0, 0, -7)
				now := time.Now()
				qxls := ""

				// I'm very sorry for that
				if dbv <= 161 {
					qxls = `SELECT l.LOGTIME, l.CLIENTIP,
                        CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
                        CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
                        p.USERNAME as OPNAME, l.TEXT
                        FROM ` + "`tc-db-main`" + `.userlog AS l
                        LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
                        LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
                        LEFT OUTER JOIN ` + "`tc-common`" + `.profiles as p ON l.USERID=p.ID
                        WHERE l.LOGTIME BETWEEN '` + weekAgo.Format(time.RFC3339) + `' AND '` + now.Format(time.RFC3339) + `' ` + summaryFilters + ` ORDER BY l.LOGTIME`
				} else {
					qxls = `SELECT l.LOGTIME, l.CLIENTIP,
                        CASE WHEN ISNULL(u.NAME) THEN '<Нет>' ELSE (u.NAME) END AS UNAME,
                        CASE WHEN ISNULL(d.NAME) THEN '<Нет>' ELSE (d.NAME) END AS DNAME,
                        l.TEXT, p.NAME as OPNAME
                        FROM ` + "`tc-db-main`" + `.userlog AS l
                        LEFT OUTER JOIN ` + "`tc-db-main`" + `.devices AS d ON l.APID=d.ID
                        LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as u ON l.OBJID=u.ID
                        LEFT OUTER JOIN ` + "`tc-db-main`" + `.personal as p ON l.USERID=p.ID
                        WHERE l.LOGTIME BETWEEN '` + weekAgo.Format(time.RFC3339) + `' AND '` + now.Format(time.RFC3339) + `' ` + summaryFilters + ` ORDER BY l.LOGTIME`
				}

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

					if t, _ := jsData.Get("settings").Get("ecU").Int(); t == 1 {
						cell = row.AddCell()
						cell.Value = v.OpName
					}
					if t, _ := jsData.Get("settings").Get("ecIP").Int(); t == 1 {
						cell = row.AddCell()
						cell.Value = v.ClientIp
					}
					if t, _ := jsData.Get("settings").Get("ecD").Int(); t == 1 {
						cell = row.AddCell()
						cell.Value = v.DName
					}
					if t, _ := jsData.Get("settings").Get("ecO").Int(); t == 1 {
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
				header["To"] = jsSettings["mail_to"].(string)
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

				err = smtp.SendMail(
					emailUser.EmailServer+": "+strconv.Itoa(emailUser.Port),
					auth,
					emailUser.Username,
					[]string{jsSettings["mail_to"].(string)},
					[]byte(message))
				checkErr(err)

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
				if t, _ := jsData.Get("settings").Get("ecU").Int(); t == 1 {
					text += fmt.Sprintf("\tПользователь: %s", v.OpName)
				}
				if t, _ := jsData.Get("settings").Get("ecIP").Int(); t == 1 {
					text += fmt.Sprintf(" (%s)", v.ClientIp)
				}
				if t, _ := jsData.Get("settings").Get("ecD").Int(); t == 1 {
					text += fmt.Sprintf("\tТочка прохода: %s", v.DName)
				}
				if t, _ := jsData.Get("settings").Get("ecO").Int(); t == 1 {
					text += fmt.Sprintf("\tОбъект: %s", v.UName)
				}
				text += fmt.Sprintf("\t%s\n", v.Text)
			}

			header := make(map[string]string)
			header["From"] = emailUser.Username
			header["To"] = jsSettings["mail_to"].(string)
			header["Subject"] = "Отчёт"

			message := ""
			for k, v := range header {
				message += fmt.Sprintf("%s: %s\r\n", k, v)
			}
			message += "\r\n" + text

			/*err = smtp.SendMail(
				emailUser.EmailServer+": "+strconv.Itoa(emailUser.Port),
				auth,
				emailUser.Username,
				[]string{jsSettings["mail_to"].(string)},
				[]byte(message))
			checkErr(err)*/

			tlc := &tls.Config{
				InsecureSkipVerify: true,
				ServerName: emailUser.EmailServer,
			}

			conn, err := tls.Dial("tcp", emailUser.EmailServer+":"+strconv.Itoa(emailUser.Port), tlc)
			checkErr(err)

			c, err := smtp.NewClient(conn, emailUser.EmailServer)
			checkErr(err)
			
			err = c.Auth(auth)
			checkErr(err)

			err = c.Mail(emailUser.Username)
			checkErr(err)

			err = c.Rcpt([]string{jsSettings["mail_to"].(string)})
			checkErr(err)

			w, err := c.Data()
			checkErr(err)

			_, err := w.Write([]byte(message))
			checkErr(err)

			err = w.Close()
			checkErr(err)

			c.Quit()

		}

		time.Sleep(5 * time.Second)
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

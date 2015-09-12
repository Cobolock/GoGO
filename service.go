package main

import (
    "os"
    "fmt"
    "time"
    "strconv"
    "net/smtp"
    _ "github.com/go-sql-driver/mysql"
    "database/sql"
)

type EmailUser struct {
    Username    string
    Password    string
    EmailServer string
    Port        int
}

type LogString struct {
    LogTime     string
    ClientIp    string
    DName       string
    UName       string
    Text        string
    OpName      int
}

func checkErr(err error) {
    if err != nil{
        panic(err)
    }
}

func getLastId(db *sql.DB, id *int){
    q := "SELECT id FROM userlog ORDER BY id DESC LIMIT 1"
    err := db.QueryRow(q).Scan(*&id)
    checkErr(err)
}

func main() {

    var id int

    emailUser := &EmailUser{
        "egorov@i-n-t.ru",
        "Idspispopd0)",
        "smtp.gmail.com",
        587,
    }

    auth := smtp.PlainAuth("",
        emailUser.Username,
        emailUser.Password,
        emailUser.EmailServer,
    )

    to := "egorov@i-n-t.ru"

    fo, err := os.Create("output.txt")
    checkErr(err)

    db, err := sql.Open("mysql",
                        "root:@tcp(127.0.0.1:3305)/tc-db-main?charset=cp1251")
    checkErr(err)

    getLastId(db, &id)

    defer func() {
        if err := fo.Close(); err != nil {
            panic(err)
        }
    }()

    q := `SELECT l.LOGTIME, l.CLIENTIP,
        CASE WHEN ISNULL(u.NAME) THEN '0' ELSE (u.NAME) END AS UNAME,
        CASE WHEN ISNULL(d.NAME) THEN '0' ELSE (d.NAME) END AS DNAME,
        l.TEXT, l.USERID
        FROM userlog AS l
        LEFT OUTER JOIN devices AS d ON l.APID=d.ID
        LEFT OUTER JOIN personal as u ON l.OBJID=u.ID
        WHERE l.ID > ?`

    for true {

        lastId := id

        rows, err := db.Query(q, id)
        checkErr(err)

        getLastId(db, &id)

        var logStrings []LogString

        for rows.Next() {
            LS  := LogString{}
            err := rows.Scan(&LS.LogTime,
                             &LS.ClientIp,
                             &LS.DName,
                             &LS.UName,
                             &LS.Text,
                             &LS.OpName)
            checkErr(err)
            logStrings = append(logStrings, LS)
            fmt.Printf("%v\n", logStrings)
        }

        if (lastId != id && lastId != 0) {

            t := time.Now()
            // text := fmt.Sprintf("%d at %s\r\n", id, t.Format(time.UnixDate))
            text := ""
            for k, v := range logStrings {
                
            }

            _, err = fo.WriteString(text)
            checkErr(err)

            header := make(map[string]string)
            header["From"] = emailUser.Username
            header["To"] = to
            header["Subject"] = "Отчёт"

            message := ""
            for k, v := range header {
                message += fmt.Sprintf("%s: %s\r\n", k, v)
            }
            message += "\r\n" + text

            err = smtp.SendMail(
                emailUser.EmailServer + ": " + strconv.Itoa(emailUser.Port),
                auth,
                emailUser.Username,
                []string{to},
                []byte(message))
            checkErr(err)
        }

        time.Sleep(5 * time.Second)
    }
}
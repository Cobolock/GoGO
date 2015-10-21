package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
)

type Config struct {
	Settings struct {
		MailFrom   string `json:"mail_from,omitempty"`
		MailPass   string `json:"mail_pass,omitempty"`
		MailServer string `json:"mail_server,omitempty"`
		MailTo     string `json:"mail_to,omitempty"`
		ServerIp   string `json:"server_ip,omitempty"`
		ServerPass string `json:"server_pass,omitempty"`
		ServerUser string `json:"server_user,omitempty"`
		Day1       int    `json:"day_1,omitempty"`
		Day2       int    `json:"day_2,omitempty"`
		Day3       int    `json:"day_3,omitempty"`
		Day4       int    `json:"day_4,omitempty"`
		Day5       int    `json:"day_5,omitempty"`
		Day6       int    `json:"day_6,omitempty"`
		Day7       int    `json:"day_7,omitempty"`
		DoorsFlag  int    `json:"doors_flag,omitempty"`
		EcD        int    `json:"ecD,omitempty"`
		EcIP       int    `json:"ecIP,omitempty"`
		EcO        int    `json:"ecO,omitempty"`
		EcU        int    `json:"ecU,omitempty"`
		StartHour  int    `json:"start_hour,omitempty"`
		StartMin   int    `json:"start_min,omitempty"`
		UsersFlag  int    `json:"users_flag,omitempty"`
		MailPort   int    `json:"mail_port,omitempty"`
		ServerPort int    `json:"server_Port,omitempty"`
	} `json:"settings"`
	Filters struct {
		Doors map[string]int
		Users map[string]int
	} `json:"filters"`
}

func getJSON(fileName string) Config {

	var C Config

	jsFile, err := ioutil.ReadFile(fileName)
	if err != nil {
		panic(err)
	}

	if err := json.Unmarshal(jsFile, &C); err != nil {
		panic(err)
	}

	return C
}

func main() {
	config := getJSON("config.json")
	fmt.Println("%+v", config)
}

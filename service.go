package main

import (
    "fmt"
    "time"
)

func main() {
    for true {
        fmt.Println("Hello, playground")
        time.Sleep(2 * time.Second)
        fmt.Println("Hello, playground 2")
    }
}
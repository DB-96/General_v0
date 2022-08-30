package main

import(
	"fmt"
	// "regexp"
	"bufio"
    "os"
    "strings"
	)

func main(){ 
	// var x int = 6
	//  y := 8 // Go will automatically infer the type
	fmt.Println("Hello Filthy Human")
	fmt.Println("Enter a string to be manipulated;")
	reader := bufio.NewReader(os.Stdin)
    fmt.Println("Simple Shell")
    fmt.Println("---------------------")
	for {
		fmt.Print("-> ")
		text, _ := reader.ReadString('\n')
		// convert CRLF to LF
		// text = RegEx(text)
		text = strings.Replace(text, "\n", "", -1)
	
		if strings.Compare("hi", text) == 0 {
		  fmt.Println("hello, Yourself")
		}
	
	  }
}

// func RegEx(input string) bool {
// 	match = False
// 	if input != "" {
// 		match,_ := regexp.MatchString("[0-9]", input)
// 		fmt.Println(match)
		
// 	}
// 	return match

// }


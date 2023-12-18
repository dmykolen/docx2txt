package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"runtime"

	"github.com/dmykolen/docx2txt"
)

const name = "docx2md"
const version = "0.0.6"

var revision = "HEAD"

func main() {
	var embed, showVersion bool
	var tablesStyle string
	flag.BoolVar(&embed, "embed", false, "embed resources")
	flag.BoolVar(&showVersion, "v", false, "Print the version")
	flag.StringVar(&tablesStyle, "ts", "pretty", "Table style: csv, pretty or md")
	flag.Parse()
	if showVersion {
		fmt.Printf("%s %s (rev: %s/%s)\n", name, version, revision, runtime.Version())
		return
	}
	if flag.NArg() == 0 {
		flag.Usage()
		os.Exit(1)
	}
	for _, arg := range flag.Args() {
		if buf, err := docx2txt.Docx2txt(arg, embed, docx2txt.StyleTbls(tablesStyle)); err != nil {
			log.Fatal(err)
		} else {
			fmt.Println(buf.String())
		}
	}
}

package docx2txt

import (
	"fmt"
	"io"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

func init() {
	os.MkdirAll("_testdata", 0755)
}

func TestEscape(t *testing.T) {
	tests := []struct {
		input  string
		escape string
		want   string
	}{
		{input: `\`, escape: `\`, want: `\\`},
		{input: `\`, escape: ``, want: `\`},
		{input: `\`, escape: `-`, want: `\`},
		{input: `\\`, escape: `\`, want: `\\\\`},
		{input: `\200`, escape: `\`, want: `\\200`},
	}
	for _, test := range tests {
		got := escape(test.input, test.escape)
		if got != test.want {
			t.Fatalf("want %v, but %v:", test.want, got)
		}
	}
}

func TestDocx2txt(t *testing.T) {
	t.Log(os.Getwd())
	tests := []struct {
		input string
		want  string
	}{
		{input: "../go-ai/_testdata/docx/Gift-short-2.docx", want: "_testdata/Gift.txt"},
	}
	for _, test := range tests {
		// buf, err := Docx2txt(test.input, false, StyleTbls("csv"), WithDebug(false), WithLogger(nil))
		buf, err := Docx2txt(test.input, false, StyleTbls("csv"), WithDebug(true))
		if assert.NoError(t, err) {
			// t.Logf("%v:",buf.String())

			f, _ := os.Create(test.want)
			buf.WriteTo(f)
		}

	}
}

func TestDocx2txt2(t *testing.T) {
	str := "HELLO \\t [WORLD]"
	t.Log("BEFORE: ", str)
	str = escape(str, `\[]`)
	t.Log("AFTER: ", str)
}

func TestLog(t *testing.T) {
	t.Log(fmt.Sprintf("format string"), "")
	// l := NewLoggerCustom(slog.New(slog.NewTextHandler(os.Stdout, &slog.HandlerOptions{Level: slog.LevelDebug})))
	l := NewLoggerCustom(os.Stdout, true)
	l.Debug("Teeeest 1")
	l.Debugf("\n\n >>>>>>>>>>>>  CBUF:[%s]\n", "Type\n")
	// l.Logger.Debug("\n\n >>>>>>>>>>>>  CBUF:[%s]\n", "Type\n")
	// l.Logger.Debug("XKMDKDMCDK %s", "00000000")
	l.Logger.Debug("XKMDKDMCDK %s {key}", "key", "00000000")
}
func TestTernary(t *testing.T) {
	v := ifTernary(true, os.Stdout, nil)
	t.Logf("Type=%T Val=%+v", v, v)
	v = ifTernary(false, os.Stdout, nil)
	t.Logf("Type=%T Val=%+v", v, v)
	v2 := ifTernary(true, io.MultiWriter(os.Stdout, os.Stderr), io.MultiWriter())
	t.Logf("Type=%T Val=%+v", v2, v2)
	v2.Write([]byte("HELLO\n"))
}

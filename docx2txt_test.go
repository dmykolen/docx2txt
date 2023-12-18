package docx2txt

import (
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

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
	tests := []struct {
		input string
		want  string
	}{
		{input: "_testdata/Gift.docx", want: "_testdata/Gift.txt"},
	}
	for _, test := range tests {
		buf, err := Docx2txt(test.input, false, StyleTbls("csv"))
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

package docx2txt

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"html"
	"io"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/mattn/go-runewidth"
)

var fp = fmt.Print
var ff = fmt.Printf

func escape(s, set string) string {
	replacer := []string{}
	for _, r := range []rune(set) {
		rs := string(r)
		replacer = append(replacer, rs, `\`+rs)
	}
	return strings.NewReplacer(replacer...).Replace(s)
}

func (zf *file) extract(rel *Relationship, w io.Writer) error {
	err := os.MkdirAll(filepath.Dir(rel.Target), 0755)
	if err != nil {
		return err
	}
	for _, f := range zf.r.File {
		if f.Name != "word/"+rel.Target {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return err
		}
		defer rc.Close()

		b := make([]byte, f.UncompressedSize64)
		n, err := rc.Read(b)
		if err != nil && err != io.EOF {
			return err
		}
		if zf.embed {
			fmt.Fprintf(w, "![](data:image/png;base64,%s)",
				base64.StdEncoding.EncodeToString(b[:n]))
		} else {
			err = os.WriteFile(rel.Target, b, 0644)
			if err != nil {
				return err
			}
			fmt.Fprintf(w, "![](%s)", escape(rel.Target, "()"))
		}
		break
	}
	return nil
}

func attr(attrs []xml.Attr, name string) (string, bool) {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value, true
		}
	}
	return "", false
}

// func (zf *file) walk(node *Node, w io.Writer) error {
// 	switch node.XMLName.Local {
// 	case "hyperlink":
// 		fmt.Fprint(w, "[")
// 		var cbuf bytes.Buffer
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, &cbuf); err != nil {
// 				return err
// 			}
// 		}
// 		fmt.Fprint(w, escape(cbuf.String(), "[]"))
// 		fmt.Fprint(w, "]")

// 		fmt.Fprint(w, "(")
// 		if id, ok := attr(node.Attrs, "id"); ok {
// 			for _, rel := range zf.rels.Relationship {
// 				if id == rel.ID {
// 					fmt.Fprint(w, escape(rel.Target, "()"))
// 					break
// 				}
// 			}
// 		}
// 		fmt.Fprint(w, ")")
// 	case "t":
// 		fmt.Fprint(w, string(node.Content))
// 	case "pPr":
// 		code := false
// 		for _, n := range node.Nodes {
// 			switch n.XMLName.Local {
// 			case "ind":
// 				if left, ok := attr(n.Attrs, "left"); ok {
// 					if i, err := strconv.Atoi(left); err == nil && i > 0 {
// 						fmt.Fprint(w, strings.Repeat("  ", i/360))
// 					}
// 				}
// 			case "pStyle":
// 				if val, ok := attr(n.Attrs, "val"); ok {
// 					if strings.HasPrefix(val, "Heading") {
// 						if i, err := strconv.Atoi(val[7:]); err == nil && i > 0 {
// 							fmt.Fprint(w, strings.Repeat("#", i)+" ")
// 						}
// 					} else if val == "Code" {
// 						code = true
// 					} else {
// 						if i, err := strconv.Atoi(val); err == nil && i > 0 {
// 							fmt.Fprint(w, strings.Repeat("#", i)+" ")
// 						}
// 					}
// 				}
// 			case "numPr":
// 				var numID, ilvl, numFmt string
// 				var start, ind int = 1, 0
// 				// numID := ""
// 				// ilvl := ""
// 				// numFmt := ""
// 				// start := 1
// 				// ind := 0
// 				for _, nn := range n.Nodes {
// 					if nn.XMLName.Local == "numId" {
// 						if val, ok := attr(nn.Attrs, "val"); ok {
// 							numID = val
// 						}
// 					}
// 					if nn.XMLName.Local == "ilvl" {
// 						if val, ok := attr(nn.Attrs, "val"); ok {
// 							ilvl = val
// 						}
// 					}
// 				}
// 				for _, num := range zf.num.Num {
// 					if numID != num.NumID {
// 						continue
// 					}
// 					for _, abnum := range zf.num.AbstractNum {
// 						if abnum.AbstractNumID != num.AbstractNumID.Val {
// 							continue
// 						}
// 						for _, ablvl := range abnum.Lvl {
// 							if ablvl.Ilvl != ilvl {
// 								continue
// 							}
// 							if i, err := strconv.Atoi(ablvl.Start.Val); err == nil {
// 								start = i
// 							}
// 							if i, err := strconv.Atoi(ablvl.PPr.Ind.Left); err == nil {
// 								ind = i / 360
// 							}
// 							numFmt = ablvl.NumFmt.Val
// 							break
// 						}
// 						break
// 					}
// 					break
// 				}

// 				fmt.Fprint(w, strings.Repeat("  ", ind))
// 				switch numFmt {
// 				case "decimal", "aiueoFullWidth":
// 					key := fmt.Sprintf("%s:%d", numID, ind)
// 					cur, ok := zf.list[key]
// 					if !ok {
// 						zf.list[key] = start
// 					} else {
// 						zf.list[key] = cur + 1
// 					}
// 					fmt.Fprintf(w, "%d. ", zf.list[key])
// 				case "bullet":
// 					fmt.Fprint(w, "* ")
// 				}
// 			}
// 		}
// 		if code {
// 			fmt.Fprint(w, "`")
// 		}
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, w); err != nil {
// 				return err
// 			}
// 		}
// 		if code {
// 			fmt.Fprint(w, "`")
// 		}
// 	case "tbl":
// 		var rows [][]string
// 		for _, tr := range node.Nodes {
// 			if tr.XMLName.Local != "tr" {
// 				continue
// 			}
// 			var cols []string
// 			for _, tc := range tr.Nodes {
// 				if tc.XMLName.Local != "tc" {
// 					continue
// 				}
// 				var cbuf bytes.Buffer
// 				if err := zf.walk(&tc, &cbuf); err != nil {
// 					return err
// 				}
// 				cols = append(cols, strings.Replace(cbuf.String(), "\n", "", -1))
// 			}
// 			rows = append(rows, cols)
// 		}
// 		maxcol := 0
// 		for _, cols := range rows {
// 			if len(cols) > maxcol {
// 				maxcol = len(cols)
// 			}
// 		}
// 		widths := make([]int, maxcol)
// 		for _, row := range rows {
// 			for i := 0; i < maxcol; i++ {
// 				if i < len(row) {
// 					width := runewidth.StringWidth(row[i])
// 					if widths[i] < width {
// 						widths[i] = width
// 					}
// 				}
// 			}
// 		}
// 		for i, row := range rows {
// 			if i == 0 {
// 				for j := 0; j < maxcol; j++ {
// 					fmt.Fprint(w, "|")
// 					fmt.Fprint(w, strings.Repeat(" ", widths[j]))
// 				}
// 				fmt.Fprint(w, "|\n")
// 				for j := 0; j < maxcol; j++ {
// 					fmt.Fprint(w, "|")
// 					fmt.Fprint(w, strings.Repeat("-", widths[j]))
// 				}
// 				fmt.Fprint(w, "|\n")
// 			}
// 			for j := 0; j < maxcol; j++ {
// 				fmt.Fprint(w, "|")
// 				if j < len(row) {
// 					width := runewidth.StringWidth(row[j])
// 					fmt.Fprint(w, escape(row[j], "|"))
// 					fmt.Fprint(w, strings.Repeat(" ", widths[j]-width))
// 				} else {
// 					fmt.Fprint(w, strings.Repeat(" ", widths[j]))
// 				}
// 			}
// 			fmt.Fprint(w, "|\n")
// 		}
// 		fmt.Fprint(w, "\n")
// 	case "r":
// 		// var bold, italic, strike bool
// 		for _, n := range node.Nodes {
// 			if n.XMLName.Local != "rPr" {
// 				continue
// 			}
// 			// for _, nn := range n.Nodes {
// 			// 	switch nn.XMLName.Local {
// 			// 	case "b":
// 			// 		bold = true
// 			// 	case "i":
// 			// 		italic = true
// 			// 	case "strike":
// 			// 		strike = true
// 			// 	}
// 			// }
// 		}
// 		// if strike {
// 		// 	fmt.Fprint(w, "~~")
// 		// }
// 		// if bold {
// 		// 	fmt.Fprint(w, "**")
// 		// }
// 		// if italic {
// 		// 	fmt.Fprint(w, "*")
// 		// }
// 		var cbuf bytes.Buffer
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, &cbuf); err != nil {
// 				return err
// 			}
// 		}
// 		fmt.Fprint(w, escape(cbuf.String(), `*~\`))
// 		// if italic {
// 		// 	fmt.Fprint(w, "*")
// 		// }
// 		// if bold {
// 		// 	fmt.Fprint(w, "**")
// 		// }
// 		// if strike {
// 		// 	fmt.Fprint(w, "~~")
// 		// }
// 	case "p":
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, w); err != nil {
// 				return err
// 			}
// 		}
// 		fmt.Fprintln(w)
// 	case "blip":
// 		if id, ok := attr(node.Attrs, "embed"); ok {
// 			for _, rel := range zf.rels.Relationship {
// 				if id != rel.ID {
// 					continue
// 				}
// 				if err := zf.extract(&rel, w); err != nil {
// 					return err
// 				}
// 			}
// 		}
// 	case "Fallback":
// 	case "txbxContent":
// 		var cbuf bytes.Buffer
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, &cbuf); err != nil {
// 				return err
// 			}
// 		}
// 		fmt.Fprintln(w, "\n```\n"+cbuf.String()+"```")
// 	default:
// 		for _, n := range node.Nodes {
// 			if err := zf.walk(&n, w); err != nil {
// 				return err
// 			}
// 		}
// 	}

// 	return nil
// }

func readFile(f *zip.File) (*Node, error) {
	rc, err := f.Open()
	defer rc.Close()

	b, _ := io.ReadAll(rc)
	if err != nil {
		return nil, err
	}

	os.WriteFile("test.xml", b, 0644)

	var node Node
	err = xml.Unmarshal(b, &node)
	if err != nil {
		return nil, err
	}
	return &node, nil
}

func findFile(files []*zip.File, target string) *zip.File {
	for _, f := range files {
		if ok, _ := path.Match(target, f.Name); ok {
			return f
		}
	}
	return nil
}

func Docx2txt(arg string, embed bool, options ...OptionsFunc) (*bytes.Buffer, error) {
	opts := NewOptions(options...)
	r, err := zip.OpenReader(arg)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	var rels Relationships
	var num Numbering

	for _, f := range r.File {
		switch f.Name {
		case "word/_rels/document.xml.rels":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}

			err = xml.Unmarshal(b, &rels)
			if err != nil {
				return nil, err
			}
		case "word/numbering.xml":
			rc, err := f.Open()
			defer rc.Close()

			b, _ := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}

			err = xml.Unmarshal(b, &num)
			if err != nil {
				return nil, err
			}
		}
	}

	f := findFile(r.File, "word/document*.xml")
	if f == nil {
		return nil, errors.New("incorrect document")
	}
	node, err := readFile(f)
	if err != nil {
		return nil, err
	}

	var buf bytes.Buffer
	zf := &file{
		r:     r,
		rels:  rels,
		num:   num,
		embed: embed,
		list:  make(map[string]int),
		opts:  opts,
	}

	// w := io.MultiWriter(&buf, ifTernary(opts.IsDebug, os.Stdout, nil))
	w := ifTernary(opts.IsDebug, io.MultiWriter(&buf, os.Stdout), io.MultiWriter(&buf))
	err = zf.walk(node, w)
	// err = zf.walk(node, &buf)
	if err != nil {
		return nil, err
	}

	str := unescapeHTMLEntities(buf.String())
	str = replaceHashSpaces(str)
	buf.Reset()
	buf.WriteString(str)

	return &buf, nil
}

var reSpace = regexp.MustCompile(`>\s+<`)

func (zf *file) walk(node *Node, w io.Writer) error {
	ff := zf.opts.Logger.Debugf
	switch node.XMLName.Local {
	case "hyperlink":
		fmt.Fprint(w, "[")
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		fmt.Fprint(w, escape(cbuf.String(), "[]"))
		fmt.Fprint(w, "]")

		fmt.Fprint(w, "(")
		if id, ok := attr(node.Attrs, "id"); ok {
			for _, rel := range zf.rels.Relationship {
				if id == rel.ID {
					fmt.Fprint(w, escape(rel.Target, "()"))
					break
				}
			}
		}
		fmt.Fprint(w, ")")
	case "t":
		fmt.Fprint(w, string(node.Content))
	case "pPr":
		code := false
		for _, n := range node.Nodes {
			switch n.XMLName.Local {
			case "ind":
				if left, ok := attr(n.Attrs, "left"); ok {
					if i, err := strconv.Atoi(left); err == nil && i > 0 {
						fmt.Fprint(w, strings.Repeat("  ", i/360))
					}
				}
			case "pStyle":
				if val, ok := attr(n.Attrs, "val"); ok {
					if strings.HasPrefix(val, "Heading") {
						if i, err := strconv.Atoi(val[7:]); err == nil && i > 0 {
							fmt.Fprint(w, strings.Repeat("#", i)+" ")
						}
					} else if val == "Code" {
						code = true
					} else {
						if i, err := strconv.Atoi(val); err == nil && i > 0 {
							fmt.Fprint(w, strings.Repeat("#", i)+" ")
						}
					}
				}
			case "numPr":
				var numID, ilvl, numFmt string
				var start, ind int = 1, 0
				for _, nn := range n.Nodes {
					if nn.XMLName.Local == "numId" {
						if val, ok := attr(nn.Attrs, "val"); ok {
							numID = val
						}
					}
					if nn.XMLName.Local == "ilvl" {
						if val, ok := attr(nn.Attrs, "val"); ok {
							ilvl = val
						}
					}
				}
				for _, num := range zf.num.Num {
					if numID != num.NumID {
						continue
					}
					for _, abnum := range zf.num.AbstractNum {
						if abnum.AbstractNumID != num.AbstractNumID.Val {
							continue
						}
						for _, ablvl := range abnum.Lvl {
							if ablvl.Ilvl != ilvl {
								continue
							}
							if i, err := strconv.Atoi(ablvl.Start.Val); err == nil {
								start = i
							}
							if i, err := strconv.Atoi(ablvl.PPr.Ind.Left); err == nil {
								ind = i / 360
							}
							numFmt = ablvl.NumFmt.Val
							break
						}
						break
					}
					break
				}

				fmt.Fprint(w, strings.Repeat("  ", ind))
				switch numFmt {
				case "decimal", "aiueoFullWidth":
					ff(">>>>>>>>>>>>  NUMFMT !!!!!!!!!")
					key := fmt.Sprintf("%s:%d", numID, ind)
					cur, ok := zf.list[key]
					if !ok {
						zf.list[key] = start
					} else {
						zf.list[key] = cur + 1
					}
					fmt.Fprintf(w, "%d. ", zf.list[key])
				case "bullet":
					fmt.Fprint(w, "* ")
				}
			}
		}
		if code {
			fmt.Fprint(w, "`")
		}
		for _, n := range node.Nodes {
			// ff("\n\n >>>>>>>>>>>>  numPr !!!!!!!!! content=%v NODE=%+v\n", string(n.Content), n)
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
		if code {
			fmt.Fprint(w, "`")
		}
		// fmt.Fprint(w, "<<##>>")
	case "tbl":
		var rows [][]string
		for _, tr := range node.Nodes {
			if tr.XMLName.Local != "tr" {
				continue
			}
			var cols []string
			for _, tc := range tr.Nodes {
				if tc.XMLName.Local != "tc" {
					continue
				}
				var cbuf bytes.Buffer
				if err := zf.walk(&tc, &cbuf); err != nil {
					return err
				}
				ff(">>>>>>>>>>>>  CBUF:[%s]", unescapeHTMLEntities(cbuf.String()))
				// zf.opts.Logger.Debugf("\n\n >>>>>>>>>>>>  CBUF:[%s]\n", unescapeHTMLEntities(cbuf.String()))
				// cols = append(cols, cbuf.String())
				bufStr := unescapeHTMLEntities(cbuf.String())
				bufStr = strings.Replace(bufStr, "\n", "", -1)
				bufStr = reSpace.ReplaceAllString(bufStr, "><")
				cols = append(cols, bufStr)
				// cols = append(cols, strings.Replace(cbuf.String(), "\n", "", -1))
				// cols = append(cols, strings.NewReplacer("\n", "", "\t", "").Replace(cbuf.String()))
			}
			rows = append(rows, cols)
		}

		maxcol := 0
		for _, cols := range rows {
			if len(cols) > maxcol {
				maxcol = len(cols)
			}
		}
		widths := make([]int, maxcol)
		for _, row := range rows {
			for i := 0; i < maxcol; i++ {
				if i < len(row) {
					width := runewidth.StringWidth(row[i])
					if widths[i] < width {
						widths[i] = width
					}
				}
			}
		}

		ff("maxcol: %d; width: %v; len_rows=%d", maxcol, widths, len(rows))
		ff("rows: %v", rows)

		for i, row := range rows {

			for j := 0; j < maxcol; j++ {
				// ff("\nj=%d; maxcol=%d; len(row)=%d; row: %v\n", j, maxcol, len(row), row)
				// fmt.Fprint(w, "|")
				switch zf.opts.StyleTbls {
				case "csv":
				case "pretty", "md":
					fmt.Fprint(w, "|")
				}

				if j < len(row) {
					rj := escape(row[j], "|")
					width := runewidth.StringWidth(row[j])

					fmt.Fprint(w, wrapIf2(rj, "\"", condCommaCotains(rj), zf.opts.isCSV()))
					// fmt.Fprint(w, wrapIf(rj, "\"", func(v interface{}) bool {
					// 	return condCommaCotains(v.(string))
					// }))

					if zf.opts.isCSV() && j != len(row)-1 {
						fmt.Fprint(w, ",")
					} else if zf.opts.isPrettyOrMD() {
						fmt.Fprint(w, strings.Repeat(" ", widths[j]-width))
					}
				} else {
					// fp("\n\n\n >>>>>>>>>>>>  JJJJJJJJJ !!!!!!!!!\n\n\n")
					fmt.Fprint(w, strings.Repeat(" ", widths[j]))
				}
			}
			switch {
			case zf.opts.isCSV():
				fmt.Fprint(w, "\n")
			case zf.opts.isPrettyOrMD() && i != 0:
				fmt.Fprint(w, "|\n")
			}

			if i == 0 && zf.opts.isPrettyOrMD() && len(rows) > 1 {
				// for j := 0; j < maxcol; j++ {
				// 	fmt.Fprint(w, "|")
				// 	fmt.Fprint(w, strings.Repeat(" ", widths[j]))
				// }
				fmt.Fprint(w, "|\n")
				for j := 0; j < maxcol; j++ {
					fmt.Fprint(w, "|")
					fmt.Fprint(w, strings.Repeat("-", widths[j]))
				}
				fmt.Fprint(w, "|\n")
				ff("FIRST LINE PASS: len=%d; row: %v", len(row), row)
			}
		}
		fmt.Fprint(w, "\n")
	case "r":
		// fp("\n\n\n >>>>>>>>>>>>  RRRRRRRRR !!!!!!!!!\n\n\n")
		for _, n := range node.Nodes {
			if n.XMLName.Local != "rPr" {
				continue
			}
		}
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		// ff("\n>>>>>>>  CBUF:[%s]\n", unescapeHTMLEntities(cbuf.String()))
		fmt.Fprint(w, escape(cbuf.String(), `*~\`))
	case "p":
		for _, n := range node.Nodes {
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
		fmt.Fprintln(w)
	case "blip":
		if id, ok := attr(node.Attrs, "embed"); ok {
			for _, rel := range zf.rels.Relationship {
				if id != rel.ID {
					continue
				}
				if err := zf.extract(&rel, w); err != nil {
					return err
				}
			}
		}
	case "Fallback":
	case "txbxContent":
		var cbuf bytes.Buffer
		for _, n := range node.Nodes {
			if err := zf.walk(&n, &cbuf); err != nil {
				return err
			}
		}
		fmt.Fprintln(w, "\n```\n"+cbuf.String()+"```")
	default:
		for _, n := range node.Nodes {
			if err := zf.walk(&n, w); err != nil {
				return err
			}
		}
	}

	return nil
}

func wrapIf2(s string, wrapper string, cond ...bool) string {
	for _, c := range cond {
		if !c {
			return s
		}
	}
	return wrapper + s + wrapper
}

func condCommaCotains(s any) bool {
	return strings.Contains(s.(string), ",")
}

// unescapeHTMLEntities converts HTML entities in the input string to their actual characters.
func unescapeHTMLEntities(input string) string {
	return html.UnescapeString(input)
}

// Function to replace spaces between # and text
func replaceHashSpaces(input string) string {
	re := regexp.MustCompile(`#\s+`)
	return re.ReplaceAllString(input, "# ")
}

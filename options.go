package docx2txt

import (
	"os"
)

type Options struct {
	StyleTbls string // "csv" or "pretty" or "md"
	IsDebug   bool
	Logger    Logger
}

func NewOptions(opts ...OptionsFunc) *Options {
	// sl := slog.NewLogLogger(slog.NewTextHandler(os.Stdout), slog.LogLevel(slog.LevelInfo))
	// slog.NewLogLogger(, level slog.Level)
	o := &Options{
		StyleTbls: "pretty",
		IsDebug:   false,
		// Logger:    NewLoggerCustom(),
		// Logger:    NewLoggerCustom(slog.Default()),
	}
	for _, opt := range opts {
		opt(o)
	}
	if o.Logger == nil {
		o.Logger = NewLoggerCustom(os.Stdout, o.IsDebug)
	}
	return o
}

func (o *Options) isCSV() bool {
	return o.StyleTbls == "csv"
}
func (o *Options) isPrettyOrMD() bool {
	return o.StyleTbls == "pretty" || o.StyleTbls == "md"
}

type OptionsFunc func(*Options)

func StyleTbls(style string) OptionsFunc {
	return func(o *Options) {
		o.StyleTbls = style
	}
}

func WithDebug(b bool) OptionsFunc {
	return func(o *Options) {
		o.IsDebug = b
	}
}

func WithLogger(l Logger) OptionsFunc {
	return func(o *Options) {
		o.Logger = l
	}
}

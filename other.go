package docx2txt

import (
	"fmt"
	"io"
	"log/slog"
)

func ifTernary[T any](cond bool, a, b T) T {
	if cond {
		return a
	}
	return b
}

type LoggerCustom struct {
	*slog.Logger
}

func NewLoggerCustom(w io.Writer, isDebug bool) *LoggerCustom {
	return &LoggerCustom{slog.New(slog.NewTextHandler(w, &slog.HandlerOptions{Level: ifTernary(isDebug, slog.LevelDebug, slog.LevelInfo)}))}
}

func (l *LoggerCustom) Debugf(format string, args ...any) {
	l.Logger.Debug(fmt.Sprintf(format, args...))
}

func (l *LoggerCustom) Infof(format string, args ...any) {
	l.Logger.Info(fmt.Sprintf(format, args...))
}

func (l *LoggerCustom) Errorf(format string, args ...any) {
	l.Logger.Error(fmt.Sprintf(format, args...))
}

// func (l *LoggerCustom) Debug(args ...any) {
// 	l.Debug(args...)
// }

// func (l *LoggerCustom) Info(args ...any) {
// 	l.Info(args...)
// }

// func (l *LoggerCustom) Error(args ...any) {
// 	l.Error(args...)
// }

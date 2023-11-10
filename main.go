package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	_ "go.uber.org/zap"
	"go.uber.org/zap/zapcore"
	"gopkg.in/natefinch/lumberjack.v2"
	"log/slog"
	"os"
	"path"
)

type Path struct {
	hz string
	mx string
}

var myPath Path

type ExcelMatch struct {
	srcPath string
	objPath string
	out     string
	cache   map[string]string
}

func (e *ExcelMatch) open(path string) (*excelize.File, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return f, nil
}

func (e *ExcelMatch) match() error {
	f, err := e.open(e.objPath)
	if err != nil {
		return err
	}

	for i, sheetName := range f.GetSheetList() {
		f.SetActiveSheet(i)
		err = f.InsertCols(sheetName, "J", 1)
		if err != nil {
			slog.Error("插入J列失败", "err", err)
		}

		err = f.SetCellStr(sheetName, "J1", "IB编码")
		if err != nil {
			slog.Error("设置表头", "err", err)
		}

		style, err := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{Type: "pattern", Color: []string{"E0EBF5"}, Pattern: 1},
		})

		if err != nil {
			slog.Error("创建样式", "err", err)
		}

		err = f.SetCellStyle(sheetName, "J1", "J1", style)
		if err != nil {
			slog.Error("设置样式", "err", err)
		}

		styleBorder, err := f.NewStyle(&excelize.Style{
			Border: []excelize.Border{
				{"left", "000000", 1},
				{"top", "000000", 1},
				{"right", "000000", 1},
				{"bottom", "000000", 1},
			},
		})

		if err != nil {
			slog.Error("创建样式", "err", err)

		}

		row, err := f.Rows(sheetName)
		if err != nil {
			slog.Error("f.rows()", "sheetName", sheetName)
		}

		row.Next()
		rowIndex := 2
		IBCode := ""
		for row.Next() {
			r, err := row.Columns()
			if err != nil {
				slog.Error("获取列", "err", err)
			}

			if len(r) >= 9 {
				IBCode = e.cache[r[8]]
			} else {
				IBCode = ""
			}

			cell := fmt.Sprintf("J%d", rowIndex)
			err = f.SetCellStr(sheetName, cell, IBCode)
			if err != nil {
				slog.Error("设置cell的值", "cell", fmt.Sprintf("J%d", rowIndex), "err", err)
			}

			slog.Debug("设置cell的值", "cell", fmt.Sprintf("J%d", rowIndex), "rowIndex", rowIndex)

			err = f.SetCellStyle(sheetName, cell, cell, styleBorder)
			if err != nil {
				slog.Error("设置cell的样式", "err", err)
			}

			rowIndex++
		}
	}

	dirPath, filename := path.Split(e.objPath)

	newFile := path.Join(dirPath, filename[:len(filename)-5]+"-new"+filename[len(filename)-5:])

	err = f.SaveAs(newFile)
	if err != nil {
		slog.Error("保存文件失败:", "err", err)
	} else {
		slog.Info("保存文件成功:", "filePath", newFile)
	}

	err = f.Close()
	if err != nil {
		slog.Error("关闭文件失败:", "err", err)
	}

	return nil
}

func (e *ExcelMatch) readExcel(srcPath string) (int, error) {
	e.cache = make(map[string]string)
	f, err := excelize.OpenFile(srcPath)
	if err != nil {
		slog.Error("打开文件失败", "err", err)
		panic(err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			slog.Error("关闭文件失败", "err", err)
		}
	}()

	rows, err := f.GetRows("汇总")
	if err != nil {
		return 0, fmt.Errorf("读取行报错: %v", err)
	}
	for _, row := range rows {
		e.cache[row[6]] = row[7]
	}
	return len(e.cache), nil
}

func newExcel(srcPath, objPath string) *ExcelMatch {
	return &ExcelMatch{srcPath: srcPath, objPath: objPath}
}

func getLogger(infoPath, errPath string) (*zap.Logger, error) {
	highPriority := zap.LevelEnablerFunc(func(level zapcore.Level) bool {
		return level >= zap.ErrorLevel
	})

	lowPriority := zap.LevelEnablerFunc(func(level zapcore.Level) bool {
		return level < zap.ErrorLevel && level >= zap.DebugLevel
	})

	prodEncoder := zap.NewProductionEncoderConfig()
	prodEncoder.EncodeTime = zapcore.TimeEncoderOfLayout("2006-01-02 15:04:05.999")

	infoHandler := &lumberjack.Logger{
		Filename:   infoPath,
		LocalTime:  true,
		MaxSize:    1,
		MaxAge:     3,
		MaxBackups: 5,
		Compress:   false,
	}

	errHandler := &lumberjack.Logger{
		Filename:   errPath,
		LocalTime:  true,
		MaxSize:    10, //10MB
		MaxAge:     60,
		MaxBackups: 300,
		Compress:   false,
	}

	lowWriteSyncer := zapcore.AddSync(infoHandler)
	highWriterSyncer := zapcore.AddSync(errHandler)

	highCore := zapcore.NewCore(zapcore.NewJSONEncoder(prodEncoder), highWriterSyncer, highPriority)
	lowCore := zapcore.NewCore(zapcore.NewJSONEncoder(prodEncoder), lowWriteSyncer, lowPriority)

	return zap.New(zapcore.NewTee(highCore, lowCore), zap.AddCaller()), nil
}

func main() {

	log, err := getLogger("info.log", "err.log")
	if err != nil {
		panic(err)
	}

	flag.StringVar(&myPath.hz, "hz", "", "")
	flag.StringVar(&myPath.mx, "mx", "", "")
	flag.Parse()

	if myPath.hz == "" {
		log.Info("请输入汇总文件路径")
		os.Exit(-1)
	}

	if myPath.mx == "" {
		log.Info("请输入明细文件路径")
		os.Exit(-1)
	}

	log.Info("汇总文件-匹配的路径:", zap.String("汇总", myPath.hz))
	log.Info("明细文件-匹配的路径:", zap.String("明细", myPath.mx))

	matcher := newExcel(myPath.hz, myPath.mx)
	n, err := matcher.readExcel(myPath.hz)
	if err != nil {
		log.Error("读取汇总文件失败", zap.Error(err))
		os.Exit(-1)
	}

	log.Info("读取汇总文件总共行数:", zap.Int("row", n))

	err = matcher.match()
	if err != nil {
		log.Error("匹配错误", zap.Error(err))
	}

}

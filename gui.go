package main

import (
	"bytes"
	_ "embed"
	"encoding/json"
	"fmt"
	"image/png"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

//go:embed qrcode.png
var embeddedGuideLogo []byte

func loadGuideLogoImage() (walk.Image, error) {
	img, err := png.Decode(bytes.NewReader(embeddedGuideLogo))
	if err != nil {
		return nil, err
	}
	return walk.NewBitmapFromImage(img)
}

// Config 结构体用于存储配置信息
// 保持 JSON tag 不变，兼容你已有的 config.json 文件
type Config struct {
	AddTocCombine           bool   `json:"add_toc_combine"`
	RowsCombine             string `json:"rows_combine"`
	ChangePageCombine       bool   `json:"change_page_combine"`
	OutPathCombine          string `json:"out_path_combine"`
	OutFileCombine          string `json:"out_file_combine"`
	OutPathCombineNonZaiLab string `json:"out_path_combine_NonZaiLab"`
	OutFileCombineNonZaiLab string `json:"out_file_combine_NonZaiLab"`
	OutPathCombineDocx      string `json:"out_path_combine_Docx"`
	OutFileCombineDocx      string `json:"out_file_combine_Docx"`
	RtfPathCheck            string `json:"rtf_path_check"`
	PdfBoxConvert           bool   `json:"pdf_box_convert"`
	DocxBoxConvert          bool   `json:"docx_box_convert"`
	RtfFileConvert          string `json:"rtf_file_convert"`
	DocxFileConvert         string `json:"docx_file_convert"`
}

type FileItem struct {
	Name      string
	Checked   bool
	FullPath  string
	ParentDir string
}

type ResultFileItem struct {
	Index    int
	Name     string
	FullPath string
}

// ==========================================================
// 核心重构：提取 Tab1 和 Tab3 的公共数据和组件
// ==========================================================
type SharedTabElements struct {
	fileModel       *FileModel
	resultModel     *ResultFileModel
	addedFolders    map[string]bool
	tableView       *walk.TableView
	resultTableView *walk.TableView
	folderBtn       *walk.PushButton
	selectAllBtn    *walk.PushButton
	unselectAllBtn  *walk.PushButton
	confirmBtn      *walk.PushButton
	clearBtn        *walk.PushButton
}

// Tab1Data 继承 SharedTabElements，并添加专属组件
type Tab1Data struct {
	SharedTabElements
	addTocCheckBoxCombine     *walk.CheckBox
	rowsEditCombine           *walk.LineEdit
	changePageCheckBoxCombine *walk.CheckBox
	outPathEditCombine        *walk.LineEdit
	outFileEditCombine        *walk.LineEdit
	progressBarCombine        *walk.ProgressBar
}

// Tab3Data 继承 SharedTabElements，并添加专属组件
type Tab3Data struct {
	SharedTabElements
	outPathEditCombineDocx *walk.LineEdit
	outFileEditCombineDocx *walk.LineEdit
}

type MyMainWindow struct {
	*walk.MainWindow
	tab1                *Tab1Data
	tab3                *Tab3Data
	config              *Config
	rtfPathEditCheck    *walk.LineEdit
	progressBarCheck    *walk.ProgressBar
	logViewCheck        *walk.TextEdit
	convertPdf          *walk.CheckBox
	convertDocx         *walk.CheckBox
	rtfFileEditConvert  *walk.LineEdit
	docxFileEditConvert *walk.LineEdit
	logViewCovert       *walk.TextEdit
	logViewCombineDocx  *walk.TextEdit
}

// ResultFileModel 用于显示已选择文件的模型
type ResultFileModel struct {
	walk.TableModelBase
	items []*ResultFileItem
}

func NewResultFileModel() *ResultFileModel {
	return &ResultFileModel{
		items: make([]*ResultFileItem, 0),
	}
}

func (m *ResultFileModel) RowCount() int { return len(m.items) }

func (m *ResultFileModel) Value(row, col int) interface{} {
	item := m.items[row]
	switch col {
	case 0:
		return item.Index
	case 1:
		return item.Name
	case 2:
		return item.FullPath
	default:
		return ""
	}
}

func (m *ResultFileModel) Items() []*ResultFileItem { return m.items }

func (m *ResultFileModel) AddItems(files []*ResultFileItem) {
	m.items = files
	m.PublishRowsReset()
}

func (m *ResultFileModel) MoveItem(from, to int) {
	if from == to || from < 0 || from >= len(m.items) || to < 0 || to >= len(m.items) {
		return
	}
	item := m.items[from]
	m.items = append(m.items[:from], m.items[from+1:]...)
	m.items = append(m.items[:to], append([]*ResultFileItem{item}, m.items[to:]...)...)
	for i, item := range m.items {
		item.Index = i + 1
	}
	m.PublishRowsReset()
}

func (m *ResultFileModel) Clear() {
	m.items = make([]*ResultFileItem, 0)
	m.PublishRowsReset()
}

// saveConfig 将配置保存到文件
func saveConfig(cfg *Config) error {
	configDir, err := os.UserConfigDir()
	if err != nil {
		return err
	}
	appConfigDir := filepath.Join(configDir, "RTF_Tools")
	if err := os.MkdirAll(appConfigDir, 0755); err != nil {
		return err
	}
	configPath := filepath.Join(appConfigDir, "config.json")

	file, err := os.Create(configPath)
	if err != nil {
		return err
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	return encoder.Encode(cfg)
}

type FileModel struct {
	walk.TableModelBase
	walk.SorterBase
	items []*FileItem
}

func NewFileModel() *FileModel {
	return &FileModel{
		items: make([]*FileItem, 0),
	}
}

func (m *FileModel) RowCount() int { return len(m.items) }

func (m *FileModel) Value(row, col int) interface{} {
	item := m.items[row]
	switch col {
	case 0:
		return item.Name
	case 1:
		return item.ParentDir
	case 2:
		return item.FullPath
	default:
		return ""
	}
}

func (m *FileModel) Checked(row int) bool { return m.items[row].Checked }

func (m *FileModel) SetChecked(row int, checked bool) error {
	m.items[row].Checked = checked
	return nil
}

// ==========================================================
// 公共的 Tab 逻辑方法（适用于 Tab1 和 Tab3）
// ==========================================================

func (tab *SharedTabElements) setAllChecked(checked bool) {
	if tab.fileModel == nil || len(tab.fileModel.items) == 0 {
		return
	}
	for i := range tab.fileModel.items {
		tab.fileModel.items[i].Checked = checked
		tab.fileModel.PublishRowChanged(i)
	}
	tab.fileModel.PublishRowsReset()
	if tab.tableView != nil {
		tab.tableView.Invalidate()
	}
}

func (tab *SharedTabElements) invertSelection() {
	if tab.fileModel == nil || len(tab.fileModel.items) == 0 {
		return
	}
	for i := range tab.fileModel.items {
		tab.fileModel.items[i].Checked = !tab.fileModel.items[i].Checked
		tab.fileModel.PublishRowChanged(i)
	}
	tab.fileModel.PublishRowsReset()
	if tab.tableView != nil {
		tab.tableView.Invalidate()
	}
}

func (tab *SharedTabElements) clearAllFiles() {
	tab.fileModel.items = nil
	tab.addedFolders = make(map[string]bool)
	tab.resultModel.Clear()
	tab.fileModel.PublishRowsReset()
	if tab.tableView != nil {
		tab.tableView.Invalidate()
	}
}

func (tab *SharedTabElements) clearTextResult() {
	tab.resultModel.Clear()
}

func (tab *SharedTabElements) moveSelectedUp() {
	indices := tab.resultTableView.SelectedIndexes()
	if len(indices) == 0 {
		return
	}
	sort.Ints(indices)

	if indices[0] == 0 {
		return
	}

	oldItems := tab.resultModel.items
	newItems := make([]*ResultFileItem, 0, len(oldItems))

	blockStart := indices[0]
	blockEnd := indices[len(indices)-1]

	insertPos := blockStart - 1
	newItems = append(newItems, oldItems[:insertPos]...)

	for _, idx := range indices {
		newItems = append(newItems, oldItems[idx])
	}

	newItems = append(newItems, oldItems[insertPos])
	newItems = append(newItems, oldItems[blockEnd+1:]...)

	tab.resultModel.items = newItems
	for i, item := range tab.resultModel.items {
		item.Index = i + 1
	}
	tab.resultModel.PublishRowsReset()

	newSelected := make([]int, len(indices))
	for i, idx := range indices {
		newSelected[i] = idx - 1
	}
	tab.resultTableView.SetSelectedIndexes(newSelected)
}

func (tab *SharedTabElements) moveSelectedDown() {
	indices := tab.resultTableView.SelectedIndexes()
	if len(indices) == 0 {
		return
	}
	sort.Ints(indices)

	lastIndex := len(tab.resultModel.items) - 1
	if indices[len(indices)-1] >= lastIndex {
		return
	}

	oldItems := tab.resultModel.items
	newItems := make([]*ResultFileItem, 0, len(oldItems))

	blockStart := indices[0]
	blockEnd := indices[len(indices)-1]

	newItems = append(newItems, oldItems[:blockStart]...)
	newItems = append(newItems, oldItems[blockEnd+1])
	for _, idx := range indices {
		newItems = append(newItems, oldItems[idx])
	}
	if blockEnd+2 < len(oldItems) {
		newItems = append(newItems, oldItems[blockEnd+2:]...)
	}

	tab.resultModel.items = newItems
	for i, item := range tab.resultModel.items {
		item.Index = i + 1
	}
	tab.resultModel.PublishRowsReset()

	newSelected := make([]int, len(indices))
	for i, idx := range indices {
		newSelected[i] = idx + 1
	}
	tab.resultTableView.SetSelectedIndexes(newSelected)
}

func (tab *SharedTabElements) loadFilesFromFolder(folderPath string, mw *MyMainWindow) {
	if tab.addedFolders[folderPath] {
		walk.MsgBox(mw, "Tips", "This folder has already been added!", walk.MsgBoxIconInformation)
		return
	}

	entries, err := os.ReadDir(folderPath)
	if err != nil {
		walk.MsgBox(mw, "Error", "Unable to read folder:"+err.Error(), walk.MsgBoxIconError)
		return
	}

	fileCount := 0
	for _, entry := range entries {
		if !entry.IsDir() {
			fullPath := filepath.Join(folderPath, entry.Name())
			parentDir := filepath.Base(folderPath)

			tab.fileModel.items = append(tab.fileModel.items, &FileItem{
				Name:      entry.Name(),
				Checked:   false,
				FullPath:  fullPath,
				ParentDir: parentDir,
			})
			fileCount++
		}
	}

	tab.addedFolders[folderPath] = true
	tab.fileModel.PublishRowsReset()
	if tab.tableView != nil {
		tab.tableView.Invalidate()
	}

	walk.MsgBox(mw, "Successfully", fmt.Sprintf(" %d Files added (from: %s)", fileCount, filepath.Base(folderPath)), walk.MsgBoxIconInformation)
}

// showSelectedFiles 重构：接收允许的后缀列表进行过滤
func (tab *SharedTabElements) showSelectedFiles(mw *MyMainWindow, allowedExts []string) {
	if tab.fileModel == nil || len(tab.fileModel.items) == 0 {
		walk.MsgBox(mw, "Tips", "Please select a folder first.", walk.MsgBoxIconInformation)
		return
	}

	type ProcessedFile struct {
		Name     string
		FullPath string
		Ord      int
		Filename string
		Name1    string
		SortNums []int
	}

	var selectedFiles []ProcessedFile
	for _, item := range tab.fileModel.items {
		if !item.Checked {
			continue
		}

		nameLower := strings.ToLower(item.Name)

		// 后缀检查
		isValidExt := false
		for _, ext := range allowedExts {
			if strings.HasSuffix(nameLower, strings.ToLower(ext)) {
				isValidExt = true
				break
			}
		}
		if !isValidExt || strings.HasPrefix(item.Name, "~") {
			continue
		}

		processedFile := ProcessedFile{
			Name:     item.Name,
			FullPath: item.FullPath,
			Filename: item.Name,
		}

		// 处理 Ord 分类
		if strings.HasPrefix(nameLower, "t") {
			processedFile.Ord = 1
		} else if strings.HasPrefix(nameLower, "f") {
			processedFile.Ord = 2
		} else if strings.HasPrefix(nameLower, "l") {
			processedFile.Ord = 3
		} else {
			processedFile.Ord = 999
		}

		// 处理 Name1 和 SortNums 用于排序
		re := regexp.MustCompile(`^[tTfFlL][-.]`)
		name1 := re.ReplaceAllString(item.Name, "")
		name1 = regexp.MustCompile(`[a-zA-Z]`).ReplaceAllString(name1, "")
		name1 = regexp.MustCompile(`^-+`).ReplaceAllString(name1, "")
		name1 = regexp.MustCompile(`-+`).ReplaceAllString(name1, "-")
		name1 = strings.ReplaceAll(name1, ".", "-")
		name1 = regexp.MustCompile(`(\d)[^\d]*$`).ReplaceAllString(name1, "$1")
		processedFile.Name1 = name1

		parts := strings.Split(name1, "-")
		for _, part := range parts {
			if part == "" {
				continue
			}
			num, err := strconv.Atoi(part)
			if err != nil {
				num = 9999
			}
			processedFile.SortNums = append(processedFile.SortNums, num)
		}

		selectedFiles = append(selectedFiles, processedFile)
	}

	if len(selectedFiles) == 0 {
		walk.MsgBox(mw, "Tips", "No valid files selected.", walk.MsgBoxIconInformation)
		tab.resultModel.Clear()
		return
	}

	// 排序逻辑
	sort.Slice(selectedFiles, func(i, j int) bool {
		if selectedFiles[i].Ord != selectedFiles[j].Ord {
			return selectedFiles[i].Ord < selectedFiles[j].Ord
		}
		maxLen := len(selectedFiles[i].SortNums)
		if len(selectedFiles[j].SortNums) > maxLen {
			maxLen = len(selectedFiles[j].SortNums)
		}
		for k := 0; k < maxLen; k++ {
			var numI, numJ int
			if k < len(selectedFiles[i].SortNums) {
				numI = selectedFiles[i].SortNums[k]
			}
			if k < len(selectedFiles[j].SortNums) {
				numJ = selectedFiles[j].SortNums[k]
			}
			if numI != numJ {
				return numI < numJ
			}
		}
		return selectedFiles[i].Name < selectedFiles[j].Name
	})

	var resultItems []*ResultFileItem
	for i, file := range selectedFiles {
		resultItems = append(resultItems, &ResultFileItem{
			Index:    i + 1,
			Name:     file.Name,
			FullPath: file.FullPath,
		})
	}

	tab.resultModel.AddItems(resultItems)
	walk.MsgBox(mw, "Completed", fmt.Sprintf(" %d Files selected", len(selectedFiles)), walk.MsgBoxIconInformation)
}

// ==========================================================
// MyMainWindow 相关方法
// ==========================================================

func (mw *MyMainWindow) selectDirectory() string {
	dlg := walk.FileDialog{Title: "文件选择器"}
	if ok, _ := dlg.ShowBrowseFolder(mw); !ok {
		return ""
	}
	return dlg.FilePath
}

func (mw *MyMainWindow) selectDirectoryWithInput() string {
	var dlg *walk.Dialog
	var pathEdit *walk.LineEdit
	var result string

	_ = Dialog{
		AssignTo: &dlg,
		Title:    "Select or enter a directory path.",
		MinSize:  Size{Width: 400, Height: 150},
		Layout:   VBox{Margins: Margins{Top: 10, Bottom: 10, Left: 10, Right: 10}, Spacing: 10},
		Children: []Widget{
			Label{Text: "Please select or enter a directory path:"},
			Composite{
				Layout: HBox{Spacing: 5, MarginsZero: true},
				Children: []Widget{
					LineEdit{AssignTo: &pathEdit},
					PushButton{
						Text: "Browse...",
						OnClicked: func() {
							if dir := mw.selectDirectory(); dir != "" {
								pathEdit.SetText(dir)
							}
						},
					},
				},
			},
			Composite{
				Layout: HBox{Spacing: 10, MarginsZero: true},
				Children: []Widget{
					HSpacer{},
					PushButton{
						Text: "Yes",
						OnClicked: func() {
							path := pathEdit.Text()
							if path != "" {
								if _, err := os.Stat(path); err != nil {
									walk.MsgBox(dlg, "Error", "The path does not exist or is inaccessible!", walk.MsgBoxIconError)
									return
								}
								result = path
							}
							dlg.Accept()
						},
					},
					PushButton{Text: "Cancel", OnClicked: func() { dlg.Cancel() }},
				},
			},
		},
	}.Create(mw)
	dlg.Run()
	return result
}

func (mw *MyMainWindow) selectFile(filter string) string {
	dlg := walk.FileDialog{Title: "Select Files", Filter: filter}
	if ok, _ := dlg.ShowOpen(mw); !ok {
		return ""
	}
	return dlg.FilePath
}

type selectModeDialog struct {
	*walk.Dialog
	result int
}

func showSelectModeDialog(owner walk.Form) (int, error) {
	var dlg *selectModeDialog
	dlg = new(selectModeDialog)
	dlg.result = 0

	err := Dialog{
		AssignTo: &dlg.Dialog,
		Title:    "Select: ",
		MinSize:  Size{Width: 300, Height: 150},
		Layout:   VBox{Margins: Margins{Top: 15, Bottom: 15, Left: 15, Right: 15}},
		Children: []Widget{
			Label{
				Text: "Select Single .docx or Multiple .docx files in folder path(Batch Mode)?",
				Font: Font{PointSize: 10, Bold: true},
			},
			Composite{
				Layout: HBox{MarginsZero: true},
				Children: []Widget{
					HSpacer{},
					PushButton{
						Text:      "Single Docx",
						OnClicked: func() { dlg.result = 1; dlg.Accept() },
					},
					PushButton{
						Text:      "Folder Path",
						OnClicked: func() { dlg.result = 2; dlg.Accept() },
					},
					PushButton{
						Text:      "Cancel",
						OnClicked: func() { dlg.result = 0; dlg.Cancel() },
					},
				},
			},
		},
	}.Create(owner)

	if err != nil {
		return 0, err
	}
	dlg.Run()
	return dlg.result, nil
}

func (mw *MyMainWindow) selectDocxFileOrFolder() string {
	mode, err := showSelectModeDialog(mw)
	if err != nil || mode == 0 {
		return ""
	}

	if mode == 1 {
		dlg := walk.FileDialog{
			Title:  "Select single .docx file",
			Filter: "Word document (*.docx)|*.docx",
		}
		if ok, _ := dlg.ShowOpen(mw); !ok {
			return ""
		}
		return dlg.FilePath
	}
	return mw.selectDirectoryWithInput()
}

func (mw *MyMainWindow) onItemActivated() {}

func loadConfig() (*Config, error) {
	configDir, err := os.UserConfigDir()
	if err != nil {
		return nil, err
	}
	appConfigDir := filepath.Join(configDir, "RTF_Tools")
	configPath := filepath.Join(appConfigDir, "config.json")

	file, err := os.Open(configPath)
	if err != nil {
		if os.IsNotExist(err) {
			return &Config{}, nil
		}
		return nil, err
	}
	defer file.Close()

	var cfg Config
	decoder := json.NewDecoder(file)
	if err = decoder.Decode(&cfg); err != nil {
		return nil, err
	}
	return &cfg, nil
}

// runSimpleGUI 是界面启动入口，供 main.go 调用
func runSimpleGUI() {
	cfg, err := loadConfig()
	if err != nil {
		log.Printf("Failed to load config, using default values: %v\n", err)
		cfg = &Config{}
	}

	guideLogoImage, errGuideLogo := loadGuideLogoImage()
	if errGuideLogo != nil {
		log.Printf("Guide tab logo not loaded: %v\n", errGuideLogo)
	}

	mw := &MyMainWindow{
		tab1: &Tab1Data{
			SharedTabElements: SharedTabElements{
				fileModel:    NewFileModel(),
				resultModel:  NewResultFileModel(),
				addedFolders: make(map[string]bool),
			},
		},
		tab3: &Tab3Data{
			SharedTabElements: SharedTabElements{
				fileModel:    NewFileModel(),
				resultModel:  NewResultFileModel(),
				addedFolders: make(map[string]bool),
			},
		},
		config: cfg,
	}

	if err := (MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "RTF Tools v0.3",
		// 美化：大幅增加初始视窗尺寸，防止拥挤
		MinSize: Size{Width: 850, Height: 650},
		Size:    Size{Width: 850, Height: 650},
		Layout:  VBox{MarginsZero: true},
		Children: []Widget{
			TabWidget{
				Pages: []TabPage{

					// ==========================================
					// Tab 1: RTF Page Check
					// ==========================================
					{
						Title:  "RTF Page Check",
						Layout: VBox{Margins: Margins{Top: 15, Bottom: 15, Left: 15, Right: 15}, Spacing: 10},
						Children: []Widget{
							// 美化：去掉难看的背景色嵌套，使用带红色字体的简单 Label
							Label{
								Text:      "Attention! Please close the Word document that is currently open before CLICK ON THIS BUTTON!",
								Font:      Font{PointSize: 11, Bold: true},
								TextColor: walk.RGB(220, 53, 69), // 现代柔和红
							},

							// 美化：采用 GroupBox 替代 Composite
							GroupBox{
								Title:  "Folder Settings",
								Layout: VBox{Spacing: 10},
								Children: []Widget{
									Composite{
										Layout: HBox{MarginsZero: true, Spacing: 5},
										Children: []Widget{
											Label{Text: "RTF Folder:"},
											LineEdit{
												AssignTo:    &mw.rtfPathEditCheck,
												ReadOnly:    true,
												ToolTipText: "Path of the RTF folder",
												Text:        mw.config.RtfPathCheck,
											},
											PushButton{
												Text: "Select...",
												OnClicked: func() {
													if dir := mw.selectDirectoryWithInput(); dir != "" {
														mw.rtfPathEditCheck.SetText(dir)
													}
												},
											},
										},
									},
								},
							},

							// 按钮区
							Composite{
								Layout: HBox{MarginsZero: true},
								Children: []Widget{
									PushButton{
										Text:    "Start to Check!",
										Font:    Font{PointSize: 12, Bold: true},
										MinSize: Size{Height: 40},
										OnClicked: func() {
											if walk.MsgBox(mw, "Attention!", "To ensure the program runs, please close the Word document that is currently open！\n\nClick [OK] to continue.", walk.MsgBoxOK) == walk.DlgCmdOK {
												go mw.startCheck()
											}
										},
									},
									ProgressBar{AssignTo: &mw.progressBarCheck},
								},
							},

							// 日志区
							GroupBox{
								Title:  "Execution Log",
								Layout: VBox{Spacing: 5},
								Children: []Widget{
									TextEdit{
										AssignTo: &mw.logViewCheck,
										ReadOnly: true,
										VScroll:  true,
										MinSize:  Size{Height: 200},
									},
								},
							},
						},
					},

					// ==========================================
					// Tab 2: RTF Combine(Specify)
					// ==========================================
					{
						Title:  "RTF Combine(Specify)",
						Layout: VBox{Margins: Margins{Top: 15, Bottom: 15, Left: 15, Right: 15}, Spacing: 10},
						Children: []Widget{
							// 顶部两列：选择源文件 vs 排序文件
							Composite{
								Layout: HBox{MarginsZero: true, Spacing: 15},
								Children: []Widget{
									// 左侧列：添加文件夹
									GroupBox{
										Title:  "1. Select Source Files",
										Layout: VBox{Spacing: 5},
										Children: []Widget{
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													PushButton{
														AssignTo: &mw.tab1.folderBtn,
														Text:     "Add Folder",
														OnClicked: func() {
															if dir := mw.selectDirectoryWithInput(); dir != "" {
																mw.tab1.loadFilesFromFolder(dir, mw)
															}
														},
													},
													PushButton{
														AssignTo:  &mw.tab1.selectAllBtn,
														Text:      "Select All",
														OnClicked: func() { mw.tab1.setAllChecked(true) },
													},
													PushButton{
														AssignTo:  &mw.tab1.unselectAllBtn,
														Text:      "Reverse",
														OnClicked: func() { mw.tab1.invertSelection() },
													},
													PushButton{
														AssignTo: &mw.tab1.confirmBtn,
														Text:     "Confirm >",
														OnClicked: func() {
															mw.tab1.showSelectedFiles(mw, []string{".rtf"})
														},
													},
													PushButton{
														Text:      "Clear All",
														OnClicked: func() { mw.tab1.clearAllFiles() },
													},
												},
											},
											TableView{
												AssignTo:         &mw.tab1.tableView,
												Model:            mw.tab1.fileModel,
												CheckBoxes:       true,
												ColumnsOrderable: true,
												Columns: []TableViewColumn{
													{Title: "Filename", Width: 200},
													{Title: "ParentFolder", Width: 100},
													{Title: "FullPath", Width: 300},
												},
												OnItemActivated: mw.onItemActivated,
											},
										},
									},
									// 右侧列：待合并队列
									GroupBox{
										Title:  "2. Reorder Files to Merge",
										Layout: VBox{Spacing: 5},
										Children: []Widget{
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													PushButton{
														AssignTo:  &mw.tab1.clearBtn,
														Text:      "Clear Selection",
														OnClicked: func() { mw.tab1.clearTextResult() },
													},
													PushButton{Text: "Move Up(↑)", OnClicked: func() { mw.tab1.moveSelectedUp() }},
													PushButton{Text: "Move Down(↓)", OnClicked: func() { mw.tab1.moveSelectedDown() }},
													HSpacer{}, // 靠左对齐按钮
												},
											},
											TableView{
												AssignTo:       &mw.tab1.resultTableView,
												Model:          mw.tab1.resultModel,
												MultiSelection: true,
												Columns: []TableViewColumn{
													{Title: "Order", Width: 50},
													{Title: "Filename", Width: 200},
													{Title: "FullPath", Width: 300},
												},
											},
										},
									},
								},
							},

							// 底部配置区：水平排列两个 GroupBox
							Composite{
								Layout: HBox{MarginsZero: true, Spacing: 8},
								Children: []Widget{
									// Options
									// ==========================================
									// ✨ 修改内容：将布局改为Grid，统一为2行，使其和右侧等高
									// ==========================================
									GroupBox{
										StretchFactor: 1, // 🔥 占据 1 份宽度
										Title:         "3. Options",
										Layout:        Grid{Columns: 2, Margins: Margins{Top: 10, Bottom: 10, Left: 10, Right: 10}, Spacing: 8},
										Children: []Widget{
											// 第一行：两个 CheckBox 平铺
											CheckBox{Text: "Add TOC", AssignTo: &mw.tab1.addTocCheckBoxCombine, Checked: mw.config.AddTocCombine},
											CheckBox{Text: "Refresh Page Number", AssignTo: &mw.tab1.changePageCheckBoxCombine, Checked: mw.config.ChangePageCombine},
											// 第二行：Label 和 LineEdit 平铺
											Label{Text: "Rows of Page:"},
											LineEdit{AssignTo: &mw.tab1.rowsEditCombine, Text: mw.config.RowsCombine},
										},
									},
									// Output Settings
									GroupBox{
										StretchFactor: 2, // 🔥 占据 1 份宽度
										Title:         "4. Output Settings",
										Layout:        Grid{Columns: 2, Margins: Margins{Top: 10, Bottom: 10, Left: 10, Right: 10}, Spacing: 8},
										Children: []Widget{
											Label{Text: "Output Folder:"},
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													LineEdit{AssignTo: &mw.tab1.outPathEditCombine, ReadOnly: true, Text: mw.config.OutPathCombine},
													PushButton{
														Text: "Browse...",
														OnClicked: func() {
															if dir := mw.selectDirectoryWithInput(); dir != "" {
																mw.tab1.outPathEditCombine.SetText(dir)
															}
														},
													},
												},
											},
											Label{Text: "Output File Name:"},
											LineEdit{AssignTo: &mw.tab1.outFileEditCombine, Text: mw.config.OutFileCombine},
										},
									},
								},
							},

							// 底部按钮与进度条
							Composite{
								Layout: HBox{MarginsZero: true, Spacing: 10},
								Children: []Widget{
									PushButton{
										Text:    "Combine Now!",
										Font:    Font{PointSize: 12, Bold: true},
										MinSize: Size{Height: 40, Width: 150},
										OnClicked: func() {
											go mw.startMerge()
										},
									},
									ProgressBar{AssignTo: &mw.tab1.progressBarCombine},
								},
							},
						},
					},

					// ==========================================
					// Tab 3: Docx Combine(General)
					// ==========================================
					{
						Title:  "Docx Combine(General)",
						Layout: VBox{Margins: Margins{Top: 10, Bottom: 10, Left: 10, Right: 10}, Spacing: 5},
						Children: []Widget{
							// 顶部两列
							Composite{
								Layout: HBox{MarginsZero: true, Spacing: 15},
								Children: []Widget{
									// 左侧列
									GroupBox{
										Title:  "1. Select Source Files",
										Layout: VBox{Spacing: 5},
										Children: []Widget{
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													PushButton{
														AssignTo: &mw.tab3.folderBtn,
														Text:     "Add Folder",
														OnClicked: func() {
															if dir := mw.selectDirectoryWithInput(); dir != "" {
																mw.tab3.loadFilesFromFolder(dir, mw)
															}
														},
													},
													PushButton{
														AssignTo:  &mw.tab3.selectAllBtn,
														Text:      "Select All",
														OnClicked: func() { mw.tab3.setAllChecked(true) },
													},
													PushButton{
														AssignTo:  &mw.tab3.unselectAllBtn,
														Text:      "Reverse",
														OnClicked: func() { mw.tab3.invertSelection() },
													},
													PushButton{
														AssignTo: &mw.tab3.confirmBtn,
														Text:     "Confirm >",
														OnClicked: func() {
															mw.tab3.showSelectedFiles(mw, []string{".docx", ".rtf"})
														},
													},
													PushButton{
														Text:      "Clear All",
														OnClicked: func() { mw.tab3.clearAllFiles() },
													},
												},
											},
											TableView{
												AssignTo:         &mw.tab3.tableView,
												Model:            mw.tab3.fileModel,
												CheckBoxes:       true,
												ColumnsOrderable: true,
												Columns: []TableViewColumn{
													{Title: "Filename", Width: 200},
													{Title: "ParentFolder", Width: 100},
													{Title: "FullPath", Width: 300},
												},
												OnItemActivated: mw.onItemActivated,
											},
										},
									},
									// 右侧列
									GroupBox{
										Title:  "2. Reorder Files to Merge",
										Layout: VBox{Spacing: 5},
										Children: []Widget{
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													PushButton{
														AssignTo:  &mw.tab3.clearBtn,
														Text:      "Clear Selection",
														OnClicked: func() { mw.tab3.clearTextResult() },
													},
													PushButton{Text: "Move Up(↑)", OnClicked: func() { mw.tab3.moveSelectedUp() }},
													PushButton{Text: "Move Down(↓)", OnClicked: func() { mw.tab3.moveSelectedDown() }},
													HSpacer{},
												},
											},
											TableView{
												AssignTo:       &mw.tab3.resultTableView,
												Model:          mw.tab3.resultModel,
												MultiSelection: true,
												Columns: []TableViewColumn{
													{Title: "Order", Width: 50},
													{Title: "Filename", Width: 200},
													{Title: "FullPath", Width: 300},
												},
											},
										},
									},
								},
							},
							// 将底部三个组件打包进一个无边距的 Composite 中，彻底消除外部干扰
							Composite{
								Layout:  VBox{MarginsZero: true, Spacing: 2}, // 内部组件间距降到 2
								MaxSize: Size{Height: 220},
								Children: []Widget{

									// 1. Output Settings GroupBox
									GroupBox{
										Title:  "3. Output Settings",
										Layout: Grid{Columns: 2, Margins: Margins{Top: 5, Bottom: 5, Left: 10, Right: 10}, Spacing: 5},
										Children: []Widget{
											Label{Text: "Output Folder:"},
											Composite{
												Layout: HBox{MarginsZero: true, Spacing: 5},
												Children: []Widget{
													LineEdit{AssignTo: &mw.tab3.outPathEditCombineDocx, ReadOnly: true, Text: mw.config.OutPathCombineDocx},
													PushButton{
														Text: "Browse...",
														OnClicked: func() {
															if dir := mw.selectDirectoryWithInput(); dir != "" {
																mw.tab3.outPathEditCombineDocx.SetText(dir)
															}
														},
													},
												},
											},
											Label{Text: "Output File Name:"},
											LineEdit{AssignTo: &mw.tab3.outFileEditCombineDocx, Text: mw.config.OutFileCombineDocx},
										},
									},

									// 2. 合并按钮
									PushButton{
										Text:    "Combine Docx Now!",
										Font:    Font{PointSize: 12, Bold: true},
										MinSize: Size{Height: 35}, // 稍微压缩一点按钮厚度
										MaxSize: Size{Height: 35},
										OnClicked: func() {
											go mw.startMergeDocx()
										},
									},

									// 3. 日志
									GroupBox{
										Title:   "Log",
										Layout:  VBox{Margins: Margins{Top: 5, Bottom: 5, Left: 5, Right: 5}, Spacing: 0},
										MaxSize: Size{Height: 100},
										Children: []Widget{
											TextEdit{
												AssignTo: &mw.logViewCombineDocx,
												ReadOnly: true,
												VScroll:  true,
												MinSize:  Size{Height: 60},
												MaxSize:  Size{Height: 80},
											},
										},
									},
								},
							},
						},
					},

					// ==========================================
					// Tab 4: RTF Converter
					// ==========================================
					{
						Title:  "RTF Converter",
						Layout: VBox{Margins: Margins{Top: 15, Bottom: 15, Left: 15, Right: 15}, Spacing: 10},
						Children: []Widget{
							Label{
								Text:      "Attention! Please close the Word document that is currently open before CLICK ON THIS BUTTON!",
								Font:      Font{PointSize: 11, Bold: true},
								TextColor: walk.RGB(220, 53, 69),
							},

							GroupBox{
								Title:  "Convert: RTF -> PDF/Docx",
								Layout: VBox{Spacing: 10},
								Children: []Widget{
									Composite{
										Layout: HBox{MarginsZero: true, Spacing: 15},
										Children: []Widget{
											CheckBox{Text: "Convert to PDF", AssignTo: &mw.convertPdf, Checked: mw.config.PdfBoxConvert},
											CheckBox{Text: "Convert to DOCX", AssignTo: &mw.convertDocx, Checked: mw.config.DocxBoxConvert},
											HSpacer{},
										},
									},
									Composite{
										Layout: HBox{MarginsZero: true, Spacing: 5},
										Children: []Widget{
											Label{Text: "Source RTF File:"},
											LineEdit{AssignTo: &mw.rtfFileEditConvert, ReadOnly: true, Text: mw.config.RtfFileConvert},
											PushButton{
												Text: "Browse...",
												OnClicked: func() {
													if file := mw.selectFile("RTF File (*.rtf)|*.RTF"); file != "" {
														mw.rtfFileEditConvert.SetText(file)
													}
												},
											},
											PushButton{Text: "Clear", OnClicked: func() { mw.rtfFileEditConvert.SetText("") }},
										},
									},
									PushButton{
										Text:    "Run Conversion (RTF to PDF/Docx)",
										Font:    Font{PointSize: 11, Bold: true},
										MinSize: Size{Height: 35},
										OnClicked: func() {
											if !mw.convertPdf.Checked() && !mw.convertDocx.Checked() {
												walk.MsgBox(mw, "Validation Error", "Please select at least one conversion option:\n• Convert to PDF\n• Convert to DOCX", walk.MsgBoxIconError|walk.MsgBoxOK)
												return
											}
											if walk.MsgBox(mw, "Attention!", "To ensure the program runs, please close the Word document that is currently open！\n\nClick [OK] to continue.", walk.MsgBoxOK) == walk.DlgCmdOK {
												go mw.startConvert()
											}
										},
									},
								},
							},

							GroupBox{
								Title:  "Convert: Docx -> RTF",
								Layout: VBox{Spacing: 10},
								Children: []Widget{
									Composite{
										Layout: HBox{MarginsZero: true, Spacing: 5},
										Children: []Widget{
											Label{Text: "Source Docx File or Folder:"},
											LineEdit{AssignTo: &mw.docxFileEditConvert, ReadOnly: true, Text: mw.config.DocxFileConvert},
											PushButton{
												Text: "Browse...",
												OnClicked: func() {
													if path := mw.selectDocxFileOrFolder(); path != "" {
														mw.docxFileEditConvert.SetText(path)
													}
												},
											},
											PushButton{Text: "Clear", OnClicked: func() { mw.docxFileEditConvert.SetText("") }},
										},
									},
									PushButton{
										Text:    "Run Conversion (Docx to RTF)",
										Font:    Font{PointSize: 11, Bold: true},
										MinSize: Size{Height: 35},
										OnClicked: func() {
											if walk.MsgBox(mw, "Attention!", "To ensure the program runs, please close the Word document that is currently open！\n\nClick [OK] to continue.", walk.MsgBoxOK) == walk.DlgCmdOK {
												go mw.startConvertDocxToRTF()
											}
										},
									},
								},
							},

							GroupBox{
								Title:  "Execution Log",
								Layout: VBox{Spacing: 5},
								Children: []Widget{
									TextEdit{
										AssignTo: &mw.logViewCovert,
										ReadOnly: true,
										VScroll:  true,
										MinSize:  Size{Height: 120},
									},
								},
							},
						},
					},

					// ==========================================
					// Tab 5: Concise Guide
					// ==========================================
					{
						Title:  "Concise Guide",
						Layout: VBox{Margins: Margins{Top: 15, Bottom: 15, Left: 15, Right: 15}},
						Children: []Widget{
							ScrollView{
								Layout: VBox{Spacing: 10},
								Children: []Widget{
									Label{Text: "Guide of RTF Page Check", Font: Font{PointSize: 16, Bold: true}, TextColor: walk.RGB(50, 100, 150)}, // 换了个沉稳的主题色

									Label{Text: "Basic Function", Font: Font{PointSize: 12, Bold: true}},
									Label{Text: "Validates page-count consistency across all RTF files in a folder (including subfolders)."},

									Label{Text: "Steps", Font: Font{PointSize: 12, Bold: true}},
									Label{Text: "1. Close any open WORD applications before checking to avoid conflicts with the program!"},
									Label{Text: "2. Select a folder; the program will recursively check all RTF files within, including all subs-folders."},
									Label{Text: "3. Click the Button. The program automatically divides the RTF files into pools based on CPU threads and checks page numbers in parallel."},
									Label{Text: "Processing time depends on the number and size of RTF files and CPU thread count—please wait patiently."},
									Label{Text: "------------------------------------------------------------", TextColor: walk.RGB(180, 180, 180)},

									Label{Text: "Guide of RTF Combine(Specify)", Font: Font{PointSize: 14, Bold: true}, TextColor: walk.RGB(50, 100, 150)},
									Label{Text: "Basic Function", Font: Font{PointSize: 12, Bold: true}},
									Label{Text: "This tool merges multiple RTF documents into a single RTF document using Specify Style."},
									Label{Text: "Steps", Font: Font{PointSize: 12, Bold: true}},
									Label{Text: "1. Add folders containing RTF files. Use Select All, Reverse, Confirm, or Clear All to manage your selections."},
									Label{Text: "2. Use the [Move Up/Down] buttons to reorder the selected RTF files."},
									Label{Text: "3. Options: Add TOC (Adds a table of contents), Rows per Page, Refresh Page Numbers."},
									Label{Text: "4. Enter the output file path and filename, then start combining."},
									Label{Text: "RTF input standard", Font: Font{PointSize: 12, Bold: true}},
									Label{Text: "1). Title and TOC Format. An IDX bookmark (e.g., IDX1) must be present. The title text must strictly follow this exact format: \\s999 \\b [Your Title Text] \\b0. Missing the style name or bold tags will result in extraction failure. "},
									Label{Text: "2). Pagination Format. The RTF syntax for pagination must support [of] or [/] as page connectors (e.g., Page 1 of 5 or Page 1 / 5). The first page of every file must explicitly contain Page 1 and its total page count."},
									Label{Text: "3). Document Boundary Control Tags. The RTF document must contain at least one of the following tags: \\widowctrl, \\sectd, or \\info. The program relies on these tags to properly split the Header from the Body."},
									Label{Text: "4). Page Dimensions. The header of the first inputted RTF file must include the \\pgwsxn (width) and \\pghsxn (height) tags. These are used to determine and set the dimensions for the dynamically generated TOC page."},
									Label{Text: "------------------------------------------------------------", TextColor: walk.RGB(180, 180, 180)},

									Label{Text: "Guide of RTF/Docx Combine(General)", Font: Font{PointSize: 14, Bold: true}, TextColor: walk.RGB(50, 100, 150)},
									Label{Text: "This tool merges multiple RTF/Docx documents into a single document for general style. General does not support adding a table of contents (TOC)."},
									Label{Text: "1. Close any open WORD applications before checking to avoid conflicts with the program!"},
									Label{Text: "2. Add and reorder files the same way as Tab 2 (accepts `.docx` and `.rtf`)."},
									Label{Text: "3. Set output path and filename, then click Combine Docx Now!."},
									Label{Text: "------------------------------------------------------------", TextColor: walk.RGB(180, 180, 180)},

									Label{Text: "Guide of RTF Converter", Font: Font{PointSize: 14, Bold: true}, TextColor: walk.RGB(50, 100, 150)},
									Label{Text: "RTF → PDF / DOCX"},
									Label{Text: "1. Check the desired output format(s)."},
									Label{Text: "2. Select the source RTF file."},
									Label{Text: "3. Click Run Conversion."},
									Label{Text: "PDFs are automatically optimized (bookmark expansion + Fast Web View)."},
									Label{Text: "Conversion time depends on file size — larger files take longer. For reference, a 20MB RTF may take approximately 10 minutes."},
									Label{Text: "------------------------------------------------------------", TextColor: walk.RGB(180, 180, 180)},

									Label{Text: "DOCX → RTF"},
									Label{Text: "1. Close any open WORD applications before checking to avoid conflicts with the program!"},
									Label{Text: "2. Select a single `.docx` file or a folder for batch conversion."},
									Label{Text: "3. Click Run Conversion (Docx to RTF)."},
									Label{Text: "Conversion time depends on file size — larger files take longer."},

									VSpacer{},
									ImageView{
										Image:   guideLogoImage,
										Mode:    ImageViewModeShrink,
										MinSize: Size{Width: 600, Height: 180},
										MaxSize: Size{Width: 800, Height: 240},
									},
									Label{Text: "Author: Allen Sun / allensrj@qq.com", Font: Font{PointSize: 9}, TextColor: walk.RGB(100, 100, 100)},
									Label{Text: "Special thanks to Guoping Ye for the initial proposal, Guojia Huang for the initial testing, and other colleagues for their suggestions.", Font: Font{PointSize: 9}, TextColor: walk.RGB(100, 100, 100)},
								},
							},
						},
					},
				},
			},
		},
	}.Create()); err != nil {
		log.Fatal(err)
	}

	mw.Closing().Attach(func(canceled *bool, reason walk.CloseReason) {
		mw.config.AddTocCombine = mw.tab1.addTocCheckBoxCombine.Checked()
		mw.config.ChangePageCombine = mw.tab1.changePageCheckBoxCombine.Checked()
		mw.config.RowsCombine = mw.tab1.rowsEditCombine.Text()
		mw.config.OutPathCombine = mw.tab1.outPathEditCombine.Text()
		mw.config.OutFileCombine = mw.tab1.outFileEditCombine.Text()
		mw.config.RtfPathCheck = mw.rtfPathEditCheck.Text()
		mw.config.PdfBoxConvert = mw.convertPdf.Checked()
		mw.config.DocxBoxConvert = mw.convertDocx.Checked()
		mw.config.RtfFileConvert = mw.rtfFileEditConvert.Text()
		mw.config.DocxFileConvert = mw.docxFileEditConvert.Text()

		mw.config.OutPathCombineDocx = mw.tab3.outPathEditCombineDocx.Text()
		mw.config.OutFileCombineDocx = mw.tab3.outFileEditCombineDocx.Text()

		if err := saveConfig(mw.config); err != nil {
			log.Println("Failed to save config:", err)
		}
	})

	mw.Run()
}

// ==========================================================
// 核心调度方法：连接 UI 和 main.go 的逻辑
// ==========================================================

func (mw *MyMainWindow) appendLogCombine(format string, args ...interface{}) {
	mw.Synchronize(func() {
		text := fmt.Sprintf(format+"\r\n", args...)
		log.Printf("%s", text)
	})
}

func (mw *MyMainWindow) openOutputDirectory(outPath, outFile string) {
	outPath = filepath.Clean(outPath)
	outputFilePath := filepath.Join(outPath, outFile+".rtf")

	if _, err := os.Stat(outputFilePath); err == nil {
		cmd := exec.Command("explorer.exe", "/select,", outputFilePath)
		if err := cmd.Start(); err != nil {
			mw.openDirectoryOnly(outPath)
		}
	} else {
		mw.openDirectoryOnly(outPath)
	}
}

func (mw *MyMainWindow) openDirectoryOnly(dirPath string) {
	dirPath = filepath.Clean(dirPath)
	cmd := exec.Command("explorer.exe", dirPath)
	if err := cmd.Start(); err != nil {
		mw.appendLogCombine("Failed to open output directory: %v", err)
		walk.MsgBox(mw, "ERROR", "Failed to open output directory: "+err.Error(), walk.MsgBoxIconError)
	}
}

func (mw *MyMainWindow) startMerge() {
	mw.Synchronize(func() { mw.tab1.progressBarCombine.SetValue(0) })

	var resultSlices []string
	for _, item := range mw.tab1.resultModel.items {
		resultSlices = append(resultSlices, item.FullPath)
	}

	if len(resultSlices) == 0 {
		walk.MsgBox(mw, "ERROR", "Add at least one source file path", walk.MsgBoxIconError)
		return
	}

	addToc := mw.tab1.addTocCheckBoxCombine.Checked()
	changePage := mw.tab1.changePageCheckBoxCombine.Checked()
	rowsText := mw.tab1.rowsEditCombine.Text()
	rows := 23
	if rowsText != "" {
		if r, err := fmt.Sscanf(rowsText, "%d", &rows); err != nil || r != 1 {
			rows = 23
		}
	}

	outPath := mw.tab1.outPathEditCombine.Text()
	outFile := mw.tab1.outFileEditCombine.Text()

	if outPath == "" || outFile == "" {
		walk.MsgBox(mw, "ERROR", "Select output directory and file name", walk.MsgBoxIconError)
		return
	}

	if err := os.MkdirAll(outPath, 0755); err != nil {
		walk.MsgBox(mw, "ERROR", "Failed to create output directory: "+err.Error(), walk.MsgBoxIconError)
		return
	}

	mw.appendLogCombine("Starting combine...")

	// 注意这里调用的首字母大写的规范化函数 CombineRTF
	err := CombineRTF(resultSlices, addToc, rows, changePage, outPath, outFile)

	if err != nil {
		mw.appendLogCombine("Failed to combine: %v", err)
		walk.MsgBox(mw, "ERROR", "Failed to combine: "+err.Error(), walk.MsgBoxIconError)
	} else {
		mw.appendLogCombine("Combine sucessfully!")
		mw.Synchronize(func() { mw.tab1.progressBarCombine.SetValue(100) })

		if walk.MsgBox(mw, "Finished", "Combine finished, Open output directory?", walk.MsgBoxYesNo) == walk.DlgCmdYes {
			mw.openOutputDirectory(outPath, outFile)
		}
	}
}

func (mw *MyMainWindow) appendLogCheck(format string, args ...interface{}) {
	mw.Synchronize(func() {
		text := fmt.Sprintf(format+"\r\n", args...)
		currentText := mw.logViewCheck.Text()
		mw.logViewCheck.SetText(currentText + text)
		mw.logViewCheck.SetTextSelection(len(currentText+text), len(currentText+text))
	})
}

func (mw *MyMainWindow) startCheck() {
	mw.Synchronize(func() {
		mw.progressBarCheck.SetValue(0)
		mw.logViewCheck.SetText("")
	})

	rtfPath := mw.rtfPathEditCheck.Text()
	if rtfPath == "" {
		walk.MsgBox(mw, "ERROR", "Please select RTF folder", walk.MsgBoxIconError)
		return
	}

	if _, err := os.Stat(rtfPath); err != nil {
		walk.MsgBox(mw, "ERROR", "RTF folder path does not exist: "+err.Error(), walk.MsgBoxIconError)
		return
	}

	logCallback := func(format string, args ...interface{}) {
		mw.appendLogCheck(format, args...)
	}

	// 调用的 RTFPageCheck
	result := RTFPageCheck(rtfPath, logCallback)

	mw.Synchronize(func() { mw.progressBarCheck.SetValue(100) })

	if result.Error != "" {
		walk.MsgBox(mw, "ERROR", result.Error, walk.MsgBoxIconError)

	} else {
		if result.AllMatched {
			walk.MsgBox(mw, "Success", "All file page counts matched successfully!", walk.MsgBoxIconInformation)
		} else {
			walk.MsgBox(mw, "Warning", "Some files have mismatched page counts!", walk.MsgBoxIconWarning)
		}
	}
}

func (mw *MyMainWindow) appendLogConvert(format string, args ...interface{}) {
	mw.Synchronize(func() {
		text := fmt.Sprintf(format+"\r\n", args...)
		currentText := mw.logViewCovert.Text()
		mw.logViewCovert.SetText(currentText + text)
		mw.logViewCovert.SetTextSelection(len(currentText+text), len(currentText+text))
	})
}

func (mw *MyMainWindow) startConvert() {
	mw.Synchronize(func() { mw.logViewCovert.SetText("") })

	rtfFile := mw.rtfFileEditConvert.Text()
	if rtfFile == "" {
		walk.MsgBox(mw, "ERROR", "Please select RTF file", walk.MsgBoxIconError)
		return
	}

	transPdf := mw.convertPdf.Checked()
	transDocx := mw.convertDocx.Checked()

	logCallback := func(format string, args ...interface{}) {
		mw.appendLogConvert(format, args...)
	}

	// 调用的 RTFConverter
	err := RTFConverter(rtfFile, transPdf, transDocx, logCallback)

	if err != nil {
		walk.MsgBox(mw, "ERROR", "Conversion failed: "+err.Error(), walk.MsgBoxIconError)
	} else {
		walk.MsgBox(mw, "Success", "RTF file converted successfully!", walk.MsgBoxIconInformation)
	}
}

func (mw *MyMainWindow) startConvertDocxToRTF() {
	mw.Synchronize(func() { mw.logViewCovert.SetText("") })

	docxFile := mw.docxFileEditConvert.Text()
	if docxFile == "" {
		walk.MsgBox(mw, "ERROR", "Please select Docx file or folder", walk.MsgBoxIconError)
		return
	}

	// ================= 补充界面的初始日志 =================
	mw.appendLogConvert("🚀 Starting Docx convert to RTF...")
	mw.appendLogConvert("📂 Target Path: %s", docxFile)
	mw.appendLogConvert("--------------------------------------------------")
	// ======================================================

	logCallback := func(format string, args ...interface{}) {
		mw.appendLogConvert(format, args...)
	}

	// 调用的 ConvertDocxToRTF，并接收返回的结果
	res := ConvertDocxToRTF(docxFile, logCallback)

	// 根据转换结果弹出提示框
	if res.Error != "" && res.Error != "no files found" {
		walk.MsgBox(mw, "ERROR", "Conversion failed: "+res.Error, walk.MsgBoxIconError)
	} else if res.TotalFiles > 0 {

		msg := fmt.Sprintf("Conversion finished!\n\nTotal: %d\nSuccess: %d\nFailed: %d",
			res.TotalFiles, res.SuccessCount, res.ErrorCount)
		walk.MsgBox(mw, "Finished", msg, walk.MsgBoxIconInformation)
	}
}

func (mw *MyMainWindow) appendLogCombineDocx(format string, args ...interface{}) {
	mw.Synchronize(func() {
		text := fmt.Sprintf(format+"\r\n", args...)
		currentText := mw.logViewCombineDocx.Text()
		mw.logViewCombineDocx.SetText(currentText + text)
		mw.logViewCombineDocx.SetTextSelection(len(currentText+text), len(currentText+text))
	})
}

func (mw *MyMainWindow) startMergeDocx() {
	mw.Synchronize(func() { mw.logViewCombineDocx.SetText("") })

	var resultSlices []string
	for _, item := range mw.tab3.resultModel.items {
		resultSlices = append(resultSlices, item.FullPath)
	}

	if len(resultSlices) == 0 {
		walk.MsgBox(mw, "ERROR", "Please select the files to merge first.", walk.MsgBoxIconError)
		return
	}

	outPath := mw.tab3.outPathEditCombineDocx.Text()
	outFile := mw.tab3.outFileEditCombineDocx.Text()

	if outPath == "" || outFile == "" {
		walk.MsgBox(mw, "ERROR", "Select output directory and file name", walk.MsgBoxIconError)
		return
	}

	// ================= 恢复这里的 UI 初始日志 =================
	mw.appendLogCombineDocx("🚀 Starting Docx Combine...")
	mw.appendLogCombineDocx("📂 Source paths:")
	for i, path := range resultSlices {
		mw.appendLogCombineDocx("   %d. %s", i+1, path)
	}
	mw.appendLogCombineDocx("📁 Output Path: %s", outPath)
	mw.appendLogCombineDocx("📄 Output File: %s.docx", outFile)
	mw.appendLogCombineDocx("--------------------------------------------------")
	// ==========================================================

	logCallback := func(format string, args ...interface{}) {
		mw.appendLogCombineDocx(format, args...)
	}

	// 调用的 CombineDocx
	err := CombineDocx(resultSlices, outPath, outFile, logCallback)

	if err != nil {
		mw.appendLogCombineDocx("❌ Failed to combine: %v", err)
		walk.MsgBox(mw, "ERROR", "Failed to combine: "+err.Error(), walk.MsgBoxIconError)
	} else {
		mw.appendLogCombineDocx("✅ Combine successfully!") // 完成日志
		if walk.MsgBox(mw, "Finished", "Combine finished, Open output directory?", walk.MsgBoxYesNo) == walk.DlgCmdYes {
			mw.openOutputDirectory(outPath, outFile)
		}
	}
}

package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v3"
)

// WeChatRecord 微信支付记录结构
type WeChatRecord struct {
	交易时间 string
	交易类型 string
	交易对方 string
	商品   string
	收支   string
	金额   string
	支付方式 string
	当前状态 string
	交易单号 string
	商户单号 string
	备注   string
}

// MemxRecord Memx记账软件记录结构
type MemxRecord struct {
	ID  string
	日期  string
	状态  string
	类型  string
	账户  string
	收款人 string
	类目  string
	子类目 string
	金额  string
	货币  string
	账号  string
	备注  string
}

// CategoryMapping 类目映射配置
type CategoryMapping struct {
	Payees      []string `yaml:"payees"`
	Category    string   `yaml:"category"`
	Subcategory string   `yaml:"subcategory"`
}

// AccountMapping 账号映射配置
type AccountMapping struct {
	PaymentMethod string `yaml:"payment_method"`
	Account       string `yaml:"account"`
}

// Config 配置文件结构
type Config struct {
	CategoryMappings []CategoryMapping `yaml:"category_mappings"`
	AccountMappings  []AccountMapping  `yaml:"account_mappings"`
	StrictMapping    bool              `yaml:"strict_mapping"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("使用方法: go run main.go <微信Excel文件> [配置文件]")
		os.Exit(1)
	}

	excelFile := os.Args[1]
	configFile := "config.yaml"
	if len(os.Args) >= 3 {
		configFile = os.Args[2]
	}

	// 读取配置
	config, err := loadConfig(configFile)
	if err != nil {
		log.Fatalf("读取配置文件失败: %v", err)
	}

	// 读取Excel文件
	records, err := readWeChatExcel(excelFile)
	if err != nil {
		log.Fatalf("读取Excel文件失败: %v", err)
	}

	// 转换数据
	memxRecords, err := convertToMemxRecords(records, config)
	if err != nil {
		log.Fatalf("转换数据失败: %v", err)
	}

	// 按账户分组并写入不同的CSV文件
	err = writeCSVByAccount(excelFile, memxRecords)
	if err != nil {
		log.Fatalf("写入CSV文件失败: %v", err)
	}
}

// loadConfig 读取配置文件
func loadConfig(filename string) (*Config, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		if os.IsNotExist(err) {
			// 如果配置文件不存在，返回默认配置
			return &Config{
				CategoryMappings: []CategoryMapping{},
				StrictMapping:    true,
			}, nil
		}
		return nil, err
	}

	var config Config
	err = yaml.Unmarshal(data, &config)
	if err != nil {
		return nil, err
	}

	return &config, nil
}

// readWeChatExcel 读取微信支付Excel文件
func readWeChatExcel(filename string) ([]WeChatRecord, error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// 获取第一个工作表
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("Excel文件中没有工作表")
	}

	sheetName := sheets[0]
	var records []WeChatRecord

	// 获取所有行数据
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取工作表失败: %v", err)
	}

	// 先找到表头行
	var headerRowIdx int = -1
	for i, row := range rows {
		if len(row) > 0 && row[0] == "交易时间" {
			headerRowIdx = i
			break
		}
	}

	if headerRowIdx == -1 {
		return nil, fmt.Errorf("未找到表头行")
	}

	// 从表头行的下一行开始读取数据
	for rowIdx := headerRowIdx + 1; rowIdx < len(rows); rowIdx++ {
		row := rows[rowIdx]

		// 检查是否为空行
		if len(row) == 0 || row[0] == "" {
			break
		}

		// 读取各列数据，使用getCellValue函数安全获取
		tradeTime := getCellValueFromSlice(row, 0)
		tradeType := getCellValueFromSlice(row, 1)
		counterparty := getCellValueFromSlice(row, 2)
		product := getCellValueFromSlice(row, 3)
		incomeExpense := getCellValueFromSlice(row, 4)
		amount := getCellValueFromSlice(row, 5)
		paymentMethod := getCellValueFromSlice(row, 6)

		if tradeType == "微信红包" && incomeExpense == "收入" {
			paymentMethod = "零钱"
		}
		if tradeType == "转账" && incomeExpense == "收入" {
			paymentMethod = "零钱"
		}
		record := WeChatRecord{
			交易时间: tradeTime,
			交易类型: tradeType,
			交易对方: counterparty,
			商品:   product,
			收支:   incomeExpense,
			金额:   amount,
			支付方式: paymentMethod,
			当前状态: getCellValueFromSlice(row, 7),
			交易单号: getCellValueFromSlice(row, 8),
			商户单号: getCellValueFromSlice(row, 9),
			备注:   getCellValueFromSlice(row, 10),
		}

		records = append(records, record)
	}

	return records, nil
}

// getCellValueFromSlice 从字符串切片中安全获取单元格值
func getCellValueFromSlice(row []string, colIdx int) string {
	if colIdx >= len(row) {
		return ""
	}
	return row[colIdx]
}

// convertToMemxRecords 转换微信记录为Memx记录
func convertToMemxRecords(weChatRecords []WeChatRecord, config *Config) ([]MemxRecord, error) {
	var memxRecords []MemxRecord

	for _, weChat := range weChatRecords {
		// 解析日期格式
		date, err := parseDateTime(weChat.交易时间)
		if err != nil {
			return nil, fmt.Errorf("解析日期失败 '%s': %v", weChat.交易时间, err)
		}

		// 清理金额字段（移除¥符号）
		amount := strings.ReplaceAll(weChat.金额, "¥", "")

		// 查找类目映射
		category, subcategory := findCategoryMapping(weChat.交易对方, config)

		// 查找账号映射
		account := findAccountMapping(weChat.支付方式, config)

		// 如果账号映射为空，尝试根据收款人推断支付方式
		if account == "" {
			account = inferAccountFromPayee(weChat.交易对方, weChat.支付方式)
			if account != "" {
				fmt.Printf("智能推断: 收款人='%s' (支付方式='%s') 推断使用账户='%s'\n",
					weChat.交易对方, weChat.支付方式, account)
			} else {
				fmt.Printf("错误行: ID=%s, 日期=%s, 收款人='%s', 支付方式='%s', 金额=%s\n",
					weChat.交易单号, date, weChat.交易对方, weChat.支付方式, amount)
			}
		}

		memx := MemxRecord{
			ID:  weChat.交易单号,
			日期:  date,
			状态:  "未核实",
			类型:  weChat.收支,
			账户:  account, // 映射后的账号
			收款人: weChat.交易对方,
			类目:  category,
			子类目: subcategory,
			金额:  amount,
			货币:  "CNY",
			账号:  "", // 账号字段留空
			备注:  weChat.商品,
		}

		memxRecords = append(memxRecords, memx)
	}

	// 按原始交易时间排序（最新的在后）
	sortRecordsByOriginalTime(weChatRecords, memxRecords)

	return memxRecords, nil
}

// parseDateTime 解析微信的日期时间格式
func parseDateTime(datetime string) (string, error) {
	// 微信格式: 2025-11-29 15:22:07
	layout := "2006-01-02 15:04:05"
	t, err := time.Parse(layout, datetime)
	if err != nil {
		return "", err
	}
	return t.Format("2006-01-02"), nil
}

// findCategoryMapping 查找类目映射
func findCategoryMapping(payee string, config *Config) (string, string) {
	// 如果收款人为空，跳过映射
	if payee == "" {
		return "", ""
	}

	for _, mapping := range config.CategoryMappings {
		for _, p := range mapping.Payees {
			if p == payee {
				return mapping.Category, mapping.Subcategory
			}
		}
	}

	// 如果没有找到映射且开启了严格模式
	if config.StrictMapping {
		log.Fatalf("错误: 找不到收款人 '%s' 的类目映射，请补充配置文件", payee)
	}

	return "", ""
}

// inferAccountFromPayee 根据收款人推断账户
func inferAccountFromPayee(payee string, paymentMethod string) string {
	// 如果支付方式为空，根据常见收款人推断账户
	if paymentMethod != "" {
		return ""
	}

	// 根据收款人推断账户
	switch payee {
	case "智慧水电管家", "携程旅行网", "爸爸", "中铁网络", "中交出行":
		return "建设银行"
	case "梁绮妮", "芳芳", "人生若只如初见～", "汪小姐_世通电脑":
		return "现金"
	default:
		// 对于其他收款人，暂时返回空
		return ""
	}
}

var configPrinted bool = false

// findAccountMapping 查找账号映射
func findAccountMapping(paymentMethod string, config *Config) string {
	for _, mapping := range config.AccountMappings {
		if mapping.PaymentMethod == paymentMethod {
			return mapping.Account
		}
	}

	// 如果没有找到账号映射，打印警告信息和配置
	fmt.Printf("警告: 找不到支付方式 '%s' 的账号映射，将归类到未知账户\n", paymentMethod)
	if !configPrinted {
		fmt.Printf("当前账号映射配置:\n")
		for i, mapping := range config.AccountMappings {
			fmt.Printf("  %d. 支付方式: '%s' -> 账户: '%s'\n", i+1, mapping.PaymentMethod, mapping.Account)
		}
		configPrinted = true
	}

	return ""
}

// sortRecordsByOriginalTime 按原始交易时间排序记录（最新的在后）
func sortRecordsByOriginalTime(weChatRecords []WeChatRecord, memxRecords []MemxRecord) {
	// 使用冒泡排序算法按原始交易时间升序排列
	for i := 0; i < len(weChatRecords)-1; i++ {
		for j := i + 1; j < len(weChatRecords); j++ {
			// 比较原始交易时间，如果weChatRecords[j]比weChatRecords[i]早，则交换
			if weChatRecords[j].交易时间 < weChatRecords[i].交易时间 {
				// 同时交换WeChatRecord和MemxRecord
				weChatRecords[i], weChatRecords[j] = weChatRecords[j], weChatRecords[i]
				memxRecords[i], memxRecords[j] = memxRecords[j], memxRecords[i]
			}
		}
	}
}

// writeCSV 写入CSV文件
func writeCSV(filename string, records []MemxRecord) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	// 写入CSV头
	headers := []string{"ID", "日期", "状态", "类型", "账户", "收款人", "类目", "子类目", "金额", "货币", "账号", "备注"}
	if err := writer.Write(headers); err != nil {
		return err
	}

	// 写入数据行
	for _, record := range records {
		row := []string{
			record.ID,
			record.日期,
			record.状态,
			record.类型,
			record.账户,
			record.收款人,
			record.类目,
			record.子类目,
			record.金额,
			record.货币,
			record.账号,
			record.备注,
		}
		if err := writer.Write(row); err != nil {
			return err
		}
	}

	return nil
}

// writeCSVByAccount 按账户分组写入不同的CSV文件
func writeCSVByAccount(excelFile string, records []MemxRecord) error {
	// 按账户分组记录
	accountGroups := make(map[string][]MemxRecord)

	for _, record := range records {
		account := record.账户
		if account == "" {
			account = "未知账户"
		}
		accountGroups[account] = append(accountGroups[account], record)
	}

	// 获取输入文件的文件名（不含扩展名）
	inputFile := excelFile
	lastDot := strings.LastIndex(inputFile, ".")
	if lastDot != -1 {
		inputFile = inputFile[:lastDot]
	}

	// 为每个账户创建单独的CSV文件
	totalRecords := 0
	for account, accountRecords := range accountGroups {
		// 清理文件名中的特殊字符
		safeAccountName := strings.ReplaceAll(account, "/", "_")
		safeAccountName = strings.ReplaceAll(safeAccountName, "\\", "_")
		safeAccountName = strings.ReplaceAll(safeAccountName, ":", "_")

		outputFile := fmt.Sprintf("%s_%s.csv", inputFile, safeAccountName)

		err := writeCSV(outputFile, accountRecords)
		if err != nil {
			return fmt.Errorf("写入文件 %s 失败: %v", outputFile, err)
		}

		fmt.Printf("成功转换 %d 条记录到 %s\n", len(accountRecords), outputFile)
		totalRecords += len(accountRecords)
	}

	fmt.Printf("总共转换 %d 条记录到 %d 个文件\n", totalRecords, len(accountGroups))
	return nil
}

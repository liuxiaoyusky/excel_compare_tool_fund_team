# Excel Compare Tool

基金团队用的 Spectra 和 HSBC 报表对比工具。上传两个 Excel，自动匹配并高亮差异。

## 快速启动

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动服务
streamlit run app.py

# 3. 浏览器打开（通常会自动打开）
# 默认地址：http://http://10.1.9.133:8501
```

上传两个文件：
- **Spectra.xls** - 左侧上传框
- **HSBC Position Appraisal Report (EXCEL).xlsx** - 右侧上传框

点 "Run Compare" 即可。

## 维护

### Security 映射管理

**问题**：有些证券在两边系统用不同的代码（ISIN 缺失或 Stack Code 不同），需要手动建立映射关系。

**操作位置**：首页的 "Security 映射管理" 展开框。

#### 增
在底部输入框填写：
- **Key**: Spectra 侧的证券标识（ISIN 或 Stack Code，不区分大小写）
- **Value**: HSBC 侧的 Security ID
- 点 "添加"

#### 删
1. 勾选表格最右侧的 "删" 列
2. 点 "删除勾选项"

#### 改
直接在表格里编辑 Value 列，然后点 "保存"。

#### 查
顶部搜索框输入关键字，按 Key/Value 模糊匹配。

**数据存储**：所有映射保存在 `mapping_override.json`，点保存后立即生效。

### 配置参数

编辑 `config.py`：

```python
# 容差设置：数值差异小于此值视为相等
TOLERANCE_ABS = 0.000

# 是否归档历史记录到 history/<timestamp>/ 目录
ENABLE_HISTORY = True

# 是否启用策略对比（legacy vs sid_map）
ENABLE_STRATEGY_COMPARISON = True

# Security ID 映射表路径（仅策略对比用）
SECURITY_ID_MAP = Path("security_id_map.csv")
```

修改后重启 Streamlit 生效。

### 历史记录

每次对比完成后，输入文件和输出结果自动归档到：

```
history/
  └── YYYYMMDD_HHMMSS/
      ├── spectra.xls              # 输入文件快照
      ├── hsbc.xlsx                # 输入文件快照
      ├── comparison_all.xlsx      # 完整对比结果（多 Sheet）
      └── comparison_all_sid.xlsx  # sid_map 策略结果（如启用）
```

**关闭归档**：在 `config.py` 设置 `ENABLE_HISTORY = False`。

## 项目结构

```
.
├── app.py                        # Streamlit 界面主程序
├── compare.py                    # 核心对比逻辑
├── extract.py                    # Excel 数据提取
├── config.py                     # 配置文件
├── seg_mapping_config.py         # 默认映射表（不要直接改这个）
├── mapping_override.json         # 自定义映射覆盖（界面操作会写这里）
├── requirements.txt              # Python 依赖
├── history/                      # 历史归档目录（自动生成）
└── security_id_map.csv           # Security ID 映射表（可选，用于策略对比）
```

**关键点**：
- 不要直接改 `seg_mapping_config.py`，用界面操作 `mapping_override.json`
- 输出 Excel 包含 4 个 Sheet：
  - **diffs**: 有差异的记录（带高亮）
  - **comparison**: 全部匹配对比
  - **unmatched**: 无法匹配的记录
  - **duplicates**: 重复记录

## 常见问题

**Q: 为什么有些证券匹配不上？**  
A: 优先级：ISIN > Stack Code > 手动映射。如果 ISIN 和 Stack Code 都不同，必须在 "Security 映射管理" 里添加映射。

**Q: 数值明明一样，为什么标黄？**  
A: 检查 `config.py` 的 `TOLERANCE_ABS`。Excel 的浮点数可能有微小差异（如 0.0001），调整容差解决。

**Q: 如何批量导入映射？**  
A: 直接编辑 `mapping_override.json`（JSON 格式），格式：
```json
{
  "KEY1": "VALUE1",
  "KEY2": "VALUE2"
}
```
保存后刷新页面。

## 技术栈

- **Streamlit 1.36+**: Web 界面
- **Pandas 2.1+**: 数据处理
- **openpyxl / xlrd**: Excel 读写
- **xlsxwriter**: Excel 格式化（高亮）

Python 3.9+ 推荐。

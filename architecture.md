# 项目目录架构

本文规定Rust实现的目录结构、模块边界和关键类型；数据字段、筛选、排序与验收规则以[project.md](project.md)为准。全部核心模块已实现并通过测试。下列模块为当前实现的架构规范。

## 1. 目录结构

```text
data-cleaning/
├── Cargo.toml                 # 项目元数据与直接依赖
├── Cargo.lock                 # 完整依赖锁定，纳入版本控制
├── rust-toolchain.toml        # 固定 Rust 工具链版本
├── .gitignore                 # 忽略 target/ 与本地样本输出
├── project.md                 # 数据处理规则与验收标准
├── architecture.md            # 本文
├── src/
│   ├── main.rs                # 程序入口，仅调用 app::cli::run()
│   ├── lib.rs                 # 声明各模块，供 main.rs 与 tests/ 使用
│   ├── test_support.rs        # 仅 src/ 内单元测试共用的辅助（唯一临时路径）
│   │
│   ├── app/                   # 交互与调度
│   │   ├── mod.rs
│   │   ├── cli.rs             # 路径输入、菜单显示、单次或批量执行调度后退出
│   │   └── runner.rs          # 编号→任务映射；单类执行或 1–9 批量执行
│   │
│   ├── model/                 # 数据模型
│   │   ├── mod.rs
│   │   ├── error.rs           # 统一错误类型 ProcessError
│   │   ├── value.rs           # 输出单元格值 Value
│   │   ├── schema.rs          # 列定义 Column / ColumnType
│   │   ├── table.rs           # 行 Row（含可选填色）与结果表 Table
│   │   └── style.rs           # 填色 Fill：黄色 #FFEB9C、粉色 #FFC7CE
│   │
│   ├── io/                    # 文件读写
│   │   ├── mod.rs
│   │   ├── paths.rs           # 路径解析与校验、文件名识别、输出路径计算
│   │   ├── xlsx_reader.rs     # calamine 读取；按绝对行列号取值
│   │   ├── xlsx_writer.rs     # rust_xlsxwriter：类型、样式、列宽、筛选与冻结表头
│   │   └── publisher.rs       # 临时文件 → 备份 → 替换 → 失败恢复
│   │
│   ├── jobs/                  # 各数据类别的处理规则
│   │   ├── mod.rs             # Job trait、Category 枚举、任务注册表
│   │   ├── invoice.rs         # 菜单 1：发票明细（第 3 节）
│   │   ├── union_subsidy.rs   # 菜单 2：银联国补明细（第 4 节）
│   │   ├── uploaded.rs        # 菜单 3、4：已上传数据（第 5 节）
│   │   ├── unionpay.rs        # 菜单 5：门店银联交易明细（第 6 节）
│   │   ├── refund.rs          # 菜单 6、7：回款明细（第 7、8 节）
│   │   ├── receipts.rs        # 菜单 8：收款单统计（第 9 节）
│   │   └── coupons.rs         # 菜单 9：销售用券情况统计（第 10 节）
│   │
│   └── utils/                 # 无状态的通用工具
│       ├── mod.rs
│       ├── dates.rs           # Excel 序列值、yyyymmdd、文本日期解析与格式化
│       ├── numbers.rs         # 两位小数、浮点尾差判断（to_cents）
│       ├── text.rs            # 空值判断、标识字段转文本、字符边界检查
│       ├── natural_sort.rs    # 文件名自然排序
│       └── doc_no.rs          # yymmdd+单据号 拼接、去“收款”前缀
│
└── tests/
    └── publishing.rs          # 集成测试：发布、覆盖与失败恢复
```

- 用户输入的源数据目录可以在项目目录内或外，名称不限，不属于上述固定结构。输出写到该源目录父目录下的`source_data/`（与源目录同级，不存在时自动创建）。
- 项目根目录下的`data/`仅存放当前样本，用于开发核对，不纳入版本控制，程序也不默认读取它。
- `target/`为编译产物目录，不纳入版本控制。

## 2. 模块职责

| 模块 | 职责 | 边界 |
|---|---|---|
| `main.rs` | 调用`app::cli::run()`，出错时打印到标准错误并以状态码1退出 | 不含任何业务逻辑 |
| `lib.rs` | 声明`app`、`model`、`io`、`jobs`、`utils` | 集成测试只能通过这里访问内部模块 |
| `app::cli` | 用`println!`与`stdin().read_line()`完成路径输入与菜单选择；执行完成后或`read_line()`返回0时正常退出 | 不含清洗规则，不直接写文件 |
| `app::runner` | 按编号查找任务并执行；空输入时依次执行1–9，用`catch_unwind`隔离单个任务的panic，汇总各类别结果 | 一次只发布一个类别的结果；自行用`println!`打印各类别结果，不回传给`cli` |
| `model::*` | 定义错误、单元格值、列定义、行、表和填色 | 不依赖`io`、`jobs` |
| `io::paths` | 解析用户路径（去首尾空白与成对引号），校验存在、是目录、可读；只列出直接子级`.xlsx`，排除`~$`；计算输出目录与文件名 | 不递归，不修改源目录 |
| `io::xlsx_reader` | 打开工作簿，枚举工作表，提供按绝对行列号访问的`SheetGrid` | 只负责读取，不做业务判断 |
| `io::xlsx_writer` | 把`Table`写成单工作表XLSX临时文件，设置数据格式、行填色、筛选、冻结表头和自适应列宽 | 不负责正式文件替换 |
| `io::publisher` | 确认临时文件已生成后替换正式文件；任一步失败恢复原文件并清理临时文件 | 只操作当前类别的结果文件 |
| `jobs::*` | 按project.md对应章节实现文件识别补充校验、表头校验、字段映射、合并、匹配、排序与行标记 | 输入为源目录，输出为`Table`；不直接写文件 |
| `utils::*` | 日期、金额、文本、自然排序、单据号等纯函数 | 不访问文件系统，便于单元测试 |

## 3. 核心类型

### 3.1 入口

```rust
// src/lib.rs
pub mod app;
pub mod io;
pub mod jobs;
pub mod model;
pub mod utils;
#[cfg(test)]
mod test_support;

// src/main.rs
fn main() {
    if let Err(error) = data_cleaning::app::cli::run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
```

`cli::run()`返回`Result<(), ProcessError>`。单类或批量执行完成后显示结果并直接退出；执行过程中的错误由`runner`捕获并打印，不会传到`main`；只有终端读写本身失败等无法继续交互的错误才向上返回。输入结束（`read_line()`返回0）时直接返回`Ok(())`正常退出。

### 3.2 结果表

每个任务只返回一个`Table`，XLSX从该`Table`写出。值的类型按单元格记录，因此第8.4节`其他支付`可以在同一列中同时出现金额和文本`-`。

```rust
// model/value.rs
pub enum Value {
    Empty,
    Text(String),
    Decimal(Decimal),        // 金额
    Integer(i64),            // 第 10 节“数量”：1 / 0 / -1
    Ratio(Decimal),          // 补贴比例：0.15 表示 15%
    Date(NaiveDate),
    Time(NaiveTime),         // 第 4 节“交易时间”
    DateTime(NaiveDateTime), // 保留到秒
}

// model/schema.rs
pub enum DecimalScale {
    Original,   // 保留原精度（第 4 节、第 5 节交易金额）
    Two,        // 固定两位小数，XLSX 格式 0.00（第 5 节补贴金额、第 6–8、10 节）
}

pub enum ColumnType {
    Text,
    Decimal(DecimalScale),
    Integer,    // XLSX 格式 0
    Ratio,      // XLSX 格式 0.00%，底层保存十进制比例，如 0.15
    Date,       // yyyy-mm-dd
    Time,       // hh:mm:ss
    DateTime,   // yyyy-mm-dd hh:mm:ss
}

pub struct Column { pub name: &'static str, pub ty: ColumnType }

// model/style.rs
pub enum Fill { Yellow, Pink }   // #FFEB9C、#FFC7CE

// model/table.rs
pub struct Row   { pub values: Vec<Value>, pub fill: Option<Fill> }
pub struct Table { pub columns: Vec<Column>, pub rows: Vec<Row> }
```

- `Integer`与`Decimal`分开，数量不以`Decimal("1")`表示；`Ratio`与`Decimal`分开，因为“0.15元”和“15%”业务含义不同。
- `xlsx_writer`只根据`ColumnType`决定数据格式，不根据字段名猜测；表头使用单行深蓝样式，普通明细使用斑马纹，`Fill`指定的黄色或粉色优先。
- 金额列需要区分“保留原精度”和“固定两位小数”，因此`Decimal`带`DecimalScale`参数。
- 同一列允许`Empty`与该列类型的值并存；除第8.4节`其他支付`的`Text("-")`外，任务不得在数值列写入文本。
- 日期类列的源值无法识别、且project.md要求“保留原值”时（第5、6节），该单元格写为`Text`原值；要求“终止处理”时（第9、10节），任务返回错误。
- `Table`提供`validate()`，写出前检查每行值的个数与列数一致。

### 3.3 任务接口

```rust
// jobs/mod.rs
#[derive(Clone, Copy)]
pub enum Category {
    Invoice = 1, UnionSubsidy, UploadedDigital, UploadedAppliance,
    UnionPay, RefundAppliance, RefundDigital, Receipts, Coupons,
}

pub trait Job {
    fn category(&self) -> Category;
    fn title(&self) -> &'static str;          // 菜单显示名，如“发票明细”
    fn output_stem(&self) -> &'static str;    // 输出基名，如“发票明细”
    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError>;
}

pub fn registry() -> &'static [&'static (dyn Job + Sync)]; // 按编号 1–9 排列

// 跨文件复用的“按某字段值查唯一命中候选”工具：coupons.rs（10.10/10.12.1/10.12.2 节）
// 与 uploaded.rs（5.7 节）共用同一套“先剔除空值、同级内去重、歧义即停不降级”判定。
pub(crate) type MultiValueIndex = HashMap<String, Vec<String>>;
pub(crate) enum PriorityOutcome { Unique(String), Ambiguous, NoHit }
pub(crate) fn resolve(hits: Vec<String>) -> PriorityOutcome;
pub(crate) fn resolve_via(index: &MultiValueIndex, key: Option<&str>) -> PriorityOutcome;
pub(crate) fn unique_hit(index: &MultiValueIndex, key: &str) -> Option<String>;
```

- `uploaded.rs`用配置结构体区分菜单3、4，合并各自的58列源数据；前25列按位置读取，第26列起按字段名称查找，支持第5.1节的“电脑”工作表变体。
- 处理明细前，按`config.category`调用对应的`RefundJob::run()`完成回款校验（菜单3↔`REFUND_DIGITAL`，菜单4↔`REFUND_APPLIANCE`）。排除`fill == Some(Fill::Pink)`的沉底记录，按`订单号`→`检索参考号`→`发票号码`建立`MultiValueIndex`。
- 每行先经`normalize_status()`归并`状态`，再用`jobs/mod.rs`的`resolve_via`/`PriorityOutcome`逐级匹配。唯一命中时覆盖为`已回款`并取回款`补贴金额`；未命中（含歧义）时保留归并后的状态，按`交易金额 × 15%`估算并保留两位小数，家电电脑封顶`1500.00`元、数码封顶`500.00`元；交易金额缺失或非数值时金额留空。结果统一追加为第59列，详见project.md第5.7节。
- `refund.rs`用配置结构体区分菜单6、7（文件名模式、同义字段表、基础字段、分组优先级、`其他支付`是否允许`-`、输出基名、`核销商编`过滤目标商户号），共用合并与沉底逻辑；读取阶段先按`核销商编`精确匹配该商户号过滤明细行，不匹配的行不进入后续校验、合并或沉底（project.md第7.1/8.1节）。`OUTPUT_FIELDS`为`pub(crate)`，供`uploaded.rs`按列位置索引匹配字段。
- 第一版不拆`uploaded/`、`refund/`子目录；代码量明显难以维护时再拆。

`unionpay.rs`分为两层接口：

```rust
// jobs/unionpay.rs
pub(crate) fn load_records(input_dir: &Path) -> Result<Vec<UnionPayRecord>, ProcessError> {
    // 文件发现 → 工作表与表头校验 → 排除首行汇总、表头、末行提示
    // → 重复导出检查 → 返回有效原始交易
}

pub struct UnionPayJob;

impl Job for UnionPayJob {
    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        let records = load_records(input_dir)?;
        // 退货识别与备注覆盖 → 稳定沉底 → 转为 Table
    }
}
```

`coupons.rs`只复用读取层，不调用`UnionPayJob::run()`、`InvoiceJob::run()`、`ReceiptsJob::run()`或
`UploadedJob::run()`，因此不会产生菜单5、菜单1、菜单8、菜单3或菜单4的结果文件：

```rust
let records = unionpay::load_records(input_dir)?;
let valid_refs: HashSet<String> = records
    .iter()
    .map(|r| r.retrieval_no.as_str())          // 源字段“检索号”
    .filter(|v| is_valid_ref_no(v))             // ^[0-9]{11}N$
    .map(str::to_owned)
    .collect();

let invoices = invoice::load_records(input_dir)?;  // 未分类、未排序、未填色的全部有效发票明细
// 按“匹配单据号”分组去重，>1 个不同“数电发票号码”视为歧义，同 10.6.2 节“歧义留空”原则。

let receipts = receipts::load_records(input_dir)?;  // 已算好第 9.5 节三阶段备注，未裁剪输出列
// 同样按“匹配单据号”分组、先剔除空备注再去重（10.12.1 节第一阶段）。

// UploadedJob::run() 本身就是未分类、未排序、未填色的完整明细表（前25列固定位置），
// 无需像 invoice/receipts 那样另外拆出 load_records；直接调用 Job::run() 复用其全部校验。
let appliance = uploaded::UPLOADED_APPLIANCE.run(input_dir)?;
let digital = uploaded::UPLOADED_DIGITAL.run(input_dir)?;
// 两组合并后按“检索参考号”“发票号码”（第7、20列）分别建索引；备注第一阶段未命中时，
// 参考号（主键）优先于数电发票号码（次键）查“状态”（第9列），10.12.2 节的两级优先级。

// 两阶段都未命中时，remark 固定为“未上传”文本（10.12.3 节）；
// 仅第一阶段实际命中的记录沉底并整行填充粉色，第二阶段及未命中记录在前且不填色（10.12.4 节）。
```

### 3.4 读取接口

```rust
// io/xlsx_reader.rs
pub struct SheetGrid { /* calamine Range 及其起始偏移 */ }

impl SheetGrid {
    pub fn name(&self) -> &str;
    pub fn cell(&self, row: u32, col: u32) -> RawCell;  // Excel 行列号，从 1 开始
    pub fn last_value_row(&self) -> Option<u32>;        // 最后一个含实际值的行
    pub fn row_texts(&self, row: u32) -> Vec<String>;   // 读取表头用
}

pub enum RawCell { Empty, Text(String), Float(f64), Int(i64), Bool(bool), DateTime(f64), Error(String) }

pub fn open_sheets(path: &Path) -> Result<Vec<SheetGrid>, ProcessError>; // 保持工作表原顺序
```

- `calamine`的`Range`从第一个非空单元格开始，`SheetGrid`用`Range::start()`换算绝对位置，效果等同project.md要求的保留开头空白行列。
- 只有格式、没有值的单元格视为`Empty`；`last_value_row`只统计实际有值的单元格。
- 数值型标识字段（如订单号被Excel存成数字）由`utils::text`转换为完整整数文本，禁止出现科学计数法或`.0`尾巴；无法无损还原时返回错误。

### 3.5 错误类型

```rust
// model/error.rs
#[derive(Debug, thiserror::Error)]
pub enum ProcessError {
    #[error("路径无效：{path}（{reason}）")]
    InvalidPath { path: PathBuf, reason: String },
    #[error("未找到符合规则的源文件：{pattern}")]
    NoInput { pattern: String },
    #[error("{file} / {sheet}：结构异常：{detail}")]
    Structure { file: String, sheet: String, detail: String },
    #[error("{file} / {sheet} 第{row}行 [{field}]：数据异常，原值“{value}”：{detail}")]
    Data { file: String, sheet: String, row: u32, field: String, value: String, detail: String },
    #[error("疑似重复导出：{detail}")]
    Duplicate { detail: String },
    #[error("读取失败：{0}")]
    Read(String),
    #[error("写入或发布失败：{0}")]
    Io(#[from] std::io::Error),
}
```

错误信息须包含project.md要求的文件名、工作表名、行号、字段名和原始值，`cli`直接打印即可。

## 4. 依赖方向

```text
main
  ↓
app ──────────► io、model（写入与发布结果）
  ↓
jobs
  ├── model
  ├── io
  └── utils
io
  ├── model
  └── utils（仅确有需要时，如读取时的标识字段转换）
model
  └── 不依赖项目其他层
utils
  └── 不依赖 jobs / app / io（可使用 model 中的错误类型）
```

- 任务间只读依赖见`project.md`第1.2节；`coupons`复用`unionpay`、`invoice`、`receipts`和`uploaded`的数据读取，`uploaded`复用对应`refund`任务的校验与结果。被引用任务不发布输出文件。
- 写文件和发布只由`app::runner`调用，任务模块不接触输出目录。

### 4.1 `utils::doc_no`的范围

只放第3、9、10节真正共用的部分：

```rust
pub fn strip_receipt_prefix(value: &str) -> &str;                   // 删除开头的“收款”
pub fn build_match_doc_no(date: NaiveDate, document_no: &str) -> String; // yymmdd + 去前缀后的单据号
```

以下第3节专有规则保留在`jobs/invoice.rs`，不得放入`doc_no.rs`：从备注中识别`销售日期`/`购机日期`与`单据号`、日期定向修正（如`2026-26-29`→`2026-06-29`）、单据号位数纠正（如`ZHLT0000524`→`ZHLT000524`）及留空判断。

## 5. 一次处理的调用顺序

```text
main.rs → app::cli → app::runner
                         │
                         ▼
                 jobs::对应任务.run(input_dir)
                   ├─ io::paths         识别源文件
                   ├─ io::xlsx_reader   读取工作表
                   └─ utils / model     转换、匹配、排序、标记
                         │
                         ▼
                    model::Table
                         │
                         ▼
        io::xlsx_writer（写入输出目录下的临时文件）
                         │
                         ▼
                   io::publisher（替换正式文件）
                         │
                         ▼
          app::cli 显示结果或错误，程序退出
```

批量模式（空输入）下，`runner`依次执行1–9，每个类别独立完成“运行→写入→发布”，失败不影响后续类别，最后逐项报告成功或失败。

## 6. 关键实现约定

### 6.1 发布

1. 在输出目录写入`.<基名>.<进程ID>.<纳秒时间戳十六进制>.xlsx.tmp`。
2. 确认临时文件存在且非空。
3. 若正式文件已存在，先重命名为`.bak`备份。
4. 将临时文件重命名为正式文件名。
5. 任一步失败：删除已放置的新文件，把`.bak`改回正式文件名，清理临时文件，返回错误。
6. 全部成功后删除`.bak`。

临时文件和正式文件位于同一目录，保证`rename`在同一文件系统内完成。正式文件被Excel占用导致替换失败时，按第5步恢复，并提示关闭文件后重试。

### 6.2 文本与正则

- `regex`不支持环视。project.md第10.6节“不得从更长数字串中截取”，用`utils::text::has_isolated_boundaries`：先匹配候选，再检查匹配前后字符是否为数字或ASCII字母；该函数只被`coupons.rs`使用。第3.4节“单据号不得吞入后续文字”由`invoice.rs`的标签正则自身的字符类边界（`[A-Za-z]*[0-9]+`，遇非字母数字即止）保证，不调用`utils::text`。
- 字段名比较使用完整字符串相等，不做大小写或空白归一化，除非project.md另有规定。
- 判断空值时，空单元格、空字符串和纯空白都视为空，但输出时保留原值。

### 6.3 数值

- 金额与比例全程使用`rust_decimal::Decimal`。从`f64`转换时直接使用`Decimal::from_f64`（最短往返表示），避免引入新的二进制误差。
- `utils::numbers::to_cents(value)`实现第10.7节的尾差规则：与最近两位小数的差不超过`0.000001`时返回该值，否则返回错误。
- 写入XLSX时把`Decimal`转换为数值单元格并套用列格式（`Decimal(Two)`固定两位小数，`Decimal(Original)`按原精度，绝不使用科学计数法）。

### 6.4 日期

- `utils::dates`统一处理Excel日期序列值、`yyyymmdd`整数或文本、`yyyy-mm-dd`、`yyyy/mm/dd`及带时间的文本。
- 第3.4节的备注日期解析（一位月日补零、两位年份、中文年月日混写、三条定向修正）单独实现在`jobs::invoice`中，不放入通用工具。

### 6.5 排序与分组

- 所有沉底操作使用稳定分区：先收集未沉底行，再收集沉底行，各自保持原顺序。
- 第4节文件顺序使用`utils::natural_sort`；第5、6节使用普通文件名升序。
- 重复分组的键同时包含“字段类型”和“字段值”，保证第7.5.1、8.6.1节“不同字段类型不得交叉匹配”。

## 7. 测试

- **纯规则单元测试**：写在`src/utils/*`等模块的`#[cfg(test)] mod tests`中，覆盖日期解析、单据号纠正、尾差判断、自然排序、参考号修正等纯函数。
- **业务规则与工作簿解析测试**：写在各`src/jobs/*.rs`的`#[cfg(test)] mod tests`中，用`rust_xlsxwriter`现场生成小样本工作簿（含开头空白行、合计行、同名列、`-`值等边界情况），运行任务后检查字段、行数、顺序、类型及填色。无需提交二进制样本文件。
- **发布集成测试**：`tests/publishing.rs`模拟目标文件已存在、临时文件缺失或为空等情况，确认原结果可恢复。
- **样本验收**：用项目根目录下的`data/`运行程序，按project.md各节“验收要求”核对当前数量与分布。此步骤不写入自动测试，因为样本数据不纳入版本控制。

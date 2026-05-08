# 客户信息Excel整合工具

一个用于将客户信息整合到Excel文档并按分类保存的Python应用程序。

## 功能特性

- 支持 JSON 和 CSV 两种输入格式
- 按客户类型自动分类保存到不同的 Excel 文件
- 支持自定义字段配置
- 生成美观的 Excel 表格，包含表头样式和边框
- 自动在输出文件名中添加时间戳

## 技术栈

- Python 3.x
- openpyxl - Excel 文件处理

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 命令行方式

```bash
python customer_excel_integrator.py <input_file> [-c <config_file>]
```

### 参数说明

- `input_file`: 输入文件路径（支持 JSON 或 CSV 格式）
- `-c, --config`: 配置文件路径（可选，默认为 config.json）

### 示例

```bash
# 使用 JSON 输入文件
python customer_excel_integrator.py sample_customers.json

# 使用 CSV 输入文件
python customer_excel_integrator.py customers.csv

# 指定自定义配置文件
python customer_excel_integrator.py customers.json -c my_config.json
```

## 输入文件格式

### JSON 格式

```json
[
    {
        "客户ID": "C001",
        "姓名": "张三",
        "性别": "男",
        "年龄": "30",
        "联系电话": "13800138001",
        "邮箱": "zhangsan@example.com",
        "地址": "北京市朝阳区",
        "城市": "北京",
        "省份": "北京市",
        "邮编": "100000",
        "客户类型": "个人客户",
        "注册日期": "2024-01-15",
        "消费金额": "5000",
        "购买产品": "产品A"
    }
]
```

### CSV 格式

```csv
客户ID,姓名,性别,年龄,联系电话,邮箱,地址,城市,省份,邮编,客户类型,注册日期,消费金额,购买产品
C001,张三,男,30,13800138001,zhangsan@example.com,北京市朝阳区,北京,北京市,100000,个人客户,2024-01-15,5000,产品A
```

## 配置文件

配置文件用于定义分类和字段结构：

```json
{
    "output_dir": "output",
    "categories": ["个人客户", "企业客户", "VIP客户"],
    "fields": {
        "基本信息": ["客户ID", "姓名", "性别", "年龄", "联系电话", "邮箱"],
        "联系信息": ["地址", "城市", "省份", "邮编"],
        "业务信息": ["客户类型", "注册日期", "消费金额", "购买产品"]
    }
}
```

### 配置项说明

- `output_dir`: 输出目录名称
- `categories`: 客户分类列表
- `fields`: 字段分组定义

## 项目结构

```
customer_excel_integrator/
├── customer_excel_integrator.py  # 主程序
├── config.json                   # 配置文件
├── requirements.txt              # 依赖列表
├── sample_customers.json         # 示例数据
└── output/                       # 输出目录（自动创建）
```

## 输出示例

运行程序后，会在 `output` 目录下生成按客户类型分类的 Excel 文件：

```
output/
├── 个人客户_20240508_123000.xlsx
├── 企业客户_20240508_123000.xlsx
└── VIP客户_20240508_123000.xlsx
```

## License

MIT
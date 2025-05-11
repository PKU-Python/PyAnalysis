# 情感分析功能说明

## 概述

`PyAnalysis.py` 是一个用于异步处理Excel文件中文本数据的Python脚本，主要功能包括：
- 使用DeepSeek API进行情感分析
- 使用jieba进行中文分词
- 批量处理Excel数据
- 提供进度监控和错误处理

## 类说明

### 1. DeepSeekClientAsync 类

DeepSeek开放平台API异步客户端，用于执行情感分析。

#### 初始化参数
| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| api_key | str | 无 | DeepSeek API密钥 |
| api_base | str | "https://api.deepseek.com" | API基础URL |
| model | str | "deepseek-chat" | 使用的模型名称 |
| max_concurrent_requests | int | 50 | 最大并发请求数 |
| request_timeout | int | 30 | 请求超时时间(秒) |
| max_retries | int | 3 | 最大重试次数 |
| retry_delay | float | 1.0 | 初始重试延迟时间(秒) |

#### 主要方法
- `sentiment_analysis(text: str, session: ClientSession) -> str`: 异步执行情感分析，包含重试机制

### 2. ExcelAnalyzerAsync 类

用于异步处理Excel文件的分析器。

#### 初始化参数
| 参数 | 类型 | 说明 |
|------|------|------|
| excel_path | str | Excel文件路径 |
| deepseek_client | DeepSeekClientAsync | DeepSeek客户端实例 |

#### 主要方法
- `read_excel()`: 读取Excel文件
- `add_columns()`: 添加情感分析和分词结果列
- `save_excel(output_path: str)`: 保存处理后的Excel文件
- `analyze_batch(texts: List[Any], indices: List[int], session: ClientSession)`: 分析一个批次的文本
- `analyze(text_column: str = "微博正文", batch_size: int = 50, rate_limit_delay: float = 0.1)`: 异步执行情感分析和分词

## 使用示例

```python
# 配置参数
EXCEL_PATH = "before/2.Data_Cleaning.xlsx"
OUTPUT_PATH = "after/2.Data_Cleaning_analyzed.xlsx"
API_KEY = "yourapikey"  # 替换为您的DeepSeek API密钥

# 创建DeepSeek异步客户端
deepseek_client = DeepSeekClientAsync(api_key=API_KEY)

# 创建Excel异步分析器
analyzer = ExcelAnalyzerAsync(
    excel_path=EXCEL_PATH,
    deepseek_client=deepseek_client
)

# 执行分析
analyzer.read_excel()
analyzer.add_columns()
asyncio.run(analyzer.analyze())
analyzer.save_excel(OUTPUT_PATH)

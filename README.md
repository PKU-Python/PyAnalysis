异步 Excel 文本情感分析与分词工具
一、简介
本项目是一个使用 Python 编写的异步程序，用于对 Excel 文件中的文本数据进行情感分析和分词处理。借助 DeepSeek 开放平台的 API 进行情感分析，并使用jieba库进行中文分词。程序采用异步编程的方式，能够高效地处理大量数据，同时通过设置并发请求数、请求延迟和重试机制等，确保程序在处理数据时的稳定性和高效性。
二、功能概述
读取 Excel 文件：从指定路径读取 Excel 文件，并将数据加载到pandas的DataFrame中。
添加列：在DataFrame中添加 “情感分析” 和 “分词结果” 两列，初始值分别为 “待分析” 和 “待分词”。
情感分析：使用 DeepSeek 开放平台的 API 对 Excel 文件中指定列的文本进行情感分析，结果分为 “非常积极”、“积极”、“中性”、“消极”、“非常消极”。
分词处理：使用jieba库对指定列的文本进行分词处理，并将结果存储在 “分词结果” 列中。
保存结果：将处理后的DataFrame保存为新的 Excel 文件。
三、环境要求
Python 3.x
所需的 Python 库：pandas、jieba、aiohttp、openpyxl
可以使用以下命令安装所需的库：
bash
pip install pandas jieba aiohttp openpyxl
四、使用步骤
1. 获取 DeepSeek API 密钥
在使用本程序之前，你需要从 DeepSeek 开放平台获取 API 密钥，并将其替换代码中的API_KEY变量。
2. 配置参数
在代码的if __name__ == "__main__":部分，可以根据需要调整以下参数：
EXCEL_PATH：输入 Excel 文件的路径。
OUTPUT_PATH：输出 Excel 文件的路径。
API_KEY：DeepSeek 开放平台的 API 密钥。
MAX_CONCURRENT_REQUESTS：最大并发请求数，根据 API 限制进行调整。
BATCH_SIZE：每个批次处理的文本数量，根据内存和 API 限制进行调整。
REQUEST_DELAY：请求间的延迟，用于控制 API 请求频率，避免被限流。
MAX_RETRIES：失败请求的最大重试次数。
RETRY_DELAY：初始重试延迟时间（秒）。
3. 运行程序
将配置好的代码保存为PyAnalysis.py，并在终端中运行以下命令：
bash
python PyAnalysis.py
五、代码结构
1. DeepSeekClientAsync类
用于与 DeepSeek 开放平台的 API 进行异步交互，实现情感分析功能，并包含重试机制。
2. ExcelAnalyzerAsync类
用于处理 Excel 文件，包括读取、添加列、分析和保存等操作。具体方法如下：
read_excel()：读取 Excel 文件。
add_columns()：添加 “情感分析” 和 “分词结果” 列。
save_excel(output_path)：保存处理后的 Excel 文件。
analyze_batch(texts, indices, session)：分析一个批次的文本，包括情感分析和分词处理。
analyze(text_column="微博正文", batch_size=50, rate_limit_delay=0.1)：优化的异步执行情感分析和分词，使用更高效的并发控制。
六、注意事项
请确保输入的 Excel 文件存在，并且包含指定的文本列（默认为 “微博正文”）。
由于使用了 DeepSeek 开放平台的 API，需要确保 API 密钥的有效性，并且遵守 API 的使用规则，避免因请求频率过高而被限流。
程序运行过程中会输出详细的进度信息，包括已完成的行数、总批次、处理速度、已用时间和预计剩余时间等，方便用户监控程序的运行状态。

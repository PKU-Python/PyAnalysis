import pandas as pd
import jieba
import asyncio
import time
import random
from aiohttp import ClientSession, ClientError
from typing import List, Dict, Any, Optional, Tuple


class DeepSeekClientAsync:
    """
    DeepSeek开放平台API异步客户端
    """

    def __init__(
            self,
            api_key: str,
            api_base: str = "https://api.deepseek.com",
            model: str = "deepseek-chat",
            max_concurrent_requests: int = 50,
            request_timeout: int = 30,
            max_retries: int = 3,
            retry_delay: float = 1.0
    ):
        """
        初始化DeepSeek API异步客户端
        """
        self.api_key = api_key
        self.api_base = api_base
        self.model = model
        self.max_concurrent_requests = max_concurrent_requests
        self.request_timeout = request_timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.semaphore = asyncio.Semaphore(max_concurrent_requests)

    async def sentiment_analysis(self, text: str, session: ClientSession) -> str:
        """
        异步执行情感分析，包含重试机制
        """
        # 确保文本是字符串类型
        text = str(text) if not isinstance(text, str) else text

        # 构建请求头
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }

        # 构建请求体
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": "你是一个情感分析专家，请分析用户提供的文本的情感倾向。"
                },
                {
                    "role": "user",
                    "content": f"请分析以下文本的情感倾向，请直接返回结果：非常积极、积极、中性、消极、非常消极：{text}"
                }
            ],
            "temperature": 0.1,
            "max_tokens": 100
        }

        # 实现重试机制
        for attempt in range(self.max_retries):
            try:
                async with self.semaphore:
                    async with session.post(
                            f"{self.api_base}/chat/completions",
                            headers=headers,
                            json=payload,
                            timeout=self.request_timeout
                    ) as response:
                        response.raise_for_status()
                        result = await response.json()

                        if "choices" in result and len(result["choices"]) > 0:
                            content = result["choices"][0]["message"]["content"].strip()
                            if "非常积极" in content:
                                return "非常积极"
                            elif "积极" in content and "非常" not in content:
                                return "积极"
                            elif "非常消极" in content:
                                return "非常消极"
                            elif "消极" in content and "非常" not in content:
                                return "消极"
                            elif "中性" in content:
                                return "中性"
                            else:
                                return content
                        else:
                            return f"未找到情感分析结果: {result}"

            except (ClientError, asyncio.TimeoutError) as e:
                if attempt < self.max_retries - 1:
                    # 使用指数退避算法进行重试
                    delay = self.retry_delay * (2 ** attempt) + random.uniform(0, 1)
                    await asyncio.sleep(delay)
                    continue
                else:
                    return f"错误: 达到最大重试次数 - {str(e)}"
            except Exception as e:
                return f"错误: {str(e)}"


class ExcelAnalyzerAsync:
    """
    ExcelAnalyzerAsync类用于异步处理Excel文件
    """

    def __init__(self, excel_path: str, deepseek_client: DeepSeekClientAsync):
        self.excel_path = excel_path
        self.deepseek_client = deepseek_client
        self.df = None

    def read_excel(self) -> None:
        """读取Excel文件"""
        try:
            self.df = pd.read_excel(self.excel_path, engine='openpyxl')
            print(f"成功读取Excel文件，共{len(self.df)}行数据")
        except Exception as e:
            print(f"读取Excel失败: {e}")
            raise

    def add_columns(self) -> None:
        """添加情感分析和分词结果列"""
        self.df["情感分析"] = "待分析"
        self.df["分词结果"] = "待分词"
        print("已添加'情感分析'和'分词结果'列")

    def save_excel(self, output_path: str) -> None:
        """保存处理后的Excel文件"""
        if not output_path.lower().endswith('.xlsx'):
            output_path = output_path.rsplit('.', 1)[0] + '.xlsx'
            print(f"注意: 自动将输出文件格式改为.xlsx: {output_path}")

        try:
            self.df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"已保存文件到{output_path}")
        except Exception as e:
            print(f"保存Excel失败: {e}")
            raise

    async def analyze_batch(self, texts: List[Any], indices: List[int], session: ClientSession) -> None:
        """
        分析一个批次的文本
        """
        # 执行情感分析
        tasks = [self.deepseek_client.sentiment_analysis(text, session) for text in texts]
        sentiments = await asyncio.gather(*tasks)

        # 更新情感分析结果
        for idx, sentiment in zip(indices, sentiments):
            self.df.loc[idx, "情感分析"] = sentiment

        # 执行分词，确保处理非字符串类型
        self.df.loc[indices, "分词结果"] = [
            " ".join(jieba.cut(str(text))) if pd.notna(text) else ""
            for text in texts
        ]

    async def analyze(self, text_column: str = "微博正文", batch_size: int = 50, rate_limit_delay: float = 0.1) -> None:
        """
        优化的异步执行情感分析和分词，使用更高效的并发控制
        :param text_column: 包含文本的列名，默认为"微博正文"
        :param batch_size: 每个批次处理的文本数量
        :param rate_limit_delay: 请求间的延迟，控制API请求频率
        """
        print("开始进行情感分析和分词...")
        start_time = time.time()

        # 准备所有批次
        valid_rows = self.df[pd.notna(self.df[text_column])]
        total_batches = (len(valid_rows) + batch_size - 1) // batch_size
        completed_batches = 0
        total_valid = len(valid_rows)

        async with ClientSession() as session:
            # 创建并执行所有批次任务
            for i in range(0, len(valid_rows), batch_size):
                batch_df = valid_rows.iloc[i:i + batch_size]
                texts = batch_df[text_column].tolist()
                indices = batch_df.index.tolist()

                # 执行当前批次
                await self.analyze_batch(texts, indices, session)

                # 更新进度
                completed_batches += 1
                completed_rows = min((i + batch_size), total_valid)

                # 计算并显示进度
                progress = completed_rows / total_valid * 100
                elapsed = time.time() - start_time
                rate = completed_rows / elapsed if elapsed > 0 else 0
                eta = (total_valid - completed_rows) / rate if rate > 0 else 0

                print(f"进度: {completed_rows}/{total_valid} ({progress:.2f}%) | "
                      f"批次: {completed_batches}/{total_batches} | "
                      f"速度: {rate:.2f}行/秒 | "
                      f"已用时间: {elapsed:.2f}s | "
                      f"预计剩余时间: {eta:.2f}s")

                # 控制API请求频率，避免被限流
                if rate_limit_delay > 0 and completed_batches < total_batches:
                    await asyncio.sleep(rate_limit_delay)

        print(f"情感分析和分词完成，共耗时: {time.time() - start_time:.2f}秒")


# 使用分析功能
if __name__ == "__main__":
    # 配置参数
    EXCEL_PATH = "before/2.Data_Cleaning.xlsx"
    OUTPUT_PATH = "after/3.Data_Cleaning_analyzed.xlsx"
    API_KEY = "yourapikey"  # 替换为您的DeepSeek API密钥

    # 性能优化配置
    MAX_CONCURRENT_REQUESTS = 50  # 根据API限制调整
    BATCH_SIZE = 50  # 根据内存和API限制调整
    REQUEST_DELAY = 0.05  # 控制API请求频率，避免被限流
    MAX_RETRIES = 3  # 失败请求的最大重试次数
    RETRY_DELAY = 1.0  # 初始重试延迟时间(秒)

    # 创建DeepSeek异步客户端
    deepseek_client = DeepSeekClientAsync(
        api_key=API_KEY,
        max_concurrent_requests=MAX_CONCURRENT_REQUESTS,
        max_retries=MAX_RETRIES,
        retry_delay=RETRY_DELAY
    )

    # 创建Excel异步分析器
    analyzer = ExcelAnalyzerAsync(
        excel_path=EXCEL_PATH,
        deepseek_client=deepseek_client
    )

    try:
        analyzer.read_excel()
        analyzer.add_columns()
        asyncio.run(analyzer.analyze(
            batch_size=BATCH_SIZE,
            rate_limit_delay=REQUEST_DELAY
        ))
        analyzer.save_excel(OUTPUT_PATH)
        print("处理完成!")
    except Exception as e:
        print(f"程序执行失败: {e}")
        import traceback

        traceback.print_exc()  # 打印详细的错误堆栈信息

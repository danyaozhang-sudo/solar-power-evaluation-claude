# Tavily API 搜索回退方案

## 适用场景

当 `web_search` 工具在网关环境（飞书、Telegram、Slack 等）不可用时，通过 Tavily API 直接调用获取网页搜索数据。

## 前提

Tavily API Key 已在 `~/.zshrc` 中配置：
```bash
export TAVILY_API_KEY="tvly-..."
```

## 完整代码模板

```python
import subprocess, json

def tavily_search(query, max_results=5):
    """通过 Tavily API 执行搜索，返回带摘要的结果列表"""
    # 获取 API Key（从 shell 环境）
    result = subprocess.run(
        ["bash", "-c", "source ~/.zshrc 2>/dev/null; echo $TAVILY_API_KEY"],
        capture_output=True, text=True
    )
    api_key = result.stdout.strip().split("\n")[-1]
    if not api_key:
        print("ERROR: TAVILY_API_KEY not found!")
        return {"results": []}
    
    payload = {
        "api_key": api_key,
        "query": query,
        "search_depth": "advanced",  # "basic" | "advanced"
        "include_answer": True,       # 包含 LLM 生成的摘要回答
        "max_results": max_results
    }
    cmd = (
        f"curl -s -X POST https://api.tavily.com/search "
        f"-H 'Content-Type: application/json' "
        f"-d '{json.dumps(payload)}'"
    )
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True, timeout=30)
    return json.loads(result.stdout)
```

## 使用示例

```python
# 单次搜索
data = tavily_search("广东 光伏 等效利用小时数 2024")
if "answer" in data:
    print(f"摘要: {data['answer']}")
for r in data.get("results", []):
    print(f"- {r.get('title')}")
    print(f"  {r.get('content')[:200]}")
    print(f"  URL: {r.get('url')}")

# 批量搜索
queries = [
    "广东 光伏 限电率 2024",
    "潮州 光伏 年发电量 利用小时数",
]
for q in queries:
    data = tavily_search(q)
    # 处理结果...
```

## 注意事项

1. Tavily 有搜索配额限制（通常每月1000次免费），避免无意义的重复搜索
2. `search_depth: "advanced"` 消耗更多额度但返回更精准结果
3. 搜索结果中 `content` 字段包含页面摘要，`url` 是原始链接
4. API Key 在 `~/.zshrc` 中，bash 子进程需要 `source ~/.zshrc` 才能读取
5. 如果 Tavily 也失败，降级到 `delegate_task` 让子代理执行搜索

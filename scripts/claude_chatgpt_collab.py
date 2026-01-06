#!/usr/bin/env python3
"""
Claude + ChatGPT Pro 協作腳本

流程:
1. 用戶提出需求
2. Claude (gemini-2.5-flash) 完成初版結果
3. ChatGPT Pro 審核並給建議
4. Claude 根據建議改進
5. 重複 3-4 步驟 N 次
6. 輸出最終結果

用法:
    python scripts/claude_chatgpt_collab.py --task "翻譯這段文字" --iterations 2
"""
import argparse
import json
import time
import sys
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import List, Optional
import requests
from openai import OpenAI

# 設定
LITELLM_API_BASE = "https://litellm.whaleforce.dev"
LITELLM_API_KEY = "sk-uI7-kCNyMyXW8QnSAbKrMg"
CHATGPT_PRO_API = "https://chatgpt-pro-api.gpu5090.whaleforce.dev"
CLAUDE_MODEL = "gemini-2.5-flash"  # 透過 LiteLLM 使用


@dataclass
class IterationResult:
    """單次迭代結果"""
    iteration: int
    claude_output: str = ""
    chatgpt_feedback: str = ""
    claude_prompt_tokens: int = 0
    claude_completion_tokens: int = 0
    claude_cost: float = 0.0
    duration_seconds: float = 0.0


@dataclass
class CollabResult:
    """協作結果"""
    task: str
    iterations: List[IterationResult] = field(default_factory=list)
    final_output: str = ""
    total_claude_tokens: int = 0
    total_claude_cost: float = 0.0
    total_duration_seconds: float = 0.0

    def to_dict(self):
        return {
            "task": self.task,
            "iterations": [asdict(it) for it in self.iterations],
            "final_output": self.final_output,
            "total_claude_tokens": self.total_claude_tokens,
            "total_claude_cost": self.total_claude_cost,
            "total_duration_seconds": self.total_duration_seconds
        }


class ClaudeChatGPTCollab:
    """Claude + ChatGPT Pro 協作器"""

    CLAUDE_INITIAL_PROMPT = """你是專業的助手。請完成以下任務：

任務：
{task}

請提供完整的結果："""

    CLAUDE_IMPROVE_PROMPT = """你是專業的助手。請根據專家的建議改進你的結果。

原始任務：
{task}

你之前的結果：
{previous_output}

專家建議：
{feedback}

請根據建議改進並提供更新後的完整結果："""

    CHATGPT_REVIEW_PROMPT = """你是資深專家審核員。請審核以下任務結果並提供具體改進建議。

任務：
{task}

當前結果：
{output}

這是第 {iteration} 次迭代（共 {total_iterations} 次）。

請提供：
1. 結果品質評估 (1-10 分)
2. 具體問題指出
3. 改進建議 (請具體且可執行)

如果結果已經很好（9分以上），可以說「結果已達標準，無需進一步修改」。"""

    def __init__(
        self,
        litellm_api_base: str = LITELLM_API_BASE,
        litellm_api_key: str = LITELLM_API_KEY,
        chatgpt_api: str = CHATGPT_PRO_API,
        claude_model: str = CLAUDE_MODEL,
        verbose: bool = True
    ):
        self.litellm_api_base = litellm_api_base
        self.litellm_api_key = litellm_api_key
        self.chatgpt_api = chatgpt_api
        self.claude_model = claude_model
        self.verbose = verbose

        # OpenAI client for Claude (via LiteLLM)
        self.claude_client = OpenAI(
            api_key=litellm_api_key,
            base_url=litellm_api_base
        )

    def _log(self, message: str):
        """輸出日誌"""
        if self.verbose:
            print(message)

    def call_claude(self, prompt: str) -> tuple:
        """
        呼叫 Claude (透過 LiteLLM)

        Returns:
            tuple: (response_text, prompt_tokens, completion_tokens, cost)
        """
        try:
            response = self.claude_client.chat.completions.create(
                model=self.claude_model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=4096,
                temperature=0.3
            )

            usage = response.usage
            prompt_tokens = usage.prompt_tokens if usage else 0
            completion_tokens = usage.completion_tokens if usage else 0
            cost = getattr(usage, 'cost', 0.0) if usage else 0.0

            return (
                response.choices[0].message.content.strip(),
                prompt_tokens,
                completion_tokens,
                cost
            )
        except Exception as e:
            self._log(f"  [ERROR] Claude 呼叫失敗: {e}")
            return "", 0, 0, 0.0

    def call_chatgpt_pro(self, prompt: str, timeout: int = 120, retries: int = 2) -> str:
        """
        呼叫 ChatGPT Pro API

        Args:
            prompt: 提示詞
            timeout: 等待超時秒數
            retries: 重試次數

        Returns:
            回應文字
        """
        task_data = None

        # 重試提交任務
        for attempt in range(retries + 1):
            try:
                response = requests.post(
                    f"{self.chatgpt_api}/chat",
                    json={"prompt": prompt},
                    timeout=30
                )
                response.raise_for_status()
                task_data = response.json()
                break
            except requests.exceptions.RequestException as e:
                if attempt < retries:
                    self._log(f"  [RETRY] ChatGPT Pro 連線失敗，第 {attempt + 1}/{retries + 1} 次嘗試...")
                    time.sleep(2)
                else:
                    self._log(f"  [ERROR] ChatGPT Pro 連線失敗: {e}")
                    return ""

        if not task_data or not task_data.get("success"):
            self._log(f"  [ERROR] ChatGPT Pro 提交失敗: {task_data}")
            return ""

        task_id = task_data["task_id"]
        self._log(f"  ChatGPT Pro 任務已提交: {task_id}")

        # 等待結果
        try:
            wait_time = min(timeout, 60)
            remaining = timeout

            while remaining > 0:
                result = requests.get(
                    f"{self.chatgpt_api}/task/{task_id}",
                    params={"wait": wait_time},
                    timeout=wait_time + 10
                ).json()

                status = result.get("status")

                if status == "completed":
                    return result.get("answer", "")
                elif status == "failed":
                    self._log(f"  [ERROR] ChatGPT Pro 任務失敗: {result.get('error')}")
                    return ""
                elif status in ["queued", "sent", "processing"]:
                    self._log(f"  ChatGPT Pro 狀態: {status}, 進度: {result.get('progress', 'N/A')}")
                    remaining -= wait_time
                    wait_time = min(remaining, 60)
                else:
                    self._log(f"  [WARNING] 未知狀態: {status}")
                    break

            self._log(f"  [ERROR] ChatGPT Pro 超時")
            return ""

        except Exception as e:
            self._log(f"  [ERROR] ChatGPT Pro 等待結果失敗: {e}")
            return ""

    def run(
        self,
        task: str,
        iterations: int = 2,
        chatgpt_timeout: int = 120
    ) -> CollabResult:
        """
        執行協作流程

        Args:
            task: 任務描述
            iterations: 迭代次數
            chatgpt_timeout: ChatGPT Pro 超時秒數

        Returns:
            CollabResult
        """
        result = CollabResult(task=task)
        start_time = time.time()

        self._log("=" * 60)
        self._log("Claude + ChatGPT Pro 協作開始")
        self._log("=" * 60)
        self._log(f"任務: {task[:100]}...")
        self._log(f"迭代次數: {iterations}")
        self._log("")

        current_output = ""

        for i in range(iterations + 1):  # +1 因為第一次是初始生成
            iter_start = time.time()
            iter_result = IterationResult(iteration=i)

            if i == 0:
                # 第一次：Claude 生成初版
                self._log(f"[迭代 {i}] Claude 生成初版結果...")
                prompt = self.CLAUDE_INITIAL_PROMPT.format(task=task)
            else:
                # 後續：根據 ChatGPT 建議改進
                self._log(f"[迭代 {i}] Claude 根據建議改進...")
                prompt = self.CLAUDE_IMPROVE_PROMPT.format(
                    task=task,
                    previous_output=current_output,
                    feedback=result.iterations[-1].chatgpt_feedback
                )

            # 呼叫 Claude
            output, prompt_tokens, completion_tokens, cost = self.call_claude(prompt)
            current_output = output

            iter_result.claude_output = output
            iter_result.claude_prompt_tokens = prompt_tokens
            iter_result.claude_completion_tokens = completion_tokens
            iter_result.claude_cost = cost

            result.total_claude_tokens += prompt_tokens + completion_tokens
            result.total_claude_cost += cost

            self._log(f"  Claude 完成 (tokens: {prompt_tokens + completion_tokens}, cost: ${cost:.6f})")
            self._log(f"  輸出預覽: {output[:200]}...")

            # 如果還有下一次迭代，呼叫 ChatGPT Pro 審核
            if i < iterations:
                self._log(f"\n[迭代 {i}] ChatGPT Pro 審核中...")

                review_prompt = self.CHATGPT_REVIEW_PROMPT.format(
                    task=task,
                    output=current_output,
                    iteration=i + 1,
                    total_iterations=iterations
                )

                feedback = self.call_chatgpt_pro(review_prompt, chatgpt_timeout)
                iter_result.chatgpt_feedback = feedback

                if feedback:
                    self._log(f"  ChatGPT Pro 建議: {feedback[:200]}...")

                    # 檢查是否已達標準
                    if "無需進一步修改" in feedback or "已達標準" in feedback:
                        self._log("  [INFO] ChatGPT Pro 認為結果已達標準，提前結束")
                        iter_result.duration_seconds = time.time() - iter_start
                        result.iterations.append(iter_result)
                        break
                else:
                    self._log("  [WARNING] ChatGPT Pro 無回應，跳過此次審核")
                    iter_result.chatgpt_feedback = "(無回應)"

            iter_result.duration_seconds = time.time() - iter_start
            result.iterations.append(iter_result)
            self._log("")

        result.final_output = current_output
        result.total_duration_seconds = time.time() - start_time

        self._log("=" * 60)
        self._log("協作完成")
        self._log("=" * 60)
        self._log(f"總迭代次數: {len(result.iterations)}")
        self._log(f"總 Claude tokens: {result.total_claude_tokens}")
        self._log(f"總 Claude 成本: ${result.total_claude_cost:.6f}")
        self._log(f"總耗時: {result.total_duration_seconds:.1f} 秒")
        self._log("")
        self._log("最終結果:")
        self._log("-" * 40)
        self._log(result.final_output)

        return result


def main():
    parser = argparse.ArgumentParser(
        description="Claude + ChatGPT Pro 協作腳本"
    )
    parser.add_argument(
        "--task", "-t",
        required=True,
        help="任務描述"
    )
    parser.add_argument(
        "--iterations", "-i",
        type=int,
        default=2,
        help="迭代次數 (預設: 2)"
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="ChatGPT Pro 超時秒數 (預設: 120)"
    )
    parser.add_argument(
        "--output", "-o",
        help="輸出 JSON 檔案路徑"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="安靜模式"
    )

    args = parser.parse_args()

    # 執行協作
    collab = ClaudeChatGPTCollab(verbose=not args.quiet)
    result = collab.run(
        task=args.task,
        iterations=args.iterations,
        chatgpt_timeout=args.timeout
    )

    # 輸出 JSON
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        print(f"\n結果已儲存至: {output_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())

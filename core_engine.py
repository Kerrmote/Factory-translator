import os
import time
import re
import logging
from typing import Dict, Optional
from openai import OpenAI, APIConnectionError, APITimeoutError, APIStatusError
from glossary_manager import GlossaryManager

class TranslationEngine:
    def __init__(self, api_key: str, base_url: str, model: str, base_dir: str, direction: str, stop_event=None):
        # 设置超时：连接超时 10s，读取超时 120s
        self.client = OpenAI(
            api_key=api_key, 
            base_url=base_url,
            timeout=120.0,
            max_retries=0 # 我们在外部手动控制重试逻辑
        )
        self.model = model
        self.base_dir = base_dir
        # 将zh2en/en2zh格式转换为zh-en/en-zh格式
        self.direction = direction.replace("2", "-") if direction else "zh-en"
        self.stop_event = stop_event
        self.glossary = GlossaryManager(base_dir)
        self.logger = logging.getLogger("TranslationEngine")

    def _get_system_prompt(self, direction: str, terms: Dict[str, str]) -> str:
        term_hint = ""
        if terms:
            term_hint = "\nSTRICT TERMINOLOGY TABLE (MUST FOLLOW):\n" + "\n".join([f"- {k} -> {v}" for k, v in terms.items()])
        
        if direction == "zh-en":
            return f"You are a professional translator for factory QMS documents. Translate Chinese to English.{term_hint}\nRequirements: Formal, consistent, no explanations, no summaries. Ensure terms in the table are used exactly as provided."
        else:
            return f"你是一名专业的工厂体系文件翻译员。请将英文翻译为中文。{term_hint}\n要求：用词正式、准确，保持一致性，不要添加额外解释或总结。必须严格遵守上述术语表中的译法。"

    def translate_segment(self, text: str, direction: Optional[str] = None, stop_event=None) -> str:
        if not text.strip(): return text
        # 使用实例变量作为默认值
        if direction is None:
            direction = self.direction
        if stop_event is None:
            stop_event = self.stop_event
        if stop_event and stop_event.is_set(): return ""

        terms = self.glossary.get_terms(direction)
        system_prompt = self._get_system_prompt(direction, terms)
        
        max_retries = 3
        last_exception = None
        
        for attempt in range(max_retries):
            if stop_event and stop_event.is_set(): return ""
            try:
                resp = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": text},
                    ],
                    temperature=0.0,
                )
                translated = resp.choices[0].message.content.strip()
                if not translated:
                    raise ValueError("API 返回了空译文")
                
                # 后处理校正
                translated = self._post_process_correction(translated, terms)
                return translated
                
            except (APITimeoutError, APIConnectionError) as e:
                last_exception = e
                self.logger.warning(f"网络异常 (尝试 {attempt+1}/{max_retries}): {e}")
                time.sleep(2 ** attempt)
            except APIStatusError as e:
                self.logger.error(f"API 状态错误: {e.status_code} - {e.message}")
                raise e
            except Exception as e:
                self.logger.error(f"翻译过程中出现未知错误: {e}")
                raise e
        
        if last_exception:
            raise last_exception
        return ""

    def _post_process_correction(self, text: str, terms: Dict[str, str]) -> str:
        sorted_keys = sorted(terms.keys(), key=len, reverse=True)
        for k in sorted_keys:
            v = terms[k]
            if re.search(r'[\u4e00-\u9fff]', k):
                text = text.replace(k, v)
            else:
                pattern = re.compile(r'\b' + re.escape(k) + r'\b', re.IGNORECASE)
                text = pattern.sub(v, text)
        return text

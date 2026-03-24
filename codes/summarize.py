import warnings
warnings.filterwarnings("ignore")

from openai import OpenAI


class Summarizer:
    def __init__(self, config: dict, model: str = "r1"):
        self.config = config
        self.model = model

    def get_model_answer(self, system_prompt: str, user_prompt: str, max_tokens: int = 8000) -> str:
        try:
            client = OpenAI(
                base_url=self.config["llm_config"]["base_url"],
                api_key=self.config["llm_config"]["api_key"]
            )
            
            completion = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                max_tokens=max_tokens,
                temperature=self.config["llm_config"]["temperature"]
            )
            
            result = str(completion.choices[0].message.content).strip('\n\n')
            return result
            
        except Exception as e:
            print(f"Error calling DeepSeek API: {str(e)}")
            return "回答生成失败，请检查API配置或网络连接。"

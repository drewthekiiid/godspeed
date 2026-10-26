from openai import OpenAI
import os

openai = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def analyze_with_gpt(query: str, content: str) -> dict:
    try:
        truncated = content[:8000]  # truncate for token safety
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a document analyst."},
                {"role": "user", "content": f"{query}\n\n{truncated}"}
            ]
        )
        return {
    "response": response.choices[0].message.content,
    "usage": {
        "prompt_tokens": response.usage.prompt_tokens,
        "completion_tokens": response.usage.completion_tokens,
        "total_tokens": response.usage.total_tokens
    }
}
    except Exception as e:
        raise RuntimeError(f"GPT-4 failed: {str(e)}")
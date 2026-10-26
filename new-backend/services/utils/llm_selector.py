import os
import random

AGENT_LLM_CONFIG = {
    "manager": [
        {"model": "o4-mini", "api_key": os.getenv("OPENAI_API_KEY_MANAGER")},
    ],
    "estimator": [
        {"model": "gpt-4o", "api_key": os.getenv("OPENAI_API_KEY")},
        {"model": "gpt-3.5-turbo", "api_key": os.getenv("OPENAI_API_KEY_2")},
    ],
    # Extend as needed for other agents (rfi, exporter, validator, etc.)
}

def select_llm(agent_name, state):
    configs = AGENT_LLM_CONFIG.get(agent_name, [])
    if not configs:
        raise Exception(f"No LLM config for agent '{agent_name}'")
    # Could add smarter selection based on state/complexity
    return random.choice(configs)

from datetime import datetime

def create_initial_state(
    query: str,
    content: str,
    files: list = None,
    metadata: dict = None,
    user_id: str = None,
    session_id: str = None
) -> dict:
    """
    Create the universal state “envelope” for every run.
    """
    return {
        "query": query,
        "content": content,
        "files": files or [],
        "metadata": metadata or {},
        "user_id": user_id,
        "session_id": session_id,
        "llm_config": {
            "model": None,
            "api_key": None,
            "params": {},
            "token_usage": None,
            "cost_estimate": None
        },
        "meeting_log": [],
        "agent_trace": [],
        "history": [],
        "result": None,
        "error": None,
        "created_at": datetime.utcnow().isoformat()
    }

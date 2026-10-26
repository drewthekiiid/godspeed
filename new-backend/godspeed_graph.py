# godspeed_graph.py

from langgraph.graph import StateGraph
from langchain_openai import ChatOpenAI
from agents.manager_agent import manager_agent

# Initialize LLM (swap to gpt-4o for multimodal soon)
llm = ChatOpenAI(model="gpt-4o", temperature=0)

# Set up the graph
graph = StateGraph()
graph.add_node("manager", manager_agent)
graph.set_entry_node("manager")

__all__ = ["graph"]
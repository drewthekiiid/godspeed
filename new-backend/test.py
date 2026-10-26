# test_state.py
from services.utils.state import create_initial_state

s = create_initial_state("Hello", "World")
print("State:", s)
assert s["query"] == "Hello"
print("✅ state utility works!")

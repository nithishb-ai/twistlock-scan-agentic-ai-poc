from langchain_ollama import OllamaLLM
from langchain_experimental.agents import create_pandas_dataframe_agent
import pandas as pd

# Load your Excel sheet
df = pd.read_excel("twistlock_report_1.xlsx")  # update path

# Set up Ollama LLM
llm = OllamaLLM(model="llama3")  # Uses the llama3 model installed in your system

# Create the agent
agent = create_pandas_dataframe_agent(llm, df, verbose=False, allow_dangerous_code=True)

# Ask a question
while True:
    query = input("Ask about the scan (or type 'exit'): ")
    if query.lower() == "exit":
        break
    response = agent.invoke(query)
    print(response)

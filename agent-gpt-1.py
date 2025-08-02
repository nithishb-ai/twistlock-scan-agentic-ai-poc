import pandas as pd
import ollama
import json

def load_excel_data(file_path):
    """Load Excel file and convert to structured text"""
    df = pd.read_excel(file_path)
    
    # Convert DataFrame to a clean text format
    data_text = f"Excel Data Summary:\n"
    data_text += f"Total Rows: {len(df)}\n"
    data_text += f"Columns: {', '.join(df.columns.tolist())}\n\n"
    
    # Add each row with clear formatting and record numbers
    data_text += "Data Records (COUNT EACH ONE CAREFULLY):\n"
    data_text += "=" * 50 + "\n"
    
    for idx, row in df.iterrows():
        data_text += f"\n[RECORD #{idx + 1}]\n"
        for col, value in row.items():
            # Handle different data types and make them more readable
            if pd.isna(value):
                value = "NULL"
            data_text += f"  {col}: {value}\n"
        data_text += "-" * 30 + "\n"
    
    # Add summary statistics for numeric columns
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        data_text += f"\nSummary Statistics:\n"
        for col in numeric_cols:
            data_text += f"  {col}: Min={df[col].min()}, Max={df[col].max()}, Mean={df[col].mean():.2f}, Count={df[col].count()}\n"
    
    # Add unique value counts for text columns
    text_cols = df.select_dtypes(include=['object']).columns
    if len(text_cols) > 0:
        data_text += f"\nUnique Value Counts:\n"
        for col in text_cols[:3]:  # Limit to first 3 text columns
            unique_count = df[col].nunique()
            data_text += f"  {col}: {unique_count} unique values\n"
    
    return data_text, df

def create_effective_prompt(data_text, user_question):
    """Create an optimized prompt for better answers"""
    prompt = f"""You are an expert data analyst. Analyze the following Excel data carefully and answer the question accurately.

EXCEL DATA:
{data_text}

INSTRUCTIONS:
- Base your answer ONLY on the provided data
- Be specific and cite exact values when possible
- If you need to calculate something, show the calculation
- If the data doesn't contain enough information, say so clearly
- Give precise, factual answers

QUESTION: {user_question}

ANSWER:"""
    
    return prompt

def verify_count_with_pandas(df, question):
    """Pre-calculate common queries to verify model answers"""
    verification = {}
    
    # Try to extract key terms from the question
    question_lower = question.lower()
    
    # Look for potential column references and values
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in question_lower:
            # Count total non-null values in this column
            verification[f"total_{col}"] = df[col].count()
            
            # If it's asking about a specific value
            for word in question.split():
                if len(word) > 2:  # Skip short words
                    matches = df[df[col].astype(str).str.contains(word, case=False, na=False)]
                    if len(matches) > 0:
                        verification[f"{col}_containing_{word}"] = len(matches)
    
    return verification

def query_excel_with_ollama(file_path, question, model_name="qwen3"):
    """Main function to query Excel data using Ollama"""
    try:
        # Load Excel data
        print("Loading Excel data...")
        data_text, df = load_excel_data(file_path)
        
        # Get verification data
        verification = verify_count_with_pandas(df, question)
        
        # Create optimized prompt
        prompt = create_effective_prompt(data_text, question)
        
        # Query Ollama with optimized settings for accuracy
        print("Querying Ollama...")
        response = ollama.generate(
            model=model_name,
            prompt=prompt,
            options={
                'temperature': 0.0,      # Zero temperature for maximum consistency
                'top_p': 0.8,           # More focused sampling
                'top_k': 20,            # Reduced for more precise answers
                'repeat_penalty': 1.15,  # Slightly higher to avoid repetition
                'num_ctx': 8192,        # Increased context window
                'num_predict': 1024     # Allow longer responses for detailed answers
            }
        )
        
        return {
            'question': question,
            'answer': response['response'],
            'data_shape': df.shape,
            'columns': df.columns.tolist(),
            'verification': verification
        }
        
    except Exception as e:
        return {'error': f"Error: {str(e)}"}

def query_excel_with_ollama(file_path, question, model_name="qwen3"):
    """Main function to query Excel data using Ollama"""
    try:
        # Load Excel data
        print("Loading Excel data...")
        data_text, df = load_excel_data(file_path)
        
        # Create optimized prompt
        prompt = create_effective_prompt(data_text, question)
        
        # Query Ollama with optimized settings for accuracy
        print("Querying Ollama...")
        response = ollama.generate(
            model=model_name,
            prompt=prompt,
            options={
                'temperature': 0.0,      # Zero temperature for maximum consistency
                'top_p': 0.8,           # More focused sampling
                'top_k': 20,            # Reduced for more precise answers
                'repeat_penalty': 1.15,  # Slightly higher to avoid repetition
                'num_ctx': 8192,        # Increased context window
                'num_predict': 1024     # Allow longer responses for detailed answers
            }
        )
        
        return {
            'question': question,
            'answer': response['response'],
            'data_shape': df.shape,
            'columns': df.columns.tolist()
        }
        
    except Exception as e:
        return {'error': f"Error: {str(e)}"}

def interactive_mode(file_path, model_name="qwen3"):
    """Interactive questioning mode"""
    print(f"\n📊 Excel Query System (Model: {model_name})")
    print("=" * 50)
    print("Type your questions about the Excel data.")
    print("Type 'quit' to exit, 'data' to see raw data preview")
    
    # Load data once
    data_text, df = load_excel_data(file_path)
    print(f"\n✅ Loaded: {df.shape[0]} rows, {df.shape[1]} columns")
    print(f"Columns: {', '.join(df.columns.tolist())}")
    
    while True:
        question = input("\n❓ Your question: ").strip()
        
        if question.lower() == 'quit':
            print("👋 Goodbye!")
            break
        elif question.lower() == 'data':
            print("\n📋 Data Preview:")
            print(df.to_string())
            continue
        elif not question:
            continue
        
        # Get answer
        result = query_excel_with_ollama(file_path, question, model_name)
        
        if 'error' in result:
            print(f"❌ {result['error']}")
        else:
            print(f"\n💡 Answer: {result['answer']}")
            if result.get('verification'):
                print(f"\n🔍 Verification data: {result['verification']}")

# Example usage
if __name__ == "__main__":
    # Configuration
    EXCEL_FILE = "twistlock_report_2.xlsx"  # Replace with your Excel file path
    MODEL_NAME = "qwen3"          # Change to your preferred Ollama model
    
    # Single question example
    # question = "What is the total count of records and what are the main categories?"
    # result = query_excel_with_ollama(EXCEL_FILE, question, MODEL_NAME)
    
    # if 'error' not in result:
    #     print("Question:", result['question'])
    #     print("Answer:", result['answer'])
    # else:
    #     print(result['error'])
    
    # Uncomment for interactive mode
    interactive_mode(EXCEL_FILE, MODEL_NAME)

# Quick test function
def quick_test():
    """Quick test with sample questions"""
    file_path = "twistlock_report_2.xlsx"  # Update this
    model = "qwen3"              # Update this
    
    test_questions = [
        "How many total records are there?",
        "What are all the column names?",
        "Show me the record with the highest value",
        "What is the average of the numeric columns?",
        "List all unique values in the first text column"
    ]
    
    print("🚀 Running Quick Test...")
    for i, q in enumerate(test_questions, 1):
        print(f"\n{i}. {q}")
        result = query_excel_with_ollama(file_path, q, model)
        if 'error' not in result:
            print(f"   → {result['answer']}")
        else:
            print(f"   ❌ {result['error']}")

# Uncomment to run quick test
# quick_test()s
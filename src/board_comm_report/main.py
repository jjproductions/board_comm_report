#!/usr/bin/env python
import sys
import warnings

from datetime import datetime

from board_comm_report.crew import BoardCommReport

warnings.filterwarnings("ignore", category=SyntaxWarning, module="pysbd")

# This main file is intended to be a way for you to run your
# crew locally, so refrain from adding unnecessary logic into this file.
# Replace with inputs you want to test with, it will automatically
# interpolate any tasks and agents information

def run():
    """
    Run the crew.
    """
    inputs = {
        'topic': 'AI LLMs',
        'current_year': str(datetime.now().year)
    }

    try:
        result = BoardCommReport().crew().kickoff(inputs=inputs)
        
        if hasattr(result, 'pydantic') and result.pydantic:
            report_data = result.pydantic
            file_name = report_data.file_name
            content = report_data.content
            
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✅ Successfully saved report to {file_name}")
        elif hasattr(result, 'json_dict') and result.json_dict:
            file_name = result.json_dict.get('file_name', 'Board_Committee_Reports_Draft.md')
            content = result.json_dict.get('content', str(result))
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✅ Successfully saved report to {file_name} from JSON dict.")
        else:
            print("Warning: Did not receive structured Pydantic output. Falling back to default filename.")
            with open("Board_Committee_Reports_Draft.md", 'w', encoding='utf-8') as f:
                f.write(str(result))
            
    except Exception as e:
        raise Exception(f"An error occurred while running the crew: {e}")


def train():
    """
    Train the crew for a given number of iterations.
    """
    inputs = {
        "topic": "AI LLMs",
        'current_year': str(datetime.now().year)
    }
    try:
        BoardCommReport().crew().train(n_iterations=int(sys.argv[1]), filename=sys.argv[2], inputs=inputs)

    except Exception as e:
        raise Exception(f"An error occurred while training the crew: {e}")

def replay():
    """
    Replay the crew execution from a specific task.
    """
    try:
        BoardCommReport().crew().replay(task_id=sys.argv[1])

    except Exception as e:
        raise Exception(f"An error occurred while replaying the crew: {e}")

def test():
    """
    Test the crew execution and returns the results.
    """
    inputs = {
        "topic": "AI LLMs",
        "current_year": str(datetime.now().year)
    }

    try:
        BoardCommReport().crew().test(n_iterations=int(sys.argv[1]), eval_llm=sys.argv[2], inputs=inputs)

    except Exception as e:
        raise Exception(f"An error occurred while testing the crew: {e}")

def run_with_trigger():
    """
    Run the crew with trigger payload.
    """
    import json

    if len(sys.argv) < 2:
        raise Exception("No trigger payload provided. Please provide JSON payload as argument.")

    try:
        trigger_payload = json.loads(sys.argv[1])
    except json.JSONDecodeError:
        raise Exception("Invalid JSON payload provided as argument")

    inputs = {
        "crewai_trigger_payload": trigger_payload,
        "topic": "",
        "current_year": ""
    }

    try:
        result = BoardCommReport().crew().kickoff(inputs=inputs)
        return result
    except Exception as e:
        raise Exception(f"An error occurred while running the crew with trigger: {e}")

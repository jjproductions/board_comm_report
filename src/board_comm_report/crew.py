from crewai import Agent, Crew, Process, Task, LLM
from crewai.project import CrewBase, agent, crew, task
from crewai.agents.agent_builder.base_agent import BaseAgent
from typing import List
import os
import yaml
from pathlib import Path
from dotenv import load_dotenv
from schema.models import CoordinatorOutput

load_dotenv()
# If you want to run a snippet of code before or after the crew starts,
# you can use the @before_kickoff and @after_kickoff decorators
# https://docs.crewai.com/concepts/crews#example-crew-class-with-decorators

# Load your externalized committee list
config_path = Path(__file__).parent / "config/committees.yaml"
with open(config_path, 'r') as f:
    committee_data = yaml.safe_load(f)
# Base URL for Ollama
base_url = os.environ.get("OLLAMA_BASE", "http://localhost:11434")
# Supervisor instructions
supervisor_instructions = """
You are the Board Operations Manager. Your goal is to oversee the creation of the Monthly Board Packet.
You have 3 specialized agents: 
1. Meeting_Coordinator: Accesses SharePoint and OneNote.
2. Document_Summarizer: Distills raw text into executive summaries.
3. Policy_Compliance_Officer: Finalizes formatting and policy alignment.

MANAGEMENT RULES:
- Always start by asking the Coordinator for the last and upcoming Board Meeting dates from SharePoint.
- You must ensure all committees ({committee_list}) are processed.
- If a committee has no meetings, instruct the Summarizer to note "No meetings held this period" for that section.
- Do not let the Policy Officer start until the Summarizer has finished all sections.
- Quality Control: If a summary is too vague, ask the Summarizer to revise it before finishing.
"""
# Higher-reasoning model for the Supervisor role
manager_llm = LLM(
    model=os.environ.get("SPV_MODEL", "ollama/mistral-small:24b-instruct-2506-q4_K_M"), 
    base_url=base_url,
    instructions=supervisor_instructions.format(committee_list=committee_data)
)
# Standard model for worker agents
worker_llm = LLM(model=os.environ.get("WKR_MODEL", "ollama/llama3.1"), base_url=base_url)


@CrewBase
class BoardCommReport():
    """BoardCommReport crew"""

    agents: List[BaseAgent]
    tasks: List[Task]

    # Learn more about YAML configuration files here:
    # Agents: https://docs.crewai.com/concepts/agents#yaml-configuration-recommended
    # Tasks: https://docs.crewai.com/concepts/tasks#yaml-configuration-recommended
    
    # If you would like to add tools to your agents, you can learn more about it here:
    # https://docs.crewai.com/concepts/agents#agent-tools
    @agent
    def document_summarizer(self) -> Agent:
        return Agent(
            config=self.agents_config['document_summarizer'],
            llm=worker_llm,
            verbose=True,
            allow_delegation=False
        )

    @agent
    def policy_compliance_officer(self) -> Agent:
        return Agent(
            config=self.agents_config['policy_compliance_officer'],
            llm=manager_llm,
            verbose=True,
            allow_delegation=True
        )

    @task
    def collection_objective(self) -> Task:
        return Task(
            config=self.tasks_config['collection_objective'],
            agent=self.meeting_coordinator(),
            output_json=CoordinatorOutput # This is the "Magic" line
        )

    @task
    def summarization_task(self) -> Task:
        return Task(
            config=self.tasks_config['summarization_task'],
            agent=self.document_summarizer()
        )

    @task
    def policy_alignment_task(self) -> Task:
        return Task(
            config=self.tasks_config['policy_alignment_task'],
            agent=self.policy_compliance_officer(),
            output_file='board_meeting_report.md' # Saves the final standardized report
        )

    @crew
    def crew(self) -> Crew:
        """Creates the Board App crew"""
        return Crew(
            agents=self.agents, # Automatically created by the @agent decorator
            tasks=self.tasks,   # Automatically created by the @task decorator
            process=Process.hierarchical, # Use Supervisor to manage tasks
            supervisor_llm=manager_llm,
            verbose=True
        )

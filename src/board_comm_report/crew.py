from crewai import Agent, Crew, Process, Task, LLM
from crewai.project import CrewBase, agent, crew, task
from typing import List
import os
import yaml
from pathlib import Path
from dotenv import load_dotenv
from schema.models import CoordinatorOutput
from board_comm_report.tools.custom_tool import get_tools

load_dotenv()

# Load external configuration for the supervisor and models
config_path = Path(__file__).parent / "config/supervisor_config.yaml"
with open(config_path, 'r') as f:
    supervisor_config = yaml.safe_load(f)

# Load your externalized committee list
committees_path = Path(__file__).parent / "config/committees.yaml"
with open(committees_path, 'r') as f:
    committee_data = yaml.safe_load(f)

# Base URL for Ollama
base_url = os.environ.get("OLLAMA_BASE", supervisor_config['ollama_base'])

# Supervisor instructions formatted with the committee list
supervisor_instructions = supervisor_config['supervisor_instructions'].format(committee_list=committee_data)

# Higher-reasoning model for the Supervisor role
manager_llm = LLM(
    model=os.environ.get("SPV_MODEL", supervisor_config['manager_model']),
    base_url=base_url,
    instructions=supervisor_instructions
)

# Standard model for worker agents
worker_llm = LLM(
    model=os.environ.get("WKR_MODEL", supervisor_config['worker_model']), 
    base_url=base_url
)

@CrewBase
class BoardCommReport():
    """A team of AI agents designed to automate the generation of board reports."""

    # If you want to run a snippet of code before or after the crew starts,
    # you can use the @before_kickoff and @after_kickoff decorators
    # https://docs.crewai.com/concepts/crews#example-crew-class-with-decorators
    
    @agent
    def meeting_coordinator(self) -> Agent:
        return Agent(
            config=self.agents_config['meeting_coordinator'],
            llm=worker_llm,
            verbose=True,
            tools=get_tools(),
            allow_delegation=False
        )

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
            output_json=CoordinatorOutput
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
            output_file='board_meeting_report.md'
        )

    @crew
    def crew(self) -> Crew:
        """Creates the Board App crew"""
        return Crew(
            agents=self.agents,
            tasks=self.tasks,
            process=Process.hierarchical,
            supervisor_llm=manager_llm,
            verbose=True
        )

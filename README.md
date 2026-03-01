# BoardCommReport Crew

Welcome to the BoardCommReport Crew project! This project is an automated AI system built upon the [crewAI](https://crewai.com) framework. Its primary objective is to streamline the creation of Monthly Board Reports for nonprofit organizations by synthesizing committee meeting minutes.

## Overview

The **Board Operations Manager** (a Supervisor AI) manages specialized agents to compile comprehensive executive summaries from raw meeting notes stored in SharePoint and OneNote. It employs a hierarchical process using local Ollama models (`mistral-small` for the manager, `llama3` for worker agents) to ensure high-quality and policy-compliant reports.

### Agents
The crew consists of three specialized AI agents (defined in `config/agents.yaml`):

1. **Meeting Coordinator (Board Data Librarian)**: 
   Navigates Microsoft 365 environments, accesses the Board SharePoint events calendar to find meeting dates, and extracts recent meeting pages from specific committee notebooks in OneNote.
2. **Document Summarizer (Board Meeting Document Specialist)**: 
   Synthesizes the raw meeting minutes extracted by the Coordinator into clear, actionable summaries. It extracts 'Motions' (Passed/Failed), identifies 'Action Items', and notes any 'High Priority' items requiring a Full Board vote.
3. **Policy Compliance Officer (Nonprofit Policy Auditor)**: 
   Performs a final audit of the report, ensuring that the documents meet formatting standards, incorporate correct version-controlled policies, and follow organizational guidelines, yielding a ready-to-upload professional Markdown draft.

### Tasks
The workflow involves three main tasks (defined in `config/tasks.yaml`):

1. **collection_objective**: Accesses SharePoint, iterates over the defined committees, and extracts the raw meeting minutes from OneNote.
2. **synthesis_objective**: Drafts a concise Board Report highlighting key motions and action items from the raw minutes.
3. **final_audit_objective**: Conducts quality assurance to produce a submission-ready Markdown document `board_meeting_report.md`.

## Configuration

### Customizing the Committees

The target committees and their corresponding OneNote Notebook IDs are externalized in `src/board_comm_report/config/committees.yaml`. By default, it covers:
- Finance (`NB_FIN_001`)
- Governance (`NB_GOV_002`)
- Human Resources (`NB_HR_003`)
- Fundraising (`NB_FUN_004`)
- Executive (`NB_EXE_005`)

### Environment Settings

Ensure you have Python >=3.10 <3.14 installed.

**Local AI Engine Requirements:**
Be sure to have local Ollama instances running to power the language models. The default setup connects to Ollama at `http://localhost:11434` and uses:
- Supervisor Manager Model: `ollama/mistral-small:24b-instruct-2506-q4_K_M`
- Worker Model: `ollama/llama3`

If needed, update the `manager_llm` and `worker_llm` instances in `src/board_comm_report/crew.py` to match your local runtime configuration.

## Installation

This project uses [UV](https://docs.astral.sh/uv/) for dependency management and package handling.

First, if you haven't already, install uv:

```bash
pip install uv
```

Next, lock the dependencies and install them by using the crewai CLI command:

```bash
crewai install
```

## Running the Project

To kickstart the crew of AI agents and launch the reporting task execution, run this from the root folder of your project:

```bash
crewai run
```

This command initializes the hierarchical `BoardCommReport` Crew, assigning tasks sequentially and letting the supervisor manage the execution. It ultimately outputs your processed summary report as `board_meeting_report.md` in the root folder.

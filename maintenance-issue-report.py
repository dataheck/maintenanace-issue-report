from dotenv import load_dotenv
from github import Github
import os

load_dotenv()

mandatory_configuration = ('GITHUB_API_KEY', 'GITHUB_ORGANIZATION','GITHUB_PROJECT_NAME', 'GITHUB_PROJECT_FINISHED_COLUMN')
config = dict()

for key in mandatory_configuration:
    value = os.environ.get(key)
    assert value is not None, f"Please set {key} before proceeding."
    config[key] = value

g = Github(config['GITHUB_API_KEY'])
org = g.get_organization(config['GITHUB_ORGANIZATION'])
project = None

for this_project in org.get_projects():
    if this_project.name == config['GITHUB_PROJECT_NAME']:
        project = this_project
        break

assert project is not None, f"Project '{config['GITHUB_PROJECT_NAME']}' was not found in the organization, please verify configuration."

project_column = None

for this_column in project.get_columns():
    if this_column.name == config['GITHUB_PROJECT_FINISHED_COLUMN']:
        project_column = this_column
        break

assert project_column is not None, f"Project column '{config['GITHUB_PROJECT_FINISHED_COLUMN']}' was not found in the project, please verify configuration."

for card in project_column.get_cards():
    this_issue = card.get_content()
    print(this_issue)
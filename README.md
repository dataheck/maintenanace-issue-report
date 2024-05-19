# Issue Reporting for GitHub Projects

I want my clients to have a good idea of what they are paying for when they have a maintenance contract with me. My primary project management tool is GitHub projects and issues, but clients tend to not be users of GitHub and their repositories are private. 

This tool makes it easy for me to share the detail stored on GitHub in an business-friendly format that can be kept for accounting records.

# Suitability

Here's my rough workflow - if it's similar to yours then you might benefit from this tool:
- Issues are added to a project specific for maintenance tracking, which is under an organization
- The project has a "Done" column which is cleared after every billing cycle
- At the end of every billing cycle a report is sent with a list of all completed issues for that cycle

This tool has two features that help with this:
1. It connects to the GitHub API and grabs all of the issues in the "Done" column and updates a Word .docx template with their information
2. Every issue's URL is followed via a Selenium *Chrome* instance and is automatically printed to PDF to a dedicated folder

The Selenium print to PDF is aided by [this script](https://gist.github.com/gaute/1357711/1c19e061c66fab71337b5b9b51f82b5abfb97f46) which improves the look of printed GitHub issues. You should install it to your Chrome before proceeding with this feature.

# Configuration

Configuration is handled via `dotenv`. Here are the configuration variables:

| Variable Name | Description | 
|---------------|-------------|
| GITHUB_API_KEY | An API key for GitHub that has access to GITHUB_ORGANIZATION and GITHUB_PROJECT_NAME. |
| GITHUB_ORGANIZATION | The name of the organization GITHUB_PROJECT_NAME is a child of |
| GITHUB_PROJECT_NUMBER | The number that appears at the end of the URL to its dashboard | 
| GITHUB_PROJECT_FINISHED_COLUMN | The column of the project that tracks finished items. For example, "Done" |
| PDF_SAVE_PATH | The path to an existing directory where all PDFs will be "printed" to |
| COVERPAGE_TEMPLATE_PATH | The path to a Word .docx file that meets our expectations. See Template Expectations, below. | 
| CLIENT_NAME | The name of your client's business |
| CLIENT_CONTACT | The name of your contact that your report is being sent to | 
| PROJECT_NAME | An identifier for your project or maintenance contract. | 
| OUTPUT_PATH | The full path including filename to the output .docx file |

# Template Expectations

The .docx template is assumed to have the following properties:

* The following custom properties, ideally linked somewhere in the text:
    * ClientName
    * ClientContact
    * ProjectName
* The following styles:
    * Closing Paragraph
    * List Paragraph

# Caveats

You must update your custom fields after producing a file with this tool. You can do this in Microsoft Word by pressing `CTRL+A` to select everything followed by pressing `F9` to update fields.

This tool uses a specific version of Chrome (103.0.5060.134.0) and a specific fork of `python-docx` (by @michael-koeller and @renejsum), for custom properties support.

Version 2.0.0 of this tool has dropped support for legacy projects in favour of the new project system accessible only via GraphQL. 

# Installation / Requirements

* Install [poetry](https://python-poetry.org/)
* Install [Chrome](https://www.google.com/intl/en_us/chrome/) v. 103.0.5060.134.0 (newer might work, untested. I'd expect printing issues with a future update.)
* Create / obtain an appropriate [GitHub API key](https://github.com/settings/tokens)
* Ensure your organization / project / column structure is appropriate for this tool
* Create a suitably formatted .docx template (see Template Expectations, above)
* Clone this repository
* Run "poetry install" in the same folder you cloned it to
* Execute `maintenance-issue-report.py` with `python`. Specify --enable_print TRUE if you would like it to export PDFs of the issues.

Selenium is not required if you do not indent to use the print to PDF functionality. 

If you do use the PDF functionality, I recommend you merge the coverletter with all of the exported files into a single PDF. No user wants to have dozens of PDFs open!

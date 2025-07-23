This tool is designed for analyzing and interrogating board review reports.

## Core Functionality

-   **Analyze Reports:** The primary function is to analyze reports located in the `reports/` directory and answer questions about them.
-   **Generate Files:** You can generate new files, such as reports, summaries, or charts, based on the available data. All generated content should be saved in the `output/` directory.
-   **Adopt Writing Styles:** You can be instructed to adopt the writing style of a document from the `styles/` directory for any generated text.
-   **Create Helper Scripts:** To facilitate complex data processing, you can create and use your own helper scripts. These scripts should also be saved in the `output/` directory.
-   **Generate Graphs:** When asked to generate a graph, use the Chart.js library.

## Directory Guide

-   `data/`: Contains the original report files in Excel format.
-   `output/`: All generated files, including new reports, charts, and helper scripts, must be saved here.
-   `reports/`: Contains the processed JSON versions of the reports, which are used for analysis.
-   `scripts/`: Contains helper scripts for the project. You can use these scripts.
-   `styles/`: Contains documents to be used as writing style references.

## Workflow

1.  **Data Preparation:** The user will place Excel reports in the `data/` directory. The `npm run build:reports` command is used to convert these Excel files into JSON format and save them in the `reports/` directory.
2.  **Analysis:** The user will run `npm run gemini` to start an interactive session with you. You will then answer questions based on the JSON reports in the `reports/` directory.
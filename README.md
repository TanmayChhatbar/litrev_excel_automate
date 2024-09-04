# Excel Literature Review Automation

## Features
- Automatic DOI Metadata Fetching
- Dynamic arrays for filtering the main table
- Comparison graphs
- Reference generation with automated/manual reference number allocation

### Automatic DOI Metadata Fetching
The included script automates the process of fetching metadata for DOIs (Digital Object Identifiers) from the CrossRef API and populating an Excel sheet with the retrieved information.

#### Features
- DOI Validation: It checks if the selected cell contains a valid DOI by ensuring it starts with "10" and contains a "/".
- CrossRef API Integration: The script fetches metadata for each DOI from the CrossRef API, including:
  - Author(s)
  - Title of the work
  - Year of publication
  - Number of citations
  - Type of publication
  - Publisher
  - Hyperlink Creation: Adds a hyperlink to the DOI in the Excel sheet that redirects to the CrossRef API page for that DOI.

#### Script Usage
- Import the script: In the Code Editor in Excel, add a 
- Selection: Select a single column in the Excel sheet containing DOIs.
- Run the Script: The script will validate each DOI, fetch the relevant metadata, and populate the adjacent cells with the retrieved information.

#### Requirements
- Excel Online with Office Scripts enabled.
- DOIs should be formatted correctly in the selected column.

### Dynamic arrays for filtering the main table
- In separate worksheets named "Table - xyz", index the rows you want to generate the table for in the order you want them. This filters the columns.
- To filter which rows to show, change the "ISNUMBER(SEARCH("Formation",Main[Target]))" in the cell that generates the table to "ISNUMBER(SEARCH("_word_that_should_be_in_cell_",Main[_column_to_check_]))"

### Reference generation
- The reference list in the references worksheet will include all the literature that has a non-black _Ref_ column.
  - If you want to allot specific indices, use unique numbers
  - If you don't care about the index of the literature, use a large value.

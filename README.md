# combinepdf
Simple utility for combining many PDFs into one

![Image of combinepdf UI](/combinepdf1.png)

## Description
This python script creates a GUI that permits users to:
- Select a source folder containing multiple files desired to be combined
- Select a destination folder where the single new PDF file will be created
- Provide a name for the new PDF file
- Gives the option to convert '.docx' files within the source folder to PDF and include in the new PDF

### Notes
- Combines files in alphabetical order
- Files that are not '.pdf' are automatically excluded from the PDF combining process and do not need to be manually sorted from the source folder
- Automatically removes created '.pdf' files derived '.docx' originals after combining
- Does not work with password protected PDFs

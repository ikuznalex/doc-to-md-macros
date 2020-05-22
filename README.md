# doc-to-md-macros
Microsoft Word macros which converts content of word document file to markdown syntax.

## How to use?

### Do these steps only the first time
1. Open your word document and then hit Alt+F11.
You should see _Microsoft Visual Basic for Applications_ window
2. Right click on **Normal** project. Insert -> Module
   - In this case the macros will be available for any document on your machine
3. Select the code of the module from [this page](https://github.com/ikuznalex/doc-to-md-macros/edit/master/ConvertDocToMarkdown.vb)
4. Close the Visial Basic for Applications window.

### Do these steps each time you want to convert a document
 - Open the word document you want to convert
 - Hit Alt + F8 (or choose View->Macros)
 - Run the macros _DocToMarkdown_ 
 - The result of conversion you should see right inside the word document
 
 ### Supported elements (will be converted to markdown)
  - Heading Styles from 1 to 5
  - Text formatting
    - _Italic_
    - **Bold**
    - <u>Underline</u>
    - Combination fo styles above
  - Links
  - First-level lists
    - Numbered
    - Bulleted
  - Tables 


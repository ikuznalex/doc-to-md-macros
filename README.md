# doc-to-md-macros
Microsoft Word macros which converts content of the document to markdown(.md) syntax.

## How to use?

### Do these steps only the first time
1. Open your word document and then hit Alt+F11.
You should see _Microsoft Visual Basic for Applications_ window
2. Right click on **Normal** project. Insert -> Module
   - In this case the macros will be available for any document on your machine
   <img width="304" alt="add-module" src="https://user-images.githubusercontent.com/5716707/82793442-92d6bf00-9e79-11ea-90b4-c747fd642e81.png">
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

### What to do if something is not working correctly or you have an enhancement idea?
 - Add the issue with _bug_ or _enhancements_ label to the [Issues](https://github.com/ikuznalex/doc-to-md-macros/issues) section. 
 - Create pull request with the fix or enhancement.
Thanks!

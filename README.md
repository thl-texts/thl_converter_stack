# THL Converter Stack

## About

**Repo Name:** `thl_converter_stack`

This is the refactored converter from Ori's original attempt at https://github.com/thl-texts/thl_converter.git
It is called "thl_converter_stack" because it uses a headstack to keep track of the current nested DIV
It also creates a TextConverter class that is the basis of the conversion. Thus, instead of having global
variables it uses class properties. So instead of keeping track of the process with global variablase such as 
`list_open`, the class has a current_el parameter, which is the last added element. Through lxml _Element properties
and methods such as element.tag (for tag name), list(element) for children, and element.getparent(), the script can 
determine its current context before doing any action.

The refactoring makes the code cleaner, neater, and easier to understand and update. But much of Ori's original logic
was ported over. So his inital pass saved much time allowing for this improvement to happen.

## Usage
After downloading this repo and installing the python required modules from the requirements.txt file, 
the following folder structure needs to be set up in the repo folder:

    workspace
      |__ in
      |__ logs
      |__ out

The `workspace` folder is not included in the git repo. 
Place the documents you want to convert in the "in" folder.

Then run the `main.py` file. The in-folder, out-folder, log-folder, and metadata XML template can all be changed by 
parameters which can be seen through `python main.py --help`

For instance, to use the old metadata form template as well as replace any existing XML output file, one would do:

`python main.py -t tib_text_old.xml --overwrite`

The converted files will be found in the out folder.

## Formatting Word Docs 
The converter takes essays or texts in Microsoft Word documents and converts them to THL XML. To do so,
the text in the Word docs must be "marked up" with THL custom styles. These are found in 
[our custom Word Template](https://drive.google.com/file/d/1RN71aJESmmQq4cQaZIVd_I8hzqJaZahx/view?usp=sharing). 

The
instructions for how to use the styles can be found in the 
[THL Text Editing Manual style guide](https://docs.google.com/document/d/1BJEwSXzXwwqgY9xPbNor-RmsZHpmVqjOb6JMwTiPVUY/edit). 


At the beginning of the Word document should be a 
[metadata table](https://drive.google.com/file/d/16pzm1cxMgGZTccU9-kY72hSKC2ihTZQd/view?usp=sharing). 
Instructions for filling out the metadata table are also found in the editing manual.


Than Grove  
Created Oct. 30, 2020  
Updated March 10, 2021


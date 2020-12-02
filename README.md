# thl_converter_stack

## About
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
        in
        logs
        out

Place the documents you want to convert in the "in" folder.

Then run the `main.py` file. The in-folder, out-folder, log-folder, and metadata XML template can all be changed by 
parameters which can be seen through --help

The converted files will be found in the out folder.

Than Grove
Created Oct. 30, 2020
Updated Dec. 2, 2020


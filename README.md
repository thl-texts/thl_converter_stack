# thl_converter_stack

This is the refactored converter from Ori's original attempt at https://github.com/thl-texts/thl_converter.git
It is called "thl_converter_stack" because it uses a headstack to keep track of the current nested DIV
It also creates a TextConverter class that is the basis of the conversion. Thus, instead of having global
variables it uses class properties. So instead of keeping track of the process with global variablase such as 
`list_open`, the class has a current_el parameter, which is the last added element. Through lxml _Element properties
and methods such as element.tag (for tag name), list(element) for children, and element.getparent(), the script can 
determine its current context before doing any action.

The refactoring makes the code cleaner, neater, and easier to understand and update. But much of Ori's original logic
was ported over. So his inital pass saved much time allowing for this improvement to happen.

Than Grove
Oct. 30, 2020

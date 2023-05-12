# VBA - Learning

### Table of Content
- [Enabling Developer Tab]('#enabling-developer-tab')
- [Recording the first macro]()


### Enabling Developer Tab
1. Right-click any part of the Ribbon and choose Customize the Ribbon from the shortcut menu.
2. In the Customize Ribbon tab of the Excel Options dialog box,
locate Developer in the box on the right.
3. Put a check mark next to Developer.
4. Click OK.

#### Recording a Macro
Ensuring that Relative References Button is active. 
1. Select a cell
2. Go to developer tab -> Record Macro option, click on it. 
3. A dialogue box will appear, fill the details like macro name (this will result as the subroutine name which will be creating as a result), description. 
4. Ensure that the 'Store Macro in ' setting has 'This Workbook' as its value, also you can assign a shortcut to your macro. 
5. Click on ok when you are done, and then perform the steps you want to be recorded.
6. Once that's done, click on the stop recording part. 
7. Now you can visit the Devleoper -> View VBA to view the codes or if you have assigned any shortcut key then you can use that to trigger the macro steps. 

**Note:** Comments in macro have line beginning by `'`. 


#### Visual Basic Editor

Visual Basic Editor or VBE is a separate excel application where you can write and edit your VBE macros. 
Beginning with Excel 2013, every workbook is displayed in new window but ony one VBE window is there which can work with all open Excel workbooks. 

**Note:** You can't run VBE separately; Excel should be running to run your VBE. 

--- 

#### Activating VBE

- Press Alt + F11 to get into VBE, while keeping excel opened, same again to go back to excel workbook. 
- Alternatively you can go, Developer -> Visual Basic 


#### VBE Components

1. Menu Bar
2. Tool Bar
3. Project Window
4. Code Window
5. Immediate Window
6. Properties Window

- You can press **Ctrl + R** to view the project window. 
- You can press **Ctril + G** to view your immediate window. 


#### Inserting a new VBA Module

**Method 1:** Using the Menu Bar
1. Select the project.
2. Click on the *Insert* menu option
3. Select *Module* Option

**Method 2:** 
1. Right click on the project. 
2. In the option list select *insert* -> *module*

#### Removing a module from the project window

**Method 2:** Using File option in Menu Bar
1. Select the module
2. Click on file -> remove <module_name>

**Method 3:**Right clicking on the module
1. Right Click on the module. 
2. Select 'remove module_name' in the option menu

#### Exporting an object 
1. Select the object in the project window
2. Go to file and click on the 'export file' option or press *Ctril + E*.

#### Importing an object
1. Select the project's name in the project window
2. File -> Import File or press *Ctrl + M*.
3. Locate the file and click on *Open*.

**Note:** You can go to *Window-> Tile Vertically/Horizontally* etc. to tile your windows in code window. 
**Note:** You can use *Ctrl+F6* to cycle through your code windows or *Ctrl+Shift+F6* to cycle in the reverse order.

### Module has three types of codes
1. Declarations
2. Sub procedures
3. Function Procedures

There's no limit on the number of either of the above but there's limit on the max number of characters per module, that is 65,000 characters. 

**Note:**To continue single line of the code onto other, you have to enter '_' at the end of the line, ensure that there's a whitespace before it and 
continue normally in the second line. You can indent to make it clearer for others that it's the continuation from the previous line...but that's optional. 

**Note:** You can press *Alt+F11* to switch between the VBA and Excel workspace windows. 

**Note:** *Application.UserName* gets you the username of the excel user currently. 

**Note:** You can press *F5* to run the procedures in the current module. 

**Note:** VBA procedures are also known as *Macros*.

**Note:** VBA concatenation is done using *&* ampersand operator. 

**Note:** *vbYesNo*, *vbYes*, *vbNo* are different type of inbuilt command values in VBA. 

#### Display all the macros in your current project. 

1. Go to *Developer* menu option
2. Click on *Macros*

#### Viewing the VBA Tools 
1. Tools -> Options

**Note:** You can type *Option Explicit* to ensure that a variable is declared before use. 
**Note:** You can use *Shift+Tab* and *Tab* to *unindent* and *indent* the codes in VBA.
**Note:** You can Drag & Drop selected text from one code window to another. 

If the settings are enabled in the Tools->Option Menu. 



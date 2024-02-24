# DataStructureControl-Plugin
Excel plugin for comparing data in ***.xlsx*** format.

![alt text](exampleXl/ribbon.png?raw=true)

The foundation of the plugin involves preparing template file, which is then converted to ***.xlcg*** format in the final stage. During the preparation process, both unchangeable (hard) values and values that can be disregarded by macros can be added. Once the file is created, it can be compared to a selected ***.xlsx*** file on a (1:1) basis or processed for multi-comparison through a GUI system (1:N).

![alt text](exampleXl/gui.png?raw=true)

Result of comparison actions, are errors that are visually added to specific workbooks. Additional features include a macro that cleans the Excel file of previously added errors and a macro that creates a workspace based on directories (for 1:N comparisons). To ensure the proper functioning of all macros, it is essential to include the class file (***.cls***) in your current workbook.

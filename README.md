# DataStructureControl-Plugin
Data Structure Control is excel plugin which makes possible to compare data in ***.xlsx*** format.

![alt text](exampleXl/ribbon.png?raw=true)

*example ribbon*

The basis of the plugin is preparing template file that on the last stage is converted to ***.xlcg*** format. Due to prepare process we can add unchangeable (hard) values and values that can be disregard by macros. After creating this file, we can compare it to selected ***.xlsx*** file (1:1) or process multi compare by gui system (1:N).

![alt text](exampleXl/gui.png?raw=true)

*gui for multi compare*

Result of compare actions are errors which are visually added to specific workbooks. Other features are including macro that cleans excel file of previously added errors and macro that creates workspace - based on directories (for 1:N compare). To get all macros working properly it is essential to include class file (***.cls***) in your current workbook.

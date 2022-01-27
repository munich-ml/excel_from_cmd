# Apply VBA Macros to Excel tables from CMD
H. Steffens, 27.01.2022

# Problem
Assuming there is some data source generating unformated Excel tables on a regular basis. Formatting the Excel tables manually can be done once, but shouldn't be done regurlary. The manual formatting operations can be recorded using the Excel Macro Redorder und exported as `.bas`-file.

This article shows how to apply those macros to new Excel tables controlled from the command line or batch scripts (`.bat`).

# Solution
![](imgs/SW_architechture.drawio.svg)

## top_level_script.bat
```cmd
cd /d "%~dp0"
call launcher_script.vbs %CD%\raw_data.xlsx %CD%\format_macros.bas "Format_and_export" %CD%\formated_export.xlsx
```
In this simple example the __top_level_script__ (`.bat`) doesn't do more than calling the __launcher_script__. Within a real-world example, the __top_level_script__ would call other scripts (e.g. the one that generated the `raw_data.xlsx`) and hold the paths as CONSTANTs.

Also in this simple example the save_path (`%CD%\formated_export.xlsx`) is given to the __launcher_script__ as an argument.

## launcher_script.vbs
Written in __Visual Basic Script__ (__VBScript__ or __VBS__) which is just similar to __VBA__, but not identical.  

- The script `launcher_script.vbs` receives 4 Arguments:
    ```VBS
    Set args = Wscript.Arguments
    raw_data = args.Item(0)
    macro_path = args.Item(1)
    macro_name = args.Item(2)
    export_path = args.Item(3)
    ```

- lauches Excel invisible:
    ```VBS
    Set xl = CreateObject("Excel.application")
    xl.Application.Visible = False
    ```

- opens an Excel workbook from the path `raw_data` and imports the macros from `macro_path`:
    ```VBS
    Set xlBook = xl.Workbooks.Open(raw_data, 0, True)
    xlBook.VBProject.VBComponents.Import macro_path
    ```
- it then calls the macro by it's `macro_name` and hands over the `export_path` as an argument:
    ```VBS
    xl.Application.run macro_name, export_path
    ```

- finally, it quits Excel

## format_marcros.bas
The `format_marcros.bas` __Visual Basic for Applications (VBA)__ contains all operations that will be applied to the data (in this example formatting). It has not been written manually, but recorded using the __Excel Macro Recorder__ (go `Developer` tab and click `Record Macro` and export as `.bas`-file afterwards).  

# Ressources
The forum discusson ["Run Excel Macro from Outside Excel Using VBScript From Command Line"](https://stackoverflow.com/questions/10232150/run-excel-macro-from-outside-excel-using-vbscript-from-command-line) was very helpful when deveoping this solution.
Set args = Wscript.Arguments
raw_data = args.Item(0)
macro_path = args.Item(1)
macro_name = args.Item(2)
export_path = args.Item(3)

Dim xl
Dim xlBook  
Set xl = CreateObject("Excel.application")
Set xlBook = xl.Workbooks.Open(raw_data, 0, True)
xlBook.VBProject.VBComponents.Import macro_path
xl.Application.Visible = False ' Show Excel Window
xl.DisplayAlerts = False  ' suppress prompts and alert messages while a macro is running
xl.Application.run macro_name, export_path
xlBook.saved = True ' suppresses the Save Changes prompt when you close a workbook
xl.activewindow.close
xl.Quit


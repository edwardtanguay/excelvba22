# excelvba22

This project contains two files that can be used to solve problems with Excel: (1) *main.xlsm* is a worksheet which has VBA code modules that enable you to more quickly built an interactive Excel sheet, and (2) *intervalProcessor.xlsm* is a worksheet which processes code at a given interval, e.g. once per second or once per minute which enables you to create solutions outside of Excel itself, e.g. monitor when files are copied to directories, display them in excel sheets, update excel sheets and other files, access databases on the network, etc.

**Explore code**

- for basic VBA interactivity, in the worksheet *Main*, try all the buttons
- for singular/plural class code, in the VBA code (F11), open the module *DataPersons* and in *TestPluralClass()* click F5
	- this reads the data from the DataPersons worksheet and builds a collection of objects, then iterates through it and writes the data out to a text file in the directory where the Excel file is

**How to use main.xlsm for a new project:**

- open the main.xlsm file
- experiment with the examples on the worksheet
- examine the code
- delete everything on the worksheet
- use for your own purposes

**How to use intervalProcessor.xlsm:**
- in module *qexc* change how often you want code executed, e.g. once per second, once per minute, etc.
- in module *tools* add to the method *tool_ObserverAction* any code you want executed regularly

**How to use the q___.vb* fils**

- use these to find VBA code without having to open the Excel file, e.g. quickly via GitHub

**Features of modules:**

- qstr_chopRight - removes text from the right side of a string

## Current Developers

* Edward Tanguay [@edwardtanguay](https://github.com/edwardtanguay)

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

This project uses the [MIT License](https://choosealicense.com/licenses/mit). Feel free to use, change, share, and distribute freely.
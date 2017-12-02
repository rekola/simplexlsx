2017-12-02 Version 0.24
	Change log:
		- Added checking of sheet names: length up to 31 characters and forbidden characters "/\[]*:?" are replaced with '_' (Excel restrictions)

2017-05-09 Version 0.23
	Change log:
		- Improved scatter charts to more "scientific" appearance, including color and symbol selection and width of the lines (thanks for this to E.Naumovich).
		- Creating charts in the worksheets by CWorkbook::AddChart. For chart sheets must be used CWorkbook::AddChartSheet (thanks for idea to Aso).
		- Now the sheets are arranged in order of creation. Using CWorkbook::SetActiveSheet can be set the active sheet (by default the first sheet).
		- CWorkbook::m_styleList was moved to private. Now must be used CWorkbook::AddStyle() and CWorkbook::GetFonts().
		- Many internal changes.
		- Added examples to the archive.

2017-02-22 Version 0.22
	Change log:
		- Support for the TDM-GCC 64-bit compiler (thanks to Eduardo Baena).
		- Minor bug fixes.

2017-01-28 Version 0.21
	Change log:
		- XmlWriter is rewritten to work with Unicode, implemented writing of strings in UTF-8.
		- Rewrote all the code to work with the new XmlWriter. All string constants for XML stored in the type const char *, which allows speed up the work and reduce the size of the executable file for the Unicode version.
		- Minimal use _stprintf functions, which is replaced by _tstringstream for the safety (no fixed buffer length).
		- Fixed bugs in the charts module, because of what Excel can not open the file or display it correctly. Compatibility is also upgraded to Office 2010 because of the charts.
		- For fields valAxisFrom and valAxisTo (from CChartsheet::Series) made standard numbering for rows (from 1)
		- For CWorkbook you can specify the name of the author. If not specified, the current user name is queried from the Operation system.
		- Many functions AddCell of the CWorksheet class can take the "raw" data without the need to create wrapper classes (for example, without CellDataStr or CellDataInt).

2013-08-31 Version 0.20
	Change log:
		- Minor bug fixes

2013-04-13 Version 0.19
	Change log:
		- Numeric, dates and times styling
			Most general formats are available. It is also possible to create custom formats
			Some information about how to create a format is here:
				http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HP001216503.aspx
		- Minor bug fixes
		- Makefile to build the project as a library (by Eduardo Zacarias)

2012-12-16 Version 0.18
	Change log:
		- Comments functionality added;
		- Font name bug fix
		- Borders functionality changed;
		- Sheet orientation feature added;
		- Some additional features in Chartsheet.

2012-10-31 Version 0.17
	Change log:
		- Creation time format bug fix.

2012-10-26 Version 0.16
	Change log:
		- Temporary directory location bug fix;
		- Creation time format bug fix.

2012-10-02 Version 0.15
	Change log:
		- Linux are supported now (tested on Ubuntu with GCC 4.6.3)
			- ZIP library ported (under Linux only compressing of files is available)

2012-09-21 Version 0.14
	Change log (bug fix):
		- Overflow in shared strings;

2012-09-21 Version 0.13
	Change log (bug fix):
		- It is enough to include just 'Xlsx/Workbook.h' instead of three files;
		- AddCell();
		- Unique directory for each new book;
		- Temporary files are deleted after book saving even if saving is failed;
		- In Worksheet/Chartsheet classes` objects 'ok' flag is set to false after saving (since stream either closed or data corrupt at another trying);
		- Some minor changes inside;

2012-09-09 Version 0.12
	Change log:
		- Proper (without repetitions) shared strings implemented;
		- Bug fix - wrong span range for row;
		- Bug fix - nonsense style id by default

2012-09-01 Version 0.11
	Change log:
		- AddCell() method to easier addition of empty cell added.

2012-08-05 Version 0.1
	First draft of the library. 
	
	Main features for data sheets formatting:
		- adjustable height for each row;
		- adjustable width for columns. Can be specified only at sheet creation;
		- frozen panes;
		- page title designation;
		- data can be added by row, by cell or by cell groups;
		- cells merging;
		- formulae recognition (without formula`s content verification);
		- styles support (fonts, fills, borders, alignment, multirow text);

	Main features for chart sheet configuration:
		- supported chart types: linear, bar, scatter;
		- customizable settings: legend presence and position, axes` labels;
		- data table (for linear and bar chart types);
		- additional axes pair (the type defines independently from main chart type);
		- data source can set for both X and Y axes for both main and additional axes pairs;
		- curve`s smoothness, node`s mark;
		- grid settings for each axis;
		- minimum and maximum values for each axis.
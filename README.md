# Excel-dumper
Processing code for Dumping Strings array into Excel file.
Core functions are fully tested.
You'll have to run this project with all the dependesis (in the folder code).

### Code snippet

Relative path need to pass via the constructor.
```processing
	ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
```

All the cells are of string type.

You can export initial table with header and body.

```processing
	ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
	ed.exportExcel(header,pass)
```

| Header1       | Header2       	|	Header3	|
| ------------- |:-------------:	|	-----:	|
| String      	| String 			|	String	|
| String      	| String			|	String 	|
| String		| String			|	String 	|

Or insert new row

```processing
	String[] toInsert = {"new","new","new"};

	ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
	ed.setLogsOn(true);
	ed.insertToExcel(toInsert);
```
| Header1       | Header2       	|	Header3	|
| ------------- |:-------------:	|	-----:	|
| String      	| String 			|	String	|
| String      	| String			|	String 	|
| String		| String			|	String 	|
| new			| new				|	new 	|


Add new column will fill all rows data in this column with nulls.

```processing
  ExcelDumper ed = new ExcelDumper(sketchPath("") + "\\Pressure\\sensor ATE Data base.xlsx");
  ed.setLogsOn(true);
  
  ed.addNewColumn("New!!");
```

| Header1       | Header2       	| Header2       	|	New!!	|
| ------------- |:-------------:	|:-------------:	|	-----:	|
| String      	| String 			| String 			|	null	|
| String      	| String			| String			|	null	|
| String		| String			| String			|	null	|
| new			| new				| new				|	null	|




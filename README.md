# Export2ExlEngine
Export datatable/dataset to excel/csv file quickly
<b>A requirement for use - a get an IgalDAL project beforehand </b>

The DLL has 2 classes - <h3>Export2Exl</h3> and <h3>Export2Csv</h3>.
Excel file is made especially to export large amounts of data in just a seconds.
<b>A prerequisite for  Export2Exl is Excel app and sql server. </b>
Both classes can get a Datatable or DataSet. 
Transferring a Dataset with several datatables in it will produce a same number of excel sheets. 
The name of each sheet will be the same as a name of a DataTable.
In CSV version - several files with adding a number after each.

A use can be very simple:

int m_exported = 0;<br>
Export2Exl XlUtil = new Export2Exl(PublicModule.sConnection);<br>
m_exported = XlUtil.ExportToExcel(dt);<br>

this will open excel application and import a data into it.

there're option to give a filename - XlUtil.XlFileName, or just XlUtil.SaveFile ([filename])
XlUtil.SilentOpen = true will just save a file opening an application in a background. BTW, closing Excel is smooth 
and doesn't leave shaddows in Windows task manager.
XlUtil.SuppressFileIfEmpty = true - will automatically check if a datatable is empty and won't create a file. 
XlUtil.CheckFileAndFolderPermissions([RemoveOldFile:false]) - for cases when creating the same filename, 
and check if there's a permitions for file/folder.  Exceptions got: Illeagal character,   No writing permissions for directory,
Directory doesn't exits, Error deleting file Probably is opened.
XlUtil.AppendToFile - in case a file is already exists.

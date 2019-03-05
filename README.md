# Departmental Administrative System (DAS) software

This an Academic Project done by me for the partial fulfilment of the requirements for the award of National Diploma(ND) in Computer Science, Federal Polytechnic Nekede, Owerri, Imo state, Nigeria.

Dated, September 2012.

The Departmental Administrative System (DAS) software is an information system that is used for management, organization and storage of student’s information. It has the capability of storing students’ bio-data, contact details and also their results and grades. It can be used for all the levels in the department where it is applied.

Other features that are included are:

- Staff Details( Bio-data and Contact info)
- Courses studied in by the students at all level and the lecturers that take them.

The methodology adopted in carrying out this project is called Structured System Analysis and Design Methodology (SSADM) and the outputs are stored in a database.

The software is a Database Management System (DBMS) which uses Microsoft Access 2007 for the Database and Visual Basic 6.0 for designing the front-end interface due to its graphical user interface (GUI) features and its platform independency.

## ODBC Configuration for 3bits Computers

The steps to follow in configuring the ODBC data source are as follow:

1. Open the control panel on your system
2. Click on Administrative tools
3. Double click on Data sources (ODBC)
4. From the dialog box that appears click on Add..., another dialog box appears
5. From the second dialog box, double click on the drivers your software is using, in this case "Microsoft Access Driver (\*.mdb,)", a third dialog box appears
6. In the third dialog box, you enter the name for the Data Source in the first input box, type in the description for the Source and click on Select button to select the database.
7. Another dialog box appears. Here you select the folder that contains the database file in the Directories list box. The database files will probably be in the program folder in the C: Drive. Once you get to the folder, the database file appears in the second list box. Select it and click on OK button
8. Click on OK button on the dialog box that appears after the previous step
9. Click on OK button on the dialog box that appears after the previous step
10. You follow the same steps in creating the ODBC data source for the three database files used by the DAS software.
11. The Data Source Name you enter for the database files are:

    - StudentSource – projectStudents.mdb
    - StaffSource – projectStaff.mdb
    - LectureSource – projectLecturers.mdb.

12. You are good to go.

## ODBC Configuration for 64bits Computers

The steps to follow in configuring the ODBC data source are as follow:

1. On your Address bar type"C:\Windows\SysWOW64\odbcad32.exe" and press your ENTER key
2. From the dialog box that appears click on Add..., another dialog box appears
3. From the second dialog box, double click on the drivers your software is using, in this case "Microsoft Access Driver (\*.mdb,)", a third dialog box appears
4. In the third dialog box, you enter the name for the Data Source in the first input box, type in the description for the Source and click on Select button to select the database.
5. Another dialog box appears. Here you select the folder that contains the database file in the Directories list box. The database files will probably be in the program folder in the C: Drive. Once you get to the folder, the database file appears in the second list box. Select it and click on OK button
6. Click on OK button on the dialog box that appears after the previous step
7. Click on OK button on the dialog box that appears after the previous step
8. You follow the same steps in creating the ODBC data source for the three database files used by the DAS software.
9. The Data Source Name you enter for the database files are:

   - StudentSource – projectStudents.mdb
   - StaffSource – projectStaff.mdb
   - LectureSource – projectLecturers.mdb

10. You are good to go.

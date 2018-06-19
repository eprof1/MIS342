<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Project 3</title>
    <style type="text/css">
        .auto-style1 {
            text-decoration: underline;
        }
        .auto-style2 {
            width: 755px;
            height: 694px;
        }
    </style>
</head>
<body bgcolor="#FFFFFF" link="blue" vlink="purple" lang="EN-US" xml:lang="EN-US">
<div class="Section1">
  <p class="TutorialNumber"><font size="3"><a href="../default.asp">MIS342&nbsp; </a></font></p>
  <h1 align="center" class="TutorialNumber style1"><font color="#CC00FF"><i>
    <% response.write(semester) %> 
  MIS 342 Advanced Business Computer Applications </i></font></h1>
  <h2 class="TutorialNumber" align="center">  Project 3- Access, SQL Server and MySQL5 Database via ODBC </h2>
  <p><strong>Overview</strong></p>
  <p class="MsoNormal">You can work together and help each other on this project, but each student must hand in their own work.&nbsp;  <br />
    <br />
  </p>
  <hr />
  <p class="MsoNormal"><strong>Complete Before Class</strong></p>
    <p class="MsoNormal">You need to use Microsoft&#39;s Hyper-V in this project. To learn more about Hyper-V you must review the following materials:<br />
        Other courses are also available.</p>
  <ol>
    <li>Configure Hyper-V-<a href="https://www.lynda.com/Windows-tutorials/Configure-Hyper-V/516571/552399-4.html" target="_blank">Lynda.com video</a> (5 minutes)</li>
    <li>Virtualization Overview- <a href="https://www.lynda.com/Windows-Server-tutorials/Virtualization/503999/539131-4.html" target="_blank">Lynda.com video</a>&nbsp;(7 minutes)</li>
  </ol>
  <p class="MsoNormal">The three Hyper-V vhd files you need are in the folder &quot;MIS342Spring2018&quot; and can be downloaded from the WSU network at:<br />
     \\mispgp\VirtualMachineImages\MIS342Spring2018<br />
      <img src="FileLocation.PNG" />
  </p>
    <p>DO NOT DOWNLOAD THESE FILES TO YOUR LOCAL WEBSITE FOLDER!&nbsp; One is almost 9GB and will cause problems if you try to upload it to your website!!!!!!&nbsp; A suggested location is ..Documents/MIS342/Project3</p>
    <hr />
    <p>

        <span class="auto-style1"><strong>Class start</strong></span>:&nbsp; Open Hyper-V and create a new virtual machine called MIS342Spring2018, using the following settings:

    </p>
    <p>

        Generation 1</p>
    <p>

        1024 MB of Dynamic Memory</p>
    <p>

        Do not connect the network</p>
    <p>

        Use the existing virtual hard disk &quot;Disk1.vhd&quot;&nbsp; &nbsp; Do not create a virtual hard disk.</p>
    <hr />
    <p>
        Once the wizard is done, make the following changes to install the Page(E:) drive and the Data(F:) drive.
    </p>
    <p>
        Hardware Settings for Controller 0,&nbsp; add a Hard Drive in Location 1, that uses &quot;...disk3.vhd&quot;</p>
    <p>
        Hardware Settings for Controller 1, add a Hard Drive in Location 1, that uses &quot;...disk2.vhd&quot;</p>
    <p>
        When done, your Hardware settings will look like this:</p>
    <p>
        <img alt="Aux Hard Drives" class="auto-style2" src="AuxHardDrives.JPG" /></p>

    <hr />
  <p class="MsoNormal">The aim of this Project  is to create connections between  databases in Access 2007, SQL Server 2005 and MySQL5.1 using ODBC. <br />
    This project illustrates the use of Microsoft Access as a front end for MySQL or other databases, such as Oracle, IBM...</p>
  <p class="MsoNormal">You will use a virtual machine   to accomplish this task.</p>
  <p class="MsoNormal">This Project is in 2 parts, you must download and complete both, then submit them, just like Assignments:<br />
    <br />
    <a href="AssignmentP3_1.asp">AssignmentP3_1</a> <br />
    <br />
  <a href="AssignmentP3_2.asp">AssignmentP3_2</a></p>
  <hr />
  <p class="MsoNormal">&nbsp;</p>
<h3 class="MsoNormal"><a name="Part1" id="Part1"></a>Part 1- Access and MySQL</h3>
<p>Note: For online help with MySQL 5.x see:   <a href="http://dev.mysql.com/doc/refman/5.7/en/index.html" target="_blank">http://dev.mysql.com/doc/refman/5.7/en/index.html</a></p>
</div>
<ol>
<li>Locate the  folder 'P3DataAndScripts' on the Virtual Machine  desktop, double click to open it up.</li>
  <li>You will use the script 'MySQL-Henry' to create a database 'Henry' in MySQL. <br />
  Note:  You could run this sql script, with some modification, in an Access Module, to create a new database also. This is why you pay $$ for Access, so that you create databases and tables using a <acronym title="Graphical User Interface"><strong>GUI</strong></acronym>, not a <acronym title="Command Line Interface"><strong>CLI</strong></acronym>.</li>
  <li>There are other scripts and text files in this folder, some add other databases, some delete database tables, some have data from which you can create tables.<br />
    Feel free to explore and examine these other files once you have completed Project 3. <br />
  </li>
  <li>In the VM choose Start&gt;All Programs&gt;MySQL&gt;MySQL Server 5.0&gt;MySQL Command Line Client&nbsp; (or click the appropriate icon in the Taskbar)</li>
  <li>Open the MySQL5.txt file that is on the VirtualPC desktop to find the password (mysql) for the root account of MySQL5, enter this in the CLI<br />
    Note: for more info on the CLI see the MySQL web site: <a href="http://dev.mysql.com/doc">http://dev.mysql.com/doc</a>/<br />
    <br />
  </li>
  <li>Type the password (<strong>mysql</strong>) into the command line client. <br />
    <img src="MySQL/CLI.png" alt="CLI password" width="669" height="338" /> <br />
  </li>
  <li>Right click the file 'MySQL-Henry' and Open with &gt;Notepad.</li>
  <li>Press Ctrl +A to select all the text, then Ctrl+C to copy it all.</li>
  <li> Right click the Command Line Client and choose Edit&gt;Paste<br />
    The Henry database and all tables with data will be created. When done it will look like this:<br />
    <img src="MySQL/CreateHenry.png" alt="CLI create database" width="669" height="338" /><br />
    <br />
    In the MySQL Command Line Client type 'status' to make sure 'Henry' is the current database.<br />
    Then type 'show tables;'
      &nbsp;&nbsp;&nbsp;You should see a list of 6 tables: author/book/branch/inventory/publisher/wrote<br />
      Then type 'help' Peruse the various commands.<br />
      Then type 'quit'
      <br />
      <br />
  </li>
  <li>Choose Start&gt;Administrative Tools&gt; 'Data Sources (ODBC); then choose the 'System DSN' tab, <br />
    <img src="MySQL/ODBCAdmin.png" alt="ODBC admin" width="461" height="377" /><br />
    <br />
  </li>
  <li>Choose 'Add...' and select the MySQL ODBC 5.1 Driver, then press 'Finish' <br />
    <img src="MySQL/ODBCdriver.png" alt="Create Data Source" width="468" height="352" /><br />
  </li>
  <li>Enter the values shown below then press 'Test Connection' and you should see a dialog stating &quot;Success; connection was made!&quot; <br />
    <br />
    <img src="MySQL/Connector.png" width="425" height="410" alt="ODBC connector" />    <br />
  </li>
  <li>Press OK twice to close the ODBC Administrator. </li>
  <li>Start up Access 2007 inside the mis342 virtual machine.&nbsp; </li>
  <li> Open up a new Access blank database.&nbsp; Name it <strong>MySQLODBC.accdb</strong>&nbsp; You can save it in the My Documents folder, which is the default. </li>
  <li>On the Access menu choose  External Data&gt;Import&gt;More&gt;ODBC Database... </li>
  <li>In the dialog box that appears, under 'Files of type:' choose 'ODBC Databases()' and a 'Select Data Source' dialog box will appear. <br />
    Click the radio button &quot;Link to the data source by creating a linked table&quot;<br />
    Then press OK.<br />
    From the 'Machine Data Source' tab select &quot;HenryMySQL&quot; then press 'OK' <br />
    <img src="MySQL/DataSource.png" alt="Select Data Source" width="449" height="395" /><br />
    <br />
  </li>
  <li>Select All, make sure to check the 'Save password' box, then press 'OK'. <br />
    <img src="MySQL/LinkTables.png" alt="Link Tables" width="467" height="248" /><br />
    <br />
    If a dialog box appears stating your password will not be encrypted click the 'Save Password' button.<br />
    If the dialog box 'Select Unique Record Identifier' appears, choose the appropriate fields.<br />
    AuthorNum for table author<br />
    BookCode for table book<br />
    BranchNum for table branch<br />
    BookCode and BranchNum for table inventory<br />
    PublisherCode for table publisher<br />
    BookCode and AuthorNum for table wrote<br />
<br />
  </li>
  <li>This provides links to all these tables appear in the Navigation Pane.<br />
    <br />
  <img src="MySQL/DatabaseWindow.png" width="144" height="190" alt="navigation pane" />  </li>
  <li>Hover over a table and the connection string will appear.<br />
    <br />
    <img src="MySQL/ConnectionString.png" width="272" height="88" alt="connection string" />    <br />
    <br />
  </li>
  <li> Even though you are in Access, any changes that you make are reflected in the MySQL database. This is because you are Linked to the MySQL table.</li>
  <li>Open the Author table and add two new records, one using your name, one using my name.<br />
    When you get an error message like this one, figure out what is causing the problem and write the resolution in question 1<a name="Q1" id="Q1"></a> of the <a href="AssignmentP3_1.asp">Project 3_1 assignment</a><br />
    <img src="MySQL/InsertFailed.png" alt="Insert failed" width="300" height="115" /> <br />
    <br />
  </li>
  <li>Once you have figure out how to insert the two records, open the MySQL CLI again and type in the following:<br />
    use henry &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;then press Enter<br />
    select * from author;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;then press Enter <br />
  </li>
  <li>Make a screen shot of the results, making sure to show the last two records you just entered, and paste it into the <a href="AssignmentP3_1.asp">Project3_1 assignment</a>, <a name="Q2" id="Q2"></a>question 2. <br />
    The result will look something like this:<br />
    <img src="MySQL/SelectAuthor.png" alt="CLI Insert" width="396" height="333" /><br />
    <br />
  </li>
  <li> Type the following command at the MySQL CLI:<br />
    Insert INTO author (AUTHORNUM, AUTHORLAST, AUTHORFIRST) Values ('28', 'Simpson', 'Homer');<br />
  </li>
  <li>Type the SQL into the CLI that will show that you have added another author, make a screen shot with this new author and paste it into <a href="AssignmentP3_1.asp">Project3_1 assignment</a>, <a name="Q3" id="Q3"></a>question 3. <br />
  </li>
  <li> Verify that you can view this new author in Access. Note you may need to open and close the table before you see the new record.<br />
  </li>
  <li> Finish <a href="AssignmentP3_1.asp">Project 3_1 assignment</a> by answering  question 4. </li>
</ol>
<p>OPTIONAL/PRACTICE</p>
<p>If you want to try creating another MySQL database, go to the MySQL web site and download the world database, available at:<br />
    <a href="https://dev.mysql.com/doc/world-setup/en/">https://dev.mysql.com/doc/world-setup/en/</a></p>
<p>This is a text file. Save the file, or copy it and paste it into the CLI to create the 'world' database. </p>
<p><span class="style1"><strong>Deliverables:</strong></span><br />
  <span class="MsoNormal">When all done, fill out and submit <a href="AssignmentP3_1.asp">Project 3_1 assignment</a>!</span> <br />
</p>
<hr />
<hr />
<p>&nbsp;</p>
<div class="Section1">
  <h2 class="TutorialNumber" align="center"><a name="Part2" id="Part2"></a>Part 2- Access and  SQL Server 2005 Database</h2>
  <p>&nbsp;</p>
  <h3><strong>Overview</strong></h3>
  <p class="MsoNormal">You can work together and help each other on this project, but each student must hand in their own work.</p>
  <p class="MsoNormal">The aim of this part  is to create a SQL Server database. <br />
    You will upsize the Microsoft Access Northwind database so that it is a SQL server database.<br />
    This project illustrates the use of Microsoft Access as a front end for SQL server. </p>
  <p class="MsoNormal">You will again use the MIS342 Virtual Machine to accomplish this project. </p>
  <p class="MsoNormal">&nbsp; </p>
  <h3 class="MsoNormal">Technical Notes</h3>
</div>
<ol>
  <li> Open your MIS342 virtual machine </li>
  <li>Open Northwind.mdb in Access 2007. (Note that the upsizing Wizard
    does not work correctly in Access 2000)<br />
    <br />
  </li>
  <li>From the menu choose Database Tools&gt;SQL Server and the Upsizing Wizard dialog box appears.<br />
    <img src="Upsize00.jpg" alt="Start Upsizing" width="252" height="198" border="2" /><br />
    <br />
  </li>
  <li>    Accept
    the default value 'Create a new database' and click Next.<br />
    <br />
    <img src="Upsize01.png" width="486" height="361" /><br />
    <br />
  </li>
  <li>Use 'misvm\sqlexpress' for the SQL server, and 'NorthwindSQL' as
    the SQL Server database name. <br />
    Make sure 
'Use Trusted Connection' is checked.<br />
        <img src="Upsize02.png" width="486" height="361" /><br />
        <br />
    Then
    click Next.<br />
    <br />
    If this does not work, you may need to open Microsoft SQL Server Management Studio Express.<br />
    Expand Security and Server Roles.<br />
    Double click 'dbcreator'<br />
    Click the 'Add' button in the lower right, a dialog box opens<br />
    Click the Browse button
    <br />
    Add 'Built-in users' and 'misvm\sqlserver2005MSSQLUser$MISVM$SQLEXPRESS<br />
    <br />
    stop and restart the SQL server service<br />
<br />
    <br />
  </li>
  <li>Click the '&gt;&gt;' button to select all tables for export to SQL server.<br />
      <img src="Upsize03.png" width="486" height="361" /><br />
    Then
    click Next.<br />
        <br />
        <br />
        <br />
  </li>
  <li>Accept the default values, click Next.<br />
<img src="Upsize04.png" width="486" height="361" /><br />
<br />
      <br />
      <br />
  </li>
  <li>Enter an appropriate path for the ADP file, make sure to save the password and userid, then click Next.<br /> <img src="Upsize05.png" width="486" height="361" /><br />
    <br />
    <br />
    <br />
  </li>
  <li>Click Yes when this message appears: <br />
    <img src="Upsize06.png" width="627" height="121" /><br />
<br />
<br />
  </li>
  <li>Click Finish. <br />
    <img src="Upsize07.png" width="486" height="361" /><br />
<br />
<br />
  </li>
  <li>This dialog box will appear. <br />
      <img src="Upsize08.png" width="464" height="176" /><br />
      <br />
  </li>
  <li>Once this is done, you will see a report of approximately 21 pages. Examine and then close the report.<br />
      <br />
  </li>
  <li>    Choose Office Button&gt;Server&gt;Server Properties.<br />
A dialog box with the SQL
    Server properties will be displayed.<br />
    Do Alt+PrtScrn and save a screen shot
    to place in your report.<br />
<img src="Upsize09.png" width="230" height="206" alt="server properties" /></li>
</ol>
<p>&nbsp;</p>
<p>When done try creating new queries or views. </p>
<p><span class="MsoNormal"><strong><font color="#FF0000"><img src="View.png" width="252" height="276" /></font></strong></span></p>
<div class="Section1">
  <h3><strong>Grading</strong></h3>
  <p class="MsoNormal"><strong><font color="#FF0000">Deliverables: </font></strong> </p>
  <p class="MsoNormal">1. Answer <a href="AssignmentP3_2.asp">these questions for Project3_2</a> &nbsp;&nbsp;&nbsp;and submit them just like a homework assignment.&nbsp; Then publish and link them on your web site. </p>
  <p class="MsoNormal">2. Make sure to submit your answers when done. </p>
  <p class="MsoNormal"><strong><font color="#FF0000">Your working SQL database in Hyper-V must be available for me to review. </font></strong> </p>
  <p class="MsoNormal"><strong><font color="#FF0000">You can also create an Access to SQL Server ODBC connection.</font></strong><br />
  </p>
</div>
<p>&nbsp;</p>
<div class="Section1">
  <hr />
        <p>
        <strong>Troubleshooting Tips:</strong> To make sure MySQL is running, startup the Virtual Machine.<br />
        Click Start>Administrative Tools>Services<br />
        Find and double click the service 'MySQL'<br />
        Click 'Start' to start the program, which will take about 2 minutes.<br />
    </p>
  <hr />

  <br />
  <font size="3"><a href="../default.asp">MIS342</a></font> </div>
</body>
</html>

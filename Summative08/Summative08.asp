<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative08</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
.HWOutput {
	color: #00F;
}
hwquestion {
	color: #00F;
}
hwquestion {
	color: #00F;
}
#HWquestion {
	color: #00F;
}
formquestion {
	color: #00F;
}
#formQueston {
	color: #00F;
}
.Question {color: #00F;
	font-weight: bold;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><i><font color="#cc00FF">MIS 342 
      <% response.write(semester) %> 
   Assignment</font></i></h2>

  <table width="100%" border="1" cellspacing="1" cellpadding="1">
    <tr bgcolor="#00FFFF">
      <td width="36%">Email Address:
      
        <!-- Enter your winona.edu email address in the VALUE field below   GROK --> 
        <input type="text" name="email" id="email" size="30" maxlength="50" value=""/><br />
      a valid  email address is required</td>
       
      <td width="40%">First Name:
      
        <!-- Enter your first name in the VALUE field below   GROK --> 
      <input type="text" name="FirstName" size="30" maxlength="50" value="EnterYour FirstName"/></td>
      
      <td width="24%">Last Name:
      
        <!-- Enter your last name in the VALUE field below   GROK --> 
      <input type="text" name="LastName" size="25" maxlength="50" value="EnterYourLastName"/></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Semester: <input type="text" name="Semester" value=<% response.write(semester)%> />  </td>
      <td>Class:  <input type="text" name="Class" value=<% response.write(xClass)%> />    </td>
      <td><input name="InstID" type="hidden" id="InstID" value="00617282" />       StarID:
        
        <!-- Enter your StarID in the VALUE field below   GROK --> 
      <input type="text" name="PIN" value="EnterYourStarID" /></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Section:  <input type="text" name="Section" value="01"/>
      </td>
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative08"/>
      </td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr bgcolor="#FF0000">
      <td colspan="3"><div align="center">
          <input type="submit" name="Submit" value="Submit" />
        </div></td>
    </tr>
  </table>

 <p><br />
     <font color="#FF0000"><b>400 points</b></font></p>
 <h1 align="center">Microsoft Access <br />
     8-Sharing, Integrating and Analyzing Data</h1>
 <p><em>You  will need  to download the <strong>LuckyTire</strong> database, LuckyTire07.accdb, along with two data files, <br />
   Service.csv<br />
   Locatons.xlsx.<br />
They are availabe at:<br />
\\store\classes\20185001248\ReadOnl<strong>y\Part08Project\</strong></em></p>
 <hr />
 <p>Len Phan is the owner of  Lucky Tire, a chain of tire and auto service centers. Len would like to create  a couple forms. You&rsquo;ll do so by completing the following:<br />
  &nbsp; </p>
 <ol start="1" type="1">
   <li>Open the <strong>LuckyTire </strong>database. </li>
 </ol>
 <ol start="2" type="1">
   <li>Export the rptTireType report as an HTML       document; do not use a template, 
   do not save the export steps. <br />
   Make sure to select 'Unicode' encoding.<br />
   Use your Web       browser to open the <strong>rptTireType </strong>HTML       document. 
   <br />
   Then scroll to the bottom of the page; use the First, Previous,       Next, and Last links to navigate through the document.<br />
       (<strong>25</strong>) Print the document as a .pdf and save it where you save the screen shots for this assignment.<br />
        <table width="95%" border="1">
       <tr>
         <td>rptTireType-Ex2 rpt as .pdf file</td>
       </tr>
     </table>

   </li>
 </ol>
 <ol start="3" type="1">
   <li>Import the CSV file named <strong>Service</strong>.csv,        as a new table in the database. Index the Invoice field with no       duplicates, choose your own primary key, name the table <strong>tblService</strong>, run the Table       Analyzer, and record the Table Analyzer&rsquo;s recommendation, but do not       accept the recommendation. Do not save the import steps.<br />
     <br />
     (<strong>25</strong>) Provide a screen shot of the Table Analyer Dialog Box<br />
     <br />
     <table width="95%" border="1">
       <tr>
         <td>Table Analyzer Wizard Screen shot-Ex3ScreenShot</td>
       </tr>
     </table>
   </li>
 </ol>
 <ol start="4" type="1">
   <li>Export the tblCustomer table as an XML file       named <strong>tblCustomer</strong>;<strong> </strong>do not create a separate XSD       file. Save the export steps.<br />
     Make sure to save the file as tblCustomer.xml in your Assignment08 folder.<br />
       (<strong>25</strong>)  Print the file as a .pdf file:<br />
             <table width="95%" border="1">
       <tr>
         <td>xml file-Ex4 .pdf file</td>
       </tr>
     </table>
       
       <br />

   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="5" type="1">
   <li>Link to the Locations.xlsx workbook,  using <strong>tblLocations </strong>as the table name. <br />
     Add a new record to the Locations       workbook as follows:<br />
LocID <strong>103<br />
</strong> Address <strong>888 W 6th St<br />
</strong> City <strong>Corona<br />
</strong> Zip <strong>92882<br />
</strong> Phone <strong>951-555-8800<br />
</strong> Type your name in the Manager cell.<br />
     <br />
     <span id="formQueston"><strong>(50) Open the linked table in Design View (ignore the warning stating the design cannot be modified), open the table properties dialog box, and paste below the contents of the Description property. It will be a long string</strong></span><strong>.<br />
     </strong> <br />
     <label>
    <input name="q5" type="text" id="q5" size="120" maxlength="120" value="description:" />
     </label>
   </li>
 </ol>
 <ol start="6" type="1">
   <li>Compact and repair the LuckyTire database,       close the database, and then exit Access.<br /></li><br />
   <li><span class="Question">(75) Briefly explain  the value of Microsoft Access in an organization that uses many programs to store data in different formats, while requiring that data to be used in multiple formats on a daily basis.</span><br />
     <label>
    <input name="q7" type="text" id="q7" size="120" maxlength="120" value="benefit: " />
     </label>
   </li>
 </ol>
<p>&nbsp;</p>
<hr />
<div>
    <article class="h3">
        <div class="col-lg-12">&nbsp;<a href="#Exercise04" class="btn btn-info" data-toggle="collapse">Assignment as pdf</a></div>
      </article>
      
    <div id="Exercise04" class="collapse">
      <article class="row">
        <div class="col-lg-12"><p>For this Exercise, please read all of these instructions.<br>
        Make sure that you have completed all the previous exercises by 
filling in your answers, and publishing them to your website.<br>
        Then browse to your http://classes.winona.edu webiste, ensure all 
the Exercises are complete, then save your assignment as a .pdf file in the same
 location that you saved all of your screen shots for this assignment.<br>
        </p>
        </div>
        <div>          
            <div class="col-lg-12 w3-label w3-text-blue">
            (<strong>100</strong>) 8. Assignment .pdf file creation.<br />
                <br>
            </div>  
        </div>     
      </article>
     </div>
     
      
      <article class="h3">
        <div class="col-lg-12"><a href="#Exercise05" class="btn btn-info" data-toggle="collapse">Combine Assignment and Screen Shots</a></div>
      </article>
      
    <div id="Exercise05" class="collapse"> 
      <article class="row">

        <div class="col-lg-12"> 
        <p>To proceed you must have Windows 10 which includes &#39;Print to PDF&#39; installed. You know it is installed when one of the print dialog box choices is &quot;Microsoft Print to PDF&quot;<br />
            <img alt="print to pdf" class="auto-style1" src="F00_PrintToPDF.JPG" /><br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative08.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 9. Upload your file 'Summative08.pdf' to the D2L Assignment Folder 'Summative08'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div><hr />
<h1 align="center">Concepts</h1>
<p>Sample focus question: In what ways can Access data be used outside of an Access database?</p>
<p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
<ul>
  <li>Exporting data</li>
  <li>HTML</li>
  <li>CSV files</li>
  <li>Table Analyzer</li>
  <li>XML</li>
  <li>PivotTable</li>
  <li>PivotChart</li>
  <li>Forms with tabs</li>
  <li>Forms with charts</li>
  <li>Linked data</li>
</ul>
<hr />
 <hr />
<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000"><td><div align="center">
		<input type="submit" name="Submit" value="Submit" />
</div></td></tr> 
</table>
</form>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative09</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
.HWOutput {
	color: #00F;
}
HWQuestion {
	color: #00F;
}
HWQuestion {
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
.formquestion {
	font-weight: bold;
	color: #00F;
}
.Question {color: #00F;
	font-weight: bold;
}
.auto-style1 {
	border-style: solid;
	border-width: 1px;
}
.auto-style2 {
	text-align: center;
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative09"/>
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
     <font color="#FF0000"><b>625 points</b></font></p>
 <h1 align="center">Microsoft Access<br />
     9-Action Queries and Advanced Table Relationships</h1>
 <p><em>You  will need  to download the <strong>LuckyTire_09</strong> database from: <br />
 \\store\classes\20185001248\ReadOnl<strong>y\</strong></em></p>
 <hr />
 <p>Len Phan is the owner of  Lucky Tire, a chain of tire and auto service centers. Len has an Access  database named LuckyTire to store information about customers, tires,  manufacturers, invoices, and credit cards. The tblCustomer table contains data  about Lucky Tire&rsquo;s customers; tblTires contains data about the tires stocked by  Lucky Tires; the tblSales contains data about tire sales; the tblEmployee  contains data about Lucky Tire&rsquo;s employees; the tblAssignment contains data  about which invoices which employees worked on; and the tblCreditCard contains  data about the credit cards used to pay invoices. Len asks you to define table  relationships and create some new queries for him. You&rsquo;ll do so by completing  the following:<br />
  &nbsp; </p>
 <ol start="1" type="1">
   <li>Open the <strong>LuckyTire </strong>database. </li>
 </ol>
 <ol start="2" type="1">
   <li>Designate the folder where you are saving your       database as a trusted folder.</li>
 </ol>
 <ol start="3" type="1">
   <li>Change the first record in the <strong>tblCustomer </strong>table datasheet so the       First Name and Last Name columns contain your first and last names. Close       the table.</li>
 </ol>
 <ol start="4" type="1">
   <li>Define a many-to-many relationship between the       tblSales and tbEmployee tables, using the tblAssignment table as the       related table. 
   <br />
   <br />
   (<strong>50</strong>) a.) Provide a screen shot of the Relationships window 
   showing the three tables. <br />
   <br />
     <table width="95%" border="2" class="auto-style1">
       <tr>
         <td valign="top"><p>screen shot of relationships with appropriate 
		 linkages:</p>
          <p>Ex4ScreenShot</p>
          <p>&nbsp;</p></td>
         <td class="auto-style2">It will look something like this, but yours 
		 must have the appropriate linkages between the tables:<br />
		 <img alt="Reltationships" src="ManyToManyRelationship.jpg" /></td>
       </tr>
     </table>
   <br />
   <br />
   b.) Define a one-to-one relationship between the primary       tblCustomer table and the related tblCreditCard table. 
   <br />
   Select the       referential integrity option and the cascade updates option for the       relationships.<br />
     <span class="formquestion">(50) Briefly explain the purpose of a one-to-one 
   relationship.</span><br />
     <label>
    <input name="q4" type="text" id="q4" size="120" maxlength="120" value="1 to 1:" />
     </label>
    .<br />
     <br />
   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="5" type="1">
   <li>Create a make-table query based on the       qryInvoiceTotal query, selecting all fields from the query and only those       records where the TireType field value is <strong>Performance</strong> or <strong>Truck</strong>.       Use <strong>tblSpecialTire </strong>as the new       table name, and store the table in the current database. Run the query,       and then save the query as qryMake_tblSpecialTire<br />
     <br />
     <span class="formquestion">(50) Copy the SQL statement for the make table query below</span><br />
     <label>
    <input name="q5" type="text" id="q5" size="120" maxlength="120" value="sql:" />
     </label>
    .</li>
 </ol>
 <ol start="6" type="1">
   <li>Create an append query based on the       qryInvoiceTotal query, selecting all fields and selecting only those       records where the ModelName field value is <strong>Harmony</strong>, <strong>Kestral</strong>,       or <strong>Peregrine</strong>. <br />
    Append the       records to the tblSpecialTire table, run the query, and then save the       query as qryAppend_tblSpecialTire.<br />
    <br />
    <span class="formquestion">(50) Copy the SQL statement for the append query below</span><br />
    <label>
    <input name="q6" type="text" id="q6" size="120" maxlength="120" value="sql:" />
    </label>
   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="7" type="1">
   <li>Create an update query to select all records in       the tblSpecialTire table where the TireType field value is <strong>Performance</strong>, changing the field       value to <strong>PerfPlus</strong>. <br />
     Run the       query, and then save the query as qryUpdate_tblSpecialTire.<br />
     <br />
     <span class="formquestion">(50) Copy the SQL statement for the update query below</span><br />
     <label>
    <input name="q7" type="text" id="q7" size="120" maxlength="120" value="sql:" />
     </label>
   </li>
 </ol>
 <ol start="8" type="1">
   <li>Create a delete query that deletes all records       in the tblSpecialTire table in which the UnitPrice field value is less       than <strong>30</strong>. <br />
     Run the query, and       then save the query as qryDelete_UnitPrice. <br />
     Open the <strong>tblSpecialTire </strong>table, resize all columns to their best fit,       and then save and close the table.<br />
     <br />
     <br />
     <span class="formquestion">(50) Copy the SQL statement for the delete query below</span><br />
     <label>
    <input name="q8" type="text" id="q8" size="120" maxlength="120" value="sql:" />
     </label>
   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="9" type="1">
   <li>Create an outer join between the tblCustomer       and tblCreditCard tables, selecting all records from the tblCreditCard       table and any matching records from the tblCustomer table. Display the       First and Last fields from the tblCustomer table, and the FirstName and       LastName fields from the tblCreditCard table. Change the column names for       the FirstName and LastName fields from the tblCreditCard table to <strong>Credit First Name </strong>and <strong>Credit Last Name</strong>. <br />
     Save the query       as <strong>qryCustomerCreditCardOuterJoin<br />
     R</strong>un the query, resize all columns to their best fit, and then save and       close the query.<br />
     <br />
     <br />
     <span class="formquestion">(50) Copy the SQL statement for qryCustomerCreditCardOuterJoin below</span><br />
     <label>
    <input name="q9" type="text" id="q9" size="120" maxlength="120" value="sql:" />
     </label>
   </li>
 </ol>
 <ol start="10" type="1">
   <li>Open the <strong>tblSpecialTire </strong>table in Design view, add an index for the CustID field, allowing       duplicates, and adding an index for the TireType field, allowing       duplicates, and then save and close the table.<br />
     <br />
   (<strong>25</strong>) Provide a screen shot of the Indexes dialog box, showing the appropriate settings.<br />
   <table width="95%" border="1">
     <tr>
       <td><p>Indexes dialog box-Ex10ScreenShot</p>
         <p>&nbsp;</p>
         <p>&nbsp;</p></td>
     </tr>
   </table>
   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="11" type="1">
   <li>Compact and repair the <strong>LuckyTire</strong> database, close the database, and then exit Access.</li><br /><br />
   <li><span class="Question">(50) Briefly explain  the danger in allowing an employee to use Microsoft Access Action Queries in an organization that uses Linked Tables 
   and other linked data sources.</span><br />
     <label>
    <input name="q12" type="text" id="q12" size="120" maxlength="120" value="danger:" />
     </label>
   </li>
  </ol>
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
            (<strong>100</strong>) 13. Assignment .pdf file creation.<br />
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
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative09.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 14. Upload your file 'Summative09.pdf' to the D2L Assignment Folder 'Summative09'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div>
 <hr />
 <h1 align="center">Concepts</h1>
 <p> Sample focus question: How can Action Queries and relationships be used to enhance the functionality of a relational database?</p>
 <p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Action queries</li>
   <li>Make-table</li>
   <li>Append</li>
   <li>Delete</li>
   <li>Update</li>
   <li>Relationships</li>
   <li>Many-to-many</li>
   <li>One-to-one</li>
   <li>Table joins</li>
   <li>Inner join</li>
   <li>Outer join</li>
   <li>Indexes</li>
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

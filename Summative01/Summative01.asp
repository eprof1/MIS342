<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Assignment01</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="../../../../global.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .auto-style1 {
            color: #0000FF;
        }
        .auto-style2 {
            height: 84px;
            width: 973px;
        }
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><i><font color="#cc00FF">MIS 342 
      <% response.write(semester) %> 
   Summative01-Basic Database Objects</font></i></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative01"/>
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
     <font color="#FF0000"><b>500 points</b></font></p>
 <h1 align="center">Microsoft  Access<br />
     </h1>
 <h1>National Parks  Database</h1>
 <p>You have been  hired by a local travel agency to set up tours for the coming season to the  National Parks in the United    States. The travel agency wants to be able  to access a database that will let them know about several features of each  park so they can best meet the needs of prospective travelers. As clients book  and then take the trips, the database will also include client information that  can be linked to the park information. Your job is to search the Internet for  information about the National Parks and set up the National Parks database.</p>
 <ol start="1" type="1">
   <li>Start Access and open a new blank       database. Name the database <strong>National       Parks</strong>.</li>
   <li>Use your favorite search engine to       research information on the National Parks in the United States. </li>
   <li>Use at least three different Web       sites, including <a href="http://www.nps.gov/">www.nps.gov</a>, to find       basic information on at least five parks. Information can include the name       of the park, the state in which the park is located, the size of the park,       adult admission fee, special features, and the date that the park opens       for the season.</li>
   <li>Create a table named <strong>Park</strong>. Accept the default ID       primary key field with the AutoNumber data type, renaming it as Park ID.       Enter  five  fields- Name, State, Fee, Hours, and       Date Opened, in the table based on the information you found. Enter other fields that you deem appropriate. Be sure to       define the data type appropriate for each field and to use appropriate       field naming techniques.</li>
   <li>Enter at least five records in the       Park table for the parks that you found during the Internet search.<br />
   (<strong>50</strong>) In the space below, provide a screen shot (use the Snipping Tool) of the Park table, datasheet view.<br />
       <table width="95%" height="99" border="1" cellpadding="3" cellspacing="3">
         <tr>
           <td>Ex5ScreenShot</td>
         </tr>
       </table>
      <br />
   </li>
   <li>Use the Simple Query Wizard to       create a query that displays only the name of the park, the state, and the       hours from the Park table in the detailed result. Give the query a       descriptive name.<br />
     (<strong>50</strong>) In the space below, provide a screen shot of the query, datasheet view.<br />
       <table width="95%" height="99" border="1" cellpadding="3" cellspacing="3">
         <tr>
           <td>Ex6ScreenShot</td>
         </tr>
       </table>
      <br />
   </li>
   <li>Use the Form tool to create a simple       form based on the Park table that you can use to enter additional records       into the database. Enter at least five more records.  <br />
   (<strong>50</strong>) In the space below, provide a screen shot of the form, form view, any record.<br />
     <table width="95%" height="99" border="1" cellpadding="3" cellspacing="3">
       <tr>
         <td>Ex7ScreenShot</td>
       </tr>
     </table>
   </li>
   <li>Use the Report tool to create a       simple report showing every field in the Park table.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the report, report view.<br />
     <table width="95%" height="99" border="1" cellpadding="3" cellspacing="3">
       <tr>
         <td>Ex8ScreenShot</td>
       </tr>
     </table>
   </li>
   <li>Preview the report.</li>
   <li>Close the database and then exit       Access.    </li>
   <li><span class="FormQuestion">(<strong>100</strong>) <span class="auto-style1"><strong>What are the primary purposes of the Access data objects: table, query, form and report? <br />
     Enter your brief explanations below.</strong></span></span><br />
     <label for="q1"></label>

       <input id="q11" type="text" class="auto-style2" maxlength="120" name="q11" size="120" 
           value="Table:
                  Query:
                   Form:
                  Report:  
           " align="top" style="background-color: #00FFFF" />
       
       <br />
     <br />
     <br />
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
            (<strong>100</strong>) 12. Assignment .pdf file creation.<br />
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
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative01.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 13. Upload your file 'Summative01.pdf' to the D2L Assignment Folder 'Summative01'</p>

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
 <p> Sample focus question: What are the elements of a relational database?</p>
 <p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Fields</li>
   <li>Records</li>
   <li>Primary Key</li>
   <li>Table</li>
   <li>Query</li>
   <li>Form</li>
   <li>Report</li>
   <li>Database Management System</li>
   <li>Compacting</li>
   <li>Backup</li>
 </ul>
<hr />
	<p>
<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000"><td><div align="center">
		<input type="submit" name="Submit" value="Submit" />
</div></td></tr> 
</table>
 </p>
 <hr />
 <hr />
</form>
</body>
</html>

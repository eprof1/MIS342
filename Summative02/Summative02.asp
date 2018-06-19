<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative02</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="../../../../global.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .auto-style1 {
            color: #0000FF;
        }
        .auto-style2 {
            height: 51px;
            width: 970px;
        }
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><font color="#cc00FF"><i>MIS 342 
      Summative02-Tables</i></font></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative02"/>
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
 <h1 align="center">Microsoft Access<br />
    </h1>
 <h1>Hollywood Movie Database</h1>
 <p>You are taking  a course on Oscar-winning movies. The scope of the course covers ten movies  from 1990 through the present that won Oscars for Best Director. In order to  prepare for the class you decide to organize the movies in a simple database  that you will create. The Internet has all the information you need for the  database; you just need to research the movies at several sites, organize, and  then enter the data. The database will include two tables. The Movie table will  include data such as the movie title, release date, number of Oscar  nominations, number of Oscar wins, and category (such as whether it is a  comedy, drama, horror, documentary, or other category you decide). The Director  table will include the data about the director of those movies.</p>
 <ol start="1" type="1">
   <li>Start Access and open a new blank       database. Name the database <strong>Oscars</strong>.</li>
   <li>Use your favorite search engine to       research information to find ten movies that won Oscars in the Best Director       category since 1990.</li>
   <li>Use Web sites, including <a href="http://www.imdb.com/">www.imdb.com</a> and <a href="http://www.oscars.org/">www.oscars.org</a> to find the information       you need for the database.</li>
   <li>Create a table called <strong>Movie</strong>. Accept the default ID       primary key field with the AutoNumber data type, renaming it as Movie ID.       Create the fields, such as title, release date, number of Oscar nominations,       number of Oscar wins, category, and director ID Be sure to define the data       type appropriate for each field and to use appropriate field naming       techniques. Enter a custom format for the release date field.</li>
   <li>Create a table called <strong>Director</strong> that has the fields director       ID, first name, last name, and any other fields of your choice based on       the research that you collected. Consider that you do not want redundant       data between the two tables. The director ID is the common field between       the two tables.</li>
   <li>Review the data types and field       properties for the fields in the two tables and change any that are not       appropriate.</li>
   <li>Enter ten records into the Movie       table.</li>
   <li>View the records in the Movie table       in datasheet view. Sort the datasheet in descending order by number of       nominations, and then save the table.<br />
     <br />
   (<strong>50</strong>) In the space below, provide a screen shot of the Movie table, <strong>datasheet view</strong>.<br />
     <table width="90%" height="84" border="2" cellpadding="2" cellspacing="2">
       <tr>
         <th scope="col">Ex8ScreenShot</th>
       </tr>
     </table>
   </li>
   <li>Enter ten records in the Director       table.</li>
   <li>View the records in the Director       table in datasheet view. In datasheet view, move the first name field to       follow the last name field, and then save the table.<br />
     <br />
   (<strong>50</strong>) In the space below, provide a screen shot of the Director table, <strong>datasheet view</strong>.<br />
     <table width="90%" height="84" border="2" cellpadding="2" cellspacing="2">
       <tr>
         <th scope="col">Ex10ScreenShot</th>
       </tr>
     </table>
   </li>
   <li>Create an Excel worksheet, adding data       for three additional directors to the worksheet. The columns in the       worksheet should be the same as the columns in the Director table and in       the same order. Save and close the Excel workbook, and then import the Excel       director data into the Director table.<br />
     <br />
   (<strong>50</strong>) In the space below, provide a screen shot of the <strong>dialog box</strong> showing the import of the Excel worksheet data into Access. <br />
     <table width="90%" height="84" border="2" cellpadding="2" cellspacing="2">
       <tr>
         <th scope="col">Ex11ScreenShot</th>
       </tr>
     </table>
   </li>
   <li>Use the Relationships window to join the two tables based on the director ID field. Enforce referential       integrity.</li>
   <li>(<strong>50</strong>) In the space below, provide a screen shot of the <strong>Relationships window</strong>. <br />
     <table width="90%" height="84" border="2" cellpadding="2" cellspacing="2">
       <tr>
         <th scope="col">Ex13ScreenShot</th>
       </tr>
     </table>
   </li>
   <li>Close the database and then exit       Access.</li>
   <li><span class="FormQuestion">(<strong>100</strong>) <span class="auto-style1"><strong>Briefly explain why it is better to spend the time creating separate Access tables (such as Movie and Director) instead of simply using one Excel spreadsheet to enter all of the Movie and Director information (aka 'flatfile').</strong></span></span><br />
     <label for="q1"></label>

    
       
       <input id="q15" type="text" class="auto-style2" maxlength="120" name="q15" size="120" value="Separate Access tables versus flatfile:  " align="top" style="background-color: #00FFFF" /><br />

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
            (<strong>100</strong>) 16. Assignment .pdf file creation.<br />
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
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative02.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 17. Upload your file 'Summative02.pdf' to the D2L Assignment Folder 'Summative02'</p>

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
 <p>Sample focus question: What are the key components and uses of  Tables?</p>
 <p>Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Fields</li>
   <li>Data Types</li>
   <li>Primary Key</li>
   <li>Table Structure</li>
   <li>Relationships</li>
   <li>Records</li>
   <li>Importing data</li>
   <li>File types</li>
   <li>Relationships</li>
   <li>One to many</li>
   <li>Referential integrity</li>
 </ul>
    <hr />

    <hr />

<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000"><td><div align="center">
		<input type="submit" name="Submit" value="Submit" />
</div></td></tr> 
</table>
<hr />


 <hr />
	
</form>
</body>
</html>

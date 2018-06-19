<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative03</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="../../../../global.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .auto-style1 {
            color: #0000FF;
        }
        .auto-style2 {
            width: 967px;
            height: 51px;
        }
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><font color="#cc00FF"><i>MIS 342 
      Summative03-Relationships and Queries</i></font></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative03"/>
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
     <font color="#FF0000"><b>650 points</b></font></p>
 <h1 align="center">Microsoft Access<br />
    </h1>
 <h1>School Club  Database</h1>
 <p>You are the  Director of Students at Wilfred Academy, a  school in New England. Your have just hired a new assistant to help  with the clubs at the school. You want to be able to have instant access to the  clubs and the faculty advisor for each club. You set up the <strong>Clubs.accdb</strong> database with two tables,  Club and Advisor. There are currently 12 clubs in the table and you need to  enter the faculty data. You research the Internet for information on the clubs  and work with this database to complete the Club and Advisor tables.</p>
 <ol start="1" type="1">
   <li>Start Access and open the <strong><a href="Clubs_Tutorial3.accdb">Clubs.accdb</a></strong> database. Save the       database as <strong>Clubs-TA3.accdb</strong>. </li>
   <li>As the director, you want the       database to include one Web site that is a primary source of information       for each club. You also want the database to include the name of one       magazine that the club should subscribe to.</li>
   <li>Use your favorite search engine to       research Web sites and magazines for each of the 12 clubs.</li>
   <li>Add a new hyperlink field in the       Club table and enter one Web site URL for each club.</li>
   <li>Add a new text field in the Club       table for the magazine name. Enter a magazine for each club. View the       table datasheet, and then resize all columns to their best fit.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the Club   table, <strong>datasheet view</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex5ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>Enter yourself as a new record in       the Advisor table, and resize all columns to their best fit. Enter at       least six more records in the table.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the Advisor   table, <strong>datasheet view</strong>.<br />
       <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
         <tbody>
           <tr>
             <th scope="col">Ex6ScreenShot</th>
           </tr>
         </tbody>
       </table>
      <br />
   </li>
   <li>Edit the advisor IDs in the Club       table records, leaving one advisor ID null and having one or more advisors       advise more than one club.</li>
   <li>Use the Relationships window to join       the two tables based on the advisor ID field. Enforce referential       integrity.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the <strong>Relationships window</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex8ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>Create a query that displays clubs       and their advisors. Include the club name, club description, first Name,       and last Name fields, and sort in ascending order by club name. Name the       query <strong>ClubsAndAdvisors</strong>.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the ClubAndAdvisors query, <strong>datasheet view</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex9ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>Create a query that displays clubs       by club type. Name the query <strong>ClubsByType</strong>.       Sort the resulting datasheet in ascending order by club type.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the ClubByType query, <strong>datasheet view</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex10ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>&nbsp;View the Advisor table to view related       records and find out which clubs display for each advisor by expanding the       plus sign.</li>
   <li>Use the Query window to create a query       displaying all advisor and club data, sorting in ascending order by       advisor ID and then in ascending order by club name. Hide the start date       field from the query. Name the query <strong>AdvisorsAndClubs</strong>.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the AdvisorsAndClubs query, <strong>datasheet view</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex12ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>View the AdvisorsAndClubs query       datasheet. Change the query to display those records with Language or       Science club types. Sort the datasheet in ascending order by club name.<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the AdvisorsAndClubs query, <strong>design</strong> <strong>view</strong>.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex13ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>Close the database and then exit       Access.</li>
   <li><span class="FormQuestion">(<strong>100</strong>) <span class="auto-style1"><strong>Briefly explain the value to a business of an employee who can create queries from verbal or written input provided by other employees or customers.</strong></span></span><br />
     <label for="q1"></label>


       <input id="q15" type="text" class="auto-style2" maxlength="120" name="q15" size="120" 

           value="Value of employee who can create queries:  
           
           " align="top" style="background-color: #00FFFF" />
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
            <img alt="print to pdf" class="auto-style1" src="F00_PrintToPDF.JPG" />
            
            <br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative03.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 17. Upload your file 'Summative03.pdf' to the D2L Assignment Folder 'Summative03'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div>
 <hr />
 <p><br />
  </p>
<h1 align="center">Concepts</h1>
 <p>Sample focus question: What are the key components and uses of queries?</p>
 <p>Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Updating data</li>
   <li>Query design</li>
   <li>Multi-table queries</li>
   <li>Sorting data</li>
   <li>Filtering data</li>
   <li>Query criteria</li>
   <li>Query format</li>
   <li>Multiple query criteria</li>
   <li>Calculated fields</li>
   <li>Aggregate functions</li>
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

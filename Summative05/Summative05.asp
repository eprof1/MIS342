<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative05</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
.Question {	color: #00F;
	font-weight: bold;
}
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><font color="#cc00FF"><i>MIS 342 
      Summative05-Queries, Advanced Tables</i></font></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative05"/>
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
  <font color="#FF0000"><b>300 points</b></font></p>
 <h1 align="center">Microsoft Access<br />
    </h1>
 <p><strong>Adventure Outdoors<br />
 </strong>Adventure  Outdoors, a youth-oriented travel company, has hired you to develop a database  that the staff can use to quickly access information about National Parks. Adventure  Outdoors uses grant money to pay the way for teens who cannot afford to take  trips. They intend to use the database to help plan teen travel during the  summer months. The database will also include the names of the United States  Senators from each of the states. Adventure Outdoors contacts senators to ask  for grant money for the programs. Three basic tables are in the rough draft of the  database to get you started. Open the database and see which fields and tables  have been created. Based on research on the Internet, you will add fields that  you think will be useful, relate the tables, and enter data for several states.  You will share your work with the staff at Adventure Outdoors and then, over  time, you will complete the database. To begin, you will research and enter  data for five states.</p>
 <ol>
   <li>Start Access and then open the <strong><a href="Adventure Outdoors_Tutorial5.accdb">Adventure Outdoors.accdb</a></strong> database provided by your instructor.</li>
   <li>Open the <strong>tblPark</strong>, <strong>tblSenator</strong>, and <strong>tblState</strong> tables. Become familiar with the fields in these tables. (<em>Note</em>: The State, Town, and Phone fields  in the tblPark table pertain to the park&rsquo;s headquarters location.)</li>
   <li>Consider and add any additional fields. Include a  Long Text field in one of the tables, and include a data validation rule for at  least one field in one of the tables. Create an input mask for the Phone field  in the tblPark and tblSenator tables<br />
   (<strong>50</strong>) In the space below, provide a screen shot of the Phone field input mask, from the table properties in design view.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex3ScreenShot</th>
         </tr>
       </tbody>
     </table>
   </li>
   <li>Open your Internet browser and search for the  information you will need to fill in the fields in the tblPar<strong>k </strong>table; this information is readily  available over the Internet. You can begin your search at <a href="http://www.nps.gov/">www.nps.gov</a>. Find information about at least ten  parks. Then use the Internet to find information for the tblSenator and tblState  tables. You should obtain information for MN and for the states  that contain the headquarters for the National Parks you selected.</li>
   <li>Enter the data you found into the tblState table,  the tblSenator table, and then the tblPark table. Resize the datasheet columns  to their best fit in each table.<br />
   (<strong>50</strong>) In the space below, provide  screen shots of the   tables tblState, tblSenator and tblPark, datasheet view.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th align="left" valign="top" scope="col">tblState-Ex5_1ScreenShot</th>
         </tr>
       </tbody>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" scope="col"><p>tblSenator-Ex5_2ScreenShot</p>
          <p>&nbsp;</p></th>
       </tr>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" valign="top" scope="col"><p>tblPark-Ex5_3ScreenShot</p>
          <p>&nbsp;</p></th>
       </tr>
     </table>
   </li>
   <li>Create a parameter query that will display the  senators&rsquo; information for a state based on the state value entered by the user.  The query should show the parks in the senators&rsquo; state. Name the query <strong>qrySenatorsParameter.</strong></li>
   <li>Run the qrySenatorsParameter query with a state parameter to make sure it works.&nbsp; Then switch to Design view.<br />
     (<strong>50</strong>) In the space below, provide a screen shot of the qrySenatorsParameter, <strong>design</strong> view.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th scope="col">Ex7ScreenShot</th>
         </tr>
       </tbody>
     </table>
    </li>
   <li>Create at least five other queries of the type indicated below, to extract  information that would be useful in planning trips and getting grants.<br />
     (<strong>100</strong>) In the space below, provide  screen shots of your five queries, <strong>design</strong> view.<br />
     <table height="84" cellspacing="2" cellpadding="2" width="90%" border="2">
       <tbody>
         <tr>
           <th align="left" valign="top" scope="col">select query1-Ex8_1ScreenShot</th>
         </tr>
       </tbody>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" scope="col"><p>select query2-Ex8_2ScreenShot</p>
             <p>&nbsp;</p></th>
       </tr>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" valign="top" scope="col"><p>select query3-Ex8_3ScreenShot</p>
             <p>&nbsp;</p></th>
       </tr>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" scope="col"><p>select query4-Ex8_4ScreenShot</p>
             <p>&nbsp;</p></th>
       </tr>
     </table>
     <br />
     <table width="90%" border="2" cellspacing="2" cellpadding="2">
       <tr>
         <th align="left" valign="top" scope="col"><p>select query5-Ex8_5ScreenShot</p>
             <p>&nbsp;</p></th>
       </tr>
     </table>
     <br />
   </li>
   <li>Create a trusted folder to store the Adventure  Outdoors database. Make a backup copy of the database, and then compact and repair  the database.</li>
   <li>Close the database and exit Access.    </li>
   <li><span class="Question">(50) What is the benefit to organizations of employees who have learned to create database queries?</span><br />
        <input id="q11"  name="q11" type="text" value="Query writing value:" size="120" style="background-color: #00FFFF"  />





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
            <img alt="print to pdf"  src="F00_PrintToPDF.JPG" /><br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative05.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 13. Upload your file 'Summative05.pdf' to the D2L Assignment Folder 'Summative05'</p>

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
 <p> Sample focus question: How can advanced queries and good table design enhance a database?</p>
 <p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Pattern match</li>
   <li>List-of-values</li>
   <li>Logical operators</li>
   <li>Filtering</li>
   <li>Parameter query</li>
   <li>Crosstab query</li>
   <li>Finding duplicate records</li>
   <li>Finding unmatched records</li>
   <li>Top values</li>
   <li>Lookup fields</li>
   <li>Input masks</li>
   <li>Data validation</li>
 </ul>
<hr />
 <hr />
<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000"><td><div align="center">
		<input type="submit" name="Submit" value="Submit" />
</div></td></tr> 
</form>
</body>
</html>

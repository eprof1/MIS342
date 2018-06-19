<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>MIS342Project3Assignment1</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <style type="text/css">
        .auto-style2 {
            width: 861px;
            height: 45px;
        }
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW">
  <h1 align="center"><i><font color="#CC00FF">MIS 342 <% response.write(semester) %> Project Assignment</font></i></h1>
 
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
      <td width="36%">Semester:
        <input type="text" name="Semester" value=<% response.write(semester)%> />
      </td>
      <td>Class:
        <input type="text" name="Class" value=<% response.write(xClass)%> />
      </td>
      <td><input name="InstID" type="hidden" id="InstID" value="00617282" />
      
        StarID:
        
        <!-- Enter your StarID in the VALUE field below   GROK --> 
      <input type="text" name="PIN" value="EnterYourStarID" /></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Section:  <input type="text" name="Section" value="01"/>
      </td>
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="P3_1"/>
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


  <p><strong><font color="#FF0000">500 points </font></strong></p>
  <p><a href="Project3.asp">Project 3 Background Info </a></p>
  <p><strong><font color="#0000FF">1.&nbsp; (100) List three difference between Access and MySQL 5.x</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="Project3.asp#q1">Refer to Project 3 instructions</a> <br />
    <br />
 

    <input name="q1" type="text" id="q1" class="auto-style2" maxlength="120"  size="120" style="background-color: #00FFFF" value="Three differences:" />

  </p>
  <p>&nbsp;  </p>
  <p>2.&nbsp; (50) Make screen shot of MySQL CLI showing two new record(s) are added: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="Project3.asp#q2">&nbsp;Refer to Project 3 instructions </a></p>
  <table width="95%"  border="1" cellspacing="2" cellpadding="2">
    <tr>
      <th scope="col"><p>Ex2_ScreenShot</p>
     </th>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>3.&nbsp;(50) Make screen shot of MySQL CLI showing Homer Simpson record was added. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="Project3.asp#q3">Refer to Project 3 instructions </a></p>
  <table width="95%"  border="1" cellspacing="2" cellpadding="2">
    <tr>
      <th scope="col"><p>Ex3_ScreenShot</p></th>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p><strong><font color="#0000FF">4. &nbsp; (100) What is the benefit of using Access as a front end for a MySQL database? <br />
  You may want to view the online MySQL5 documentation to get some ideas at this url: <a href="http://dev.mysql.com/doc/refman/5.0/en/index.html">http://dev.mysql.com/doc/refman/5.0/en/index.html</a><br />
  or feel free to browse the web for information.
    </font></strong><br />


          <input name="q4" type="text" id="q4" class="auto-style2" maxlength="120"  size="120" style="background-color: #00FFFF" value="Benefit of Access front end:" />

  </p>
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
            (<strong>100</strong>) 5. Assignment .pdf file creation.<br />
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
            <img alt="print to pdf" src="F00_PrintToPDF.JPG" /><br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as P3_A1.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 6. Upload your file 'P3_A1.pdf' file to the D2L Assignment Folder 'Project3'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div>
  <hr />
  <hr />
  
  <table width="100%" border="0" cellspacing="1" cellpadding="1">
    <tr bgcolor="#FF0000">
      <td><div align="center">
        <input type="submit" name="Submit" value="Submit" />
      </div></td>
    </tr>
  </table>
</form>
  
  </body>
</html>

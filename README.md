

# Welcome to Excel Reporting! 🤨

Instructions are always very boring, but I will try to keep it simply simple **😏**

## Purpose this project  🤯
While executing any large number of **Automation scripts** or to use run time generated data for the next script. Which I call it as **Post Batch** script, it's very difficult to track and organize the data. 

Example: Every company has it's customers. Now, by using the customer data you commit a transaction. If you are an **Uber** 🚗 Automation Tester, booking an Cab is your transaction. Validating whether the booking is available in my **booking history** 🗒️ is your task.

In an ideal scenario, to make the automation script run effortlessly we either provide booking id by opening your excel file or feature file, is not so **automated** 🖥️. Else, we will commit the transaction once again for this script. What if you search for booking history based on your past script which booked the cab for you.  

## Solution to above problem  🤔
This Jar, solves all your problems. Logs the data where ever its been created or generated and saves it in an excel file for further use. Please go through our demo below for further instructions or how to do it. I guess our flight ✈️ is now booked lets fly. 

# ###Demo### 😎

<p align="center">
<a href="https://www.linkedin.com/company/sspart"><img align="center" src="http://sspart.org/wp-content/uploads/2019/12/ExcelReporting_Thumbnail.png" width="580"></a>
</p>


## Instructions on how to use this Jar 👍 
 
 - Step 1 😝: Import this jar using following Maven dependency
 - Step 2 (Most boring step 😑): Configure yourself based on your requirement, these following lines
 
 ```
 	/*
 	* Add these two lines at the end of each test case or where you feel its correct.
 	* Either in cucumber hooks or TestNG Listener onSuccess and onFailure.
 	* First line, appends the data into excel and the second line clears the global data.
 	* 
 	* import below two lines ->
 	* import static com.ramSabhanam.excelReport.ExcelReport.reporter;
 	* import static com.ramSabhanam.excelReport.GlobalData.data;
 	*
 	*/
 	reporter().flushExcel(data());
 	data().clear();
	
 ```
 
 - Step 3 (An Intelligent step 😬): 
 
 ```
 	/*
 	* Import below line -> 
 	* import static com.ramSabhanam.configuration.Configuration.*;
 	*/
 	reporter().setDatabase("xlsx", getBundle().get("reports.location") + "/" + fileNameWithOutExtension);
 ```
 
 - Step 4 (Final step 😌):

## Sample Report Screenshot 🤓🤓

<p align="center">
<img src="http://sspart.org/wp-content/uploads/2019/12/ExcelReport.png" width="780">
</p>

## Connect with me 🤝

<p align="center">
<a href="https://www.facebook.com/SSPART.ORG/"><img src="http://sspart.org/wp-content/uploads/2019/11/Facebook_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="https://www.instagram.com/sspart_org/"><img src="http://sspart.org/wp-content/uploads/2019/11/Instagram_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="mailto:contact@sspart.org"><img src="http://sspart.org/wp-content/uploads/2019/11/Mail_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="https://www.youtube.com/channel/UCyNXuAWqDjMIoSXj5I1NqaA"><img src="http://sspart.org/wp-content/uploads/2019/11/YouTube_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="https://wa.me/919515093965"><img src="http://sspart.org/wp-content/uploads/2019/11/WhatsApp_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="https://github.com/sabhanam"><img src="http://sspart.org/wp-content/uploads/2019/11/GitHub_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="http://sspart.org/"><img src="http://sspart.org/wp-content/uploads/2019/11/WebSite_Circle.png" width="48"></a><span style="padding-left:10px;"/>
<a href="https://www.linkedin.com/in/ram-sabhanam/"><img src="http://sspart.org/wp-content/uploads/2019/11/LinkedIn_Circle-e1574599074500.png" width="48"></a>
</p>


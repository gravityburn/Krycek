# Krycek
Java Middleware Program

![krycek](https://user-images.githubusercontent.com/17135519/32349996-efac30b2-bfe6-11e7-8c3f-ea02238ef8f0.jpg)

Overview

•	For X-Files fans out there, this program is named after Alexander Krycek, the antagonist known for passing information between the shadowy government agencies and the FBI.

•	A lightweight, multithreaded Java middleware program for connecting two differentiated language API’s.  The program was initially created to bridge the lack of functionality between the web application Mavenlink to the CRM EFI Pace and was intent on delivering employee information concerning the employee number, job ID, employee hours, activity code, and finally the date any information was entered into on the Mavenlink side, which would then be imported into Pace as an .xls file.  The entire scope of the project was to automate this process in place of manual entry.

•	Utilizes a simple Process Builder method that is inclusive of the well-known cURL command-line tool; this program is capable of requesting and returning large amounts of JSON data.

•	An OAuth 2.0 authorization header is built in to the Process Builder, ultimately passing the authorization token to the resource server of your choice and returning the requested information back to the client.

•	The JSON data collected is then extracted, parsed, and built into an Excel Worksheet for importation into another MIS of your choice, or can be set up for data dumps.

•	Utilizes the JExcel API and the org.json package.  The JExcel API supports .xls extension formats only.  For .xlsx, you’ll need to utilize Apache’s POI API for Java found here: https://poi.apache.org/.


To learn more about the JXL API and for code samples, please see: http://jexcelapi.sourceforge.net/.  For more information about the JSON package in Java, see: https://github.com/stleary/JSON-java.  For more information about JSON itself, see: http://www.JSON.org/.

# Krycek
Java Middleware Program

Krycek Overview

•	For X-Files fans out there, this program is named after Alexander Krycek, the antagonist known for passing information between the shadowy government agencies and the FBI.

•	A lightweight, multithreaded Java middleware program for connecting two differentiated language API’s.

•	Utilizes a simple Process Builder method that is inclusive of the well-known cURL command-line tool; this program is capable of requesting and returning large amounts of JSON data.

•	An OAuth 2.0 authorization header is built in to the Process Builder, ultimately passing the authorization token to the resource server of your choice and returning the requested information back to the client.

•	The JSON data collected is then extracted, parsed, and built into an Excel Worksheet for importation into another MIS of your choice, or can be set up for data dumps.

•	Utilizes the JExcel API and the org.json package.  The JExcel API supports .xls extension formats only.  For .xlsx, you’ll need to utilize Apache’s POI API for Java found here: https://poi.apache.org/.


To learn more about the JXL API and for code samples, please see: http://jexcelapi.sourceforge.net/.  For more information about the JSON package in Java, see: https://github.com/stleary/JSON-java.  For more information about JSON itself, see: http://www.JSON.org/.

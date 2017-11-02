package krycek;

import java.awt.Desktop;
import java.io.FileNotFoundException;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.Locale;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.annotate.JsonAutoDetect;
import org.codehaus.jackson.annotate.JsonMethod;
import org.codehaus.jackson.map.ObjectMapper;
import org.codehaus.jackson.map.SerializationConfig.Feature;

import org.json.JSONException;
import org.json.JSONObject;
import org.json.JSONTokener;
import org.json.simple.parser.ParseException;

/**
 *
 * @author Brent Golke
 * 
 */
public class Krycek {

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     * @throws java.io.FileNotFoundException
     * @throws jxl.write.WriteException
     * @throws org.json.JSONException
     * @throws java.text.ParseException
     * @throws org.json.simple.parser.ParseException
     */
    public static void main(String[] args)throws IOException,
            FileNotFoundException, WriteException, JSONException, 
            java.text.ParseException, ParseException {
        
        // Methods for extracting Mavenlink data and output JSON
        // data to specified file location for Pace pickup
        ProcessBuilder pb = new ProcessBuilder
            ("curl", "-s", "-H", "Authorization: "
            + "Bearer "
            + "yourOAuth2.0TokenHere",
            "https:// + yourApiRequestHere");
        pb.directory(new File("yourFileDirLocationHere"));
        pb.redirectErrorStream(true);
        Process p = pb.start();
        InputStream is = p.getInputStream();
        ObjectMapper mapper = new ObjectMapper();
        mapper.enable(Feature.INDENT_OUTPUT);
        mapper.setVisibility(JsonMethod.SETTER, JsonAutoDetect.Visibility.ANY);
        JsonNode json = mapper.readTree(is);
        File output = new File("yourOutputFileLocationHere");
        mapper.writeValue(output, json);
        
        // JSON File Retrieval and Iteration Methods
        try {
            
            // For obtaining and parsing JSON values
            FileReader fr = new FileReader("readFromYourOutputFile");
            JSONTokener tokener = new JSONTokener(fr);
            JSONObject obj = new JSONObject(tokener);
            JSONObject stories = (JSONObject) obj.get("stories");
            Iterator iter = stories.keys();
            int i;
            int j;
            
            // Iterates each "story" to produce requested String value
            while(iter.hasNext()) {

                // Builds Excel file from JSON data
                String filename = "yourExcelFile";

                // Workbook settings
                WorkbookSettings wbSettings = new WorkbookSettings();
                wbSettings.setLocale(new Locale("en", "EN"));

                // Creates new Workbook
                WritableWorkbook workbook = Workbook.createWorkbook(new File(filename), wbSettings);

                // Creates new Sheet
                WritableSheet ws = (WritableSheet)workbook.createSheet("Sheet1", 0);

                // Sets Auto Cell Width
                for(int c = 0; c < 6; c++) {
                    
                    CellView cv = ws.getColumnView(c);
                    cv.setAutosize(true);
                    ws.setColumnView(c, cv);
                    
                    // Adds Static Cells to the Sheet with Column Names
                    Label empID = new Label(0, 0, "Employee");
                    Label job = new Label(1, 0, "Job ID");
                    Label empHours = new Label(2, 0, "Employee Hours");
                    Label actCode = new Label(3, 0, "Activity Code");
                    Label created = new Label(4, 0, "Date");

                    // Creates Static Cells in the Sheet for Column Names
                    ws.addCell(empID);
                    ws.addCell(job);
                    ws.addCell(empHours);
                    ws.addCell(actCode);
                    ws.addCell(created);
                }

                    // Iterates nested objects for their values
                    for(i = 0; i < stories.names().length(); i++) {
                        
                        String key = (String)iter.next();
                        JSONObject story = stories.getJSONObject(key);
                        String jobId = story.getString("title");
                        String userId = story.getString("creator_id");
                        String empId[] = userId.split(" ");
                        
                        for(j = 0; j < empId.length; j++) {
                            
                            if("yourUserId".equals(userId)) {
                                
                                System.out.println((userId));
                          
                        
                                double employeeHours = story.getDouble("logged_billable_time_in_minutes")/60;
                                String actCode = story.getString("description");

                                // This field in Mavenlink cannot be empty, throws parsing exception here
                                String createdAt = story.optString("due_date");

                                // Converts Mavenlink Date format to other Date format
                                SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
                                SimpleDateFormat outputFormat = new SimpleDateFormat("MM-dd-yyyy");
                                Date inputDate = inputFormat.parse(createdAt);
                                String datePerformed = outputFormat.format(inputDate);

                                // Adds values to the appropriate cell
                                Label label = new Label(0, 1+i, userId);
                                Label label1 = new Label(1, 1+i, jobId);
                                Number label2 = new Number(2, 1+i, employeeHours);
                                Label label3 = new Label(3, 1+i, actCode);
                                Label label4 = new Label(4, 1+i, datePerformed);
                                ws.addCell(label);
                                ws.addCell(label1);
                                ws.addCell(label2);
                                ws.addCell(label3);
                                ws.addCell(label4);
                                
                                // Prints all available task components within this pull request
                                System.out.println("Employee: " + userId);
                                System.out.println("Job ID: " + jobId);
                                System.out.println("Employee Hours: " + employeeHours);
                                System.out.println("Activity Code: " + actCode);
                                System.out.println("Date: " + datePerformed);
                                LocalDateTime ldt = LocalDateTime.now();
                                System.out.println("Date and Time of Pull Request: " + ldt);
                                System.out.println("\n");
                        
                            }
                        }
                    }
                    
                // Writes and Closes Workbook
                workbook.write();
                workbook.close();
                
                // Opens new instance of the file
                Desktop.getDesktop().open(new File(filename));
            }
        }
        catch(FileNotFoundException | RowsExceededException e){
        }
        catch(JSONException ex){
            Logger.getLogger(Krycek.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
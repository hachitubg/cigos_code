import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ConvertExcel2To3 {

    public static String base_url = "F:/Upload/new/";
    public static String nameExcelExport = "final/Format Reviewers.xlsx";

    public static String firstName = "First name";
    public static String lastName = "Last name";
    public static String email = "Nom d'utilisateur";

    public static String topicPaper = "TOPIC";
    public static String title = "TITLE";

    public static String countryReviewer = "Country";
    public static String countryPaper = "COUNTRY";

    public static String maxReview = "Max of papers to review";
    public static String paperId = "DOCID";

    public static String other = "Other";

    // Write excel
    public static String paperFinal = "TITLE";
    public static String emailReviewer = "EMAIL";
    public static String nameReviewer = "NAME";
    public static String reviewerMax = "MAX REVIEW";
    public static String topic = "TOPIC";
    public static String docid = "DOCID";
    public static String paperCountry = "PAPER COUNTRY";
    public static String reviewerCountry = "REVIEWER COUNTRY";
    public static String timesPaperReview = "TIMES PAPER REVIEW";
    public static String timesReview = "TIMES REVIEW";

    // Convert 2
    public static String Id = "Id";
    public static String Topics = "Topics";
    public static String Papers = "Papers";
    public static String Country = "Country";
    public static String Reviewer1 = "Reviewer 1";
    public static String Reviewer2 = "Reviewer 2";
    public static String Reviewer3 = "Reviewer 3";

    // To 3
    public static String Email = "Email";
    public static String MaxToReview = "Max to review";


    public static List<JSONObject> readXLSFile(String fileName) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(base_url + fileName);
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        Iterator rows = sheet.rowIterator();
        DataFormatter dataFormatter = new DataFormatter();

        JSONObject json;
        List<JSONObject> jsonObjects = new ArrayList<>();
        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            Iterator cells = row.cellIterator();

            if (row.getRowNum() == 0) {
                continue; //just skip the rows if row number is 0
            }

            int index = 0;
            json = new JSONObject();
            while (cells.hasNext()) {
                cell = (HSSFCell) cells.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                try {
                    String key = sheet.getRow(0).getCell(cell.getColumnIndex()).toString();
                    json.put(key, cellValue);
                } catch (Exception ex) {
                    System.out.println(ex);
                }
                index++;
            }
            jsonObjects.add(json);
        }

        System.out.println("Count rows " + fileName + " : " + sheet.getPhysicalNumberOfRows());
        return jsonObjects;
    }

    public static void writeExcel(List<JSONObject> map) throws IOException {

        String url = base_url + nameExcelExport;
        //name of excel file
        String excelFileName = url;

        //name of sheet
        String sheetName = "Sheet1";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        XSSFRow rowHeader = sheet.createRow(0);
        for (int c = 0; c <= 5; c++) {
            XSSFCell cell = rowHeader.createCell(c);
            if (c == 0) cell.setCellValue("No.");
            if (c == 1) cell.setCellValue("Email");
            if (c == 2) cell.setCellValue("Id");
            if (c == 3) cell.setCellValue("Paper");
            if (c == 4) cell.setCellValue("Topic");
            if (c == 5) cell.setCellValue("Max to review");
        }

        //iterating r number of rows
        int r = 1;
        for (JSONObject jsonObject: map) {
            XSSFRow row = sheet.createRow(r);
            for (int c = 0; c <= 5; c++) {
                XSSFCell cell = row.createCell(c);
                if (c == 0) cell.setCellValue(r);
                if (c == 1) cell.setCellValue(jsonObject.getString(Email));
                if (c == 2) cell.setCellValue(jsonObject.getString(Id));
                if (c == 3) cell.setCellValue(jsonObject.getString(Papers));
                if (c == 4) cell.setCellValue(jsonObject.getString(Topics));
                if (c == 5) cell.setCellValue(jsonObject.getString(MaxToReview));
            }
            r++;
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        System.out.println("------------------------");
        System.out.println("Export excel success !!!");
        System.out.println("Export file to url: " + url);
    }

    public static void saveToObject(JSONObject reviewer,
                                    JSONObject paper,
                                    JSONObject finalObject,
                                    List<JSONObject> finalObjects,
                                    Map<JSONObject, Integer> mapPaper,
                                    Map<JSONObject, Integer> mapReviewer,
                                    int max,
                                    int maxReviewPaper) {

        if (reviewer.has(firstName)) {
            String name = "";
            if (!reviewer.has(lastName)) name = reviewer.getString(firstName);
            else name = reviewer.getString(firstName) + " " + reviewer.getString(lastName);
            finalObject.put(nameReviewer, name);
        } else {
            finalObject.put(nameReviewer, "NULL");
        }

        finalObject.put(emailReviewer, reviewer.getString(email));
        finalObject.put(reviewerMax, reviewer.getString(maxReview));

        // INFO PAPER
        finalObject.put(paperFinal, paper.getString(title));
        finalObject.put(topic, paper.getString(topicPaper));
        finalObject.put(docid, paper.getString(paperId));

        // COUNTRY
        finalObject.put(paperCountry, paper.has(countryPaper) ? paper.getString(countryPaper) : other);
        finalObject.put(reviewerCountry, reviewer.has(countryReviewer) ? reviewer.getString(countryReviewer) : other);

        // TOTAL
        finalObject.put(timesReview, max + 1);
        finalObject.put(timesPaperReview, maxReviewPaper + 1);

        finalObject.put("STT PAPER", paper.getInt("STT"));
        finalObject.put("STT REVIEWER", reviewer.getInt("STT"));

        if (finalObject.length() != 0) finalObjects.add(finalObject);
        mapReviewer.put(reviewer, max + 1);
        mapPaper.put(paper, maxReviewPaper + 1);
    }


    public static void main(String[] args) throws IOException {

        // Đọc file
        List<JSONObject> jsonObjectList = readXLSFile("editors_updated.xls");


        List<JSONObject> jsonObjects1 = new ArrayList<>();

        for (JSONObject jsonObject: jsonObjectList) {
            JSONObject jsonObject1 = new JSONObject();

            String reviewer1 = jsonObject.getString(Reviewer1);
            String country = "";
            String email = "";
            String max = "";

            if (reviewer1 != "") {
                String[] reviewers = reviewer1.split(",");
                if (reviewers.length >= 0) country = reviewers[0];
                if (reviewers.length >= 1) email = reviewers[1];
                if (reviewers.length >= 3)  max = reviewers[3];
                jsonObject1.put(Email , email);
                jsonObject1.put(Country, country);
                jsonObject1.put(MaxToReview, max);
                jsonObject1.put(Papers, jsonObject.getString(Papers));
                jsonObject1.put(Id, jsonObject.getString(Id));
                jsonObject1.put(Topics, jsonObject.getString(Topics));
                jsonObjects1.add(jsonObject1);
            }

            String reviewer2 = jsonObject.getString(Reviewer2);
            if (reviewer2 != "") {
                jsonObject1 = new JSONObject();
                String[] reviewers = reviewer2.split(",");
                if (reviewers.length >= 0) country = reviewers[0];
                if (reviewers.length >= 1) email = reviewers[1];
                if (reviewers.length >= 3)  max = reviewers[3];
                jsonObject1.put(Email , email);
                jsonObject1.put(Country, country);
                jsonObject1.put(MaxToReview, max);
                jsonObject1.put(Papers, jsonObject.getString(Papers));
                jsonObject1.put(Id, jsonObject.getString(Id));
                jsonObject1.put(Topics, jsonObject.getString(Topics));
                jsonObjects1.add(jsonObject1);
            }

            String reviewer3 = jsonObject.getString(Reviewer3);
            if (reviewer3 != "") {
                try {
                    jsonObject1 = new JSONObject();
                    String[] reviewers = reviewer3.split(",");
                    if (reviewers.length >= 0) country = reviewers[0];
                    if (reviewers.length >= 1) email = reviewers[1];
                    if (reviewers.length >= 3)  max = reviewers[3];
                    jsonObject1.put(Email , email);
                    jsonObject1.put(Country, country);
                    jsonObject1.put(MaxToReview, max);
                    jsonObject1.put(Papers, jsonObject.getString(Papers));
                    jsonObject1.put(Id, jsonObject.getString(Id));
                    jsonObject1.put(Topics, jsonObject.getString(Topics));
                    jsonObjects1.add(jsonObject1);
                } catch (Exception ex) {
                    System.out.println(ex);
                }

            }

        }

        System.out.println(jsonObjects1.size());
        writeExcel(jsonObjects1);
    }

}
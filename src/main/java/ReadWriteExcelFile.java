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
import java.util.*;

public class ReadWriteExcelFile {

    public static String base_url = "F:/Upload/";
    public static String nameExcelExport = "paper_reviewer.xlsx";

    public static String firstName = "First name";
    public static String lastName = "Last name";
    public static String email = "Nom d'utilisateur";
    public static String topicReviewer = "If YES, please choose the topic(s) you can review";

    public static String topicPaper = "TOPIC";
    public static String title = "TITLE";
    public static String AUTHOR = "AUTHORS";

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

    public static void writeFileFinal(List<JSONObject> jsonObjects) throws IOException {

        String url = base_url + nameExcelExport;
        //name of excel file
        String excelFileName = url;

        //name of sheet
        String sheetName = "Sheet1";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        XSSFRow rowHeader = sheet.createRow(0);
        for (int c = 0; c <= 10; c++) {
            XSSFCell cell = rowHeader.createCell(c);
            if (c == 0) cell.setCellValue("STT");
            if (c == 1) cell.setCellValue(nameReviewer);
            if (c == 2) cell.setCellValue(emailReviewer);
            if (c == 3) cell.setCellValue(paperFinal);
            if (c == 4) cell.setCellValue(topicPaper);
            if (c == 5) cell.setCellValue(paperId);
            if (c == 6) cell.setCellValue(maxReview);
            if (c == 7) cell.setCellValue(timesReview);
            if (c == 8) cell.setCellValue(paperCountry);
            if (c == 9) cell.setCellValue(reviewerCountry);
            if (c == 10) cell.setCellValue(timesPaperReview);
        }

        //iterating r number of rows
        int r = 1;
        for (JSONObject jsonObject : jsonObjects) {
            XSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c <= 10; c++) {
                XSSFCell cell = row.createCell(c);
                if (c == 0) cell.setCellValue(r);
                if (c == 1) cell.setCellValue(jsonObject.getString(nameReviewer));
                if (c == 2) cell.setCellValue(jsonObject.getString(emailReviewer));
                if (c == 3) cell.setCellValue(jsonObject.getString(paperFinal));
                if (c == 4) cell.setCellValue(jsonObject.getString(topicPaper));
                if (c == 5) cell.setCellValue(jsonObject.getString(paperId));
                if (c == 6) cell.setCellValue(jsonObject.getString(reviewerMax));
                if (c == 7) cell.setCellValue(jsonObject.getInt(timesReview));
                if (c == 8) cell.setCellValue(jsonObject.getString(paperCountry));
                if (c == 9) cell.setCellValue(jsonObject.getString(reviewerCountry));
                if (c == 10) cell.setCellValue(jsonObject.getInt(timesPaperReview));
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
        finalObject.put(reviewerCountry, paper.has(countryReviewer) ? paper.getString(countryReviewer) : other);

        // TOTAL
        finalObject.put(timesReview, max + 1);
        finalObject.put(timesPaperReview, maxReviewPaper + 1);

        if (finalObject.length() != 0) finalObjects.add(finalObject);
        mapReviewer.put(reviewer, max + 1);
        mapPaper.put(paper, maxReviewPaper + 1);
    }


    public static void main(String[] args) throws IOException {

        List<JSONObject> finalObjects = new ArrayList<>();

        List<JSONObject> lstReviewer = readXLSFile("reviewer.xls");
        List<JSONObject> lstPaper = readXLSFile("paper.xls");

        Map<JSONObject, Integer> mapPaper = new HashMap<>();
        Map<JSONObject, Integer> mapReviewer = new HashMap<>();

        // Paper - Số lần chấm
        for (JSONObject paper : lstPaper) {
            mapPaper.put(paper, 0);
        }

        // Reviewer - Bài đã chấm
        for (JSONObject reviewer : lstReviewer) {
            mapReviewer.put(reviewer, 0);
        }

        JSONObject finalObject;

        for (Map.Entry<JSONObject, Integer> entryPaper : mapPaper.entrySet()) {
            JSONObject paper = entryPaper.getKey();
            int maxReviewPaper = entryPaper.getValue();

            int plusVN = 0;

            for (Map.Entry<JSONObject, Integer> entry : mapReviewer.entrySet()) {
                JSONObject reviewer = entry.getKey();
                int max = entry.getValue();

                finalObject = new JSONObject();
                if (reviewer.getString(topicReviewer).contains(paper.getString(topicPaper)) && !(reviewer.getString(email).contains(paper.getString(AUTHOR)))) {

                    if (max < reviewer.getInt(maxReview) && maxReviewPaper < 2) {

                        if (paper.has(countryPaper) && paper.getString(countryPaper).equals("VN")) {
                            if (plusVN == 0) {
                                if (!reviewer.has(countryReviewer)) {

                                    plusVN++;
                                    saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, max, maxReviewPaper);
                                    maxReviewPaper++;

                                }
                            } else if (plusVN == 1) {

                                saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, max, maxReviewPaper);
                                maxReviewPaper++;

                            }
                        } else {

                            saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, max, maxReviewPaper);
                            maxReviewPaper++;

                        }


                    }
                }
                plusVN = 0;

            }
        }

        for (Map.Entry<JSONObject, Integer> entryPaper : mapPaper.entrySet()) {
            JSONObject paper = entryPaper.getKey();
            int maxReviewPaper = entryPaper.getValue();

            for (Map.Entry<JSONObject, Integer> entry : mapReviewer.entrySet()) {
                JSONObject reviewer = entry.getKey();
                int max = entry.getValue();

                finalObject = new JSONObject();
                if (reviewer.getString(topicReviewer).contains(paper.getString(topicPaper)) && !(reviewer.getString(email).contains(paper.getString(AUTHOR)))) {

                    if (max < reviewer.getInt(maxReview) && maxReviewPaper < 3) {

                        saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, max, maxReviewPaper);
                        maxReviewPaper++;

                    }
                }
            }
        }
        System.out.println("Count rows final: " + finalObjects.size());
        writeFileFinal(finalObjects);
    }

}
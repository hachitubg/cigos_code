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

public class WriteExcel2 {

    public static String base_url = "F:/Upload/new/";
    public static String topicReviewer = "If YES, please choose the topic(s) you can review";
    public static String topicPaper = "TOPIC";
    public static String countryReviewer = "Country";
    public static String coquanPaper = "LABOS";

    public static String max = "Max of papers to review";


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

    public static void outputFile1(List<JSONObject> jsonObjects) throws IOException {

        String url = base_url + "final/output-file-1.xlsx";

        //name of excel file
        String excelFileName = url;

        //name of sheet
        String sheetName = "Sheet1";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        XSSFRow rowHeader = sheet.createRow(0);

        for (int c = 0; c <= 10; c++) {
            XSSFCell cell = rowHeader.createCell(c);
            if (c == 0) cell.setCellValue("Stt");
            if (c == 1) cell.setCellValue("Topics");
            if (c == 2) cell.setCellValue("Total of papers");
            if (c == 3) cell.setCellValue("Papers in Vietnam");
            if (c == 4) cell.setCellValue("Total of reviewer");
            if (c == 5) cell.setCellValue("Reviewer others than VN");
            if (c == 6) cell.setCellValue("Reviews from others than VN");
            if (c == 7) cell.setCellValue("Priority to review (From 1 to the total of topics)");
        }

        //iterating r number of rows
        int r = 1;
        for (JSONObject jsonObject : jsonObjects) {
            XSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c <= 10; c++) {
                XSSFCell cell = row.createCell(c);
                if (c == 0) cell.setCellValue(r);
                if (c == 1) cell.setCellValue(jsonObject.getString("Topics"));
                if (c == 2) cell.setCellValue(jsonObject.getInt("Total of papers"));
                if (c == 3) cell.setCellValue(jsonObject.getInt("Papers in Vietnam"));
                if (c == 4) cell.setCellValue(jsonObject.getInt("Total of reviewer"));
                if (c == 5) cell.setCellValue(jsonObject.getInt("Reviewer others than VN"));
                if (c == 6) cell.setCellValue(jsonObject.getInt("Reviews from others than VN"));
                if (c == 7) cell.setCellValue(jsonObject.getInt("Priority to review (From 1 to the total of topics)"));
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


    public static void main(String[] args) throws IOException {

        // Lấy danh sách reviewer
        List<JSONObject> lstReviewer = readXLSFile("reviewer.xls");
        List<JSONObject> lstReviewerSc = readXLSFile("reviewer_sc.xls");

        List<JSONObject> lstPaper = readXLSFile("paper.xls");

        Map<String, String> topics = new HashMap<>();
        topics.put("TOPIC1", "AMCS");
        topics.put("TOPIC2", "APSC");
        topics.put("TOPIC3", "BDM");
        topics.put("TOPIC4", "CEM");
        topics.put("TOPIC5", "DTIT");
        topics.put("TOPIC6", "GEEE");
        topics.put("TOPIC7", "STBCP");
        topics.put("TOPIC8", "GTE");
        topics.put("TOPIC9", "SCMT");

        List<JSONObject> finalObjects = new ArrayList<>();
        JSONObject finalObject;
        int index = 1;
        for (Map.Entry<String, String> entry : topics.entrySet()) {
            String topicName = entry.getValue();

            finalObject = new JSONObject();

            int totalReviewer = 0;
            int totalReviewerSc = 0;

            int totalReviewerOther = 0;
            int totalReviewerOtherSc = 0;

            int totalPapers = 0;
            int totalPapersVN = 0;

            for (JSONObject reviewer : lstReviewer) {
                if (reviewer.getString(topicReviewer).contains(topicName)) {
                    totalReviewer += 1;
                    if (!(reviewer.has(countryReviewer) && reviewer.getString(countryReviewer).equals("VN"))) totalReviewerOther += Integer.valueOf(reviewer.getString(max));
                }
            }

            for (JSONObject reviewer : lstReviewerSc) {
                if (reviewer.getString(topicReviewer).contains(topicName)) {
                    totalReviewerSc += 1;
                    if (!(reviewer.has(countryReviewer) && reviewer.getString(countryReviewer).equals("VN"))) totalReviewerOtherSc += Integer.valueOf(reviewer.getString(max));
                }
            }

            for (JSONObject paper : lstPaper) {
                try {
                    if (paper.getString(topicPaper).contains(topicName)) {
                        totalPapers += 1;
                        if (paper.getString(coquanPaper).contains("Vietnam")) totalPapersVN += 1;
                    }
                } catch (Exception ex){
                    System.out.println(ex);
                }
            }

            finalObject.put("Topics", topicName);
            finalObject.put("Total of papers", totalPapers);
            finalObject.put("Papers in Vietnam", totalPapersVN);
            finalObject.put("Total of reviewer", totalReviewer + totalReviewerSc);
            finalObject.put("Reviewer others than VN", totalReviewerOther + totalReviewerOtherSc);
            finalObject.put("Total of reviews", (totalReviewer * 3) + (totalReviewerSc * 6));
            finalObject.put("Reviews from others than VN", totalReviewerOther + totalReviewerOtherSc);
            finalObject.put("Priority to review (From 1 to the total of topics)", index);

            if (finalObject.length() != 0) finalObjects.add(finalObject);
            index++;
        }

        outputFile1(finalObjects);

    }

}
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

public class PhanBoPaper {
    public static String base_url = "./file/";

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

    public static String coquanPaper = "LABOS";
    public static String coquanReviewer = "Affiliation:";

    public static String endEmail = "endEmail";

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
                    if (!cellValue.equals("")) {
                        json.put(key, cellValue);
                    }
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

    public static void writeExcel(Map<JSONObject, List<JSONObject>> map) throws IOException {

        String url = base_url + "final/Format_editors.xlsx";

        //name of excel file
        String excelFileName = url;

        //name of sheet
        String sheetName = "Sheet1";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        XSSFRow rowHeader = sheet.createRow(0);
        for (int c = 0; c <= 10; c++) {
            XSSFCell cell = rowHeader.createCell(c);
            if (c == 0) cell.setCellValue("No.");
            if (c == 1) cell.setCellValue("Id");
            if (c == 2) cell.setCellValue("Papers");
            if (c == 3) cell.setCellValue("Topics");
            if (c == 4) cell.setCellValue("Country");
            if (c == 5) cell.setCellValue("Authors");
            if (c == 6) cell.setCellValue("Affiliation (Labos)");
            if (c == 7) cell.setCellValue("Reviewer 1");
            if (c == 8) cell.setCellValue("Reviewer 2");
            if (c == 9) cell.setCellValue("Reviewer 3");
        }

        //iterating r number of rows
        int r = 1;
        for (Map.Entry<JSONObject, List<JSONObject>> entry : map.entrySet()) {

            JSONObject paper = entry.getKey();
            List<JSONObject> reviewers = entry.getValue();

            XSSFRow row = sheet.createRow(r);

            for (int c = 0; c <= 10; c++) {
                XSSFCell cell = row.createCell(c);

                if (c == 0) cell.setCellValue(r);
                if (c == 1) cell.setCellValue(paper.getString(paperId));
                if (c == 2) cell.setCellValue(paper.getString(paperFinal));
                if (c == 3) cell.setCellValue(paper.getString(topicPaper));
                if (c == 4) {
                    if (paper.getString(coquanPaper).contains("Vietnam") || paper.getString(coquanPaper).contains("VIETNAM"))
                        cell.setCellValue("Vietnam");
                    else cell.setCellValue("Other");
                }
                if (c == 5) cell.setCellValue(paper.getString(AUTHOR));
                if (c == 6) cell.setCellValue(paper.getString(coquanPaper));
                if (c == 7 && reviewers != null) {
                    JSONObject reviewer = reviewers.get(0);
                    cell.setCellValue(
                            (reviewer.has(reviewerCountry) ? reviewer.getString(reviewerCountry) : other)
                                    + ", " + reviewer.getString("EMAIL")
                                    + ", " + reviewer.getInt(reviewerMax)
                                    + ", " + reviewer.getInt(timesReview)
                    );
                }
                if (c == 8 && reviewers != null && reviewers.size() > 1) {
                    JSONObject reviewer = reviewers.get(1);
                    cell.setCellValue(
                            (reviewer.has(reviewerCountry) ? reviewer.getString(reviewerCountry) : other)
                                    + ", " + reviewer.getString("EMAIL")
                                    + ", " + reviewer.getInt(reviewerMax)
                                    + ", " + reviewer.getInt(timesReview)
                    );
                }
                if (c == 9 && reviewers != null && reviewers.size() > 2) {
                    JSONObject reviewer = reviewers.get(2);
                    cell.setCellValue(
                            (reviewer.has(reviewerCountry) ? reviewer.getString(reviewerCountry) : other)
                                    + ", " + reviewer.getString("EMAIL")
                                    + ", " + reviewer.getInt(reviewerMax)
                                    + ", " + reviewer.getInt(timesReview)
                    );
                }
            }
            r++;

        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        System.out.println("------------------------");
        System.out.println("Format editors để quản lý papers:");
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
        List<JSONObject> finalObjects = new ArrayList<>();

        List<JSONObject> lstPaper = readXLSFile("paper.xls");
        List<JSONObject> lstReviewer =  readXLSFile("reviewers.xls");;
        Collections.shuffle(lstReviewer);

        Map<JSONObject, Integer> mapPaper = new LinkedHashMap<>();
        Map<JSONObject, Integer> mapPaperVN = new LinkedHashMap<>();
        Map<JSONObject, Integer> mapPaperOther = new LinkedHashMap<>();
        Map<JSONObject, Integer> mapReviewer = new LinkedHashMap<>();

        // Paper - Số lần chấm
        int i = 1;
        for (JSONObject paper : lstPaper) {
            String topic = paper.getString(topicPaper);
            String[] topic1 = topic.split("\\(");
            String topic2 = topic1[1].replace(")", "");
            paper.put(topicPaper, topic2);
            paper.put("STT", i);
            if (paper.getString(coquanPaper).contains("Vietnam") || paper.getString(coquanPaper).contains("VIETNAM")) {
                mapPaperVN.put(paper, 0);
            } else {
                mapPaperOther.put(paper, 0);
            }
            i++;
        }
        mapPaper.putAll(mapPaperVN);
        mapPaper.putAll(mapPaperOther);

        // Reviewer - Bài đã chấm
        int j = 1;
        for (JSONObject reviewer : lstReviewer) {
            String[] emails = reviewer.getString(email).split("@");
            reviewer.put(endEmail, emails[1]);
            reviewer.put("STT", j);
            mapReviewer.put(reviewer, 0);
            j++;
        }

        // Map các TOPIC theo độ ưu tiên
        Map<String, String> topics = new TreeMap<>();
        topics.put("TOPIC1", "APSC");
        topics.put("TOPIC2", "GTE");
        topics.put("TOPIC3", "BDM");
        topics.put("TOPIC4", "CEM");
        topics.put("TOPIC5", "GEEE");
        topics.put("TOPIC6", "AMCS");
        topics.put("TOPIC7", "DTIT");
        topics.put("TOPIC8", "STBCP");
        topics.put("TOPIC9", "SCMT");

        // Vòng lặp các TOPIC theo độ ưu tiên từ 1 -> 9
        for (Map.Entry<String, String> entryTopic : topics.entrySet()) {
            String topicName = entryTopic.getValue();
            JSONObject finalObject;

            // Vòng lặp các REVIEWER
            for (Map.Entry<JSONObject, Integer> entryReviewer : mapReviewer.entrySet()) {
                JSONObject reviewer = entryReviewer.getKey();
                int plusVN = 0;

                // Vòng lặp các PAPER
                for (Map.Entry<JSONObject, Integer> entryPaper : mapPaper.entrySet()) {
                    JSONObject paper = entryPaper.getKey();

                    int maxReviewPaper = entryPaper.getValue();
                    int maxReviewer = entryReviewer.getValue();

                    // Nếu PAPER thuộc TOPIC
                    if (paper.getString(topicPaper).contains(topicName)) {
                        finalObject = new JSONObject();

                        // Check các điều kiện
                        if (
                                // 1. TOPIC của REVIEWER == TOPIC
                                reviewer.getString(topicReviewer).contains(topicName)
                                // 2. REVIEWER không thuộc nhóm tác giả của PAPER (Check theo email và đuôi email)
                                && !(paper.getString(AUTHOR).contains(reviewer.getString(email)))
                                && !(paper.getString(AUTHOR).contains(reviewer.getString(endEmail)))
                                // 3. REVIEWER không cùng cơ quan với PAPER
                                && !(reviewer.has(coquanReviewer) && reviewer.getString(coquanReviewer).contains(paper.getString(coquanPaper)))
                        ) {
                            // Check tiếp các điều kiện
                            // 1. Số lần chấm điểm của reviewer chưa max và PAPER chưa được chấm 2 lần
                            if ((maxReviewer < reviewer.getInt(maxReview)) && maxReviewPaper < 2) {
                                // 2. Nếu PAPER là bài của VIETNAM
                                if (paper.getString(coquanPaper).contains("Vietnam") || paper.getString(coquanPaper).contains("VIETNAM")) {
                                    // 3. Kiểm tra xem PAPER này đã có người nước ngoài chấm hay chưa ?
                                    if (plusVN == 0) {
                                        // Nếu chưa có người nước ngoài chấm thì thêm người nước ngoài chấm bài
                                        if (!reviewer.has(countryReviewer)) {
                                            plusVN++;
                                            saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, maxReviewer, maxReviewPaper);
                                            maxReviewPaper++;
                                        }
                                    } else if (plusVN == 1) {
                                        // Nếu bài đã có người nước ngoài chấm thì thêm ai vào cũng được
                                        saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, maxReviewer, maxReviewPaper);
                                        maxReviewPaper++;
                                    }
                                } else {
                                    // Nếu bài không phải của VIETNAM thì thêm ai chấm cũng được
                                    saveToObject(reviewer, paper, finalObject, finalObjects, mapPaper, mapReviewer, maxReviewer, maxReviewPaper);
                                    maxReviewPaper++;
                                }
                            }
                        }
                        plusVN = 0;

                    }

                }

            }
        }

        System.out.println("Count rows final: " + finalObjects.size());

        Map<JSONObject, List<JSONObject>> map101 = new HashMap<>();
        for (JSONObject paper : lstPaper) {
            map101.put(paper, null);
            for (JSONObject objectFinal : finalObjects) {
                if (paper.getString(docid).equals(objectFinal.getString(docid))) {
                    List<JSONObject> jsonObjects101 = map101.get(paper);
                    if (jsonObjects101 == null) {
                        List<JSONObject> jsonObjects102 = new ArrayList<>();
                        jsonObjects102.add(objectFinal);
                        map101.put(paper, jsonObjects102);
                    } else {
                        jsonObjects101.add(objectFinal);
                        map101.put(paper, jsonObjects101);
                    }
                }
            }
        }

        writeExcel(map101);
    }
}

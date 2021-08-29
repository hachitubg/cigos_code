import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bouncycastle.asn1.cms.AuthenticatedData;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Check {
    public static String base_url = "Upload/new/";

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

        String url = base_url + "final/check_v2.xlsx";

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
            if (c == 1) cell.setCellValue("Email");
            if (c == 2) cell.setCellValue("Topic");
        }

        //iterating r number of rows
        int r = 1;
        for (JSONObject jsonObject : jsonObjects) {
            XSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c <= 10; c++) {
                XSSFCell cell = row.createCell(c);
                if (c == 0) cell.setCellValue(r);
                if (c == 1) cell.setCellValue(jsonObject.getString("Email"));
                if (c == 2) cell.setCellValue(jsonObject.getString("Topic"));
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
        // Check members bên phải xem có tồn tại ở cột email bên trái hay không ?
        // Xuất ra danh sách những người không thuộc bên trái

        // ĐỌc file
        List<JSONObject> checks = readXLSFile("check_v2.xls");

        // Tạo 2 Object để so sánh
        List<JSONObject> jsonObjects1 = new ArrayList<>();
        List<JSONObject> jsonObjects2 = new ArrayList<>();

        // Thêm phần tử vào Object
        for (JSONObject check : checks) {
            if (check.has("Members Assigned")) {
                JSONObject jsonObject1 = new JSONObject();
                jsonObject1.put("emailLeft", check.getString("Members Assigned"));
                jsonObjects1.add(jsonObject1);
            }
            if (check.has("List of sicentific members")) {
                JSONObject jsonObject2 = new JSONObject();
                jsonObject2.put("emailRight", check.getString("List of sicentific members"));
                jsonObject2.put("topic", check.getString("Topic"));
                jsonObjects2.add(jsonObject2);
            }
        }

        // So sánh 2 Object
        List<JSONObject> jsonObjects3 = new ArrayList<>();
        for (JSONObject right : jsonObjects2) {
            int dem = 0;
            for (JSONObject left : jsonObjects1) {
                if (right.getString("emailRight").contains(left.getString("emailLeft")))
                    dem ++;
            }
            if (dem == 0) {
                JSONObject jsonObject = new JSONObject();
                jsonObject.put("Email", right.getString("emailRight"));
                jsonObject.put("Topic", right.getString("topic"));
                jsonObjects3.add(jsonObject);
            }
        }

        outputFile1(jsonObjects3);
    }

}
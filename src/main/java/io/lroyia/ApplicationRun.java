package io.lroyia;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.util.List;

/**
 * 主启动类
 * @author <a href="https://blog.lroyia.top">lroyia</a>
 * @since 2021/3/22 14:46
 **/
public class ApplicationRun {

    public static void main(String[] args) throws IOException {
        String chartSet = "utf8";
        String fileUri = null;
        for (String arg : args) {
            if(arg.startsWith("--")) {
                String[] argSplit = arg.split("=");
                if(argSplit[0].equals("--f")){
                    fileUri = argSplit[1];
                }else if(argSplit[0].equals("--c")){
                    chartSet = argSplit[1];
                }
            }
        }
        if(fileUri == null){
            System.out.println("please select a convert file.");
            return;
        }
        File file = new File(fileUri);
        String fileName = file.getName();
        String newFileName = fileName.replace(".csv", ".xls");
        try(CSVParser csv = CSVParser.parse(file, Charset.forName(chartSet), CSVFormat.DEFAULT);
            OutputStream os = new FileOutputStream(newFileName)){
            List<String> headerNames = csv.getHeaderNames();
            System.out.println(headerNames);
            HSSFWorkbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.createSheet();
            Row header = sheet.createRow(0);
            int colLength = headerNames.size();
            System.out.println(colLength);
            List<CSVRecord> records = csv.getRecords();
            System.out.println(records.size());
            for (int i = 0; i < colLength; i++) {
                createCell(header, i, headerNames.get(i));
            }
            long lastCol = csv.getRecordNumber();
            for(int i = 0; i < lastCol; i++){
                Row curRow = sheet.createRow(i+1);
                CSVRecord curRecord = records.get(i);
                for(int j = 0; j < curRecord.size(); j++){
                    createCell(curRow, j, curRecord.get(j));
                }
            }
            workbook.write(os);
        }

    }

    /**
     * 创建单元格并设置值
     * @param row   行
     * @param colIndex  列序号
     * @param value 值
     * @return  创建结果
     * @author lroyia
     * @since 2021年3月22日 16:06:57
     */
    public static Cell createCell(Row row, int colIndex, String value){
        Cell cell = row.createCell(colIndex);
        cell.setCellValue(value);
        return cell;
    }
}

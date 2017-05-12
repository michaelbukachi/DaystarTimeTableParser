package com.dita.dev.parser;

import com.dita.dev.parser.model.Event;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Parser {
    private final String filename;
    private final String shift;
    private OPCPackage pkg;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private DataFormatter formatter;
    

    public Parser(String filename, String shift) {
        this.filename = filename;
        this.shift = shift;
    }
    
    
    public void start() throws InvalidFormatException, IOException {
        System.out.println("Opening "+filename);
        pkg = OPCPackage.open(new File(filename));
        workbook = new XSSFWorkbook(pkg);
        int sheets = workbook.getNumberOfSheets();
        sheet =  null;
        System.out.println("Selecting shift...");
        for (int i = 0; i < sheets; i++) {
            sheet = workbook.getSheetAt(i);
            if (sheet.getSheetName().toLowerCase().contains(shift.toLowerCase())) {
                break;
            }
        } 
        formatter = new DataFormatter();
    }
    
    public Event getDetails(String course) throws ParseException {
        int _row, col;
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell);
                if (!text.isEmpty() && text.toLowerCase().contains(course.toLowerCase())) {
                    _row = cell.getRowIndex();
                    col = cell.getColumnIndex();
                    Event event = new Event();
                    event.date = getDate(_row, col);
                    event.location = formatter.formatCellValue(row.getCell(0));
                    return event;
                }
            }
        }
        return null;
    }
    
    private Date getDate(int row, int col) throws ParseException {
        SimpleDateFormat timeFormatter = new SimpleDateFormat("h:mma");
        SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yy");
        Calendar time = Calendar.getInstance(); 
        Calendar result = Calendar.getInstance();
        Pattern pattern;
        Matcher matcher;
        
        for (int i = row; row >= 0; i--) {
            XSSFRow tempRow = sheet.getRow(i);
            if (tempRow != null) {
                
                XSSFCell cell = tempRow.getCell(col);
                
                if (cell != null) {
                    String text = formatter.formatCellValue(cell);
                    pattern = Pattern.compile("([\\d]+:[\\d]+[apm]+)", Pattern.CASE_INSENSITIVE);
                    matcher = pattern.matcher(text);
                    if (matcher.find()) {
                        time.setTime(timeFormatter.parse(matcher.group()));
                        pattern = Pattern.compile("[\\w]+day\\s([\\d]+\\/[\\d]+\\/[\\d]+)", Pattern.CASE_INSENSITIVE);
                        i--;
                        for (int j = col; j >= 0; j--) {
                            cell = sheet.getRow(i).getCell(j);
                            text = formatter.formatCellValue(cell);
                            matcher = pattern.matcher(text);
                            if (matcher.find()) {
                                result.setTime(dateFormatter.parse(matcher.group(1)));
                                result.set(Calendar.HOUR_OF_DAY, time.get(Calendar.HOUR_OF_DAY));
                                result.set(Calendar.MINUTE, time.get(Calendar.MINUTE));
                                return result.getTime();
                            }
                        }
                    }
                }
            }
        }
        return null;
    }
    
    public void close() throws IOException {
        pkg.close();
    }
    
}

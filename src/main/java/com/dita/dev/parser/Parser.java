package com.dita.dev.parser;

import com.dita.dev.parser.model.Event;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
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
    
    
    public void start() throws InvalidFormatException, IOException, Exception {
        System.out.println("Opening "+filename);
        pkg = OPCPackage.open(new File(filename));
        workbook = new XSSFWorkbook(pkg);
        int sheets = workbook.getNumberOfSheets();
        boolean invalidSheet = true;
        sheet =  null;
        System.out.println("Selecting shift...");
        for (int i = 0; i < sheets; i++) {
            sheet = workbook.getSheetAt(i);
            if (sheet.getSheetName().toLowerCase().contains(shift.toLowerCase())) {
                invalidSheet = false;
                break;
            }
        } 
        if (invalidSheet) {
            throw new Exception("Ivalid shift");
        }
        formatter = new DataFormatter();
    }
    
    public List<Event> getDetails(String... units) throws ParseException {
        List<Event> result = new ArrayList<>();
        int _row, col;
        boolean isFound;
        for (String unit : units) {
            unit = unit.trim();
            for (Row row : sheet) {
                isFound = false;
                for (Cell cell : row) {
                    String text = formatter.formatCellValue(cell);
                    if (!text.isEmpty() && text.toLowerCase().contains(unit.toLowerCase())) {
                        _row = cell.getRowIndex();
                        col = cell.getColumnIndex();
                        Event event = new Event();
                        event.title = unit;
                        event.date = getDate(_row, col);
                        event.location = formatter.formatCellValue(row.getCell(0));
                        result.add(event);
                        isFound = true;
                        break;
                    }
                }
                if (isFound) {
                    break;
                }
            }
        }
        
        return result;
    }
    
    private Date getDate(int row, int col) throws ParseException {
        SimpleDateFormat timeFormatter = new SimpleDateFormat("h:mma");
        SimpleDateFormat dateFormatter = new SimpleDateFormat("dd/MM/yy");
        Calendar time = Calendar.getInstance(); 
        Calendar result = Calendar.getInstance();
        Pattern pattern;
        Matcher matcher;
        XSSFRow tempRow;
        XSSFCell cell;
        String text;
        
        for (int i = row; row >= 0; i--) {
            tempRow = sheet.getRow(i);
            if (tempRow != null) {
                
                cell = tempRow.getCell(col);
                
                if (cell != null) {
                    text = formatter.formatCellValue(cell);
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

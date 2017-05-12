package com.dita.dev.parser;

import com.dita.dev.parser.model.Event;
import java.io.IOException;
import java.text.ParseException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {

    public static void main(String[] args) {
        try {
            // TODO code application logic here
            Parser parser = new Parser("timetable.xlsx", "athi");
            parser.start();
            String course = "acs251";
            System.out.println("Searching for "+course);
            Event result = parser.getDetails(course);
            if (result != null) {
                System.out.println(result);
            } else {
                System.out.println("Not found");
            }
            parser.close();
        } catch (InvalidFormatException | IOException | ParseException ex) {
            ex.printStackTrace();
        }
    }
    
    
}

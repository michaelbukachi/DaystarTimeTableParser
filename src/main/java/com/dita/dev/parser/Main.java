package com.dita.dev.parser;

import com.dita.dev.parser.model.Event;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;

public class Main {

    public static void main(String[] args) {
        String file;
        String shift;
        String units;
        
        if (args.length > 0) {
            file = args[0];
            shift = args[1];
            units = args[2];
        } else {
            Scanner input = new Scanner(System.in);
            System.out.println("Enter excel file path: ");
            file = input.nextLine();
            System.out.println("Enter shift");
            shift = input.nextLine();
            System.out.println("Enter units e.g acs101, acs102");
            units = input.nextLine();
        }
        
        String[] unitsList = units.split(",");
             
        try {
            // TODO code application logic here
            Parser parser = new Parser(file, shift);
            parser.start();
            List<Event> result = parser.getDetails(unitsList);
            if (!result.isEmpty()) {
                for (Event event : result) {
                    System.out.println(event);
                }
            } else {
                System.out.println("Not found");
            }
            parser.close();
        } catch (InvalidFormatException | IOException | ParseException ex) {
            ex.printStackTrace();
        } catch (InvalidOperationException ex) {
            System.err.println("File not found");
        } catch (Exception ex) {
            System.err.println(ex.getMessage());
        }
    }
    
    
}

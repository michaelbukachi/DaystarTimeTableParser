package com.dita.dev.parser.model;

import java.util.Date;


public class Event {
    public String title;
    public Date date;
    public String day;
    public String location;

    @Override
    public String toString() {
        return "Event{" + "title=" + title + ", date=" + date + ", location=" + location + '}';
    }

    
}

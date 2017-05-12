package com.dita.dev.parser.model;

import java.util.Date;


public class Event {
    public Date date;
    public String day;
    public String location;

    @Override
    public String toString() {
        return "Event{" + "date=" + date + ", location=" + location + '}';
    }
}

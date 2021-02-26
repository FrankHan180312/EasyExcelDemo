package model;

import com.fasterxml.jackson.databind.annotation.JsonSerialize;

import java.util.Date;

public class Event {
    public String name;

    @JsonSerialize(using = CustomDateSerializer.class)
    public Date eventDate;
}


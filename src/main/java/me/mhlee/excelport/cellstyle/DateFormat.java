package me.mhlee.excelport.cellstyle;

import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;

public enum DateFormat {
    NONE(DateTimeFormatter.ofLocalizedDate(FormatStyle.LONG)),
    YYYYMMDD(DateTimeFormatter.ofPattern("yyyyMMdd")),
    YYYYMMDD_WITH_DASH(DateTimeFormatter.ofPattern("yyyy-MM-dd")),
    YYYYMMDD_WITH_DOT(DateTimeFormatter.ofPattern("yyyy.MM.dd")),
    YYYYMMDD_WITH_SLASH(DateTimeFormatter.ofPattern("yyyy/MM/dd")),
    YYYYMMDD_HHMISS(DateTimeFormatter.ofPattern("yyyyMMddHHmmss")),
    YYYYMMDD_HHMISS_WITH_DASH_COLON(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")),
    YYYYMMDD_HHMISS_WITH_DOT_COLON(DateTimeFormatter.ofPattern("yyyy.MM.dd HH:mm:ss")),
    YYYYMMDD_HHMISS_WITH_SLASH_COLON(DateTimeFormatter.ofPattern("yyyy/MM.dd HH:mm:ss"));

    private DateTimeFormatter format;

    DateFormat(DateTimeFormatter ofPattern) {
        this.format = ofPattern;
    }

    public DateTimeFormatter getFormat() {
        return format;
    }

}
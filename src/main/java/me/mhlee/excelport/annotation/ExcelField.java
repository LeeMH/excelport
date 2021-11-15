package me.mhlee.excelport.annotation;

import me.mhlee.excelport.cellstyle.Align;
import me.mhlee.excelport.cellstyle.DateFormat;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ToStringExclude;
import org.apache.commons.lang3.math.NumberUtils;

public class ExcelField implements Comparable<ExcelField> {
    // 컬럼 타이틀
    private String name;

    // 필드명 (컬럼 타이틀이 없는경우, 출력)
    private String fieldName;

    // 컬럼 출력 순번 (낮을수록 먼저 출력됨, 디폴트 999)
    private int order = 999;

    // 정렬
    private Align align;

    // 포맷
    private DateFormat dateFormat;


    @Override
    public int compareTo(ExcelField o) {
        return  getOrder() - o.getOrder();
    }

    public static ExcelField create(String fieldName, String name, int order, Align align, DateFormat dateFormat) {
        return new ExcelField()
                .setFieldName(fieldName)
                .setName(name)
                .setOrder(order)
                .setAlign(align)
                .setDateFormat(dateFormat);
    }

    public String getFieldName() { return fieldName; }

    public String getName() { return name; }

    public int getOrder() { return order; }

    public Align getAlign() { return align; }

    public DateFormat getDateFormat() { return dateFormat; }

    public String getHeader() {
        return (StringUtils.isNotEmpty(getName())) ? getName() : getFieldName();
    }

    public ExcelField setFieldName(String fieldName) {
        this.fieldName = fieldName;
        return this;
    }

    public ExcelField setName(String name) {
        this.name = name;
        return this;
    }

    public ExcelField setOrder(int order) {
        this.order = order;
        return this;
    }

    public ExcelField setAlign(Align align) {
        this.align = align;
        return this;
    }

    public ExcelField setDateFormat(DateFormat dateFormat) {
        this.dateFormat = dateFormat;
        return this;
    }

    @Override
    public String toString() {
        return String.format("< fieldName=[%s], name=[%s], order=[%d], align=[%s], dateFormat=[%s] >",
                fieldName, name, order, align, dateFormat);
    }

    public void setFieldByStringTemplate(String template) {
        String[] kv = template.split("=");

        if (kv.length < 2) return;
        kv[0] = StringUtils.strip(kv[0]);
        kv[1] = StringUtils.strip(kv[1]);


        if (StringUtils.equalsIgnoreCase("name", kv[0])) {
            this.name = kv[1];
        } else if (StringUtils.equalsIgnoreCase("fieldName", kv[0])) {
            this.fieldName = kv[1];
        } else if (StringUtils.equalsIgnoreCase("order", kv[0])) {
            this.order = NumberUtils.toInt(kv[1], 999);
        } else if (StringUtils.equalsIgnoreCase("align", kv[0])) {
            this.align = Align.valueOf(StringUtils.upperCase(kv[1]));
        } else if (StringUtils.equalsIgnoreCase("dateFormat", kv[0])) {
            this.dateFormat = DateFormat.valueOf(StringUtils.upperCase(kv[1]));
        } else {
            System.out.println("unknown string key = [" + kv[0] + "]");
        }
    }


}

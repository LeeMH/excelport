package me.mhlee.excelport.annotation;

import me.mhlee.excelport.cellstyle.Align;
import me.mhlee.excelport.cellstyle.DateFormat;

import java.time.LocalDateTime;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

class AnnotationParserTest {

    public static class SampleDto {
        @Excel(name = "이름", order = 2, align = Align.CENTER)
        private String name;

        @Excel(name = "나이", order = 1, align = Align.RIGHT)
        private int age;

        @Excel(order = 3, format = DateFormat.YYYYMMDD)
        private LocalDateTime joined_at = LocalDateTime.now();

        @Excel
        private Long nullValueField;

        private String noExcelAnnotationField;

        public static SampleDto create(String name, int age) {
            SampleDto obj = new SampleDto();
            obj.name = name;
            obj.age = age;

            return obj;
        }

        public String getName() { return name + " 회원님"; }

        public int getAge() { return age; }

        public Long getNullValueField() { return (nullValueField != null) ? nullValueField : 999; }
    }

    @Test
    public void testExtractExcelColumns() {
        SampleDto dto = SampleDto.create("홍길동", 25);
        List<ExcelField> result = AnnotationParser.extractExcelColumns(dto);

        assertTrue(result.size() == 4);

        //first column's information
        assertTrue(result.get(0).getFieldName().equals("age"));
        assertTrue(result.get(0).getAlign().equals(Align.RIGHT));
        assertTrue(result.get(0).getName().equals("나이"));

        //last column's information
        assertTrue(result.get(result.size() -1).getFieldName().equals("nullValueField"));
        assertTrue(result.get(result.size() -1).getAlign().equals(Align.NONE));
        assertTrue(result.get(result.size() -1).getName().equals(StringUtils.EMPTY));

    }

    @Test
    public void testToRow() {
        SampleDto dto = SampleDto.create("홍길동", 25);
        List<ExcelField> parsedExcel = AnnotationParser.extractExcelColumns(dto);
        List<Object> rows = AnnotationParser.toRow(dto, parsedExcel);

        assertTrue(rows.size() == 4);

        // 필드에서 값을 추출
        assertTrue(rows.get(0) instanceof Integer);
        assertTrue(dto.getAge() == (int)rows.get(0));

        // getter 메소드를 호출하여 값을 추출
        assertTrue(rows.get(1) instanceof String);
        assertTrue(dto.getName().equals(rows.get(1)));
    }
}
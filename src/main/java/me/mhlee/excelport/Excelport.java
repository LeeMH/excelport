package me.mhlee.excelport;

import me.mhlee.excelport.annotation.AnnotationParser;
import me.mhlee.excelport.annotation.ExcelField;
import me.mhlee.excelport.cellstyle.Align;
import me.mhlee.excelport.cellstyle.CellStyleManager;
import me.mhlee.excelport.cellstyle.DateFormat;

import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.google.gson.Gson;

public class Excelport {

    private static final String NEW_LINE = System.lineSeparator();
    private static final long MAX_PER_SHEET = 60_000L;
    private static final Gson gson = new Gson();

    public static void toCsv(OutputStream os, Iterator iterator) {
        List<String> lines = new LinkedList<>();

        // pick first item for extracting excel column
        if (!iterator.hasNext()) return;
        Object firstItem = iterator.next();

        // write column header
        List<ExcelField> parsedExcels = AnnotationParser.extractExcelColumns(firstItem);
        List<String> headers = AnnotationParser.toHeader(parsedExcels);
        lines.add(toCsvString(parsedExcels, headers, true));

        // write first item
        lines.add(toCsvString(parsedExcels, AnnotationParser.toRow(firstItem, parsedExcels), false));

        while(iterator.hasNext()) {
            Object obj = iterator.next();
            lines.add(toCsvString(parsedExcels, AnnotationParser.toRow(obj, parsedExcels), false));
        }

        try {
            for(int ii = 0; ii < lines.size(); ii++) {
                os.write(lines.get(ii).getBytes());
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            if (os != null) try {
                os.close();
            } catch (Exception e) {
            }
        }
    }

    public static void toJson(OutputStream os, Iterator iterator) {
        List<byte[]> lines = new LinkedList<>();

        // pick first item for extracting excel column
        if (!iterator.hasNext()) return;
        Object firstItem = iterator.next();

        // write column header
        List<ExcelField> parsedExcels = AnnotationParser.extractExcelColumns(firstItem);

        // write first item
        lines.add(toJsonByte(parsedExcels, AnnotationParser.toRow(firstItem, parsedExcels)));

        while(iterator.hasNext()) {
            Object obj = iterator.next();

            lines.add(toJsonByte(parsedExcels, AnnotationParser.toRow(obj, parsedExcels)));
        }

        try {
            for(int ii = 0; ii < lines.size(); ii++) {
                os.write(lines.get(ii));
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            if (os != null) try {
                os.close();
            } catch (Exception e) {
            }
        }
    }

    private static byte[] toJsonByte(List<ExcelField> parsedExcels, List<Object> columns) {
        Map<String, Object> map = new HashMap<>();
        for(int ii = 0; ii < parsedExcels.size(); ii++) {
            String key = parsedExcels.get(ii).getHeader();
            Object value = getMatchedType(columns.get(ii), parsedExcels.get(ii).getDateFormat());
            map.put(key, value);
        }

        String json = gson.toJson(map) + NEW_LINE;

        return json.getBytes();
    }

    private static String toCsvString(List<ExcelField> parsedExcels, List<? extends Object> columns, boolean isHeader) {
        List<String> row = new ArrayList<>();
        for(int ii = 0; ii < parsedExcels.size(); ii++) {
            String column = String.format("\"%s\"", getMatchedType(columns.get(ii), parsedExcels.get(ii).getDateFormat()));
            row.add(column);
        }

        return row.stream().collect(Collectors.joining(",")) + NEW_LINE;
    }

    /** 입력된 iterator 객체를 엑셀로 변환
     * @param os
     * @param iterator
     */
    public static void toExcel(OutputStream os, Iterator iterator, String[] template) {
        SXSSFWorkbook workbook = null;
        workbook = new SXSSFWorkbook(1000);
        SXSSFSheet sheet = null;

        if (!iterator.hasNext()) return;

        // pick first item for extracting excel column
        Object firstItem = iterator.next();
        List<ExcelField> parsedExcels = AnnotationParser.extractExcelColumnsFromString(template);

        boolean writeFirstLine = true;
        do {
            sheet = workbook.createSheet();
            sheet.trackAllColumnsForAutoSizing();
            toExcel(workbook, sheet, iterator, firstItem, parsedExcels, writeFirstLine);
            writeFirstLine = false;
        } while(iterator.hasNext());

            try {
            workbook.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (os != null) try { os.flush(); os.close(); } catch(Exception e) {}
            if (workbook != null) try { workbook.dispose(); } catch(Exception e) {}
            if (workbook != null) try { workbook.close(); } catch(Exception e) {}
        }
    }

    /** 입력된 iterator 객체를 엑셀로 변환
     * @param os
     * @param iterator
     */
    public static void toExcel(OutputStream os, Iterator iterator, Class template) {
        SXSSFWorkbook workbook = null;
        workbook = new SXSSFWorkbook(1000);
        SXSSFSheet sheet = null;

        if (!iterator.hasNext()) return;

        // pick first item for extracting excel column
        Object firstItem = iterator.next();
        List<ExcelField> parsedExcels = null;
        try {
            parsedExcels = AnnotationParser.extractExcelColumns(template.newInstance());
        } catch(InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        }

        boolean writeFirstLine = true;
        do {
            sheet = workbook.createSheet();
            sheet.trackAllColumnsForAutoSizing();
            toExcel(workbook, sheet, iterator, firstItem, parsedExcels, writeFirstLine);
            writeFirstLine = false;
        } while(iterator.hasNext());

            try {
            workbook.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (os != null) try { os.flush(); os.close(); } catch(Exception e) {}
            if (workbook != null) try { workbook.dispose(); } catch(Exception e) {}
            if (workbook != null) try { workbook.close(); } catch(Exception e) {}
        }
    }

    /** 입력된 iterator 객체를 엑셀로 변환
     * @param os
     * @param iterator
     */
    public static void toExcel(OutputStream os, Iterator iterator) {
        SXSSFWorkbook workbook = null;
        workbook = new SXSSFWorkbook(1000);
        SXSSFSheet sheet = null;

        if (!iterator.hasNext()) return;

        // pick first item for extracting excel column
        Object firstItem = iterator.next();
        List<ExcelField> parsedExcels = AnnotationParser.extractExcelColumns(firstItem);

        boolean writeFirstLine = true;
        do {
            sheet = workbook.createSheet();
            sheet.trackAllColumnsForAutoSizing();
            toExcel(workbook, sheet, iterator, firstItem, parsedExcels, writeFirstLine);
            writeFirstLine = false;
        } while(iterator.hasNext());

        try {
            workbook.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (os != null) try { os.flush(); os.close(); } catch(Exception e) {}
            if (workbook != null) try { workbook.dispose(); } catch(Exception e) {}
            if (workbook != null) try { workbook.close(); } catch(Exception e) {}
        }
    }

    private static void toExcel(Workbook workbook, Sheet sheet, Iterator iterator, Object firstItem, List<ExcelField> parsedExcels, boolean writeFirstLine) {
        int rowNum = 0;

        //write column header
        List<String> headers = AnnotationParser.toHeader(parsedExcels);
        processHeaderRow(workbook, sheet.createRow(rowNum++), headers);

        //write first record
        if (writeFirstLine) {
            processDataRow(workbook, sheet.createRow(rowNum++), AnnotationParser.toRow(firstItem, parsedExcels), parsedExcels);
        }

        while (iterator.hasNext()) {
            Object obj = iterator.next();

            processDataRow(workbook, sheet.createRow(rowNum++), AnnotationParser.toRow(obj, parsedExcels), parsedExcels);

            // sheet 당 최대 출력 line 조정
            if (rowNum % MAX_PER_SHEET == 1) {  //rowNum = 데이터 컬럼(n) + 헤더 컬럼(1)
                break;
            }
        }

        //컬럼 넓이 조절
        for (int ii = 0; ii < parsedExcels.size(); ii++) {
            sheet.autoSizeColumn(ii);
        }
    }

    /** 엑셀 헤더 row 출력
     * @param workbook
     * @param row
     * @param columns
     */
    private static void processHeaderRow(Workbook workbook, Row row, List<? extends Object> columns) {
        CellStyle cellStyle = CellStyleManager.getMatchedCellStyle(workbook, StringUtils.EMPTY, Align.CENTER);

        for (int ii = 0; ii < columns.size(); ii++) {
            Cell cell = row.createCell(ii);
            fillCellMatchedType(cell, columns.get(ii), DateFormat.NONE);
            cell.setCellStyle(cellStyle);
        }
    }

    /** 엑셀 데이터 row 출력
     * @param workbook
     * @param row
     * @param columns
     * @param columnsInfo
     */
    private static void processDataRow(Workbook workbook, Row row, List<? extends Object> columns, List<ExcelField> columnsInfo) {
        for (int ii = 0; ii < columns.size(); ii++) {
            Cell cell = row.createCell(ii);
            fillCellMatchedType(cell, columns.get(ii), columnsInfo.get(ii).getDateFormat());
            cell.setCellStyle(CellStyleManager.getMatchedCellStyle(
                    workbook, columns.get(ii), columnsInfo.get(ii).getAlign()));
        }
    }


    /** 엑셀의 cell 에 데이터 출력
     * Object 타입과 format 을 조합하여 적절한 형태로 출력
     * @param cell
     * @param obj
     * @param dateFormat
     */
    private static void fillCellMatchedType(Cell cell, Object obj, DateFormat dateFormat) {
        if (Objects.isNull(obj)) return;

        if (obj instanceof String) {
            cell.setCellValue((String)obj);
        }
        else if (obj instanceof Number) {
            cell.setCellValue(NumberUtils.toDouble(obj.toString()));
        }
        else if (obj instanceof Date) { //Date 타입인 경우, LocalDateTime 타입으로 변환후 처리
            LocalDate localDate = ((Date)obj).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            cell.setCellValue(localDate.format(dateFormat.getFormat()));
        }
        else if (obj instanceof LocalDate) { //LocalDate 타입인 경우, LocalDateTime 타입으로 변환후 처리
            LocalDate localDate = ((LocalDate)obj);
            cell.setCellValue(localDate.format(dateFormat.getFormat()));
        }
        else if (obj instanceof LocalDateTime) {
            // format 의 디폴트는 Date 포맷이다. LocalDateTime 타입이지만 format 이 없는 경우, DateTime 에 대한 포맷으로 적용
            if (DateFormat.NONE.equals(dateFormat)) {
                cell.setCellValue(((LocalDateTime)obj).format(DateTimeFormatter.ofLocalizedDateTime(FormatStyle.MEDIUM)) );
            } else
                cell.setCellValue(((LocalDateTime)obj).format(dateFormat.getFormat()));
        }
        else if (obj instanceof Boolean) cell.setCellValue((Boolean) obj);
        else cell.setCellValue(obj.toString());
    }

    /** Object 타입과 format 을 조합하여 적절한 형태로 출력
     * @param obj
     * @param dateFormat
     */
    private static Object getMatchedType(Object obj, DateFormat dateFormat) {
        if (Objects.isNull(obj)) return null;

        if (obj instanceof String) {
            return obj;
        }
        else if (obj instanceof Integer || obj instanceof Long) {
            return NumberUtils.toLong(obj.toString());
        }
        else if (obj instanceof Float || obj instanceof Double) {
            return NumberUtils.toDouble(obj.toString());
        }
        else if (obj instanceof Date) { //Date 타입인 경우, LocalDate 타입으로 변환후 처리
            LocalDate localDate = ((Date)obj).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            return localDate.format(dateFormat.getFormat());
        }
        else if (obj instanceof LocalDate) {
            LocalDate localDate = ((LocalDate)obj);
            return localDate.format(dateFormat.getFormat());
        }
        else if (obj instanceof LocalDateTime) {
            // format 의 디폴트는 Date 포맷이다. LocalDateTime 타입이지만 format 이 없는 경우, DateTime 에 대한 포맷으로 적용
            if (DateFormat.NONE.equals(dateFormat)) {
                return ((LocalDateTime)obj).format(DateTimeFormatter.ofLocalizedDateTime(FormatStyle.MEDIUM));
            } else
                return ((LocalDateTime)obj).format(dateFormat.getFormat());
        }
        else if (obj instanceof Boolean)
            return obj;
        else if (obj instanceof Map)
            return obj;
        else
            return obj.toString();
    }

}

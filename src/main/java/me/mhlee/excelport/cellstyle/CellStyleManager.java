package me.mhlee.excelport.cellstyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleManager {
    private static CellStyle left = null;  //좌측정렬
    private static CellStyle center = null;  //중앙정렬
    private static CellStyle right = null;  //우측정렬
    private static CellStyle rightWithNumberFormat = null; //우측정렬 천단위 컴마
    private static CellStyle rightWithFloatFormat = null; //우측정렬 천단위 컴마 + 소숫점

    private static CellStyle getLeft(Workbook workbook) {
        if (left == null) {
            left = workbook.createCellStyle();
            left.setAlignment(HorizontalAlignment.LEFT);
        }

        return left;
    }

    private static CellStyle getCenter(Workbook workbook) {
        if (center == null) {
            center = workbook.createCellStyle();
            center.setAlignment(HorizontalAlignment.CENTER);
        }

        return center;
    }

    private static CellStyle getRight(Workbook workbook) {
        if (right == null) {
            right = workbook.createCellStyle();
            right.setAlignment(HorizontalAlignment.RIGHT);
        }

        return right;
    }

    private static CellStyle getRightWithNumberFormat(Workbook workbook) {
        if (rightWithNumberFormat == null) {
            rightWithNumberFormat = workbook.createCellStyle();
            rightWithNumberFormat.setAlignment(HorizontalAlignment.RIGHT);
            rightWithNumberFormat.setDataFormat(workbook.createDataFormat().getFormat("#,###"));
        }

        return rightWithNumberFormat;
    }

    private static CellStyle getRightWithFloatFormat(Workbook workbook) {
        if (rightWithFloatFormat == null) {
            rightWithFloatFormat = workbook.createCellStyle();
            rightWithFloatFormat.setAlignment(HorizontalAlignment.RIGHT);
            rightWithFloatFormat.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        }

        return rightWithFloatFormat;
    }


    /** 데이터 타입과 정렬 타입으로 적절한 CellType 리턴
     * @param workbook
     * @param obj
     * @param align
     * @return
     */
    public static CellStyle getMatchedCellStyle(Workbook workbook, Object obj, Align align) {
        // 정렬 None 인경우, 포맷별 적절한 값 적용
        if (Align.NONE.equals(align)) {
            if (obj instanceof Integer || obj instanceof  Long) return getRightWithNumberFormat(workbook);
            else if (obj instanceof Float || obj instanceof Double) return getRightWithFloatFormat(workbook);
            else if (obj instanceof Number) return getRight(workbook);
            else return getCenter(workbook);
        } else if (Align.LEFT.equals(align)) {
            return getLeft(workbook);
        } else if (Align.CENTER.equals(align)) {
            return getCenter(workbook);
        } else if (Align.RIGHT.equals(align)) {
            // 우측 정렬인 경우, 데이터가 숫자관련이라면 적절한 포맷팅 부여
            if (obj instanceof Integer || obj instanceof  Long) return getRightWithNumberFormat(workbook);
            else if (obj instanceof Float || obj instanceof Double) return getRightWithFloatFormat(workbook);
            else return getRight(workbook);
        } else return getCenter(workbook);
    }
}

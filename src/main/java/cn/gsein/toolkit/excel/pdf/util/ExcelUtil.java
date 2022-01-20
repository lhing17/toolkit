package cn.gsein.toolkit.excel.pdf.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelUtil {

    /**
     * 获取excel中列数最多的行号
     */
    public static int getRowNumOfMaxColumnCount(Sheet sheet) {
        int rowNum = 0;
        int maxCol = 0;
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row != null && maxCol < row.getLastCellNum() + 1) {
                maxCol = row.getLastCellNum() + 1;
                rowNum = r;
            }
        }
        return rowNum;
    }

    public static int[] getColumnWidths(Sheet sheet) {

        // 获取最多列的行号
        int rowNum = getRowNumOfMaxColumnCount(sheet);
        Row row = sheet.getRow(rowNum);
        short firstCellNum = row.getFirstCellNum();
        int cellCount = row.getLastCellNum() + 1 - firstCellNum;
        int[] widths = new int[cellCount];

        for (int i = firstCellNum; i < cellCount + firstCellNum; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                widths[i - firstCellNum] = sheet.getColumnWidth(i);
            }
        }
        return widths;

    }

    public static int[] getRowHeights(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        int rowNum = sheet.getLastRowNum() + 1 - firstRowNum;

        int[] rowHeights = new int[rowNum];
        for (int i = firstRowNum; i < rowNum + firstRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                rowHeights[i - firstRowNum] = row.getHeight();
            }
        }

        return rowHeights;
    }
}

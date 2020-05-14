package cn.gsein.toolkit.excel.util;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Objects;


/**
 * 将EXCEL转为PDF的工具类
 *
 * @author G. Seinfeld
 * @since 2020-05-11
 */
public final class ExcelToPdfPdfBoxUtil {
    private ExcelToPdfPdfBoxUtil() {
    }

    private static final String XLS = ".xls";
    private static final String XLSX = ".xlsx";
    private static final String XLSM = ".xlsm";
    private static final String PERCENT = "%";

    private static final float POINTS_PER_MM = 2.8346457f;
    private static final float PAGE_WIDTH = 210 * POINTS_PER_MM;
    private static final float PAGE_HEIGHT = 297 * POINTS_PER_MM;


    /**
     * 获取字体
     */
    private static org.apache.poi.ss.usermodel.Font getExcelFont(Workbook workbook, Cell cell, String excelName) {
        if (excelName.endsWith(XLS)) {
            return ((HSSFCell) cell).getCellStyle().getFont(workbook);
        }
        return ((XSSFCell) cell).getCellStyle().getFont();
    }


    /**
     * 获取excel单元格数据显示格式
     */
    private static String getNumStyle(String dataFormat) throws Exception {
        if (dataFormat == null || dataFormat.length() == 0) {
            throw new Exception("");
        }
        if (dataFormat.contains(PERCENT)) {
            return dataFormat;
        } else {
            return dataFormat.substring(0, dataFormat.length() - 2);
        }

    }

    /**
     * 判断单元格是否是合并单元格
     */
    private static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }


    /**
     * 计算合并单元格合并的跨行跨列数
     */
    private static int[] getMergedSpan(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        int[] span = {1, 1};
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstColumn == column && firstRow == row) {
                span[0] = lastRow - firstRow + 1;
                span[1] = lastColumn - firstColumn + 1;
                break;
            }
        }
        return span;
    }

    /**
     * 获取excel中列数最多的行号
     */
    private static int getMaxColRowNum(Sheet sheet) {
        int rowNum = 0;
        int maxCol = 0;
        for (int r = sheet.getFirstRowNum(); r < sheet.getPhysicalNumberOfRows(); r++) {
            Row row = sheet.getRow(r);
            if (row != null && maxCol < row.getPhysicalNumberOfCells()) {
                maxCol = row.getPhysicalNumberOfCells();
                rowNum = r;
            }
        }
        return rowNum;
    }

//    /**
//     * excel垂直对齐方式映射到pdf对齐方式
//     */
//    private static int getVerticalAlignment(VerticalAlignment align) {
//        switch (align) {
//            case BOTTOM:
//                return com.itextpdf.text.Element.ALIGN_BOTTOM;
//            case TOP:
//                return com.itextpdf.text.Element.ALIGN_TOP;
//            default:
//                return com.itextpdf.text.Element.ALIGN_MIDDLE;
//        }
//    }
//
//
//    /**
//     * excel水平对齐方式映射到pdf水平对齐方式
//     */
//    private static int getHorizontalAlignment(HorizontalAlignment align) {
//        switch (align) {
//            case RIGHT:
//                return com.itextpdf.text.Element.ALIGN_RIGHT;
//            case LEFT:
//                return com.itextpdf.text.Element.ALIGN_LEFT;
//            default:
//                return com.itextpdf.text.Element.ALIGN_CENTER;
//        }
//    }

    /**
     * 格式化数字
     */
    private static String numFormat(String pattern, double num) {
        DecimalFormat format = new DecimalFormat(pattern);
        return format.format(num);
    }

    /**
     * 获取excel中每列宽度的占比
     */
    private static float[] getColWidth(Sheet sheet) {
        int rowNum = getMaxColRowNum(sheet);
        Row row = sheet.getRow(rowNum);
        int cellCount = row.getPhysicalNumberOfCells();
        int[] colWidths = new int[cellCount];
        int sum = 0;

        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                colWidths[i] = sheet.getColumnWidth(i);
                sum += sheet.getColumnWidth(i);
            }
        }

        float[] colWidthPer = new float[cellCount];
        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            colWidthPer[i] = (float) colWidths[i] / sum * 100;
        }
        return colWidthPer;
    }

    private static float[] getRowHeight(Sheet sheet) {
        int rowNum = sheet.getPhysicalNumberOfRows();

        float[] rowHeight = new float[rowNum];
        float sum = 0;
        for (int i = sheet.getFirstRowNum(); i < rowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                rowHeight[i] = row.getHeightInPoints();
                sum += rowHeight[i];
            }
        }

        float[] rowHeightPer = new float[rowNum];
        for (int i = sheet.getFirstRowNum(); i < rowNum; i++) {
            rowHeightPer[i] = (float) rowHeight[i] / sum * 100;
        }
        return rowHeightPer;
    }

    public static void excelToPdf(String excelPath, String pdfPath) throws Exception {
        InputStream in = new FileInputStream(excelPath);

        // 获取工作簿
        Workbook workbook = getWorkbook(excelPath, in);


        int sheetCount = workbook.getNumberOfSheets();
        //新建PDF文档
        PDDocument document = new PDDocument();
        for (int i = 0; i < sheetCount; i++) {
            // 获取第i张工作表
            Sheet sheet = workbook.getSheetAt(i);

            // 设置页面大小为A4纸
            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);
            PDPageContentStream stream = new PDPageContentStream(document, page);
            InputStream inputStream = ExcelToPdfPdfBoxUtil.class.getClassLoader().getResourceAsStream("font/STXIHEI.TTF");
            PDType0Font font = PDType0Font.load(document, inputStream, true);
            stream.setFont(font, 14);
            stream.setStrokingColor(Color.BLACK);

            float[] excelWidths = getColWidth(sheet);
            float[] widths = handleWidths(excelWidths);

            float[] excelHeights = getRowHeight(sheet);
            float[] heights = handleHeights(excelHeights);

            int colCount = widths.length;
            for (int rowNum = sheet.getFirstRowNum(); rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = row.getFirstCellNum(); (cellNum < row.getLastCellNum() || cellNum < colCount) && cellNum > -1; cellNum++) {
                        if (cellNum >= row.getPhysicalNumberOfCells()) {
                            continue;
                        }
                        Cell excelCell = row.getCell(cellNum);
                        // 获取单元格的值
                        String value = getCellValue(excelCell);

                        //设置单元格字体
                        org.apache.poi.ss.usermodel.Font excelFont = getExcelFont(workbook, excelCell, excelPath);
//                        Font pdFont = new Font(baseFont, excelFont.getFontHeightInPoints(),
//                                excelFont.getBold() ? Font.BOLD : Font.NORMAL, BaseColor.BLACK);


//                         处理图片
                        List<PicturesInfo> infos = PoiExtend.getAllPictureInfos(sheet, rowNum, rowNum, cellNum, cellNum, false);
                        if (!infos.isEmpty()) {
                            PicturesInfo info = infos.get(0);

                            float x = 29.76f;
                            float y = PAGE_HEIGHT - 15f;
                            for (int j = 0; j < cellNum; j++) {
                                x += widths[j];
                            }

                            float width = 0;
                            for (int j = cellNum; j < cellNum + info.getMaxCol() - info.getMinCol() + 1; j++) {
                                width += widths[j];
                            }

                            for (int j = 0; j < rowNum; j++) {
                                y -= heights[j];
                            }
                            float height = 0;
                            for (int j = rowNum; j < rowNum + info.getMaxRow() - info.getMinRow() + 1; j++) {
                                height += heights[j];
                            }

                            stream.drawImage(PDImageXObject.createFromByteArray(document, info.getPictureData(), null), x, y - height, width, height);
                        }


                        int border = getBorder(Objects.requireNonNull(excelCell));
                        if (isMergedRegion(sheet, rowNum, cellNum)) {
                            int[] span = getMergedSpan(sheet, rowNum, cellNum);
                            //忽略合并过的单元格
                            if (span[0] == 1 && span[1] == 1) {
                                continue;
                            }
                            border = getPdfCellBorderForMerged(sheet, rowNum, cellNum, span);
                            //合并过的列直接跳过
                            PdfUtil.drawRect(stream, widths, heights, rowNum, cellNum, span[0], span[1], border);
                            PdfUtil.drawString(font, stream, excelFont, widths, heights, rowNum, cellNum, span[0], span[1], value);
                            cellNum = cellNum + span[1] - 1;
                        } else {
                            PdfUtil.drawRect(stream, widths, heights, rowNum, cellNum, 1, 1, border);
                            PdfUtil.drawString(font, stream, excelFont, widths, heights, rowNum, cellNum, 1, 1, value);
                        }
                    }
                }
            }
            stream.stroke();
            stream.close();
        }

        document.save(pdfPath);
        document.close();

    }

    private static float[] handleHeights(float[] excelHeights) {
        float[] heights = new float[excelHeights.length];
        for (int i = 0; i < excelHeights.length; i++) {
            heights[i] = excelHeights[i] / 100 * (PAGE_HEIGHT - 30);
        }
        return heights;
    }

    private static float[] handleWidths(float[] excelWidths) {
        float[] widths = new float[excelWidths.length];
        for (int i = 0; i < excelWidths.length; i++) {
            widths[i] = excelWidths[i] / 100 * PAGE_WIDTH * 0.94f;
        }
        return widths;
    }


    private static String getCellValue(Cell excelCell) {
        String value = "";

        if (excelCell != null) {
            value = excelCell.toString().trim();
            if (value.length() != 0) {
                //获取excel单元格数据显示样式
                String dataFormat = excelCell.getCellStyle().getDataFormatString();
                //noinspection AlibabaUndefineMagicConstant
                if (!"General".equals(dataFormat) && !"@".equals(dataFormat)) {
                    try {
                        String numStyle = getNumStyle(dataFormat);
                        value = numFormat(numStyle, excelCell.getNumericCellValue());
                    } catch (Exception ignored) {

                    }
                }
            }
        }
        return value;
    }

    private static int getPdfCellBorderForMerged(Sheet sheet, int rowNum, int cellNum, int[] span) {
        int rowSpan = span[0];
        int colSpan = span[1];
        Cell leftTopCell = sheet.getRow(rowNum).getCell(cellNum);
        Cell rightTopCell = sheet.getRow(rowNum).getCell(cellNum + colSpan - 1);
        Cell leftBottomCell = sheet.getRow(rowNum + rowSpan - 1).getCell(cellNum);
        Cell rightBottomCell = sheet.getRow(rowNum + rowSpan - 1).getCell(cellNum + colSpan - 1);
        return getBorder(leftTopCell, rightTopCell, leftBottomCell, rightBottomCell);
    }

    private static int getBorder(Cell leftTopCell, Cell rightTopCell, Cell leftBottomCell, Cell rightBottomCell) {
        int border = 0;
        if (leftTopCell.getCellStyle().getBorderTopEnum().getCode() > 0) {
            border += 1;
        }
        if (rightBottomCell.getCellStyle().getBorderBottomEnum().getCode() > 0) {
            border += 4;
        }
        if (leftBottomCell.getCellStyle().getBorderLeftEnum().getCode() > 0) {
            border += 8;
        }
        if (rightTopCell.getCellStyle().getBorderRightEnum().getCode() > 0) {
            border += 2;
        }
        return border;
    }

    private static int getBorder(Cell excelCell) {
        return getBorder(excelCell, excelCell, excelCell, excelCell);
    }

    private static Workbook getWorkbook(String excelPath, InputStream in) throws IOException {
        Workbook workbook;
        if (excelPath.endsWith(XLSX) || excelPath.endsWith(XLSM)) {
            workbook = new XSSFWorkbook(in);
        } else {
            workbook = new HSSFWorkbook(in);
        }
        return workbook;
    }
}

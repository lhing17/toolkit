package cn.gsein.toolkit.excel.util;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Objects;

/**
 * 将EXCEL转为PDF的工具类
 *
 * @author G. Seinfeld
 * @since 2020-05-11
 */
public final class ExcelToPdfItextUtil {
    private ExcelToPdfItextUtil() {
    }

    private static final String XLS = ".xls";
    private static final String XLSX = ".xlsx";
    private static final String XLSM = ".xlsm";
    private static final String PERCENT = "%";


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

    /**
     * excel垂直对齐方式映射到pdf对齐方式
     */
    private static int getVerticalAlignment(VerticalAlignment align) {
        switch (align) {
            case BOTTOM:
                return com.itextpdf.text.Element.ALIGN_BOTTOM;
            case TOP:
                return com.itextpdf.text.Element.ALIGN_TOP;
            default:
                return com.itextpdf.text.Element.ALIGN_MIDDLE;
        }
    }


    /**
     * excel水平对齐方式映射到pdf水平对齐方式
     */
    private static int getHorizontalAlignment(HorizontalAlignment align) {
        switch (align) {
            case RIGHT:
                return com.itextpdf.text.Element.ALIGN_RIGHT;
            case LEFT:
                return com.itextpdf.text.Element.ALIGN_LEFT;
            default:
                return com.itextpdf.text.Element.ALIGN_CENTER;
        }
    }


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

    public static void excelToPdf(String excelPath, String pdfPath) throws Exception {
        InputStream in = new FileInputStream(excelPath);

        // 获取工作簿
        Workbook workbook = getWorkbook(excelPath, in);

        ByteArrayOutputStream stream = new ByteArrayOutputStream();

        int sheetCount = workbook.getNumberOfSheets();
        //此处根据excel大小设置pdf纸张大小
        Document document = new Document(PageSize.A4);
        PdfWriter.getInstance(document, stream);
        //设置页边距
        document.setMargins(0, 0, 15, 15);
        document.open();

        for (int i = 0; i < sheetCount; i++) {
            // 获取第一张工作表
            Sheet sheet = workbook.getSheetAt(i);

            float[] widths = getColWidth(sheet);
            PdfPTable table = new PdfPTable(widths);
            table.setWidthPercentage(90);
            int colCount = widths.length;
            //设置基本字体
            URL resource = ExcelToPdfItextUtil.class.getClassLoader().getResource("font/STXIHEI.TTF");
            BaseFont baseFont = BaseFont.createFont(Objects.requireNonNull(resource).toString(), BaseFont.IDENTITY_H,
                    BaseFont.EMBEDDED);

            for (int rowNum = sheet.getFirstRowNum(); rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = row.getFirstCellNum(); (cellNum < row.getLastCellNum() || cellNum < colCount) && cellNum > -1; cellNum++) {
                        if (cellNum >= row.getPhysicalNumberOfCells()) {
                            PdfPCell pCell = new PdfPCell(new Phrase(""));
                            pCell.setBorder(0);
                            table.addCell(pCell);
                            continue;
                        }
                        Cell excelCell = row.getCell(cellNum);
                        // 获取单元格的值
                        String value = getCellValue(excelCell);

                        //设置单元格字体
                        org.apache.poi.ss.usermodel.Font excelFont = getExcelFont(workbook, excelCell, excelPath);
                        Font pdFont = new Font(baseFont, excelFont.getFontHeightInPoints(),
                                excelFont.getBold() ? Font.BOLD : Font.NORMAL, BaseColor.BLACK);

                        PdfPCell pCell = new PdfPCell(new Phrase(value, pdFont));
                        List<PicturesInfo> infos = PoiExtend.getAllPictureInfos(sheet, rowNum, rowNum, cellNum, cellNum, false);
                        if (!infos.isEmpty()) {
                            PicturesInfo info = infos.get(0);
                            pCell = new PdfPCell(Image.getInstance(infos.get(0).getPictureData()), true);
                            pCell.setRowspan(info.getMaxRow() - info.getMinRow() + 1);
                            pCell.setColspan(info.getMaxCol() - info.getMinCol() + 1);
                            System.out.println("最大行：" + info.getMaxRow() + "最小行：" + info.getMinRow() + "最大列:" + info.getMaxCol() + "最小列：" + info.getMinCol());
                        }

                        setPdfCellBorder(Objects.requireNonNull(excelCell), pCell);
                        pCell.setHorizontalAlignment(getHorizontalAlignment(excelCell.getCellStyle().getAlignmentEnum()));
                        pCell.setVerticalAlignment(getVerticalAlignment(excelCell.getCellStyle().getVerticalAlignmentEnum()));

                        pCell.setMinimumHeight(row.getHeightInPoints());
                        if (isMergedRegion(sheet, rowNum, cellNum)) {
                            int[] span = getMergedSpan(sheet, rowNum, cellNum);
                            //忽略合并过的单元格
                            if (span[0] == 1 && span[1] == 1) {
                                continue;
                            }
                            pCell.setRowspan(span[0]);
                            pCell.setColspan(span[1]);
                            setPdfCellBorderForMerged(sheet, rowNum, cellNum, span, pCell);
                            //合并过的列直接跳过
                            cellNum = cellNum + span[1] - 1;
                        }

                        table.addCell(pCell);

                    }
                } else {
                    PdfPCell pCell = new PdfPCell(new Phrase(""));
                    pCell.setBorder(0);
                    pCell.setMinimumHeight(13);
                    table.addCell(pCell);
                }
            }
            document.add(table);
            document.newPage();
        }
        document.close();

        writeToFile(pdfPath, stream);

    }

    private static void writeToFile(String pdfPath, ByteArrayOutputStream stream) throws IOException {
        byte[] pdfByte = stream.toByteArray();
        stream.flush();
        stream.reset();
        stream.close();

        FileOutputStream outputStream = new FileOutputStream(pdfPath);
        outputStream.write(pdfByte);
        outputStream.flush();
        outputStream.close();
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

    private static void setPdfCellBorderForMerged(Sheet sheet, int rowNum, int cellNum, int[] span, PdfPCell pCell) {
        int rowSpan = span[0];
        int colSpan = span[1];
        Cell leftTopCell = sheet.getRow(rowNum).getCell(cellNum);
        Cell rightTopCell = sheet.getRow(rowNum).getCell(cellNum + colSpan - 1);
        Cell leftBottomCell = sheet.getRow(rowNum + rowSpan - 1).getCell(cellNum);
        Cell rightBottomCell = sheet.getRow(rowNum + rowSpan - 1).getCell(cellNum + colSpan - 1);
        doSetBorders(pCell, leftTopCell, rightTopCell, leftBottomCell, rightBottomCell);
    }

    private static void doSetBorders(PdfPCell pCell, Cell leftTopCell, Cell rightTopCell, Cell leftBottomCell, Cell rightBottomCell) {
        int border = 0;
        if (leftTopCell.getCellStyle().getBorderTopEnum().getCode() > 0) {
            border += PdfPCell.TOP;
        }
        if (rightBottomCell.getCellStyle().getBorderBottomEnum().getCode() > 0) {
            border += PdfPCell.BOTTOM;
        }
        if (leftBottomCell.getCellStyle().getBorderLeftEnum().getCode() > 0) {
            border += PdfPCell.LEFT;
        }
        if (rightTopCell.getCellStyle().getBorderRightEnum().getCode() > 0) {
            border += PdfPCell.RIGHT;
        }
        pCell.setBorder(border);
    }

    private static void setPdfCellBorder(Cell excelCell, PdfPCell pCell) {
        doSetBorders(pCell, excelCell, excelCell, excelCell, excelCell);
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

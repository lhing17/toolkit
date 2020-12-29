package cn.gsein.toolkit.excel.util;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.io.*;
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
     * 内容所占宽度比例
     */
    private static final float CONTENT_WIDTH_PERCENT = 0.94f;


    /**
     * 获取字体
     */
    private static Font getExcelFont(Workbook workbook, Cell cell) {
        int fontIndex = cell.getCellStyle().getFontIndexAsInt();
        return workbook.getFontAt(fontIndex);
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
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row != null && maxCol < row.getLastCellNum() + 1) {
                maxCol = row.getLastCellNum() + 1;
                rowNum = r;
            }
        }
        return rowNum;
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
        // 获取最多列的行号
        int rowNum = getMaxColRowNum(sheet);
        Row row = sheet.getRow(rowNum);
        int cellCount = row.getLastCellNum() + 1;
        int[] colWidths = new int[cellCount];
        int sum = 0;

        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                colWidths[i] = sheet.getColumnWidth(i);
                sum += sheet.getColumnWidth(i);
            }
        }

        // 宽度百分比
        float[] colWidthPer = new float[cellCount];
        for (int i = row.getFirstCellNum(); i < cellCount; i++) {
            colWidthPer[i] = (float) colWidths[i] / sum * 100;
        }
        return colWidthPer;
    }

    private static float[] getRowHeight(Sheet sheet) {
        int rowNum = sheet.getLastRowNum() + 1;

        float[] rowHeight = new float[rowNum];
        float sum = 0;
        for (int i = sheet.getFirstRowNum(); i < rowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                rowHeight[i] = row.getHeightInPoints();
                sum += rowHeight[i];
            }
        }

        // 高度百分比
        float[] rowHeightPer = new float[rowNum];
        for (int i = sheet.getFirstRowNum(); i < rowNum; i++) {
            rowHeightPer[i] = rowHeight[i] / sum * 100;
        }
        return rowHeightPer;
    }

    public static void excelToPdf(InputStream excelInput, OutputStream pdfOutput) throws Exception {

        // 获取工作簿
        Workbook workbook = WorkbookFactory.create(excelInput);

        // 获取工作表的数量
        int sheetCount = workbook.getNumberOfSheets();

        //新建PDF文档
        PDDocument document = new PDDocument();
        for (int i = 0; i < sheetCount; i++) {

            // 获取第i张工作表
            Sheet sheet = workbook.getSheetAt(i);
            // 不处理没有内容的sheet
            if (sheet.getPhysicalNumberOfRows() == 0) {
                continue;
            }

            PDPageContentStream stream = createPageAndContentStream(document);

            // 从classpath中读取字体文件
            PDType0Font font = loadFontFromClasspath(document);

            // 设置默认字体、字号、颜色
            setDefaultFontAndColor(stream, font);

            // 获取每列的宽度（从excel宽度转为pdf宽度）
            float[] excelWidths = getColWidth(sheet);
            float[] widths = handleWidths(excelWidths);

            // 获取每行的高度（从excel高度转为pdf高度）
            float[] excelHeights = getRowHeight(sheet);
            float[] heights = handleHeights(excelHeights);

            // 先遍历行，再遍历列，对每一个单元格进行处理
            int colCount = widths.length;
            for (int rowNum = sheet.getFirstRowNum(); rowNum < sheet.getLastRowNum() + 1; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = row.getFirstCellNum(); (cellNum < row.getLastCellNum() + 1 || cellNum < colCount) && cellNum > -1; cellNum++) {
                        Cell excelCell = row.getCell(cellNum);
                        // 获取单元格的值
                        String value = getCellValue(excelCell);

                        //设置单元格字体
                        org.apache.poi.ss.usermodel.Font excelFont = getExcelFont(workbook, excelCell);

                        // 单独处理图片
                        List<PicturesInfo> infos = PoiExtend.getAllPictureInfos(sheet, rowNum, rowNum, cellNum, cellNum, false);
                        drawImage(document, stream, widths, heights, rowNum, cellNum, infos);

                        // 判断是否有边框
                        int border = getBorder(Objects.requireNonNull(excelCell));

                        // 判断是否为合并单元格，分别处理
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

                            // 非合并的单元格直接绘制
                            PdfUtil.drawRect(stream, widths, heights, rowNum, cellNum, 1, 1, border);
                            PdfUtil.drawString(font, stream, excelFont, widths, heights, rowNum, cellNum, 1, 1, value);
                        }
                    }
                }
            }
            stream.stroke();
            stream.close();
        }

        document.save(pdfOutput);
        document.close();

    }

    private static void setDefaultFontAndColor(PDPageContentStream stream, PDType0Font font) throws IOException {
        stream.setFont(font, 14);
        stream.setStrokingColor(Color.BLACK);
    }

    private static PDType0Font loadFontFromClasspath(PDDocument document) throws IOException {
        InputStream inputStream = ExcelToPdfPdfBoxUtil.class.getClassLoader().getResourceAsStream("font/STXIHEI.TTF");
        return PDType0Font.load(document, inputStream, true);
    }

    private static PDPageContentStream createPageAndContentStream(PDDocument document) throws IOException {
        // 设置页面大小为A4纸
        PDPage page = new PDPage(PDRectangle.A4);

        // 将页面加入文档中
        document.addPage(page);

        // 创建处理pdf内容的流
        return new PDPageContentStream(document, page);
    }

    private static void drawImage(PDDocument document, PDPageContentStream stream, float[] widths, float[] heights, int rowNum, int cellNum, List<PicturesInfo> infos) throws IOException {
        if (!infos.isEmpty()) {
            // 目前只处理第一张图片
            PicturesInfo info = infos.get(0);

            // 左上角坐标
            float x = 29.76f;
            float y = PAGE_HEIGHT - 15f;

            // 图片起始X坐标
            for (int j = 0; j < cellNum; j++) {
                x += widths[j];
            }

            // 图片宽度
            float width = 0;
            for (int j = cellNum; j < cellNum + info.getMaxCol() - info.getMinCol() + 1; j++) {
                width += widths[j];
            }

            // 图片起始Y坐标
            for (int j = 0; j < rowNum; j++) {
                y -= heights[j];
            }

            // 图片高度
            float height = 0;
            for (int j = rowNum; j < rowNum + info.getMaxRow() - info.getMinRow() + 1; j++) {
                height += heights[j];
            }

            // 绘制图片
            stream.drawImage(PDImageXObject.createFromByteArray(document, info.getPictureData(), null), x, y - height, width, height);
        }
    }

    public static void excelToPdf(String excelPath, String pdfPath) throws Exception {
        InputStream in = new FileInputStream(excelPath);
        OutputStream out = new FileOutputStream(pdfPath);
        excelToPdf(in, out);
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
            widths[i] = excelWidths[i] / 100 * PAGE_WIDTH * CONTENT_WIDTH_PERCENT;
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

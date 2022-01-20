package cn.gsein.toolkit.excel.pdf;

import cn.gsein.toolkit.excel.pdf.util.ExcelUtil;

import cn.gsein.toolkit.excel.pdf.util.PdfUtil;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.*;
import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class PdfboxConverter extends Converter {

    private static final String XLS = ".xls";
    private static final String XLSX = ".xlsx";
    private static final String XLSM = ".xlsm";
    private static final String PERCENT = "%";

    private static final float POINTS_PER_MM = 2.8346457f;
    private static final float PAGE_WIDTH = 210 * POINTS_PER_MM;
    private static final float PAGE_HEIGHT = 297 * POINTS_PER_MM;

    public PdfboxConverter(Configuration configuration) {
        super(configuration);
    }

    @Override
    public void convert(InputStream excelInput, OutputStream pdfOutput) throws IOException {
        // 根据输入流创建工作簿对象，会自动识别xls和xlsx
        Workbook workbook = WorkbookFactory.create(excelInput);
        PDDocument document = new PDDocument();

        // 处理每个sheet
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            writeExcelSheetToPdfPages(workbook.getSheetAt(i), document);
        }

        document.save(pdfOutput);
        document.close();
    }

    private PDRectangle getRect() {
        PDRectangle rect = new PDRectangle();

        switch (pageSize) {
            case A0:
                rect = PDRectangle.A0;
                break;
            case A1:
                rect = PDRectangle.A1;
                break;
            case A2:
                rect = PDRectangle.A2;
                break;
            case A3:
                rect = PDRectangle.A3;
                break;
            case A4:
                rect = PDRectangle.A4;
                break;
            case A5:
                rect = PDRectangle.A5;
                break;
            case A6:
                rect = PDRectangle.A6;
                break;
        }
        if (mode == Mode.PORTRAIT) {
            return rect;
        } else {
            // 如果是横版，将长和宽交换一下
            return new PDRectangle(rect.getHeight(), rect.getWidth());
        }

    }

    public void writeExcelSheetToPdfPages(Sheet sheet, PDDocument document) throws IOException {
        int maxCountIndex = ExcelUtil.getRowNumOfMaxColumnCount(sheet);
        Row row0 = sheet.getRow(maxCountIndex);
        if (row0 == null) {
            return;
        }

        int[] widths = ExcelUtil.getColumnWidths(sheet);
        int[] heights = ExcelUtil.getRowHeights(sheet);
        PDRectangle rect = getRect();

        List<PDFPage> pdfPageList = getPdfPages(sheet, widths, heights, rect, row0);

        // 从classpath中读取字体文件
        PDType0Font font = loadFontFromClasspath(document);
        for (PDFPage pdfPage : pdfPageList) {
            PDPage page = new PDPage(rect);
            document.addPage(page);

            // 创建处理pdf内容的流
            PDPageContentStream stream = new PDPageContentStream(document, page);

            // 设置默认字体、字号、颜色
            setDefaultFontAndColor(stream, font);

            for (int rowNum = pdfPage.getStartRowNum(); rowNum < pdfPage.getStartRowNum() + pdfPage.getRowCount(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    for (int cellNum = pdfPage.getStartColumnNum(); cellNum < pdfPage.getStartColumnNum() + pdfPage.getColumnCount(); cellNum++) {
                        Cell excelCell = row.getCell(cellNum);
                        // 获取单元格的值
                        String value = getCellValue(excelCell);

                        //设置单元格字体
                        org.apache.poi.ss.usermodel.Font excelFont = getExcelFont(sheet.getWorkbook(), excelCell);

//                        // 单独处理图片
//                        List<PicturesInfo> infos = PoiExtend.getAllPictureInfos(sheet, rowNum, rowNum, cellNum, cellNum, false);
//                        drawImage(document, stream, widths, heights, rowNum, cellNum, infos);

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

                            PdfUtil.drawRect(stream, widths, heights, pdfPage.getStartRowNum(), rowNum, pdfPage.getStartColumnNum(), cellNum, span[0], span[1], border);
                            PdfUtil.drawString(font, stream, excelFont, widths, heights, pdfPage.getStartRowNum(), rowNum, pdfPage.getStartColumnNum(), cellNum, span[0], span[1], value);
                            cellNum = cellNum + span[1] - 1;
                        } else {
                            // 非合并的单元格直接绘制
                            PdfUtil.drawRect(stream, widths, heights, pdfPage.getStartRowNum(), rowNum, pdfPage.getStartColumnNum(), cellNum, 1, 1, border);
                            PdfUtil.drawString(font, stream, excelFont, widths, heights, pdfPage.getStartRowNum(), rowNum, pdfPage.getStartColumnNum(), cellNum, 1, 1, value);
                        }
                    }
                }
            }

            stream.stroke();
            stream.close();
        }


    }

    private List<PDFPage> getPdfPages(Sheet sheet, int[] widths, int[] heights, PDRectangle rect, Row row0) {
        int firstRowNum = sheet.getFirstRowNum();

        short firstCellNum = row0.getFirstCellNum();


        float accumulateWidth = 0;
        float accumulateHeight = 0;
        boolean createPdf = false;


        float pageWidth = rect.getWidth() * 0.94f;
        float pageHeight = rect.getHeight() - 30;
        int pageNum = 0;

        List<PDFPage> pdfPageList = new ArrayList<>();
        int startRowNum = firstRowNum;
        int startColumnNum = firstCellNum;

        for (int i = firstRowNum; i < firstRowNum + heights.length; i++) {
            if (createPdf) {
                for (int j = firstCellNum; j < firstCellNum + widths.length; j++) {
                    accumulateWidth += widths[j - firstCellNum] * 1.0f / 256 * 8;
                    if (accumulateWidth > pageWidth) {
                        PDFPage pdfPage = new PDFPage();
                        pdfPage.setPageNum(pageNum++);
                        pdfPage.setStartColumnNum(startColumnNum);
                        pdfPage.setColumnCount(j - startColumnNum);
                        pdfPage.setStartRowNum(startRowNum);
                        pdfPage.setRowCount(i - startRowNum);
                        pdfPageList.add(pdfPage);
                        accumulateWidth = widths[j - firstCellNum] * 1.0f / 256 * 8;
                        startColumnNum = j;
                    }
                }
                createPdf = false;
                startRowNum = i;
                startColumnNum = 0;
                accumulateWidth = 0;
            }
            accumulateHeight += heights[i - firstRowNum] * 1.5f / 256 * 8;
            if (accumulateHeight > pageHeight) {
                createPdf = true;
                accumulateHeight = heights[i - firstRowNum] * 1.5f / 256 * 8;
            }
        }
        return pdfPageList;
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

    /**
     * 获取字体
     */
    private static Font getExcelFont(Workbook workbook, Cell cell) {
        int fontIndex = cell.getCellStyle().getFontIndexAsInt();
        return workbook.getFontAt(fontIndex);
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
     * 格式化数字
     */
    private static String numFormat(String pattern, double num) {
        DecimalFormat format = new DecimalFormat(pattern);
        return format.format(num);
    }


    private static PDType0Font loadFontFromClasspath(PDDocument document) throws IOException {
        InputStream inputStream = PdfboxConverter.class.getClassLoader().getResourceAsStream("font/STXIHEI.TTF");
        return PDType0Font.load(document, inputStream, true);
    }

    private static void setDefaultFontAndColor(PDPageContentStream stream, PDType0Font font) throws IOException {
        stream.setFont(font, 14);
        stream.setStrokingColor(Color.BLACK);
    }
}

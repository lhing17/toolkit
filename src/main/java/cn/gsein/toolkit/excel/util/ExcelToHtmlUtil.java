package cn.gsein.toolkit.excel.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * EXCEL转为pdf的工具类
 *
 * @author G. Seinfeld
 * @since 2020-05-11
 */
public final class ExcelToHtmlUtil {
    private ExcelToHtmlUtil() {
    }

    /**
     * @ClassName: PoiExcelToHtmlUtil
     * @Description: TODO(poi转excel为html)
     */
    private static Map<String, Object>[] map;

    /**
     * 程序入口方法（读取指定位置的excel，将其转换成html形式的字符串，并保存成同名的html文件在相同的目录下，默认带样式）
     *
     * @param sourcePath 文件路径
     * @return <table>...</table> 字符串
     */
    public static String excelWriteToHtml(String sourcePath, String uploadPath) {
        File sourceFile = new File(sourcePath);
        try {
            InputStream fis = new FileInputStream(sourceFile);
            return readExcelToHtml(fis, true, uploadPath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 程序入口方法（将指定路径的excel文件读取成字符串）
     *
     * @param filePath    文件的路径
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public static String readExcelToHtml(String filePath, boolean isWithStyle, String uploadPath) {
        InputStream is = null;
        String htmlExcel = null;
        try {
            File sourcefile = new File(filePath);
            is = new FileInputStream(sourcefile);
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb, isWithStyle, uploadPath);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }

    /**
     * 程序入口方法（将指定路径的excel文件读取成字符串）
     *
     * @param is          excel转换成的输入流
     * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
     * @return <table>...</table> 字符串
     */
    public static String readExcelToHtml(InputStream is, boolean isWithStyle, String uploadPath) {
        String htmlExcel = null;
        try {
            Workbook wb = WorkbookFactory.create(is);
            htmlExcel = readWorkbook(wb, isWithStyle, uploadPath);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return htmlExcel;
    }

    /**
     * 根据excel的版本分配不同的读取方法进行处理
     */
    private static String readWorkbook(Workbook wb, boolean isWithStyle, String uploadPath) {
        String htmlExcel = "";
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            htmlExcel = getExcelInfo(xWb, isWithStyle, uploadPath);
        } else if (wb instanceof HSSFWorkbook) {
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            htmlExcel = getExcelInfo(hWb, isWithStyle, uploadPath);
        }
        return htmlExcel;
    }

    /**
     * 读取excel成string
     */
    public static String getExcelInfo(Workbook wb, boolean isWithStyle, String uploadPath) {

        StringBuilder sb = new StringBuilder();
        int sheetCount = wb.getNumberOfSheets();
        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = wb.getSheetAt(i);//获取第一个Sheet的内容
            // map等待存储excel图片
            Map<String, PictureData> sheetIndexPicMap = getSheetPictures(i, sheet, wb);
            //临时保存位置，正式环境根据部署环境存放其他位置
            try {
                if (sheetIndexPicMap != null) {
                    printImg(sheetIndexPicMap, uploadPath);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            //读取excel拼装html
            int lastRowNum = sheet.getLastRowNum();
            map = getRowSpanColSpanMap(sheet);
            createTableHtml(wb, isWithStyle, sb, sheet, sheetIndexPicMap, lastRowNum, i);
        }

        return sb.toString();
    }

    private static void createTableHtml(Workbook wb, boolean isWithStyle, StringBuilder sb, Sheet sheet, Map<String, PictureData> sheetIndexPicMap, int lastRowNum, int i) {
        sb.append("<table style='border-collapse:collapse;width:80%;'>");
        Row row;
        Cell cell;

        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            if (rowNum > 1000) {
                break;
            }
            row = sheet.getRow(rowNum);

            int lastColNum = getColsOfTable(sheet)[0];
            int rowHeight = getColsOfTable(sheet)[1];

            if (row != null) {
                lastColNum = row.getLastCellNum();
                rowHeight = row.getHeight();
            }

            if (row == null) {
                sb.append("<tr><td >  </td></tr>");
                continue;
            } else if (row.getZeroHeight()) {
                continue;
            } else if (0 == rowHeight) {
                continue;     //针对jxl的隐藏行（此类隐藏行只是把高度设置为0，单getZeroHeight无法识别）
            } else if (rowNum == lastRowNum && lastColNum == 1) {
                continue;
            }
            sb.append("<tr>");

            for (int colNum = 0; colNum < lastColNum; colNum++) {
                if (sheet.isColumnHidden(colNum)) {
                    continue;
                }
                String imageRowNum = i + "_" + rowNum + "_" + colNum;
                String imageHtml = "";
                cell = row.getCell(colNum);
                if ((sheetIndexPicMap == null || !sheetIndexPicMap.containsKey(imageRowNum)) && cell == null) {    //特殊情况 空白的单元格会返回null+//判断该单元格是否包含图片，为空时也可能包含图片
                    sb.append("<td>  </td>");
                    continue;
                }
                if (sheetIndexPicMap != null && sheetIndexPicMap.containsKey(imageRowNum)) {
                    //待修改路径
                    String imagePath = "/upload/pic" + imageRowNum + ".jpeg";

                    imageHtml = "<img src='" + imagePath + "' style='height:" + rowHeight / 20 + "px;'>";
                }
                String stringValue = getCellValue(cell);
                if (map[0].containsKey(rowNum + "," + colNum)) {
                    String pointString = (String) map[0].get(rowNum + "," + colNum);
                    int bottomeRow = Integer.parseInt(pointString.split(",")[0]);
                    int bottomeCol = Integer.parseInt(pointString.split(",")[1]);
                    int rowSpan = bottomeRow - rowNum + 1;
                    int colSpan = bottomeCol - colNum + 1;
                    if (map[2].containsKey(rowNum + "," + colNum)) {
                        rowSpan = rowSpan - (Integer) map[2].get(rowNum + "," + colNum);
                    }
                    sb.append("<td rowspan= '").append(rowSpan).append("' colspan= '").append(colSpan).append("' ");
                    if (map.length > 3 && map[3].containsKey(rowNum + "," + colNum)) {
                        //此类数据首行被隐藏，value为空，需使用其他方式获取值
                        stringValue = getMergedRegionValue(sheet, rowNum, colNum);
                    }
                } else if (map[1].containsKey(rowNum + "," + colNum)) {
                    map[1].remove(rowNum + "," + colNum);
                    continue;
                } else {
                    sb.append("<td ");
                }

                //判断是否需要样式
                if (isWithStyle) {
                    dealExcelStyle(wb, sheet, cell, sb);//处理单元格样式
                }

                sb.append(">");
                if (sheetIndexPicMap != null && sheetIndexPicMap.containsKey(imageRowNum)) {
                    sb.append(imageHtml);
                }
                if (stringValue == null || "".equals(stringValue.trim())) {
                    sb.append("   ");
                } else {
                    // 将ascii码为160的空格转换为html下的空格（ ）
                    sb.append(stringValue.replace(String.valueOf((char) 160), " "));
                }
                sb.append("</td>");
            }
            sb.append("</tr>");
        }
        sb.append("</table>");
        sb.append("<br>");
    }

    /**
     * 分析excel表格，记录合并单元格相关的参数，用于之后html页面元素的合并操作
     *
     * @param sheet
     * @return
     */
    private static Map<String, Object>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<>();    //保存合并单元格的对应起始和截止单元格
        Map<String, String> map1 = new HashMap<>();    //保存被合并的那些单元格
        Map<String, Integer> map2 = new HashMap<>();    //记录被隐藏的单元格个数
        Map<String, String> map3 = new HashMap<>();    //记录合并了单元格，但是合并的首行被隐藏的情况
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range;
        Row row;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            /*
             * 此类数据为合并了单元格的数据
             * 1.处理隐藏（只处理行隐藏，列隐藏poi已经处理）
             */
            if (topRow != bottomRow) {
                int zeroRoleNum = 0;
                int tempRow = topRow;
                for (int j = topRow; j <= bottomRow; j++) {
                    row = sheet.getRow(j);
                    if (row.getZeroHeight() || row.getHeight() == 0) {
                        if (j == tempRow) {
                            //首行就进行隐藏，将rowTop向后移
                            tempRow++;
                            continue;//由于top下移，后面计算rowSpan时会扣除移走的列，所以不必增加zeroRoleNum;
                        }
                        zeroRoleNum++;
                    }
                }
                if (tempRow != topRow) {
                    map3.put(tempRow + "," + topCol, topRow + "," + topCol);
                    topRow = tempRow;
                }
                if (zeroRoleNum != 0) {
                    map2.put(topRow + "," + topCol, zeroRoleNum);
                }
            }
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, topRow + "," + topCol);
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = {map0, map1, map2, map3};
        System.err.println(map0);
        return map;
    }


    /**
     * 获取合并单元格的值
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getCellValue(fCell);
                }
            }
        }
        return null;
    }

    /**
     * 获取表格单元格Cell内容
     */
    private static String getCellValue(Cell cell) {
        String result;
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if ("General".equals(temp)) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理单元格样式
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuilder sb) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            HorizontalAlignment alignment = cellStyle.getAlignmentEnum();
            sb.append("align='").append(convertAlignToHtml(alignment)).append("' ");//单元格内容的水平对齐方式
            VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignmentEnum();
            sb.append("valign='").append(convertVerticalAlignToHtml(verticalAlignment)).append("' ");//单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {
                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                boolean bold = xf.getBold();
                sb.append("style='");
                if (bold) {
                    // 字体加粗
                    sb.append("font-weight: bold;");
                }
                // 字体大小
                sb.append("font-size: ").append(xf.getFontHeight() / 2).append("%;");

                int topRow = cell.getRowIndex(), topColumn = cell.getColumnIndex();
                int columnWidth;
                if (map[0].containsKey(topRow + "," + topColumn)) {
                    //该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                    String value = (String) map[0].get(topRow + "," + topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);

                    if (topColumn != bottomColumn) {
                        //合并列，需要计算相应宽度
                        columnWidth = 0;
                        for (int i = topColumn; i <= bottomColumn; i++) {
                            columnWidth += sheet.getColumnWidth(i);
                        }
                    } else {
                        columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                    }
                } else {
                    columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                }
                sb.append("width:").append(columnWidth / 256 * xf.getFontHeight() / 20).append("pt;");

                // 字体颜色
                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc.toString())) {
                    sb.append("color:#").append(xc.getARGBHex().substring(2)).append(";");
                }

                // 背景颜色
                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor.toString())) {
                    sb.append("background-color:#").append(bgColor.getARGBHex().substring(2)).append(";");
                }

                // 边框
                sb.append("border:solid #000000 1px;");
            } else if (wb instanceof HSSFWorkbook) {
                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                boolean bold = hf.getBold();
                short fontColor = hf.getColor();
                sb.append("style='");

                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                if (bold) {
                    sb.append("font-weight: bold;"); // 字体加粗
                }
                sb.append("font-size: ").append(hf.getFontHeight() / 2).append("%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:").append(fontColorStr).append(";"); // 字体颜色
                }

                int topRow = cell.getRowIndex(), topColumn = cell.getColumnIndex();
                int columnWidth;
                //该单元格为合并单元格，宽度需要获取所有单元格宽度后合并
                if (map[0].containsKey(topRow + "," + topColumn)) {
                    String value = (String) map[0].get(topRow + "," + topColumn);
                    String[] ary = value.split(",");
                    int bottomColumn = Integer.parseInt(ary[1]);

                    if (topColumn != bottomColumn) {//合并列，需要计算相应宽度
                        columnWidth = 0;
                        for (int i = topColumn; i <= bottomColumn; i++) {
                            columnWidth += sheet.getColumnWidth(i);
                        }
                    } else {
                        columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                    }
                } else {
                    columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                }
                sb.append("width:").append(columnWidth / 256 * hf.getFontHeight() / 20).append("pt;");

                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:").append(bgColorStr).append(";");        // 背景颜色
                }
                sb.append("border:solid #000000 1px;");
            }
            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     */
    private static String convertAlignToHtml(HorizontalAlignment alignment) {
        String align = "left";
        switch (alignment) {
            case LEFT:
                align = "left";
                break;
            case CENTER:
                align = "center";
                break;
            case RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式
     */
    private static String convertVerticalAlignToHtml(VerticalAlignment verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case BOTTOM:
                valign = "bottom";
                break;
            case CENTER:
                valign = "center";
                break;
            case TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {
        StringBuilder sb = new StringBuilder();
        if (hc != null) {
            if (HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex() == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }


    /**
     * 获取Excel图片公共方法
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictures(int sheetNum, Sheet sheet, Workbook workbook) {
        if (workbook instanceof HSSFWorkbook) {
            return getSheetPictrues03(sheetNum, (HSSFSheet) sheet, (HSSFWorkbook) workbook);
        } else if (workbook instanceof XSSFWorkbook) {
            return getSheetPictrues07(sheetNum, (XSSFSheet) sheet);
        } else {
            return null;
        }
    }

    /**
     * 获取Excel2003图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictrues03(int sheetNum, HSSFSheet sheet, HSSFWorkbook workbook) {

        Map<String, PictureData> sheetIndexPicMap = new HashMap<>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.size() != 0) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                shape.getLineWidth();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = sheetNum + "_"
                            + anchor.getRow1() + "_"
                            + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    /**
     * 获取Excel2007图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictrues07(int sheetNum, XSSFSheet sheet) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<>();

        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = sheetNum + "_"
                            + ctMarker.getRow() + "_" + ctMarker.getCol();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }

        return sheetIndexPicMap;
    }

    public static void printImg(Map<String, PictureData> map, String uploadPath) throws IOException {

        for (Map.Entry<String, PictureData> entry : map.entrySet()) {
            // 获取图片流
            PictureData pic = entry.getValue();
            // 获取图片索引
            String picName = entry.getKey();
            // 获取图片格式
            String ext = pic.suggestFileExtension();
            byte[] data = pic.getData();

            File file = new File(uploadPath);
            if (!file.exists()) {
                file.mkdirs();
            }

            FileOutputStream out = new FileOutputStream(uploadPath + "pic" + picName + "." + ext);
            out.write(data);
            out.flush();
            out.close();
        }
    }


    private static int[] getColsOfTable(Sheet sheet) {
        int[] data = {0, 0};
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null) {
                data[0] = sheet.getRow(i).getLastCellNum();
                data[1] = sheet.getRow(i).getHeight();
            }
        }
        return data;
    }
}

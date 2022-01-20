package cn.gsein.toolkit.excel.pdf;

import lombok.Data;

@Data
public class PDFPage {

    /**
     * 第几页，从0开始
     */
    private int pageNum;
    /**
     * 开始列数
     */
    private int startColumnNum;
    /**
     * 当前页列数
     */
    private int columnCount;
    /**
     * 开始行数
     */
    private int startRowNum;
    /**
     * 当前页行数
     */
    private int rowCount;
}

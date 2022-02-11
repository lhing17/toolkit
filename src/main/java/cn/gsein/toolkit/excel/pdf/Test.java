package cn.gsein.toolkit.excel.pdf;

import cn.gsein.toolkit.excel.pdf.util.ExcelUtil;
import cn.gsein.toolkit.excel.util.ExcelToPdfPdfBoxUtil;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

public class Test {

    public static void main(String[] args) throws IOException {
        Path excelPath = Paths.get("/Users", "lianghao", "Documents", "a.xlsx");
        System.out.println(excelPath);
        try (InputStream excelInput = Files.newInputStream(excelPath);
             OutputStream pdfOutput = Files.newOutputStream(Paths.get("/Users", "lianghao", "Documents", "a.pdf"))) {

            Converter converter = new PdfboxConverter(Configuration.builder().build());
            converter.convert(excelInput, pdfOutput);
        }

    }
}

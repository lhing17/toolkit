package cn.gsein.toolkit.excel.pdf;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public abstract class Converter {

    protected final PageSize pageSize;

    protected final Mode mode;

    public Converter(Configuration configuration) {
        this.pageSize = configuration.pageSize;
        this.mode = configuration.mode;
    }

    public abstract void convert(InputStream excelInput, OutputStream pdfOutput) throws IOException;
}

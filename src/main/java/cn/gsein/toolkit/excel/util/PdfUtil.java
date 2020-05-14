package cn.gsein.toolkit.excel.util;

import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;
import org.apache.pdfbox.util.Matrix;
import org.apache.poi.ss.usermodel.Font;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * PDF工具类
 *
 * @author G. Seinfeld
 * @since 2020-05-14
 */
public final class PdfUtil {

    private static final int TOP = 1;
    private static final int RIGHT = 2;
    private static final int BOTTOM = 4;
    private static final int LEFT = 8;

    private static final float POINTS_PER_MM = 2.8346457f;
    private static final float PAGE_WIDTH = 210 * POINTS_PER_MM;
    private static final float PAGE_HEIGHT = 297 * POINTS_PER_MM;

    public static void drawRect(PDPageContentStream stream, float[] widths, float[] heights, int rowNum, int cellNum, int rowSpan, int colSpan, int border) {
        float x = PAGE_WIDTH * 0.03f;
        float y = PAGE_HEIGHT - 15f;
        for (int i = 0; i < cellNum; i++) {
            x += widths[i];
        }

        float width = 0;
        for (int i = cellNum; i < cellNum + colSpan; i++) {
            width += widths[i];
        }

        for (int i = 0; i < rowNum; i++) {
            y -= heights[i];
        }
        float height = 0;
        for (int i = rowNum; i < rowNum + rowSpan; i++) {
            height += heights[i];
        }

        try {
            stream.moveTo(x, y);
            if ((border & TOP) != 0) {
                stream.lineTo(x + width, y);
            }
            stream.moveTo(x + width, y);
            if ((border & RIGHT) != 0) {
                stream.lineTo(x + width, y - height);
            }
            stream.moveTo(x + width, y - height);
            if ((border & BOTTOM) != 0) {
                stream.lineTo(x, y - height);
            }
            stream.moveTo(x, y - height);
            if ((border & LEFT) != 0) {
                stream.lineTo(x, y);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void drawString(PDFont font, PDPageContentStream stream, Font excelFont, float[] widths, float[] heights, int rowNum, int cellNum, int rowSpan, int colSpan, String value) {
        float x = PAGE_WIDTH * 0.03f;
        float y = PAGE_HEIGHT - 15f;
        for (int i = 0; i < cellNum; i++) {
            x += widths[i];
        }

        float width = 0;
        for (int i = cellNum; i < cellNum + colSpan; i++) {
            width += widths[i];
        }

        for (int i = 0; i < rowNum; i++) {
            y -= heights[i];
        }
        float height = 0;
        for (int i = rowNum; i < rowNum + rowSpan; i++) {
            height += heights[i];
        }

        try {
            float fontSize = excelFont.getFontHeightInPoints();
            stream.setFont(font, fontSize);


            List<String> realLines = new ArrayList<>(Arrays.asList(value.split("[\n\r]")));
            List<String> lines = new ArrayList<>();
            for (String realLine : realLines) {
                float stringWidth = font.getStringWidth(realLine) / 1000f * fontSize;
                if (stringWidth > width - fontSize) {
                    List<String> fakeLines = splitRealLine(realLine, font, fontSize, width - fontSize);
                    lines.addAll(fakeLines);
                } else {
                    lines.add(realLine);
                }
            }

            int lineCount = lines.size();
            for (int i = 0; i < lineCount; i++) {
                String line = lines.get(i).trim();
                float stringWidth = font.getStringWidth(line) / 1000f * fontSize;
                float centeredX = x + 0.5f * (width - stringWidth);
                float centeredY = y - 0.5f * (height + fontSize) + (0.75f * (lineCount - 1) - 1.5f * i) * fontSize;
                if (excelFont.getBold()) {
                    stream.setRenderingMode(RenderingMode.STROKE);
                } else {
                    stream.setRenderingMode(RenderingMode.FILL);
                }
                stream.beginText();
                stream.setTextMatrix(new Matrix(1, 0, 0, 1, centeredX, centeredY));
                stream.showText(line);
                stream.endText();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> splitRealLine(String realLine, PDFont font, float fontSize, float width) throws IOException {
        List<String> fakeLines = new ArrayList<>();
        // 按空白字符拆分
        String[] words = realLine.split("\\s");
        StringBuilder builder = new StringBuilder();
        float nowWidth = 0;
        for (int i = 0; i < words.length; i++) {
            float wordWidth = font.getStringWidth(words[i]) / 1000f * fontSize;
            if (wordWidth > width) {
                // 单词本身超过一行
                if (builder.length() != 0) {
                    fakeLines.add(builder.toString());
                }
                fakeLines.addAll(furtherDealFakeLine(words[i]));
                builder = new StringBuilder();
                nowWidth = 0;
            } else if (i == words.length - 1) {
                // 最后一行
                builder.append(" ").append(words[i]);
                fakeLines.add(builder.toString());
            } else if (nowWidth + wordWidth > width) {
                fakeLines.add(builder.toString());
                builder = new StringBuilder(words[i]);
                // 下一行当前宽度
                nowWidth = wordWidth;
            } else {
                nowWidth += wordWidth;
                builder.append(" ").append(words[i]);
            }
        }
        return fakeLines;
    }

    private static List<String> furtherDealFakeLine(String word) {
        List<String> fakeLines = new ArrayList<>();
        if (word.contains("/")) {
            String[] split = word.split("/");
            for (int i = 0; i < split.length; i++) {
                if (i == 0) {
                    fakeLines.add(split[i]);
                } else {
                    fakeLines.add("/" + split[i]);
                }
            }
        } else {
            fakeLines.add(word);
        }
        return fakeLines;
    }
}

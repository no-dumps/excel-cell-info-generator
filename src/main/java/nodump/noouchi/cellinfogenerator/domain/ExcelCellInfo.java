package nodump.noouchi.cellinfogenerator.domain;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Arrays;

public class ExcelCellInfo {
    public String cellRange;
    public String value;
    public String alignment;
    public String borderBottom;
    public String borderLeft;
    public String borderRight;
    public String borderTop;
    public String bottomBorderColor;
    public String leftBorderColor;
    public String rightBorderColor;
    public String topBorderColor;
    public String dataFormat;
    public String fillBackgroundColor;
    public String fillForegroundColor;
    public String fillPattern;
    public Font font;
    public String hidden;
    public String indention;
    public String locked;
    public String rotation;
    public String verticalAlignment;

    public ExcelCellInfo(XSSFCell cell) {
        cellRange = cell.getReference();
        value = getCellValueToString(cell);
        alignment = cell.getCellStyle().getAlignment().name();
        borderBottom = cell.getCellStyle().getBorderBottom().name();
        borderLeft = cell.getCellStyle().getBorderLeft().name();
        borderRight = cell.getCellStyle().getBorderRight().name();
        borderTop = cell.getCellStyle().getBorderTop().name();
        bottomBorderColor = String.valueOf(cell.getCellStyle().getBottomBorderColor());
        leftBorderColor = String.valueOf(cell.getCellStyle().getLeftBorderColor());
        rightBorderColor = String.valueOf(cell.getCellStyle().getRightBorderColor());
        topBorderColor = String.valueOf(cell.getCellStyle().getTopBorderColor());
        dataFormat = cell.getCellStyle().getDataFormatString();
        fillBackgroundColor = String.valueOf(cell.getCellStyle().getFillBackgroundColorColor());
        fillForegroundColor = String.valueOf(cell.getCellStyle().getFillForegroundColorColor());
        fillPattern = cell.getCellStyle().getFillPattern().name();
        font = new Font(cell.getCellStyle().getFont());
        hidden = String.valueOf(cell.getCellStyle().getHidden());
        indention = String.valueOf(cell.getCellStyle().getIndention());
        locked = String.valueOf(cell.getCellStyle().getLocked());
        rotation = String.valueOf(cell.getCellStyle().getRotation());
        verticalAlignment = cell.getCellStyle().getVerticalAlignment().name();
    }

    private static String getCellValueToString(XSSFCell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return String.valueOf(cell.getDateCellValue());
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return String.valueOf(cell.getCellFormula());
            case ERROR:
                return String.valueOf(cell.getErrorCellValue());
            case BLANK:
            case _NONE:
                return "";
            default:
                throw new IllegalStateException("Unexpected value: " + cell.getCellType());
        }
    }

    private static class Font {
        public String bold;
        public String charSet;
        public String xssfColor;
        public String themeColor;
        public String fontHeight;
        public String fontHeightInPoints;
        public String fontName;
        public String italic;
        public String strikeout;
        public String typeOffset;
        public String underline;

        public Font(XSSFFont font) {
            bold = String.valueOf(font.getBold());
            charSet = String.valueOf(font.getCharSet());
            xssfColor = font.getXSSFColor() == null ? null : Arrays.toString(font.getXSSFColor().getRGB());
            themeColor = String.valueOf(font.getThemeColor());
            fontHeight = String.valueOf(font.getFontHeight());
            fontHeightInPoints = String.valueOf(font.getFontHeightInPoints());
            fontName = String.valueOf(font.getFontName());
            italic = String.valueOf(font.getItalic());
            strikeout = String.valueOf(font.getStrikeout());
            typeOffset = String.valueOf(font.getTypeOffset());
            underline = String.valueOf(font.getUnderline());
        }
    }
}

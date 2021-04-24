package nodump.noouchi.cellinfogenerator.domain.excel;

import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
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
        XSSFCellStyle cellStyle = cell.getCellStyle();
        alignment = cellStyle.getAlignment().name();
        borderBottom = cellStyle.getBorderBottom().name();
        borderLeft = cellStyle.getBorderLeft().name();
        borderRight = cellStyle.getBorderRight().name();
        borderTop = cellStyle.getBorderTop().name();
        bottomBorderColor = String.valueOf(cellStyle.getBottomBorderColor());
        leftBorderColor = String.valueOf(cellStyle.getLeftBorderColor());
        rightBorderColor = String.valueOf(cellStyle.getRightBorderColor());
        topBorderColor = String.valueOf(cellStyle.getTopBorderColor());
        dataFormat = cellStyle.getDataFormatString();
        fillBackgroundColor = cellStyle.getFillBackgroundColorColor() == null ? null : Arrays.toString(cellStyle.getFillBackgroundColorColor().getRGB());
        fillForegroundColor = cellStyle.getFillForegroundColorColor() == null ? null : Arrays.toString(cellStyle.getFillForegroundColorColor().getRGB());
        fillPattern = cellStyle.getFillPattern().name();
        font = new Font(cellStyle.getFont());
        hidden = String.valueOf(cellStyle.getHidden());
        indention = String.valueOf(cellStyle.getIndention());
        locked = String.valueOf(cellStyle.getLocked());
        rotation = String.valueOf(cellStyle.getRotation());
        verticalAlignment = cellStyle.getVerticalAlignment().name();
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
                // FIXME Want to improve the accuracy of detecting formulas that are not supported by POI.
                Object[] notSupportedFormula = WorkbookEvaluator.getNotSupportedFunctionNames().stream().filter(cell.getCellFormula()::contains).toArray();
                if (notSupportedFormula.length == 0) {
                    return String.valueOf(cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator().evaluateInCell(cell));
                }
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

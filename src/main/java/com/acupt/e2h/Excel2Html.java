package com.acupt.e2h;


import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Created by liujie on 2018/8/11.
 */
public class Excel2Html {

    private static ThreadLocal<SimpleDateFormat> defaultDateFormat = new ThreadLocal<SimpleDateFormat>() {
        @Override
        protected SimpleDateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
    };

    public String excel2html(File excelFile) throws Excel2HtmlException {
        try (InputStream inputStream = new FileInputStream(excelFile)) {
            Workbook wb = WorkbookFactory.create(inputStream);
            if (wb instanceof XSSFWorkbook) {
                return excel2html((XSSFWorkbook) wb);
            } else if (wb instanceof HSSFWorkbook) {
                throw new Excel2HtmlException("not support '.xls' yet");
            } else {
                throw new Excel2HtmlException("unknown Workbook type");
            }
        } catch (Exception e) {
            throw new Excel2HtmlException(e);
        }
    }

    private String excel2html(XSSFWorkbook wb) {
        StringBuilder buf = new StringBuilder(head4html()).append(head4table());
        int sheetNo = 0;
        Sheet sheet = wb.getSheetAt(sheetNo);
        for (Row row : sheet) {
            buf.append(head4tr());
            for (Cell cell : row) {
                CellInfo info = new CellInfo(cell);
                if (!info.exist()) {
                    continue;
                }
                buf.append(head4td(info.getRowspan(), info.getColspan()));
                buf.append(cell2html(cell));
                buf.append(tail4td());
            }
            buf.append(tail4tr());
        }
        buf.append(tail4table()).append(tail4html());
        return buf.toString();
    }

    private String cell2html(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                return defaultDateFormat.get().format(cell.getDateCellValue());
            }
            return NumberFormat.getInstance().format(cell.getNumericCellValue());
        }
        String value = cell.toString();
        if (cell.getHyperlink() != null) {
            return String.format("<a target=\"_blank\" href=\"%s\">%s</a>",
                    cell.getHyperlink().getAddress(), value);
        }
        return value;
    }

    private String head4html() {
        return "<!DOCTYPE html>\n"
                + "<html>\n"
                + "<head>\n"
                + "    <meta charset=\"UTF-8\">\n"
                + "    <title></title>\n"
                + "</head>\n"
                + "<body>";
    }

    private String tail4html() {
        return "</body>\n"
                + "</html>";
    }

    private String head4table() {
        return "<table border=\"1\" cellspacing=\"0\">";
    }

    private String tail4table() {
        return "</table>";
    }

    private String head4tr() {
        return "<tr>";
    }

    private String tail4tr() {
        return "</tr>";
    }

    private String head4td(int rowspan, int colspan) {
        return String.format(" <td rowspan=\"%d\" colspan=\"%d\">", rowspan, colspan);
    }

    private String tail4td() {
        return "</td>";
    }

    private static class CellInfo {

        private int rowspan;
        private int colspan;
        private boolean mergeRegion;

        public CellInfo(Cell cell) {
            Sheet sheet = cell.getSheet();
            int row = cell.getRowIndex();
            int column = cell.getColumnIndex();
            int sheetMergeCount = sheet.getNumMergedRegions();
            for (int i = 0; i < sheetMergeCount; i++) {
                CellRangeAddress range = sheet.getMergedRegion(i);
                int firstColumn = range.getFirstColumn();
                int lastColumn = range.getLastColumn();
                int firstRow = range.getFirstRow();
                int lastRow = range.getLastRow();
                if (row >= firstRow && row <= lastRow) {
                    if (column >= firstColumn && column <= lastColumn) {
                        mergeRegion = true;
                        if (row == firstRow) {
                            rowspan = lastRow - firstRow + 1;
                        }
                        if (column == firstColumn) {
                            colspan = lastColumn - firstColumn + 1;
                        }
                    }
                }
            }
            if (!mergeRegion) {
                rowspan = 1;
                colspan = 1;
            }
        }

        private boolean exist() {
            return rowspan > 0 && colspan > 0;
        }

        public int getRowspan() {
            return rowspan;
        }

        public int getColspan() {
            return colspan;
        }

        public boolean isMergeRegion() {
            return mergeRegion;
        }
    }

    public static class Excel2HtmlException extends Exception {

        private static final long serialVersionUID = -8249276025603068631L;

        public Excel2HtmlException(String message) {
            super(message);
        }

        public Excel2HtmlException(Throwable cause) {
            super(cause);
        }
    }

    public static void main(String[] args) throws Excel2HtmlException {
        if (args.length == 0) {
            System.out.println("[excel_file_path]");
            System.exit(1);
        }
        String excelPath = args[0];
        File excelFile = new File(excelPath);
        Excel2Html excel2Html = new Excel2Html();
        String html = excel2Html.excel2html(excelFile);
        System.out.println(html);
    }
}

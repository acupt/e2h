package com.acupt.e2h;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;


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

    /**
     * excel to html
     *
     * @param excelFile xls/xlsx file path
     * @return html content
     */
    public String excel2html(File excelFile, String title) throws Excel2HtmlException {
        try (InputStream inputStream = new FileInputStream(excelFile)) {
            Workbook wb = WorkbookFactory.create(inputStream);
            if (title == null || title.isEmpty()) {
                title = excelFile.getName();
                if (title.contains(".")) {
                    title = title.substring(0, title.lastIndexOf("."));
                }
            }
            return excel2html(title, wb);
        } catch (Exception e) {
            throw new Excel2HtmlException(e);
        }
    }

    private String excel2html(String title, Workbook wb) {
        StringBuilder buf = new StringBuilder(head4html(title));
        int n = wb.getNumberOfSheets();
        for (int i = 0; i < n; i++) {
            Sheet sheet = wb.getSheetAt(i);
            buf.append(sheet2htmlTable(sheet, n > 1 ? sheet.getSheetName() : null))
                    .append("<br>\n");
        }
        buf.append(tail4html());
        return buf.toString();
    }

    private String sheet2htmlTable(Sheet sheet, String title) {
        StringBuilder buf = new StringBuilder(head4table());// <table>
        if (title != null) {
            buf.append("<caption>" + title + "</caption>\n");
        }
        for (Row row : sheet) {
            StringBuilder rowsb = new StringBuilder();
            for (Cell cell : row) {
                CellInfo info = new CellInfo(cell);
                rowsb.append(info.tdHtml());// <td> text </td>
            }
            if (rowsb.length() > 0) {
                buf.append(head4tr()).append(rowsb).append(tail4tr());
            }
        }
        return buf.append(tail4table()).toString();// </table>
    }

    private String head4html(String title) {
        return "<!DOCTYPE html>\n"
                + "<html>\n"
                + "<head>\n"
                + "\t<meta charset=\"UTF-8\">\n"
                + "\t<title>" + title + "</title>\n"
                + "</head>\n"
                + "<body>\n";
    }

    private String tail4html() {
        return "</body>\n</html>\n";
    }

    private String head4table() {
        return "<table border=\"1\" cellspacing=\"0\">\n";
    }

    private String tail4table() {
        return "</table>\n";
    }

    private String head4tr() {
        return "\t<tr>\n";
    }

    private String tail4tr() {
        return "\t</tr>\n";
    }

    private static class CellInfo {

        private Cell cell;
        private int rowspan;
        private int colspan;
        private boolean mergeRegion;

        public CellInfo(Cell _cell) {
            cell = _cell;
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

        public String tdHtml() {
            if (!exist()) {
                return "";
            }
            StringBuilder sb = new StringBuilder("\t\t<td");
            if (rowspan > 1) {
                sb.append(" rowspan=\"").append(rowspan).append("\"");
            }
            if (colspan > 1) {
                sb.append(" colspan=\"").append(colspan).append("\"");
            }
            sb.append(">");
            if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    sb.append(defaultDateFormat.get().format(cell.getDateCellValue()));
                } else {
                    sb.append(NumberFormat.getInstance().format(cell.getNumericCellValue()));
                }
            } else {
                String value = cell.toString();
                if (cell.getHyperlink() != null) {
                    sb.append(String.format("<a target=\"_blank\" href=\"%s\">%s</a>",
                            cell.getHyperlink().getAddress(), value));
                } else {
                    sb.append(value);
                }
            }
            sb.append("</td>\n");
            return sb.toString();
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
            System.out.println("<excel_file_path> [html_file_path] [title]");
            System.exit(1);
        }
        String excelPath = args[0];
        String htmlPath = args.length > 1 ? args[1] : null;
        String title = args.length > 2 ? args[2] : null;
        File excelFile = new File(excelPath);
        Excel2Html excel2Html = new Excel2Html();
        String html = excel2Html.excel2html(excelFile, title);
        if (htmlPath == null || htmlPath.isEmpty() || "std".equals(htmlPath.toLowerCase())) {
            System.out.println(html);
            return;
        }
        try (FileWriter writer = new FileWriter(htmlPath)) {
            writer.write(html);
            writer.flush();
            System.out.println("SUCCESS");
        } catch (IOException e) {
            throw new Excel2HtmlException(e);
        }
    }
}

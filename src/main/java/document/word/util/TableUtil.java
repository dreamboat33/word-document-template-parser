package document.word.util;

import java.lang.reflect.Field;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public final class TableUtil {

    private TableUtil() {
        throw new UnsupportedOperationException(getClass() + " cannot be instantiated");
    }

    public static void copyRow(XWPFTableRow target, XWPFTableRow source) {
        target.getCtRow().set(source.getCtRow());
        try {
            // This is to reset the internal tableCells field, which forces an update and
            // circumvent the bug where target.getTableCells() returns an outdated list of table cells
            // after executing the above line, which updates some internal structure.

            // Another way to circumvent this, without resorting to reflection, is to not use
            // target.getTableCells() for the XWPFTableRow and use target.getTableICells() instead.

            // Another another way, first create the new row with the following code
            // XWPFTableRow newRow = new XWPFTableRow(CTRow.Factory.parse(rowTemplate.getCtRow().newInputStream()), table);
            // then POPULATE the new row, and then ADD the row to table
            // table.addRow(newRow);
            // Very importantly, and potentially a bug, it doesn't work if you ADD the row first and then UPDATE the new row.
            // Even though the XWPFTable, XWPFTableRow and XWPFTableCell objects look fine under inspection,
            // when calling doc.write(), it still writes old row data.
            // Reference: https://isurunuwanthilaka.medium.com/lets-play-with-apache-poi-186aa8d8ec71

            Field field = XWPFTableRow.class.getDeclaredField("tableCells");
            field.setAccessible(true);
            field.set(target, null);
        } catch (ReflectiveOperationException e) {
            throw new RuntimeException(e);
        }
    }
}

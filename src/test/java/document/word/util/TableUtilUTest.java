package document.word.util;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class TableUtilUTest {

    private XWPFTable table;
    private XWPFTableRow row;

    @SuppressWarnings("resource")
	@BeforeEach
    public void initDummyDocument() {
        XWPFDocument doc = new XWPFDocument();
        table = doc.createTable();
        row = table.createRow();
        row.setCantSplitRow(true);
        row.createCell().setText("some text");
        getFirstRunInCell(row.getTableCells().get(1)).setFontFamily("Arial");
    }

    @Test
    public void copyTest() {
        XWPFTableRow newRow = table.createRow();
        assertEquals(1, newRow.getTableCells().size());
        assertEquals(1, newRow.getTableICells().size());

        // verify copy works
        TableUtil.copyRow(newRow, row);
        assertEquals(2, newRow.getTableCells().size());
        assertEquals(2, newRow.getTableICells().size());
        assertEquals("some text", newRow.getTableCells().get(1).getText());
        assertEquals("Arial", getFirstRunInCell(newRow.getTableCells().get(1)).getFontFamily());

        // verify the two rows do not share the same table cell list or other references
        getFirstRunInCell(newRow.getTableCells().get(1)).setFontFamily("Verdana");
        newRow.createCell().setText("Other text");
        assertEquals(3, newRow.getTableCells().size());
        assertEquals(3, newRow.getTableICells().size());
        assertEquals(2, row.getTableCells().size());
        assertEquals(2, row.getTableICells().size());
        assertEquals("Arial", getFirstRunInCell(row.getTableCells().get(1)).getFontFamily());
    }

    private XWPFRun getFirstRunInCell(XWPFTableCell cell) {
        return cell.getParagraphArray(0).getRuns().get(0);
    }
}

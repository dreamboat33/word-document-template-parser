package document.word.util;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class RunUtilUTest {

    private XWPFRun run1, run2;

    @SuppressWarnings("resource")
	@BeforeEach
    public void initDummyDocument() {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph initialParagraph = doc.createParagraph();
        run1 = initialParagraph.createRun();
        run1.setText("Here");
        run1.addBreak();
        run1.setText("Have");
        run1.setText(" some");
        run2 = initialParagraph.createRun();
        run2.setText(" text");
    }

    @Test
    public void copyStyleTest() {
        run1.setBold(true);
        run1.setFontFamily("Arial");
        assertNotEquals(run1.isBold(), run2.isBold());
        assertNotEquals(run1.getFontFamily(), run2.getFontFamily());
        assertNotEquals(true, run2.isBold());
        assertNotEquals("Arial", run2.getFontFamily());

        // verify copyStyle works
        RunUtil.copyStyle(run2, run1);
        assertEquals(run1.isBold(), run2.isBold());
        assertEquals(run1.getFontFamily(), run2.getFontFamily());
        assertEquals(true, run2.isBold());
        assertEquals("Arial", run2.getFontFamily());

        // verify the two runs do not share the same style configuration reference
        run1.setBold(false);
        run1.setFontFamily("Verdana");
        assertNotEquals(run1.isBold(), run2.isBold());
        assertNotEquals(run1.getFontFamily(), run2.getFontFamily());
        assertEquals(true, run2.isBold());
        assertEquals("Arial", run2.getFontFamily());
    }

    @Test
    public void copyTest() {
        run1.setBold(true);
        assertNotEquals(run1.isBold(), run2.isBold());
        assertNotEquals(run1.text(), run2.text());
        assertNotEquals(true, run2.isBold());
        assertNotEquals("Here\nHave some", run2.text());

        // verify copy works
        RunUtil.copy(run2, run1);
        assertEquals(run1.isBold(), run2.isBold());
        assertEquals(run1.text(), run2.text());
        assertEquals(true, run2.isBold());
        assertEquals("Here\nHave some", run2.text());

        // verify the two runs do not share the same configuration reference
        run1.setBold(false);
        run1.setText("something else");
        assertNotEquals(run1.isBold(), run2.isBold());
        assertNotEquals(run1.text(), run2.text());
        assertEquals(true, run2.isBold());
        assertEquals("Here\nHave some", run2.text());
    }

    @Test
    public void replaceRunTextTest() {
        assertEquals("Here\nHave some", run1.text());
        assertEquals(3, run1.getCTR().getTList().size());
        assertEquals(1, run1.getCTR().getBrList().size());

        RunUtil.replaceRunText(run1, "run 1");
        assertEquals("run 1", run1.text());
        assertEquals(1, run1.getCTR().getTList().size());
        assertEquals(0, run1.getCTR().getBrList().size());
    }

    @Test
    public void spliceRunTail_affectSingleTextNodeTest() {
        RunUtil.spliceRunTail(run1, 4, "none");
        assertEquals("Here\nHave none", run1.text());
        assertEquals(3, run1.getCTR().getTList().size());
        assertEquals(1, run1.getCTR().getBrList().size());
    }

    @Test
    public void spliceRunTail_affectMultipleTextNodesTest() {
        RunUtil.spliceRunTail(run1, 8, "ello");
        assertEquals("Here\nHello", run1.text());
        assertEquals(2, run1.getCTR().getTList().size());
        assertEquals(1, run1.getCTR().getBrList().size());
    }

    @SuppressWarnings("resource")
	@Test
    public void spliceRunTail_replaceAllTest() {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Part 1");
        run.setText(", part 2");
        run.setText(", part 3");

        RunUtil.spliceRunTail(run, 999, "Something else");
        assertEquals("Something else", run.text());
        assertEquals(1, run.getCTR().getTList().size());
        assertEquals(0, run.getCTR().getBrList().size());
    }

    @Test
    public void spliceRunHead_affectSingleTextNodeTest() {
        RunUtil.spliceRunHead(run1, 1, "Th");
        assertEquals("There\nHave some", run1.text());
        assertEquals(3, run1.getCTR().getTList().size());
        assertEquals(1, run1.getCTR().getBrList().size());
    }

    @Test
    public void spliceRunHead_affectMultipleTextNodesTest() {
        RunUtil.spliceRunHead(run1, 6, "I'");
        assertEquals("\nI've some", run1.text());
        assertEquals(2, run1.getCTR().getTList().size());
        assertEquals(1, run1.getCTR().getBrList().size());
    }

    @SuppressWarnings("resource")
	@Test
    public void spliceRunHead_replaceAllTest() {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Part 1");
        run.setText(", part 2");
        run.setText(", part 3");

        RunUtil.spliceRunHead(run, 999, "Something else");
        assertEquals("Something else", run.text());
        assertEquals(1, run.getCTR().getTList().size());
        assertEquals(0, run.getCTR().getBrList().size());
    }
}

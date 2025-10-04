package document.word.util;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import java.util.List;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class ParagraphUtilUTest {

    private XWPFDocument doc;
    private XWPFParagraph initialParagraph;

    @BeforeEach
    public void initDummyDocument() {
        doc = new XWPFDocument();
        initialParagraph = doc.createParagraph();
        initialParagraph.createRun().setText("initial text");
    }

    @Test
    public void newParagraphBeforeTest() {
        XWPFParagraph paragraph = ParagraphUtil.newParagraphBefore(doc, initialParagraph);
        paragraph.createRun().setText("before");
        assertEquals("before", doc.getParagraphArray(0).getText());
        assertEquals("initial text", doc.getParagraphArray(1).getText());
    }

    @Test
    public void newParagraphAfterLastTest() {
        XWPFParagraph paragraph = ParagraphUtil.newParagraphAfter(doc, initialParagraph);
        paragraph.createRun().setText("after");
        assertEquals("initial text", doc.getParagraphArray(0).getText());
        assertEquals("after", doc.getParagraphArray(1).getText());
    }

    @Test
    public void newParagraphAfterNonLastTest() {
        XWPFParagraph lastParagraph = doc.createParagraph();
        lastParagraph.createRun().setText("last");

        XWPFParagraph paragraph = ParagraphUtil.newParagraphAfter(doc, initialParagraph);
        paragraph.createRun().setText("after");

        assertEquals("initial text", doc.getParagraphArray(0).getText());
        assertEquals("after", doc.getParagraphArray(1).getText());
        assertEquals("last", doc.getParagraphArray(2).getText());
    }

    @Test
    public void copyStyleTest() {
        XWPFParagraph newParagraph = doc.createParagraph();
        initialParagraph.setAlignment(ParagraphAlignment.CENTER);
        assertNotEquals(initialParagraph.getAlignment(), newParagraph.getAlignment());
        assertNotEquals(ParagraphAlignment.CENTER, newParagraph.getAlignment());

        // verify copyStyle works
        ParagraphUtil.copyStyle(newParagraph, initialParagraph);
        assertEquals(initialParagraph.getAlignment(), newParagraph.getAlignment());
        assertEquals(ParagraphAlignment.CENTER, newParagraph.getAlignment());

        // verify the two paragraphs do not share the same style configuration reference
        initialParagraph.setAlignment(ParagraphAlignment.RIGHT);
        assertNotEquals(initialParagraph.getAlignment(), newParagraph.getAlignment());
        assertEquals(ParagraphAlignment.CENTER, newParagraph.getAlignment());
    }

    @Test
    public void updateRunText_emptyArrayTest() {
        initialParagraph.getRuns().get(0).setText(" more text ");
        initialParagraph.createRun().setText("end text");
        assertEquals("initial text more text end text", doc.getParagraphArray(0).getText());
        assertEquals(2, doc.getParagraphArray(0).getRuns().get(0).getCTR().getTArray().length);

        // action
        int runsAdded = ParagraphUtil.updateRunText(initialParagraph, 0, new String[0]);

        // verify
        assertEquals(0, runsAdded);
        assertEquals("end text", doc.getParagraphArray(0).getText());

        List<XWPFRun> runs = doc.getParagraphArray(0).getRuns();
        assertEquals(2, runs.size());

        assertEquals(1, runs.get(0).getCTR().getTArray().length);
        assertEquals(0, runs.get(0).getCTR().getBrArray().length);
        assertEquals("", runs.get(0).text());

        assertEquals(1, runs.get(1).getCTR().getTArray().length);
        assertEquals(0, runs.get(1).getCTR().getBrArray().length);
        assertEquals("end text", runs.get(1).text());
    }

    @Test
    public void updateRunText_NormalArrayTest() {
        initialParagraph.getRuns().get(0).setText(" more text ");
        initialParagraph.createRun().setText("end text");
        assertEquals("initial text more text end text", doc.getParagraphArray(0).getText());
        assertEquals(2, doc.getParagraphArray(0).getRuns().get(0).getCTR().getTArray().length);

        // action
        int runsAdded = ParagraphUtil.updateRunText(initialParagraph, 0, new String[] { "line 1", "line 2" });

        // verify
        assertEquals(2, runsAdded);
        assertEquals("line 1\nline 2end text", doc.getParagraphArray(0).getText());

        List<XWPFRun> runs = doc.getParagraphArray(0).getRuns();
        assertEquals(4, runs.size());

        assertEquals(1, runs.get(0).getCTR().getTArray().length);
        assertEquals(0, runs.get(0).getCTR().getBrArray().length);
        assertEquals("line 1", runs.get(0).text());

        assertEquals(0, runs.get(1).getCTR().getTArray().length);
        assertEquals(1, runs.get(1).getCTR().getBrArray().length);
        assertEquals("\n", runs.get(1).text());

        assertEquals(1, runs.get(2).getCTR().getTArray().length);
        assertEquals(0, runs.get(2).getCTR().getBrArray().length);
        assertEquals("line 2", runs.get(2).text());

        assertEquals(1, runs.get(3).getCTR().getTArray().length);
        assertEquals(0, runs.get(3).getCTR().getBrArray().length);
        assertEquals("end text", runs.get(3).text());
    }
}

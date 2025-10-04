package document.word.util;

import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

public final class ParagraphUtil {

    private ParagraphUtil() {
        throw new UnsupportedOperationException(getClass() + " cannot be instantiated");
    }

    public static XWPFParagraph newParagraphBefore(IBody parent, XWPFParagraph target) {
        return parent.insertNewParagraph(target.getCTP().newCursor());
    }

    public static XWPFParagraph newParagraphAfter(IBody parent, XWPFParagraph target) {
        XmlCursor cursor = target.getCTP().newCursor();
        boolean hasNext = cursor.toNextSibling();
        if (!hasNext) {
            cursor.toParent();
            cursor.toEndToken();
        }
        return parent.insertNewParagraph(cursor);
    }

    public static void copyStyle(XWPFParagraph target, XWPFParagraph source) {
        target.getCTP().setPPr(source.getCTP().getPPr());
    }

    public static int updateRunText(XWPFParagraph paragraph, int runIndex, String[] linesWithBr) {
        XWPFRun run = paragraph.getRuns().get(runIndex);
        RunUtil.replaceRunText(run, linesWithBr.length == 0 ? "" : linesWithBr[0]);
        for (int i = 1; i < linesWithBr.length; i++) {
            XWPFRun newRunWithBr = paragraph.insertNewRun(++runIndex);
            RunUtil.copyStyle(newRunWithBr, run);
            newRunWithBr.addBreak();

            XWPFRun newRunWithText = paragraph.insertNewRun(++runIndex);
            RunUtil.copyStyle(newRunWithText, run);
            newRunWithText.setText(linesWithBr[i]);
        }
        int runsAdded = linesWithBr.length == 0 ? 0 : (linesWithBr.length - 1) * 2;
        return runsAdded;
    }
}

package document.word.util;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public final class RunUtil {

    private RunUtil() {
        throw new UnsupportedOperationException(getClass() + " cannot be instantiated");
    }

    public static void copyStyle(XWPFRun target, XWPFRun source) {
        target.getCTR().setRPr(source.getCTR().getRPr());
    }

    public static void copy(XWPFRun target, XWPFRun source) {
        target.getCTR().set(source.getCTR());
    }

    /*
     * Replace multiple text and br nodes under the same run with the given text.
     *
     * Apache POI has limited abstraction methods to effectively manipulate nodes under the same run,
     * possibly because a typical Word document program will break those into multiple runs anyway.
     *
     * For example, inserting a br node between two existing text nodes is impossible with only high level methods from XWPFRun or CTR.
     * (You need to access the xmlbeans directly via CTRImpl.get_store, and even then it is very messy and might break other stuff.)
     *
     * It is easier to break those nodes into multiple runs, each with individual text and br nodes.
     */
    public static void replaceRunText(XWPFRun run, String text) {
        CTR ctr = run.getCTR();
        for (int i = 0, len = ctr.getBrList().size(); i < len; i++) {
            ctr.removeBr(0);
        }
        for (int i = 1, len = ctr.getTList().size(); i < len; i++) {
            ctr.removeT(1);
        }
        run.setText(text, 0);
    }

    /*
     * Remove given number of characters from the end of the run, then insert text at the end of the run.
     * This operates under the assumption that there are no br nodes between affected text nodes.
     */
    public static void spliceRunTail(XWPFRun run, int numOfChar, String textToInsert) {
        CTR ctr = run.getCTR();
        List<CTText> texts = new ArrayList<>(ctr.getTList());
        int i = texts.size() - 1;
        while (i >= 0 && numOfChar > 0) {
            String text = texts.get(i).getStringValue();
            if (text.length() >= numOfChar) {
                texts.get(i).setStringValue(text.substring(0, text.length() - numOfChar) + textToInsert);
                return;
            }
            if (i == 0) {
                texts.get(0).setStringValue(textToInsert);
                return;
            }
            ctr.removeT(i);
            numOfChar -= text.length();
            i--;
        }
    }

    /*
     * Remove given number of characters from the head of the run, then insert text at the head of the run.
     * This operates under the assumption that there are no br nodes between affected text nodes.
     */
    public static void spliceRunHead(XWPFRun run, int numOfChar, String textToInsert) {
        CTR ctr = run.getCTR();
        List<CTText> texts = new ArrayList<>(ctr.getTList());
        int i = 0;
        while (i < texts.size() && numOfChar > 0) {
            String text = texts.get(i).getStringValue();
            if (text.length() >= numOfChar) {
                texts.get(i).setStringValue(textToInsert + text.substring(numOfChar));
                return;
            }
            if (i == texts.size() - 1) {
                texts.get(i).setStringValue(textToInsert);
                return;
            }
            ctr.removeT(0);
            numOfChar -= text.length();
            i++;
        }
    }
}

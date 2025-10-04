package document.word;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import document.word.exception.MissingTemplateVariableException;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;

public class WordDocumentTemplateParserTest {

    @SuppressWarnings("unchecked")
	@Test
    public void wordDocumentTemplateParserFillTest() throws IOException {
        // setup
        File wordFile = new File("src/test/resources/test-input.docx");
        File variableFile = new File("src/test/resources/test-variables.json");
        File outputFile = new File("src/test/resources/test-output.docx");

        // action
        Map<String, Object> variables = new ObjectMapper().readValue(variableFile, Map.class);
        WordDocumentTemplateParser parser = new WordDocumentTemplateParser(wordFile, variables, true);
        parser.fill(outputFile);

        // verify
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(outputFile))) {
            // header, foot and footnotes are replaced
            assertEquals(1, findParagraphsInHeaders(doc, "Page Header: some header value").size());
            assertEquals(1, findParagraphsInFooters(doc, "Page Footer: some footer value").size());
            assertEquals(1, findParagraphsInFootnotes(doc, " Reference: some reference value").size());

            // template rows are created, for both list in mapped variables and list in default value
            assertEquals(5, doc.getTableArray(0).getRows().size());
            assertEquals(1, findParagraphs(doc, "Key: some test key 1").size());
            assertEquals(1, findParagraphs(doc, "Key: some test key 2").size());
            assertEquals(1, findParagraphs(doc, "Key: some test key 3").size());
            assertEquals(1, findParagraphs(doc, "some test value 1").size());
            assertEquals(1, findParagraphs(doc, "some test value 2").size());
            assertEquals(1, findParagraphs(doc, "some test value 3").size());

            // new bullet point paragraphs are created for array variables
            assertEquals(1, findParagraphs(doc, "some list value A").size());
            assertEquals(1, findParagraphs(doc, "some list value B").size());
            assertEquals(1, findParagraphs(doc, "some list value C").size());

            // unmapped variables remain unchanged
            assertEquals(1, findParagraphs(doc, "Name: \\$\\{unmappedName\\}").size());

            // unmapped variables with default values are replaced
            assertEquals(1, findParagraphs(doc, "If contact is unknown, call no one.").size());

            // unmapped variables with assigned default values are replaced
            assertEquals(1, findParagraphs(doc, "Title: Programmer").size());
            assertEquals(1, findParagraphs(doc, "Title reminder: Programmer").size());

            // \n in variables are mapped to line breaks
            assertEquals(1, findParagraphs(doc, "Description: paragraph 1 line 1\\.\\nparagraph 1 line 2").size());

            // new paragraphs are created for array variables
            assertEquals(1, findParagraphs(doc, "paragraph 2\\..*").size());
        }
    }

    @SuppressWarnings("unchecked")
	@Test
    public void wordDocumentTemplateParserThrowsExceptionTest() throws IOException {
        // setup
        File wordFile = new File("src/test/resources/test-input-with-error.docx");
        File variableFile = new File("src/test/resources/test-variables.json");
        File outputFile = new File("src/test/resources/test-output-error.docx");
        outputFile.delete();

        // action
        Map<String, Object> variables = new ObjectMapper().readValue(variableFile, Map.class);
        WordDocumentTemplateParser parser = new WordDocumentTemplateParser(wordFile, variables, true);
        MissingTemplateVariableException exception = assertThrows(MissingTemplateVariableException.class, () -> parser.fill(outputFile));
        assertEquals("Missing template variable: name (The name is missing)", exception.getMessage());
        assertFalse(outputFile.exists());
    }

    private List<XWPFParagraph> findParagraphsInHeaders(XWPFDocument doc, String regex) {
        List<XWPFParagraph> result = new ArrayList<>();
        for (XWPFHeader header : doc.getHeaderList()) {
            _findParagraphs(header, regex, result);
        }
        return result;
    }

    private List<XWPFParagraph> findParagraphsInFooters(XWPFDocument doc, String regex) {
        List<XWPFParagraph> result = new ArrayList<>();
        for (XWPFFooter footer : doc.getFooterList()) {
            _findParagraphs(footer, regex, result);
        }
        return result;
    }

    private List<XWPFParagraph> findParagraphsInFootnotes(XWPFDocument doc, String regex) {
        List<XWPFParagraph> result = new ArrayList<>();
        for (XWPFFootnote footnote : doc.getFootnotes()) {
            _findParagraphs(footnote, regex, result);
        }
        return result;
    }

    private List<XWPFParagraph> findParagraphs(IBody body, String regex) {
        return _findParagraphs(body, regex, new ArrayList<>());
    }

    private List<XWPFParagraph> _findParagraphs(IBody body, String regex, List<XWPFParagraph> result) {
        for (XWPFParagraph paragraph : body.getParagraphs()) {
            if (paragraph.getText().matches(regex)) {
                result.add(paragraph);
            }
        }
        for (XWPFTable table : body.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    _findParagraphs(cell, regex, result);
                }
            }
        }
        return result;
    }
}

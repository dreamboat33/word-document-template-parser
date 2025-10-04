package document.word;

import document.word.exception.MissingTemplateVariableException;
import document.word.util.ParagraphUtil;
import document.word.util.RunUtil;
import document.word.util.TableUtil;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.MatchResult;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFEndnote;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordDocumentTemplateParser {

    /*
     * Pattern breakdown:
     *
     *     \$\{ ------------------- start with a dollar sign and an open curly bracket
     *
     *     ( ---------------------- start of group 1
     *
     *     [a-zA-Z0-9.$_-]+? ------ non-greedy match for a non empty string that consists of alphanumerics, dots, dollar signs, underscores or hyphens
     *
     *     (?:\[\])? -------------- an optional non-capturing group that matches an open square bracket and a close square bracket
     *
     *     ) ---------------------- end of group 1
     *
     *     ( ---------------------- start of group 2
     *
     *     :[-=?] ----------------- matches a colon, followed by a hyphen, an equal sign, or a question mark
     *
     *     .*? -------------------- non-greedy match for a nullable string
     *
     *     )? --------------------- end of optional group 2
     *
     *     \} --------------------- end with a close curly bracket
     *
     * Pattern examples:
     *     - ${var_name}
     *     - ${var_name:-default value}
     *     - ${var_name:=assign if missing}
     *     - ${var_name:?error if missing}
     *     - ${table_row_array[]}
     *     - ${table_row_array[]:-["default value 1","default value 2"]}
     */
    private static final Pattern PATTERN = Pattern.compile("\\$\\{([a-zA-Z0-9.$_-]+?(?:\\[\\])?)(:[-=?].*?)?\\}");
    private static final Pattern PATTERN_FORCE_MATCH = Pattern.compile("^|" + PATTERN.pattern());

    private final File source;
    private final Map<String, Object> variables;
    private final boolean checkEnvVar;

    public WordDocumentTemplateParser(File source, Map<String, Object> variables, boolean checkEnvVar) {
        this.source = source;
        this.variables = new HashMap<>(variables);
        this.variables.put("$", "$");
        this.checkEnvVar = checkEnvVar;
    }

    public void fill(File output) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(source))) {

            handleParagraphsAndTables(doc, this.variables);

            for (XWPFHeader header : doc.getHeaderList()) {
                handleParagraphsAndTables(header, this.variables);
            }
            for (XWPFFooter footer : doc.getFooterList()) {
                handleParagraphsAndTables(footer, this.variables);
            }
            for (XWPFFootnote footnote : doc.getFootnotes()) {
                handleParagraphsAndTables(footnote, this.variables);
            }
            for (XWPFEndnote endnote : doc.getEndnotes()) {
                handleParagraphsAndTables(endnote, this.variables);
            }

            try (FileOutputStream outputStream = new FileOutputStream(output)) {
                doc.write(outputStream);
            }
        }
    }

    private void handleParagraphsAndTables(IBody body, Map<String, Object> variables) {
        for (XWPFParagraph paragraph : new ArrayList<>(body.getParagraphs())) {
            replaceTemplateVariableInText(body, paragraph, variables);
        }
        for (XWPFTable table : body.getTables()) {
            int currentRowIndex = 0;
            for (XWPFTableRow row : new ArrayList<>(table.getRows())) {
                int rowsToRepeat = 1;
                Set<MatchResult> matches = retrieveAllTemplateVariableMatchesForTableRow(new HashSet<>(), row);
                for (MatchResult match : matches) {
                    boolean isRepeatRowVariable = match.group(1).endsWith("[]");
                    if (isRepeatRowVariable && resolveVariable(match, variables) instanceof List<?> list && list.size() > rowsToRepeat) {
                        rowsToRepeat = list.size();
                    }
                }

                for (int i = 1; i < rowsToRepeat; i++) {
                    XWPFTableRow newRow = table.insertNewTableRow(++currentRowIndex);
                    TableUtil.copyRow(newRow, row);
                    for (XWPFTableCell cell : newRow.getTableCells()) {
                        handleParagraphsAndTables(cell, computeRowVariables(variables, matches, i));
                    }
                }
                for (XWPFTableCell cell : row.getTableCells()) {
                    handleParagraphsAndTables(cell, computeRowVariables(variables, matches, 0));
                }
                currentRowIndex++;
            }
        }
    }

    private void replaceTemplateVariableInText(IBody context, XWPFParagraph paragraph, Map<String, Object> variables) {
        combineReplacePatternAcrossMultipleRuns(paragraph);

        int index = 0;
        for (XWPFRun run : new ArrayList<>(paragraph.getRuns())) {
            boolean replaced = false;
            StringBuilder replacedText = new StringBuilder();
            Matcher matcher = PATTERN.matcher(run.text());
            while (matcher.find()) {
                replaced = true;
                Object substitution = resolveVariable(matcher, variables);
                if (substitution instanceof List<?> substitutions) {
                    for (int i = 0, len = substitutions.size(); i < len - 1; i++) {
                        XWPFParagraph newParagraph = ParagraphUtil.newParagraphBefore(context, paragraph);
                        ParagraphUtil.copyStyle(newParagraph, paragraph);

                        if (i == 0) {
                            matcher.appendReplacement(replacedText, Matcher.quoteReplacement(String.valueOf(substitutions.get(i))));
                            for (XWPFRun prevRun : paragraph.getRuns().subList(0, index)) {
                                XWPFRun newRun = newParagraph.createRun();
                                RunUtil.copy(newRun, prevRun);
                            }
                            XWPFRun newRun = newParagraph.createRun();
                            RunUtil.copyStyle(newRun, run);
                            ParagraphUtil.updateRunText(newParagraph, index, replacedText.toString().split("\n"));
                            while (index > 0) {
                                paragraph.removeRun(0);
                                index--;
                            }
                            matcher = PATTERN_FORCE_MATCH.matcher(matcher.appendTail(new StringBuilder()));
                            matcher.find();
                            replacedText = new StringBuilder();
                        } else {
                            XWPFRun newRun = newParagraph.createRun();
                            RunUtil.copyStyle(newRun, run);
                            ParagraphUtil.updateRunText(newParagraph, 0, String.valueOf(substitutions.get(i)).split("\n"));
                        }
                    }
                    String lastSubstitution = substitutions.size() == 0 ? "" : String.valueOf(substitutions.get(substitutions.size() - 1));
                    matcher.appendReplacement(replacedText, Matcher.quoteReplacement(lastSubstitution));
                } else {
                    matcher.appendReplacement(replacedText, Matcher.quoteReplacement(String.valueOf(substitution)));
                }
            }

            if (!replaced) {
                index++;
                continue;
            }

            matcher.appendTail(replacedText);
            index += 1 + ParagraphUtil.updateRunText(paragraph, index, replacedText.toString().split("\n"));
        }
    }

    private Object resolveVariable(MatchResult matcher, Map<String, Object> variables) {
        String name = matcher.group(1);
        Object value = variables.get(name);
        if (value != null) return value;

        if (checkEnvVar) {
            String env = System.getenv(name);
            if (env != null) {
                value = parseJsonValue(env);
                variables.put(name, value);
                return value;
            }
        }

        String defaultValue = matcher.group(2);
        if (defaultValue != null) {
            if (defaultValue.startsWith(":?")) {
                String customMessage = defaultValue.substring(":?".length());
                StringBuilder fullMessage = new StringBuilder("Missing template variable: ").append(name);
                if (customMessage.length() > 0) {
                    fullMessage.append(" (").append(customMessage).append(")");
                }
                throw new MissingTemplateVariableException(fullMessage.toString());
            }
            value = parseJsonValue(defaultValue.substring(":-".length()));
            if (defaultValue.startsWith(":=")) {
                variables.put(name, value);
            }
            return value;
        }

        return matcher.group(0);
    }

    private Object parseJsonValue(String value) {
        try {
            return new ObjectMapper().readValue(value, Object.class);
        } catch (JsonProcessingException e) {
            return value;
        }
    }

    private Map<String, Object> computeRowVariables(Map<String, Object> allVariables, Set<MatchResult> matches, int rowIndex) {
        Map<String, Object> result = new HashMap<>();
        for (MatchResult match : matches) {
            String name = match.group(1);
            Object value = resolveVariable(match, allVariables);
            boolean isRepeatRowVariable = name.endsWith("[]");
            if (isRepeatRowVariable && value instanceof List<?> list) {
                result.put(name, rowIndex < list.size() ? list.get(rowIndex) : "");
            } else {
                result.put(name, rowIndex == 0 ? value : "");
            }
        }
        return result;
    }

    private Set<MatchResult> retrieveAllTemplateVariableMatchesForTableRow(Set<MatchResult> result, XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            for (XWPFParagraph p : cell.getParagraphs()) {
                Matcher matcher = PATTERN.matcher(p.getText());
                while (matcher.find()) {
                    result.add(matcher.toMatchResult());
                }
            }
            for (XWPFTable t : cell.getTables()) {
                for (XWPFTableRow r : t.getRows()) {
                    retrieveAllTemplateVariableMatchesForTableRow(result, r);
                }
            }
        }
        return result;
    }

    private void combineReplacePatternAcrossMultipleRuns(XWPFParagraph paragraph) {
        StringBuilder full = new StringBuilder();
        int start = 0;
        List<RunWrapper> wrappers = new ArrayList<>();
        for (XWPFRun run : new ArrayList<>(paragraph.getRuns())) {
            String text = run.text();
            full.append(text);
            wrappers.add(new RunWrapper(run, start, start + text.length()));
            start += text.length();
        }

        int runsDeleted = 0;
        int wrapperIndex = 0;
        Matcher matcher = PATTERN.matcher(full);
        while (matcher.find()) {
            while (matcher.start() >= wrappers.get(wrapperIndex).end) {
                wrapperIndex++;
            }
            int runStartIndex = wrapperIndex;

            while (matcher.end() > wrappers.get(wrapperIndex).end) {
                wrapperIndex++;
            }
            int runEndIndex = wrapperIndex;

            if (runStartIndex != runEndIndex) {
                RunWrapper runStart = wrappers.get(runStartIndex);
                RunUtil.spliceRunTail(runStart.run, runStart.end - matcher.start(), matcher.group(0));
                while (++runStartIndex < runEndIndex) {
                    paragraph.removeRun(runStartIndex - runsDeleted);
                    runsDeleted++;
                }
                RunWrapper runEnd = wrappers.get(runEndIndex);
                RunUtil.spliceRunHead(runEnd.run, matcher.end() - runEnd.start, "");
            }
        }
    }

    private static class RunWrapper {
        XWPFRun run;
        int start, end;
        RunWrapper(XWPFRun run, int start, int end) {
            this.run = run;
            this.start = start;
            this.end = end;
        }
    }
}

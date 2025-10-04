package document.word;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class WordDocumentTemplateParserMain {

    private static void printUsage(String extraMessage) {
        if (extraMessage != null) {
            System.out.println(extraMessage);
        }
        System.out.println("Usage: java " + WordDocumentTemplateParserMain.class.getName() + " -i <input_docx_file> -o <output_docx_file> -v <json_or_file>");
        System.out.println("Flags:");
        System.out.println("    -h, --help           print this help");
        System.out.println("    -E, --no-env-var     do not use environment variables for the template");
        System.out.println("    -i, --input          the input docx file");
        System.out.println("    -o, --output         the output docx file");
        System.out.println("    -v, --variables      either a json object for resolving template variables");
        System.out.println("                             e.g. '{\"author\":\"Andy\"}'");
        System.out.println("                         or a json file path, prefixed by the symbol @");
        System.out.println("                             e.g. '@/home/user/template-variables.json'");
        System.exit(1);
    }

    @SuppressWarnings("unchecked")
    public static void main(String[] args) throws IOException {
        final ObjectMapper objectMapper = new ObjectMapper();
        boolean checkEnvVar = true;
        File input = null;
        File output = null;
        Map<String, Object> variables = new HashMap<>();

        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
                case "-h", "--help" -> {
                    printUsage(null);
                }
                case "-i", "--input" -> {
                    if (input != null) printUsage("Multiple input files");
                    if (i == args.length - 1) printUsage("Expected input file");
                    input = new File(args[++i]);
                }
                case "-o", "--output" -> {
                    if (output != null) printUsage("Multiple output files");
                    if (i == args.length - 1) printUsage("Expected output file");
                    output = new File(args[++i]);
                }
                case "-v", "--variables" -> {
                    if (i == args.length - 1) printUsage("Expected json or file");
                    String ref = args[++i];
                    if (ref.startsWith("@")) {
                        variables.putAll(objectMapper.readValue(new File(ref.substring(1)), Map.class));
                    } else {
                        variables.putAll(objectMapper.readValue(ref, Map.class));
                    }
                }
                case "-E", "--no-env-var" -> {
                    checkEnvVar = false;
                }
                default -> {
                    printUsage("Unknown flag: " + args[i]);
                }
            }
        }

        if (input == null) {
            printUsage("Missing input file");
        }
        if (output == null) {
            printUsage("Missing output file");
        }

        new WordDocumentTemplateParser(input, variables, checkEnvVar).fill(output);
    }
}

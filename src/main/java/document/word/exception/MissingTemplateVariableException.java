package document.word.exception;

public class MissingTemplateVariableException extends IllegalArgumentException {

	private static final long serialVersionUID = -4349613558002209829L;

	public MissingTemplateVariableException(String message) {
        super(message);
    }
}

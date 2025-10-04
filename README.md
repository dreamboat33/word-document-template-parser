# word-document-template-parser

## Description

This java program parses an existing .docx word document and replaces template variable pattern ```${xxx}``` with provided values.

For example, given the following json mapping:
```
{
    "name": "Pet Store",
    "description": "Hello!\nWelcome!",
    "pets": ["cat", "dog", "bird"]
}
```

In the output document:

- ```${name}``` will be replaced with text "Pet Store".

- ```${description}``` will be replaced with "Hello!", followed by a line break, and then "Welcome!".

- ```${pets}``` will be replaced by "cat", followed by a new paragraph "dog", and another new paragraph "bird". Paragraph style will be kept so if the original paragraph is an ordered list, the inserted paragraphs will also be part of the ordered list.

Extra notes:

- Table rows will be programmatically generated if there is any template variable with its name ending in "[]" in the original row that gets mapped to a list of values, e.g. ```${rows[]}``` being mapped to ```["value for row1", "value for row 2"]```.

- Default value of each template variable can be configured directly in the template via the pattern suffix ```${name:-Default Name}``` or```${name:=Default Name}```.

- A template variable will be deemed mandatory when given the pattern suffix ```${name:?Name variable is missing}```. An exception will be thrown with the provided error message if the template variable is missing.

- Optional missing template variables will be untouched.

- The dollar sign can be escaped with ```${$}```.

- Environment variables are used as a fallback for missing variable mappings, e.g.
  ```
  export name='Pet Store'
  export description='"Hello!\nWelcome!"'
  export pets='["cat","dog","bird"]'
  ```
  This behaviour can be disabled by using the ```--no-env-var``` flag.

You can run WordDocumentTemplateParserTest and compare test-output.docx with test-input.docx to study the program's behaviour.

## How to build jar and run

```
$ mvn clean install

$ echo '{"name":"Pet Store","description":"Hello!\nWelcome!",pets":["cat","dog","bird"]}' > variables.json

$ java -jar word-document-template-parser-1.0.0-SNAPSHOT-jar-with-dependencies.jar -i input.docx -o output.docx -v @variables.json
```

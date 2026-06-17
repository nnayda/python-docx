Feature: Add a footnote to a document
  In order to cite references in a Word document
  As a developer using docx
  I need to add footnotes to paragraphs and runs

  Scenario: Paragraph.add_footnote adds a footnote
    Given a document with a paragraph
     When I add a footnote with text "Smith, J. (2024)."
     Then the saved document has 1 footnote
      And footnote 1 has text "Smith, J. (2024)."

  Scenario: Run.add_footnote adds a footnote after the run
    Given a document with a paragraph containing a run
     When I add a footnote to the run with text "Jones, A. (2023)."
     Then the saved document has 1 footnote
      And footnote 1 has text "Jones, A. (2023)."

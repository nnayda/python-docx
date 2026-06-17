Feature: Add a hyperlink to a document
  In order to link to external resources from a Word document
  As a developer using docx
  I need to add hyperlinks to paragraphs

  Scenario: Paragraph.add_hyperlink adds an external hyperlink
    Given a document with a paragraph
     When I add a hyperlink with text "Google" and url "https://google.com/"
     Then the saved document has 1 hyperlink
      And hyperlink 1 has text "Google"
      And hyperlink 1 has url "https://google.com/"
      And the saved document defines the "Hyperlink" character style

  Scenario: Hyperlink.add_run composes the visible text
    Given a document with a paragraph
     When I add an empty hyperlink with url "https://example.com/"
      And I add a run "read more" to the hyperlink
     Then the saved document has 1 hyperlink
      And hyperlink 1 has text "read more"
      And hyperlink 1 has url "https://example.com/"

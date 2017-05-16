# Script formatter

#### Reads scripts and formats properly

This Python script reads a specially-formatted .docx file that parses documents based on '<' and '>' tags. Margins and spacing are automatically formatted and placed into a new .docx file which is similar to a screenplay.

Tags for Interior, Exterior, Transitions by default. Any non-default names are interpreted as character names. Paragraphs without tags are treated as descriptions.

Each paragraph is treated as a different section. When a tag is included at the beginning of the paragraph, it follows the entirety of the paragraph.

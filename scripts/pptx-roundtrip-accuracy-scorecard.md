# PPTX round-trip accuracy — composed deck class

Attribute = a content/structure fact the fixture baked in. intact = present
with its original value. The engine row is openPptx → savePptx (every save);
the LibreOffice row is its own open + re-save of the same fixture.

| attribute                | genoffice | soffice                                       |
| ------------------------ | --------- | --------------------------------------------- |
| slide count              | ✓         | ✓                                             |
| slide size               | ✓         | ✓                                             |
| text frames survive      | ✓         | ✓                                             |
| mixed runs: sizes + bold | ✓         | ✓                                             |
| large text keeps 2800    | ✓         | ✓                                             |
| shape geometry roundRect | ✓         | ✓                                             |
| shape fill 4472C4        | ✓         | ✓                                             |
| frame EMU geometry       | ✓         | degraded: frames present, coordinates changed |
| table part survives      | ✓         | ✓                                             |
| table 2 columns          | ✓         | ✓                                             |
| table cell texts         | ✓         | ✓                                             |
| picture element          | ✓         | ✓                                             |
| image media part         | ✓         | ✓                                             |
| image bytes verbatim     | ✓         | ✓                                             |
| speaker notes            | ✓         | ✓                                             |

genoffice intact 15/15; LibreOffice intact 14/15.

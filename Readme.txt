=== phppptx 4.0 ===
https://www.phppptx.com/

phppptx is a PHP library designed to dynamically generate presentations in PowerPoint format (PresentationML).

=== What's new on phppptx 4.0? ===

- HTML Extended and CSS Extended (Premium licenses):
    · Call phppptx methods from custom HTML tags.
    · Apply phppptx styles from custom CSS.
- PPTX to HTML (Advanced and Premium licenses):
    · Native PHP classes to do the transformations.
    · Customize tags and styles.
    · Use custom plugins to handle the conversions.
- Create table styles.
- PHP 8.5 support.
- New method to add text box connectors: addTextBoxConnector (Advanced and Premium licenses).
- New method and options to improve the performance when working with templates:
    · processTemplate method.
    · preprocessed option in CreatePptxFromTemplate and setTemplateSymbol.
- Fill cell tables with pictures in addTable.
- textDirection, verticalAlign and shapeGuide options in addShape.
- marginRight option in addText.
- HTML to PPTX:
    · Improved adding alpha color styles.
    · Supported #rgba format to apply colors.
- PptxCustomizer (Premium licenses):
    · marginRight option in paragraph and table-cell-paragraph types.
- Merge PPTX files (Advanced and Premium licenses):
    · 3d models.
    · customXml contents.
    · inks.
    · objects.
    · tags.
- getActiveSlideInformation returns ID information.
- flipH and flipV options in addShape.
- Removed PHP Warnings when images to be added do not exist.
- Extra check to avoid duplicated shape internal ids.
- name option moved to the position array option in addTextBox.
- Increased internal ids maximum values.
- forceNotTidy option available in replaceVariableHtml.

3.5 VERSION

- PptxCustomizer (Premium licenses):
    · Change styles of existing contents on the fly in presentations created from scratch and templates.
- PptxPath (Advanced and Premium licenses):
    · Clone elements.
    · Get elements.
    · Move elements.
    · Remove elements.
- New getTemplateVariablesType method to return template variables and their types (Advanced and Premium licenses).
- New removeTiming method to remove timing tags.
- New type option in replaceVariableText and removeVariableText to do block and inline replacements.
- New paragraph styles: distributed align, before text indentation, left margin, before and after spacing, rtl.
- Indexer (Advanced and Premium licenses):
    · Supported PPTM, POTX and POTM files.
    · Audios.
    · Videos.
- Merge PPTX files (Advanced and Premium licenses):
    · Supported images and medias in themes.
    · Supported diagrams (layout, data, drawing, color and quickStyles).
    · Fix non-standard internal paths automatically.
- HTML to PPTX:
    · List tags (ul, ol, li).
    · Improved handle br tags.
- Conversion plugin based on LibreOffice (Advanced and Premium licenses):
    · Added "--norestore" to all conversions.
    · New documentation in the macros-libreoffice folder to enable and use lossless compression without adding a macro.
    · New path option in transform to set the path to libreoffice.
    · New escapeshellarg option.
- Conversion plugin based on MS PowerPoint (Advanced and Premium licenses):
    · PPT to PPTX.
    · PPTX to PPT.
    · PPT to PDF.
- Supported POTX and PPTM files as templates.
- New show option in setSlideSettings to show or hide slides.
- Add descr value in addAudio, addTable and addVideo methods.
- New rtl option in addTable.
- Theme charts:
    · Set custom title layout.
    · Apply font styles to series and values labels (Premium licenses).
    · New valueDataLabels option to customize labels by position (Premium licenses).
- The listLevel and listStyles options can be used with addText.
- splitPptx allows setting a custom path for the target documents distinct than the script folder (Advanced and Premium licenses).
- Improved cleaning timing tags related to the removed shape in removeShapeSlide, removeVariableAudio, removeVariableImage and removeVariableVideo methods.
- Removed @ (shut-up) operator in replaceChartData to get legends and categories.
- Removed ~E_STRICT from the default logger error levels when using PHP 8.0 or newer.
- Corrections in the internal phpdoc documentation.

3.0 VERSION

- MS PowerPoint 2024 support.
- PHP 8.4 support.
- Merge PPTX files (Advanced and Premium licenses).
- Notes:
    · Add new notes in slides.
    · Replace variables.
    · Remove notes.
- New targets to use with replace and remove template methods: notesSlides, slideLayouts and slideMasters.
- New method to insert SVG contents: addSvg.
- WebP image format supported.
- HTML to PPTX:
    · CSS variables.
    · Supported root and only-child selectors.
    · Improved CSS media query handling.
    · CSS 8-digit HEX colors are added as 6-digit HEX colors.
- New method to add a macro: addMacroFromPptx.
- New methods to remove contents in a template: removeVariableAudio, removeVariableImage, removeVariableVideo.
- New method addLink to add links.
- Handle PPTX, PPTM, POTX and POTM extensions automatically.
- Supported ofPie as chart type in addChart.
- Indexer (Advanced and Premium licenses):
    · Notes.
- replaceVariableImage includes sizeX and sizeY options to set custom sizes, and the descr option.
- removeShapeSlide handles timing tags related to the removed shape.
- New cleanLayoutParagraphContents option not to clean paragraph contents from the layout when adding a new slide with addSlide.
- Encrypt PPTX supports files bigger than 6.5MB (Premium licenses).
- GdImage objects supported when adding images.
- Theme charts supports applying custom colors to lines (Premium licenses).
- setPresentationShettings applies RTL settings if enabled.
- Included a sample_composer.json file in the plugins folder of the namespaces package (Advanced and Premium licenses).
- PHP GD Extension is checked in phppptx-cli and check scripts.
- New enableHugeXmlMode static function in XmlUtilities to parse huge XML contents enabling the XML_PARSE_HUGE flag (LIBXML_PARSEHUGE predefined constant).

2.5 VERSION

- phppptx CLI command (Premium licenses):
    · Speed up development by generating phppptx code skeletons automatically.
    · Skeletons generated for presentations from scratch and using templates.
    · Check server settings.
    · Show automatic recommendations.
    · Return phppptx information.
- New PptxUtilities method: replaceChartData to replace the data associated with a given chart (Advanced and Premium licenses).
- Add shapes: addShape.
- Add footers in slide: addFooterSlide (dateAndTime, slideNumber and textContent).
- Replace placeholders with HTML in PPTX templates: replaceVariableHtml.
- Supported new chart types in addChart: bubble, radar, scatter, surface.
- replaceVariablePptxFragment supports contents with external relationships.
- Added a default HTML content when adding an empty content with embedHTML and replaceVariableByHTML to avoid throwing a PHP error.
- Remove invalid UTF-8 XML characters automatically.
- PptxFragment class moved to the Elements namespace in the namespaces package.

2.0 VERSION

- Add sections: addSection.
- Add comments and comment authors: addComment and addCommentAuthor.
- Add math equations: addMathEquation.
- Replace placeholders with PptxFragments in PPTX templates: replaceVariablePptxFragment.
- PHP 8.3 support.
- HTML to PPTX:
    · Add breaks.
- The addText method adds a parseLineBreaks option to parse line breaks ('\n', "\n", '\n\r', "\n\r" and others).
- replaceVariableList and replaceVariableTable methods allow using PptxFragment as content values.
- addSlide allows setting specific positions and sections to add slides into the presentation.
- New option in setPresentationSettings to set readOnly.
- Indexer (Advanced and Premium licenses):
    · Sections.
    · Comments.
    · Comment authors.
- New PhppptxLogger::$errorReporting public static variable to set a custom error reporting value. Default as E_ALL & ~E_NOTICE & ~E_STRICT & ~E_DEPRECATED.
- Reordered internal XML files in the default base template to return Microsoft PowerPoint 2007+ mime type.
- Corrections in the internal phpdoc documentation.
- Parsing a PPTX shows the file name in the Exception if a file can't be read as a ZIP file.

1.0 VERSION

- Support for all MS PowerPoint versions from MS PowerPoint 2007 to MS PowerPoint 2021. Other PPTX readers such as LibreOffice and Google Docs are supported too (the support of these programs reading PPTX files may vary).
- Generate PPTX files from scratch and using templates.
- Content methods: addAudio, addChart, addHtml, addImage, addList, addSlide, addTable, addText, addTextBox, addVideo.
- Layout and general methods: addBackgroundImage, addProperties, getActiveSlide, getActiveSlideInformation, removeShapeSlide, setBackgroundColor, setMarkAsFinal, setPresentationSettings, setRtl, setSlideSettings.
- Template methods: getTemplateVariables, replaceVariableAudio, replaceVariableImage, replaceVariableList, replaceVariableTable, replaceVariableText, replaceVariableVideo, removeVariableText, setTemplateSymbol.
- Transform PPT to PPTX, PPTX to PDF, PPTX to ODP, ODP to PPTX (Advanced and Premium licenses).
- Indexer: return information from a PPTX (Advanced and Premium licenses).
- Crypto: encrypt PPTX files (Premium licenses).
- Sign PPTX files (Premium licenses).
- Save and download PPTX files.
- Stream mode (Premium licenses).
- PptxUtilities: removeSlide, searchAndReplace, split (Advanced and Premium licenses).
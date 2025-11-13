<?php
// generate a PPTX with text contents and transform it to HTML

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$content = array(
    'text' => 'My title',
);
// the getActiveSlideInformation method returns information about placeholders in the current active slide to add text contents
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

$content = array(
    'text' => 'My subtitle',
    'bold' => true,
    'underline' => 'single',
);
$pptx->addText($content, array('placeholder' => array('name' => 'Subtitle 2')));

// append a new content in the same placeholder position
$content = array(
    array(
        'text' => 'At vero',
        'highlight' => '628A54',
    ),
    array(
        'text' => ' eos et ',
        'bold' => true,
        'italic' => true,
        'font' => 'Times New Roman',
    ),
    array(
        'text' => 'accusamus et iusto.',
        'strikethrough' => true,
        'color' => '628A54',
        'italic' => true,
    ),
);
$pptx->addText($content, array('placeholder' => array('name' => 'Subtitle 2')));

// add new slide using the Title and Content layout
$pptx->addSlide(array('layout' => 'Title and Content', 'active' => true));

$content = array(
    'text' => 'My custom title',
    'bold' => true,
    'font' => 'Arial',
    'fontSize' => 60,
);
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

// replace the text contents in the placeholder position. By default appends the text contents
$content = array(
    'text' => "My new \ntitle",
);
$paragraphStyles = array(
    'align' => 'center',
    'parseLineBreaks' => true,
);
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')), $paragraphStyles, array('insertMode' => 'replace'));

$pptx->savePptx(__DIR__ . '/example_transformNativeHtml_1');

$transformHtmlPlugin = new TransformNativeHtmlDefaultPlugin();
$transform = new TransformNativeHtml(__DIR__ . '/example_transformNativeHtml_1.pptx');
$html = $transform->transform($transformHtmlPlugin);

echo $html;
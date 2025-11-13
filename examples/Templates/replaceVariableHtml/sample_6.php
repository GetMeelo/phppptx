<?php
// replace text variables (placeholders) with HTML doing block replacements in notesSlides, slides, slideLayouts and slideMasters targets

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// replace variables
$pptx->replaceVariableHtml(array('VAR_TITLE' => '<p style="font-size: 48px;"><b>php</b><em style="color:#C060E0;">pptx</em></p>'));
$pptx->replaceVariableHtml(array('VAR_TITLE' => '<p style="font-size: 32px;"><i>content in</i> <u>layout</u></p>'), array('target' => 'slideLayouts'));
$pptx->replaceVariableHtml(array('VAR_MASTER' => '<p style="font-size: 32px;"><i>content in</i><em> <b>slide master</b></em></p>'), array('target' => 'slideMasters'));
$html = '<style>
    p {
        font-family: Arial;
        font-weight: bold;
        text-decoration: underline;
    }
</style>
<p>Content <br>in <em>note</em></p>';
$pptx->replaceVariableHtml(array('VAR_NOTE_1' => $html), array('target' => 'notesSlides'));
$html = '<style>
    p {
        font-family: Arial;
        font-weight: bold;
        text-decoration: underline;
    }
</style>
<p>Content in <em>note</em></p>';
$pptx->replaceVariableHtml(array('VAR_NOTE_2' => $html, 'VAR_NOTE_3' => $html), array('target' => 'notesSlides', 'type' => 'inline'));

$pptx->savePptx(__DIR__ . '/example_replaceVariableHtml_6');
<?php
// replace text variables (placeholders) with HTML in the internal active slide doing block type replacements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$htmlA = '<p style="font-size: 48px;"><b>php</b><em style="color:#C060E0;">pptx</em></p>';

// replace variables in the internal active slide
$pptx->replaceVariableHtml(array('VAR_TITLE' => $htmlA), array('activeSlide' => true));

// change the internal active slide
$pptx->setActiveSlide(array('position' => 1));

$htmlB = '<style>
    p {
        background-color: #AEBF7A;
        text-align: center;
    }
    .c1 {
        text-decoration: underline;
        font-size: 16pt;
    }
</style>
<p>
    <b>phppptx</b> can transform <span class="c1">HTML to PPTX</span>.
    <a href="https://www.phppptx.com">Link to phppptx</a>
</p>';

// replace variables in the internal active slide
$pptx->replaceVariableHtml(array('VAR_TITLE' => $htmlB), array('activeSlide' => true));

$pptx->savePptx(__DIR__ . '/example_replaceVariableHtml_3');
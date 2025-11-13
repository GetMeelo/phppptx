<?php
// return information from an existing PPTX

require_once __DIR__ . '/../../classes/CreatePptx.php';

$indexer = new Indexer(__DIR__ . '/../files/sample_indexer.pptx');
$output = $indexer->getOutput();

print_r('presentation: ');
print_r($output['presentation']);

print_r('slides: ');
print_r($output['slides']);

print_r('layouts: ');
print_r($output['layouts']);

print_r('core properties: ');
print_r($output['properties']['core']);

print_r('comment authors: ');
print_r($output['comments']['authors']);
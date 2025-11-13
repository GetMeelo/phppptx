<?php

/**
 * This class offers some utilities to work with existing PowerPoint (.pptx) documents
 *
 * @category   Phppptx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
require_once __DIR__ . '/CreatePptx.php';

class PptxUtilities
{
    /**
     * Removes a slide from a PowerPoint presentation
     *
     * @access public
     * @param string|PptxStructure $source path to the presentation
     * @param string $target path to the output presentation
     * @param array $options
     *        'slideNumber' (array): slide numbers to remove
     */
    public function removeSlide($source, $target, $options)
    {
        if ($source instanceof PptxStructure) {
            // PptxStructure object
            $pptxFile = $source;
        } else {
            // file
            $pptxFile = new PptxStructure();
            $pptxFile->parsePptx($source);
        }

        $xmlUtilities = new XmlUtilities();

        $contentTypesXML = $pptxFile->getContent('[Content_Types].xml');
        $contentTypesDOM = $xmlUtilities->generateDomDocument($contentTypesXML);
        $contentTypesXPath = new DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        // get slides from $mainXMLPathNodes to get the correct order of the slides
        $mainXML = $pptxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
        $mainDOM = $xmlUtilities->generateDomDocument($mainXML);

        $mainXPath = new DOMXPath($mainDOM);
        $mainXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $query = '//xmlns:sldIdLst/xmlns:sldId';
        // query by slide number if set
        if (isset($options['slideNumber'])) {
            $query .= '[';
            foreach ($options['slideNumber'] as $slideNumber) {
                $query .= 'position()='.$slideNumber.' or ';
            }
            $query = substr($query, 0, -4);
            $query .= ']';
        }

        $slideNodes = $mainXPath->query($query);

        foreach ($slideNodes as $slideNode) {
            $slideNode->parentNode->removeChild($slideNode);
        }

        // save the data in the PPTX file
        $pptxFile->addContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1), $mainDOM->saveXML());

        // save file
        $pptxFile->savePptx($target);

        // free DOMDocument resources
        $contentTypesDOM = null;
        $mainDOM = null;
    }

    /**
     * Replaces chart data from a PowerPoint presentation
     *
     * @access public
     * @param string|PptxStructure $source path to the PPTX
     * @param string $target path to the output PPTX
     * @param array $chartData key (int): number of the chart to replace
     * Values:
     *     legends (array): chart legends
     *     categories (array): chart categories
     *     values (array): chart values
     *     title (string): chart title
     * Data must exist in the chart before being replaced
     * @param array $options
     *        'slideNumber' (int) slide number to replace the chart. Default as 1 (first slide)
     * @return PptxStructure
     */
    public function replaceChartData($source, $target, $chartData, $options = array())
    {
        if ($source instanceof PptxStructure) {
            // PptxStructure object
            $pptxFile = $source;
        } else {
            // file
            $pptxFile = new PptxStructure();
            $pptxFile->parsePptx($source);
        }

        // default values
        if (!isset($options['slideNumber'])) {
            $options['slideNumber'] = 1;
        }

        $xmlUtilities = new XmlUtilities();

        $contentTypesXML = $pptxFile->getContent('[Content_Types].xml');
        $contentTypesDOM = $xmlUtilities->generateDomDocument($contentTypesXML);
        $contentTypesXPath = new DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        // get slides from $mainXMLPathNodes to get the correct order of the slides
        $mainXML = $pptxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
        $mainDOM = $xmlUtilities->generateDomDocument($mainXML);
        $mainXPath = new DOMXPath($mainDOM);
        $mainXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $query = '//xmlns:sldIdLst/xmlns:sldId';
        $query .= '['.$options['slideNumber'].']';
        $slideNodes = $mainXPath->query($query);

        $chartIdData = array();

        if ($slideNodes->length > 0) {
            $slideNode = $slideNodes->item(0);
            // get slide rels to get the slide contents
            $mainRelsXML = $pptxFile->getContent(str_replace('ppt/', 'ppt/_rels/', substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1)) . '.rels');
            $mainRelsDOM = $xmlUtilities->generateDomDocument($mainRelsXML);
            $mainRelsXPath = new DOMXPath($mainRelsDOM);
            $mainRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

            $query = '//xmlns:Relationship[@Id="'.$slideNode->getAttribute('r:id').'"]';
            $slideContentNodes = $mainRelsXPath->query($query);
            $slideContent = $pptxFile->getContent('ppt/' . $slideContentNodes->item(0)->getAttribute('Target'));

            $slideContentDOM = $xmlUtilities->generateDomDocument($slideContent);
            $slideContentXPath = new DOMXPath($slideContentDOM);
            $slideContentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
            $slideContentXPath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');

            $iChart = 0;
            $slideChartNodes = $slideContentXPath->query('//a:graphic//c:chart');
            foreach ($slideChartNodes as $slideChartNode) {
                if (!isset($chartData[$iChart])) {
                    //  add new data only if the index key is set
                    $iChart++;
                    continue;
                }
                if ($slideChartNode->hasAttribute('r:id')) {
                    // get chart elements
                    $slideRelsContent = $pptxFile->getContent('ppt/'. str_replace('slides/', 'slides/_rels/', $slideContentNodes->item(0)->getAttribute('Target')) . '.rels');
                    $slideRelsContentDOM = $xmlUtilities->generateDomDocument($slideRelsContent);
                    $slideRelsContentXPath = new DOMXPath($slideRelsContentDOM);
                    $slideRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                    $slideChartNodes = $slideRelsContentXPath->query('//xmlns:Relationship[@Id="'.$slideChartNode->getAttribute('r:id').'"]');
                    if ($slideChartNodes->length > 0 && $slideChartNodes->item(0)->hasAttribute('Target')) {
                        $chartPath = 'ppt' . str_replace('../', '/', $slideChartNodes->item(0)->getAttribute('Target'));
                        $chartRelsPath = 'ppt' . str_replace('../charts/', '/charts/_rels/', $slideChartNodes->item(0)->getAttribute('Target')) . '.rels';
                        $chartContent = $pptxFile->getContent($chartPath);
                        $chartRelsContent = $pptxFile->getContent($chartRelsPath);
                        $chartContentDOM = $xmlUtilities->generateDomDocument($chartContent);
                        $chartRelsContentDOM = $xmlUtilities->generateDomDocument($chartRelsContent);
                        $chartContentXPath = new DOMXPath($chartContentDOM);
                        $chartContentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                        $chartContentXPath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                        $chartRelsContentXPath = new DOMXPath($chartRelsContentDOM);
                        $chartRelsContentXPath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');

                        $xmlWP = $chartContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'plotArea');
                        $nodePlotArea = $xmlWP->item(0);
                        $type = '';
                        foreach ($nodePlotArea->childNodes as $node) {
                            if (strpos($node->nodeName, 'Chart') !== false) {
                                list($namespace, $type) = explode(':', $node->nodeName);
                                break;
                            }
                        }
                        // generate the chart class to be used from the chart type
                        $type = ucwords(str_replace(array('3D', 'Col'), array('', 'Bar'), ucwords($type)));
                        // remove subtype strings
                        $type = str_replace(array('Cylinder', 'Cone', 'Pyramid', 'Chart'), '', $type);
                        $chartClass = 'CreateChart' . $type;
                        $chartType = new $chartClass();
                        $onlyData = $chartType->prepareData($chartData[$iChart]['values']);

                        $tags = $chartType->dataTag();

                        // replace title
                        if (isset($chartData[$iChart]['title']) && !empty($chartData[$iChart]['title'])) {
                            $i = 0;
                            $query = '//c:title/c:tx/c:rich/a:p/a:r/a:t';
                            $xmlSeries = $chartContentXPath->query($query);
                            // the title can have more than one a:t, replace only the first one and empty the others
                            foreach ($xmlSeries as $entry) {
                                if ($i > 0) {
                                    $entry->nodeValue = '';
                                } else {
                                    $entry->nodeValue = $chartData[$iChart]['title'];
                                }
                                $i++;
                            }
                        }

                        // replace legends values
                        if (isset($chartData[$iChart]['legends']) && count($chartData[$iChart]['legends']) > 0) {
                            $i = 0;
                            $query = '//c:tx/c:strRef/c:strCache/c:pt/c:v';
                            $xmlSeries = $chartContentXPath->query($query);
                            foreach ($xmlSeries as $entry) {
                                if (isset($chartData[$iChart]['legends'][$i])) {
                                    $entry->nodeValue = $chartData[$iChart]['legends'][$i];
                                }  else {
                                    $entry->nodeValue = '';
                                }
                                $i++;
                            }
                        }

                        // replace categories values
                        if (isset($chartData[$iChart]['categories']) && count($chartData[$iChart]['categories']) > 0) {
                            $i = 0;
                            $query = '//c:cat/c:strRef/c:strCache/c:pt/c:v';
                            $xmlLegends = $chartContentXPath->query($query);
                            foreach ($xmlLegends as $entry) {
                                if (isset($chartData[$iChart]['categories'][$i])) {
                                    $entry->nodeValue = $chartData[$iChart]['categories'][$i];
                                } else {
                                    $entry->nodeValue = '';
                                }
                                $i++;
                            }
                        }

                        // replace chart values
                        if (isset($chartData[$iChart]['values']) && count($chartData[$iChart]['values']) > 0) {
                            $i = 0;
                            foreach ($tags as $tag) {
                                $query = '//c:' . $tag . '/c:numRef/c:numCache/c:pt/c:v';
                                $xmlGraphics = $chartContentXPath->query($query);
                                foreach ($xmlGraphics as $entry) {
                                    $entry->nodeValue = $onlyData[$i];
                                    $i++;
                                }
                            }
                        }

                        $chartXml = $chartContentDOM->saveXML();
                        $chartXml = str_replace('Hoja', 'Sheet', $chartXml);
                        $pptxFile->addContent($chartPath, $chartXml);

                        //prepare the new excel file
                        $excel = $chartType->getXlsxType();

                        // generate the data chart structure
                        $charDataNew = array();
                        $charDataNew['data'] = array();
                        $i = 0;
                        foreach ($chartData[$iChart]['values'] as $dataValue) {
                            $charDataNew['data'][$i]['values'] = $dataValue;
                            if (isset($chartData[$iChart]['categories'][$i])) {
                                $charDataNew['data'][$i]['name'] = $chartData[$iChart]['categories'][$i];
                            }
                            $i++;
                        }
                        if (isset($chartData[$iChart]['legends'])) {
                            foreach ($chartData[$iChart]['legends'] as $legend) {
                                $charDataNew['legend'][] = $legend;
                            }
                        }

                        // generate uniqid value to be used to save the new XLSX files
                        $tempRnd = uniqid((string)mt_rand(999, 9999));

                        $charts = $slideContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'chart');
                        $idChart = $charts->item($iChart)->attributes->getNamedItemNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'id')->nodeValue;

                        $chartStructure = $excel->createChartXlsx('chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx', $charDataNew);
                        $chartStructure->savePptx(TempDir::getTempDir() . '/' . 'chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx');
                        rename(TempDir::getTempDir() . '/chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx.pptx', TempDir::getTempDir() . '/chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx');

                        // add the new XLSX to the PPTX. This XLSX allows editing the chart
                        $chartTarget = $chartRelsContentXPath->query('//rel:Relationship')->item(0)->getAttribute('Target');
                        $pptxFile->addFile(str_replace('../', 'ppt/', $chartTarget), TempDir::getTempDir() . '/chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx');
                        // keep temp file path to remove it after creating the new PPTX
                        $chartIdData[$tempRnd] = $idChart;
                    }
                }

                $iChart++;
            }
        }

        // save file
        $pptxStructure = $pptxFile->savePptx($target);

        // remove temp chart files
        foreach ($chartIdData as $tempRnd => $idChart) {
            unlink(TempDir::getTempDir() . '/chartData' . $tempRnd . str_replace('rId', '', $idChart) . '.xlsx');
        }

        // free DOMDocument resources
        $contentTypesDOM = null;
        $mainDOM = null;

        return $pptxStructure;
    }

    /**
     * Search and replace text in a PowerPoint presentation
     *
     * @param string|PptxStructure $source path to the presentation
     * @param string $target path to the output presentation
     * @param array $data strings to be searched and replaced
     * @param array $options
     *        'slideNumber' : slide number to replace the value. All if not set
     */
    public function searchAndReplace($source, $target, $data, $options = array())
    {
        if ($source instanceof PptxStructure) {
            // PptxStructure object
            $pptxFile = $source;
        } else {
            // file
            $pptxFile = new PptxStructure();
            $pptxFile->parsePptx($source);
        }

        $xmlUtilities = new XmlUtilities();

        $contentTypesXML = $pptxFile->getContent('[Content_Types].xml');
        $contentTypesDOM = $xmlUtilities->generateDomDocument($contentTypesXML);
        $contentTypesXPath = new DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        // get slides from $mainXMLPathNodes to get the correct order of the slides
        $mainXML = $pptxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
        $mainDOM = $xmlUtilities->generateDomDocument($mainXML);
        $mainXPath = new DOMXPath($mainDOM);
        $mainXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $query = '//xmlns:sldIdLst/xmlns:sldId';
        // query by slide number if set
        if (isset($options['slideNumber'])) {
            $query .= '['.$options['slideNumber'].']';
        }
        $slideNodes = $mainXPath->query($query);

        // get slide rels to get the slide contents
        $mainRelsXML = $pptxFile->getContent(str_replace('ppt/', 'ppt/_rels/', substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1)) . '.rels');
        $mainRelsDOM = $xmlUtilities->generateDomDocument($mainRelsXML);
        $mainRelsXPath = new DOMXPath($mainRelsDOM);
        $mainRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $slidesData = array();
        foreach ($slideNodes as $slideNode) {
            $query = '//xmlns:Relationship[@Id="'.$slideNode->getAttribute('r:id').'"]';
            $slideContentNodes = $mainRelsXPath->query($query);
            $slidesData['ppt/' . $slideContentNodes->item(0)->getAttribute('Target')] = $pptxFile->getContent('ppt/' . $slideContentNodes->item(0)->getAttribute('Target'));
        }

        // replace the data
        foreach ($slidesData as $slideKey => $slideValue) {
            $slideDataDOM = $xmlUtilities->generateDomDocument($slideValue);
            $slideXPath = new DOMXPath($slideDataDOM);
            $slideXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            foreach ($data as $dataKey => $dataValue) {
                $this->searchToReplace($slideXPath, $dataKey, $dataValue);
            }

            $slidesData[$slideKey] = $slideDataDOM->saveXML();

            // free DOMDocument resources
            $slideDataDOM = null;
        }

        // save the data in the PPTX file
        foreach ($slidesData as $slideKey => $slideValue) {
            $pptxFile->addContent($slideKey, $slideValue);
        }

        // save file
        $pptxFile->savePptx($target);

        // free DOMDocument resources
        $contentTypesDOM = null;
        $mainDOM = null;
        $mainRelsDOM = null;
    }

    /**
     * Splits a PowerPoint document
     *
     * @access public
     * @param string|PptxStructure $source Path to the presentation
     * @param string $target Path to the resulting PPTX (a new file will be created per slide)
     * @param array $options
     *      'optimizeOutput' (bool): remove not needed contents in the file outputs
     */
    public function splitPptx($source, $target, $options = array())
    {
        if ($source instanceof PptxStructure) {
            // PptxStructure object
            $pptxFile = $source;
        } else {
            // file
            $pptxFile = new PptxStructure();
            $pptxFile->parsePptx($source);
        }

        $xmlUtilities = new XmlUtilities();

        $targetInfo = pathinfo($target);

        $contentTypesXML = $pptxFile->getContent('[Content_Types].xml');
        $contentTypesDOM = $xmlUtilities->generateDomDocument($contentTypesXML);
        $contentTypesXPath = new DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        // get slides from $mainXMLPathNodes to get the correct order of the slides
        $mainXML = $pptxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
        $mainDOM = $xmlUtilities->generateDomDocument($mainXML);
        $mainXPath = new DOMXPath($mainDOM);
        $mainXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $query = '//xmlns:sldIdLst/xmlns:sldId';

        $slideNodes = $mainXPath->query($query);

        // counter used for each new file name
        $i = 0;
        foreach ($slideNodes as $slideNode) {
            // increment the file counter
            $i++;

            $filePptxPath = $targetInfo['dirname'] . '/' . $targetInfo['filename'] . $i . '.' . $targetInfo['extension'];

            $pptxFile = new PptxStructure();
            $pptxFile->parsePptx($source);

            // remove other slides from the PPTX content
            $mainXMLNew = $pptxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
            $mainDOMNew = $xmlUtilities->generateDomDocument($mainXMLNew);
            $mainXPathNew = new DOMXPath($mainDOMNew);
            $mainXPathNew->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            $queryNew = '//xmlns:sldIdLst/xmlns:sldId';
            $slideNodesNew = $mainXPathNew->query($queryNew);

            $j = 1;
            // keep current slide id
            // keep removed slide ids
            $currentSlideRid = null;
            $cleanedSlidesRid = array();
            foreach ($slideNodesNew as $slideNodeNew) {
                if ($i != $j) {
                    $cleanedSlidesRid[] = $slideNodeNew->getAttribute('r:id');
                    $slideNodeNew->parentNode->removeChild($slideNodeNew);
                } else {
                    $currentSlideRid = $slideNodeNew->getAttribute('r:id');
                }
                $j++;
            }

            $pptxFile->addContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1), $mainDOMNew->saveXml());

            // clean output
            if (isset($options['optimizeOutput']) && $options['optimizeOutput']) {
                // clear contents from removed slides

                // get rels content
                $mainXMLRelsCleaned = $pptxFile->getContent(str_replace('ppt/', 'ppt/_rels/', substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1)) . '.rels');
                $mainDOMRelsCleaned = $xmlUtilities->generateDomDocument($mainXMLRelsCleaned);
                $mainXMLRelsCleanedXPath = new DOMXPath($mainDOMRelsCleaned);
                $mainXMLRelsCleanedXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                $filesToBeRemoved = array();

                foreach ($cleanedSlidesRid as $cleanedSlideRid) {
                    $queryNewRid = '//xmlns:Relationship[@Id="'.$cleanedSlideRid.'"]';
                    $nodeRelSlide = $mainXMLRelsCleanedXPath->query($queryNewRid);
                    if ($nodeRelSlide->length > 0) {
                        $targetSlide = $nodeRelSlide->item(0)->getAttribute('Target');
                        $contentSlideRels = $pptxFile->getContent('ppt/' . $targetSlide);
                        // get contents to be removed
                        $contentSlideRelsDOM = $xmlUtilities->generateDomDocument($contentSlideRels);
                        $contentSlideRelsXPath = new DOMXPath($contentSlideRelsDOM);
                        $contentSlideRelsXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                        $contentSlideRelsXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                        $contentSlideRelsXPath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

                        // rels content
                        $contentSlideRelsContent = $pptxFile->getContent('ppt/' . str_replace('slides/', 'slides/_rels/', $targetSlide . '.rels'));
                        if ($contentSlideRelsContent) {
                            $contentSlideRelsContentDOM = $xmlUtilities->generateDomDocument($contentSlideRelsContent);
                            $contentSlideRelsContentXPath = new DOMXPath($contentSlideRelsContentDOM);
                            $contentSlideRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                            // remove images
                            $slideImagesRemove = $contentSlideRelsXPath->query('//a:blip');
                            foreach ($slideImagesRemove as $slideImageRemove) {
                                $queryNodesContentTarget = $contentSlideRelsContentXPath->query('//xmlns:Relationship[@Id="'.$slideImageRemove->getAttribute('r:embed').'"]');
                                if ($queryNodesContentTarget->length > 0) {
                                    foreach ($queryNodesContentTarget as $queryNodeContentTarget) {
                                        $filesToBeRemoved[] = str_replace('../', 'ppt/', $queryNodeContentTarget->getAttribute('Target'));
                                    }
                                }
                            }

                            // free DOMDocument resources
                            $contentSlideRelsContentDOM = null;
                        }

                        // free DOMDocument resources
                        $contentSlideRelsDOM = null;
                    }
                }

                if (count($filesToBeRemoved) > 0) {
                    // keep files to don't be removed
                    $filesToKeep = array();

                    // check if the files to be removed are not used in the slide to be kept
                    $queryCurrentRid = '//xmlns:Relationship[@Id="'.$currentSlideRid.'"]';
                    $nodeRelCurrentSlide = $mainXMLRelsCleanedXPath->query($queryCurrentRid);
                    if ($nodeRelCurrentSlide->length > 0) {
                        $targetCurrentSlide = $nodeRelCurrentSlide->item(0)->getAttribute('Target');
                        // rels content
                        $contentSlideRelsContent = $pptxFile->getContent('ppt/' . str_replace('slides/', 'slides/_rels/', $targetCurrentSlide . '.rels'));
                        if ($contentSlideRelsContent) {
                            $contentSlideRelsContentDOM = $xmlUtilities->generateDomDocument($contentSlideRelsContent);
                            $contentSlideRelsContentXPath = new DOMXPath($contentSlideRelsContentDOM);
                            $contentSlideRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                            foreach ($filesToBeRemoved as $fileToBeRemoved) {
                                $queryNodesContentTarget = $contentSlideRelsContentXPath->query('//xmlns:Relationship[@Target="'.$fileToBeRemoved.'"]');
                                if ($queryNodesContentTarget->length > 0) {
                                    // the file exists in the slide, don't remove it
                                    $filesToKeep[] = $filesToBeRemoved;
                                }
                            }

                            // free DOMDocument resources
                            $contentSlideRelsContentDOM = null;
                        }
                    }

                    // delete not needed files
                    foreach ($filesToBeRemoved as $fileToBeRemoved) {
                        // delete only if the file must not be kept
                        if (!in_array($fileToBeRemoved, $filesToKeep)) {
                            $pptxFile->deleteContent($fileToBeRemoved);
                        }
                    }
                }

                // free DOMDocument resources
                $mainDOMRelsCleaned = null;
            }

            // save file
            $pptxFile->savePptx($filePptxPath);

            // free DOMDocument resources
            $mainDOMNew = null;
        }

        // free DOMDocument resources
        $contentTypesDOM = null;
        $mainDOM = null;
    }

    /**
     * This is the method that selects the nodes that need to be manipulated
     *
     * @access private
     * @param DOMXPath $xPath the node to be changed
     * @param string $searchTerm
     * @param string $replaceTerm
     */
    private function searchToReplace($xPath, $searchTerm, $replaceTerm)
    {
        $query = '//xmlns:t';
        $tNodes = $xPath->query($query);
        $xmlUtilities = new XmlUtilities();
        $searchTerm = $xmlUtilities->parseAndCleanTextString($searchTerm);
        $replaceTerm = $xmlUtilities->parseAndCleanTextString($replaceTerm);

        foreach ($tNodes as $tNode) {
            if (strstr($tNode->nodeValue, $searchTerm)) {
                $tNode->nodeValue = str_replace($searchTerm, $replaceTerm, $tNode->nodeValue);
            }
        }
    }
}
<?php

/**
 * Return information of a PPTX file
 *
 * @category   Phppptx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
require_once __DIR__ . '/CreatePptx.php';

class Indexer
{
    /**
     * @var PptxStructure
     */
    private $documentPptx;

    /**
     * @var XmlUtilities XML Utilities classes
     */
    private $xmlUtilities;

    /**
     * @var array Stores the file internal structure
     */
    private $pptxStructure;

    /**
     * Class constructor
     *
     * @param mixed $source File path or PptxStructure
     */
    public function __construct($source)
    {
        if ($source instanceof PptxStructure) {
            $this->documentPptx = $source;
        } else {
            $this->documentPptx = new PptxStructure();
            $this->documentPptx->parsePptx($source);
        }

        // XMLUtilites class
        $this->xmlUtilities = new XmlUtilities();

        // init the document structure array as empty
        $this->pptxStructure = array(
            'comments' => array(
                'authors' => array(),
                'comments' => array(),
            ),
            'layouts' => array(
                'masters' => array(),
                'slides' => array(),
            ),
            'presentation' => array(
                'sizes' => array(),
                'slides' => array(),
            ),
            'properties' => array(
                'core' => array(),
                'custom' => array(),
            ),
            'slides' => array(),
            'signatures' => array(),
        );

        // parse the document
        $this->parse($source);
    }

    /**
     * Return a file as array or JSON
     *
     * @param string $output Output type: 'array' (default), 'json'
     * @return mixed $output
     * @throws Exception If the output type format not supported
     */
    public function getOutput($output = 'array')
    {
        // if the chosen output type is not supported throw an exception
        if (!in_array($output, array('array', 'json'))) {
            throw new Exception('The output "' . $output . '" is not supported');
        }

        // output the document after index
        return $this->output($output);
    }

    /**
     * Extract comment authors from an XML string
     *
     * @param string $xml XML string
     */
    protected function extractCommentAuthors($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        $cmAuthorTags = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cmAuthor');
        if ($cmAuthorTags->length > 0) {
            foreach ($cmAuthorTags as $cmAuthorTag) {
                $this->pptxStructure['comments']['authors'][] = array(
                    'id' => $cmAuthorTag->getAttribute('id'),
                    'name' => $cmAuthorTag->getAttribute('name'),
                );
            }
        }
    }

    /**
     * Extract master layout from an XML string
     *
     * @param string $xml XML string
     * @param string $contentTarget Target
     */
    protected function extractLayoutMaster($xml, $contentTarget)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // load XML rels information
        $contentRelsTarget = str_replace('slideMasters/', 'slideMasters/_rels/', $contentTarget) . '.rels';
        $contentRels = $this->documentPptx->getContent($contentRelsTarget);
        $contentRelsDOM = null;
        if (!empty($contentRels)) {
            $contentRelsDOM = $this->xmlUtilities->generateDomDocument($contentRels);
            $contentRelsXPath = new DOMXPath($contentRelsDOM);
            $contentRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        }

        // layouts
        $layoutsMaster = array();
        $sldLayoutIdLstTags = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutIdLst');
        if ($sldLayoutIdLstTags->length > 0) {
            foreach ($sldLayoutIdLstTags as $sldLayoutIdLstTag) {
                $sldLayoutIdTags = $sldLayoutIdLstTag->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutId');
                if ($sldLayoutIdTags->length > 0) {
                    foreach ($sldLayoutIdTags as $sldLayoutIdTag) {
                        $targetRelationship = '';
                        if (isset($contentRelsDOM)) {
                            $nodesRelationship = $contentRelsXPath->query('//xmlns:Relationship[@Id="' . $sldLayoutIdTag->getAttribute('r:id') . '"]');
                            if ($nodesRelationship->length > 0) {
                                if ($nodesRelationship->item(0)->hasAttribute('Target')) {
                                    $targetRelationship = $nodesRelationship->item(0)->getAttribute('Target');
                                }
                            }
                        }

                        $layoutsMaster[] = array(
                            'rId' => $sldLayoutIdTag->getAttribute('r:id'),
                            'target' => $targetRelationship
                        );
                    }
                }
            }
        }

        $this->pptxStructure['layouts']['masters'][] = array(
            'layouts' => $layoutsMaster,
        );

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract slide layout from an XML string
     *
     * @param string $xml XML string
     * @param string $contentTarget Target
     */
    protected function extractLayoutSlide($xml, $contentTarget)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);
        $contentXPath = new DOMXPath($contentDOM);
        $contentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // name
        $layoutName = '';
        $nodesClSd = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cSld');
        if ($nodesClSd->length > 0) {
            if ($nodesClSd->item(0)->hasAttribute('name')) {
                $layoutName = $nodesClSd->item(0)->getAttribute('name');
            }
        }

        // placeholders
        $layoutPlaceholders = array();
        $nodesCNvPr = $contentXPath->query('//p:sp/p:nvSpPr/p:cNvPr[@name]');
        if ($nodesCNvPr->length > 0) {
            foreach ($nodesCNvPr as $nodeCNvPr) {
                $placeholderId = '';
                $placeholderName = '';
                if ($nodeCNvPr->hasAttribute('id')) {
                    $placeholderId = $nodeCNvPr->getAttribute('id');
                }
                if ($nodeCNvPr->hasAttribute('name')) {
                    $placeholderName = $nodeCNvPr->getAttribute('name');
                }
                $layoutPlaceholders[] = array(
                    'id' => $placeholderId,
                    'name' => $placeholderName,
                );
            }
        }

        $this->pptxStructure['layouts']['slides'][] = array(
            'name' => $layoutName,
            'placeholders' => $layoutPlaceholders,
        );

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract presentation contents from an XML string
     *
     * @param string $xml XML string
     */
    protected function extractPresentation($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // sldMasterIdLst/sldMasterId tags
        $sldMasterIdLstTags = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
        if ($sldMasterIdLstTags->length > 0) {
            foreach ($sldMasterIdLstTags as $sldMasterIdLstTag) {
                $sldMasterIdTags = $sldMasterIdLstTag->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterId');
                if ($sldMasterIdTags->length > 0) {
                    foreach ($sldMasterIdTags as $sldMasterIdTag) {
                        $this->pptxStructure['presentation']['mastersLayouts'][] = array(
                            'id' => $sldMasterIdTag->getAttribute('id'),
                            'rId' => $sldMasterIdTag->getAttribute('r:id'),
                        );
                    }
                }
            }
        }

        // sldIdLst/sldId tags
        $sldIdLstTags = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
        if ($sldIdLstTags->length > 0) {
            foreach ($sldIdLstTags as $sldIdLstTag) {
                $sldIdTags = $sldIdLstTag->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldId');
                if ($sldIdTags->length > 0) {
                    foreach ($sldIdTags as $sldIdTag) {
                        $this->pptxStructure['presentation']['slides'][] = array(
                            'id' => $sldIdTag->getAttribute('id'),
                            'rId' => $sldIdTag->getAttribute('r:id'),
                        );
                    }
                }
            }
        }

        // p14:sectionLst tags
        $sectionLstTags = $contentDOM->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
        if ($sectionLstTags->length > 0) {
            $this->pptxStructure['presentation']['sections'] = array();
            foreach ($sectionLstTags as $sectionLstTag) {
                $sectionTags = $sectionLstTag->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'section');
                if ($sectionTags->length > 0) {
                    foreach ($sectionTags as $sectionTag) {
                        // section name
                        $nameSection = '';
                        if ($sectionTag->hasAttribute('name')) {
                            $nameSection = $sectionTag->getAttribute('name');
                        }

                        // section slides
                        $slidesSection = array();
                        $slideSldIdLstTags = $sectionTag->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldIdLst');
                        if ($slideSldIdLstTags->length > 0) {
                            foreach ($slideSldIdLstTags as $slideSldIdLstTag) {
                                $slideSldIdTags = $slideSldIdLstTag->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldId');
                                if ($slideSldIdTags->length > 0) {
                                    foreach ($slideSldIdTags as $slideSldIdTag) {
                                        if ($slideSldIdTag->hasAttribute('id')) {
                                            $slidesSection[] = $slideSldIdTag->getAttribute('id');
                                        }
                                    }
                                }
                            }
                        }
                        $this->pptxStructure['presentation']['section'][] = array(
                            'name' => $nameSection,
                            'slides' => $slidesSection,
                        );
                    }
                }
            }
        }

        // sldSz tag
        $sldSzTags = $contentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldSz');
        if ($sldSzTags->length > 0) {
            $this->pptxStructure['presentation']['sizes'] = array();
            if ($sldSzTags->item(0)->hasAttribute('cx')) {
                $this->pptxStructure['presentation']['sizes']['width'] = $sldSzTags->item(0)->getAttribute('cx');
            }
            if ($sldSzTags->item(0)->hasAttribute('cy')) {
                $this->pptxStructure['presentation']['sizes']['height'] = $sldSzTags->item(0)->getAttribute('cy');
            }
            if ($sldSzTags->item(0)->hasAttribute('type')) {
                $this->pptxStructure['presentation']['sizes']['type'] = $sldSzTags->item(0)->getAttribute('type');
            }
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract document properties from an XML string
     *
     * @param string $xml XML string
     * @param string $target Properties target: core, custom
     */
    protected function extractProperties($xml, $target)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        if ($target == 'core') {
            // do a global xpath query getting only text tags
            $contentXpath = new DOMXPath($contentDOM);
            $contentXpath->registerNamespace('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
            $propertiesEntries = $contentXpath->query('//cp:coreProperties');

            if ($propertiesEntries->item(0)->childNodes->length > 0) {
                foreach ($propertiesEntries->item(0)->childNodes as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $propertyEntryFullName = explode(':', $propertyEntry->tagName);
                    $nameProperty = $propertyEntryFullName[1];

                    $this->pptxStructure['properties']['core'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        } else if ($target == 'custom') {
            // do a global xpath query getting only property tags
            $contentXpath = new DOMXPath($contentDOM);
            $contentXpath->registerNamespace('ns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
            $propertiesEntries = $contentXpath->query('//ns:Properties//ns:property');

            if ($propertiesEntries->length > 0) {
                foreach ($propertiesEntries as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $nameProperty = $propertyEntry->getAttribute('name');

                    $this->pptxStructure['properties']['custom'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract slide contents from an XML string
     *
     * @param string $xml XML string
     * @param string $contentTarget Target
     */
    protected function extractSlide($xml, $contentTarget)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        $contentXPath = new DOMXPath($contentDOM);
        $contentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $contentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // rels content
        $relsFilePath = str_replace('slides/', 'slides/_rels/', $contentTarget) . '.rels';
        $contentSlideRels = $this->documentPptx->getContent($relsFilePath);
        if (!empty($contentSlideRels)) {
            $contentSlideRelsDom = $this->xmlUtilities->generateDomDocument($contentSlideRels);
            $contentSlideRelsXPath = new DOMXPath($contentSlideRelsDom);
            $contentSlideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        }

        // text
        $slideTexts = $this->extractTexts($xml);

        // comments
        $slideComments = array();
        if (!empty($contentSlideRels)) {
            $nodesComment = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]');
            if ($nodesComment->length > 0) {
                foreach ($nodesComment as $nodeComment) {
                    $targetComment = $nodeComment->getAttribute('Target');
                    $targetCommentFilePath = str_replace('../', 'ppt/', $targetComment);
                    $contentComment = $this->documentPptx->getContent($targetCommentFilePath);
                    if (!empty($contentComment)) {
                        $contentCommentDOM = $this->xmlUtilities->generateDomDocument($contentComment);
                        $nodesCm = $contentCommentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cm');
                        if ($nodesCm->length > 0) {
                            foreach ($nodesCm as $nodeCm) {
                                $slideComment = array();

                                if ($nodeCm->hasAttribute('authorId')) {
                                    $slideComment['authorId'] = $nodeCm->getAttribute('authorId');
                                }
                                if ($nodeCm->hasAttribute('date')) {
                                    $slideComment['date'] = $nodeCm->getAttribute('date');
                                }
                                if ($nodeCm->hasAttribute('idx')) {
                                    $slideComment['idx'] = $nodeCm->getAttribute('idx');
                                }

                                $nodesPosition = $nodeCm->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'pos');
                                if ($nodesPosition->length > 0) {
                                    $slideComment['position'] = array(
                                        'coordinateX' => $nodesPosition->item(0)->getAttribute('x'),
                                        'coordinateY' => $nodesPosition->item(0)->getAttribute('y'),
                                    );
                                }

                                $nodesText = $nodeCm->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'text');
                                if ($nodesText->length > 0) {
                                    $slideComment['text'] = $nodesText->item(0)->textContent;
                                }

                                $slideComments[] = $slideComment;
                            }
                        }
                    }
                }
            }
        }

        // layout
        $slideLayout = array();
        if (!empty($contentSlideRels)) {
            $nodesSlideLayout = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
            if ($nodesSlideLayout->length > 0) {
                foreach ($nodesSlideLayout as $nodeSlideLayout) {
                    $targetSlideLayout = '';
                    if ($nodeSlideLayout->hasAttribute('Target')) {
                        $targetSlideLayout = $nodeSlideLayout->getAttribute('Target');
                    }
                    $nameSlideLayout = '';
                    if (!empty($targetSlideLayout)) {
                        $slideLayoutFilePath = str_replace('../', 'ppt/', $targetSlideLayout);
                        $contentSlideLayout = $this->documentPptx->getContent($slideLayoutFilePath);
                        if (!empty($contentSlideLayout)) {
                            $contentDOMSlideLayout = $this->xmlUtilities->generateDomDocument($contentSlideLayout);
                            $nodesClSd = $contentDOMSlideLayout->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cSld');
                            if ($nodesClSd->length > 0) {
                                if ($nodesClSd->item(0)->hasAttribute('name')) {
                                    $nameSlideLayout = $nodesClSd->item(0)->getAttribute('name');
                                }
                            }
                        }
                    }

                    $slideLayout = array(
                        'name' => $nameSlideLayout,
                        'Target' => $targetSlideLayout,
                    );
                }
            }
        }

        // notes
        $slideNotes = array();
        if (!empty($contentSlideRels)) {
            $nodesSlideNote = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]');
            if ($nodesSlideNote->length > 0) {
                foreach ($nodesSlideNote as $nodeSlideNote) {
                    $targetSlideNote = '';
                    if ($nodeSlideNote->hasAttribute('Target')) {
                        $targetSlideNote = $nodeSlideNote->getAttribute('Target');
                    }
                    $slideNoteFilePath = str_replace('../', 'ppt/', $targetSlideNote);
                    $contentSlideNote = $this->documentPptx->getContent($slideNoteFilePath);
                    if (!empty($contentSlideNote)) {
                        $slideNoteTexts = $this->extractTexts($contentSlideNote);

                        $slideNotes = array(
                            'text' => $slideNoteTexts,
                            'Target' => $targetSlideNote,
                        );
                    }
                }
            }
        }

        // images
        $slideImages = array();
        if (!empty($contentSlideRels)) {
            $slideImages = $this->getPicInformation($contentXPath, $contentSlideRelsXPath, 'a:blip', 'r:embed');
        }

        // audios
        $slideAudios = array();
        if (!empty($contentSlideRels)) {
            $slideAudios = $this->getPicInformation($contentXPath, $contentSlideRelsXPath, 'a:audioFile', 'r:link');
        }

        // videos
        $slideVideos = array();
        if (!empty($contentSlideRels)) {
            $slideVideos = $this->getPicInformation($contentXPath, $contentSlideRelsXPath, 'a:videoFile', 'r:link');
        }

        // links
        $slideLinks = array();
        $nodesHlinkClick = $contentXPath->query('//a:hlinkClick');
        if ($nodesHlinkClick->length > 0) {
            foreach ($nodesHlinkClick as $nodeHlinkClick) {
                if ($nodeHlinkClick->hasAttribute('r:id')) {
                    if (!empty($contentSlideRelsDom) && isset($contentSlideRelsXPath)) {
                        $nodesRelationship = $contentSlideRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeHlinkClick->getAttribute('r:id') . '"]');
                        if ($nodesRelationship->length > 0) {
                            $target = '';
                            if ($nodesRelationship->item(0)->hasAttribute('Target')) {
                                $target = $nodesRelationship->item(0)->getAttribute('Target');
                            }

                            $slideLinks[] = array(
                                'target' => $target,
                            );
                        }
                    }
                }
            }
        }

        // placeholders
        $slidePlaceholders = array();
        $nodesPSp = $contentXPath->query('//p:sp');
        if ($nodesPSp->length > 0) {
            foreach ($nodesPSp as $nodePSp) {
                $nodesCNvPr = $contentXPath->query('.//p:cNvPr', $nodePSp);
                $placeholderPosition = array();
                if ($nodesCNvPr->length > 0) {
                    if ($nodesCNvPr->item(0)->hasAttribute('name')) {
                        $placeholderPosition['name'] = $nodesCNvPr->item(0)->getAttribute('name');
                    }
                    if ($nodesCNvPr->item(0)->hasAttribute('descr')) {
                        $placeholderPosition['descr'] = $nodesCNvPr->item(0)->getAttribute('descr');
                    }
                }
                $nodesPh = $contentXPath->query('.//p:ph', $nodePSp);
                if ($nodesPh->length > 0) {
                    if ($nodesPh->item(0)->hasAttribute('type')) {
                        $placeholderPosition['type'] = $nodesPh->item(0)->getAttribute('type');
                    }
                }
                $slidePlaceholders[] = $placeholderPosition;
            }
        }

        $this->pptxStructure['slides'][] = array(
            'audios' => $slideAudios,
            'comments' => $slideComments,
            'images' => $slideImages,
            'links' => $slideLinks,
            'layout' => $slideLayout,
            'notes' => $slideNotes,
            'placeholders' => $slidePlaceholders,
            'text' => $slideTexts,
            'videos' => $slideVideos,
        );

        // add comments to the comments key
        if (count($slideComments) > 0) {
            //$slideComments = array_merge($slideComments);
            $this->pptxStructure['comments']['comments'][] = $slideComments;
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract signature contents from an XML string
     *
     * @param string $xml XML string
     */
    protected function extractSignature($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // get X509Certificate
        $contentXpath = new DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $x509CertificateEntry = $contentXpath->query('//xmlns:X509Certificate');
        $x509CertificateContent = null;
        if ($x509CertificateEntry->length > 0) {
            $x509Reader = openssl_x509_read("-----BEGIN CERTIFICATE-----\n" . $x509CertificateEntry->item(0)->textContent . "\n-----END CERTIFICATE-----\n");
            if ($x509Reader) {
                $x509CertificateContent = openssl_x509_parse($x509Reader);
            }
        }

        // get SignatureProperties time
        $contentXpath = new DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $contentXpath->registerNamespace('mdssi', 'http://schemas.openxmlformats.org/package/2006/digital-signature');
        $signatureTime = $contentXpath->query('//xmlns:SignatureProperties//mdssi:SignatureTime/mdssi:Value');
        $signatureTimeContent = null;
        if ($signatureTime->length > 0) {
            $signatureTimeContent = $signatureTime->item(0)->textContent;
        }

        // get SignatureProperties comment
        $contentXpath = new DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $contentXpath->registerNamespace('xmlnsdigsig', 'http://schemas.microsoft.com/office/2006/digsig');
        $signatureComment = $contentXpath->query('//xmlns:SignatureProperties//xmlnsdigsig:SignatureInfoV1//xmlnsdigsig:SignatureComments');
        $signatureCommentContent = null;
        if ($signatureComment->length > 0) {
            $signatureCommentContent = $signatureComment->item(0)->textContent;
        }

        $this->pptxStructure['signatures'][] = array(
            'SignatureComment' => $signatureCommentContent,
            'SignatureTime' => $signatureTimeContent,
            'X509Certificate' => $x509CertificateContent,
        );

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract text contents from an XML string
     *
     * @param string $xml XML string
     * @return string Text content
     */
    protected function extractTexts($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // do a global xpath query getting only text tags
        $contentXpath = new DOMXPath($contentDOM);
        $contentXpath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $paragraphEntries = $contentXpath->query('//a:p[./a:r/a:t]');

        // iterate text content and extract text strings. Add a blank space to separate each string
        $content = '';
        foreach ($paragraphEntries as $paragraphEntry) {
            // if empty paragraph avoid adding the content
            if (empty($paragraphEntry->textContent)) {
                continue;
            }
            $textEntries = $paragraphEntry->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');
            foreach ($textEntries as $textEntry) {
                $content .= $textEntry->textContent;
            }
            $content .= ' ';
        }

        // free DOMDocument resources
        $contentDOM = null;

        return trim($content);
    }

    /**
     * Get pic information
     *
     * @param DOMXPath $contentXPath Content XPath
     * @param DOMXPath $contentSlideRelsXPath Content rels XPath
     * @param string $type Output type: 'array' (default), 'json'
     * @return array Pic information
     */
    protected function getPicInformation($contentXPath, $contentSlideRelsXPath, $type, $attribute)
    {
        $slidePic = array();

        $nodesBlip = $contentXPath->query('//p:pic//' . $type);
        if ($nodesBlip->length > 0) {
            foreach ($nodesBlip as $nodeBlip) {
                if ($nodeBlip->hasAttribute($attribute)) {
                    $nodesRelationship = $contentSlideRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeBlip->getAttribute($attribute) . '"]');

                    // path
                    $pathPic = '';
                    if ($nodesRelationship->length > 0 && $nodesRelationship->item(0)->hasAttribute('Target')) {
                        $pathPic = $nodesRelationship->item(0)->getAttribute('Target');
                    }

                    // sizes
                    $heightPic = '';
                    $heightOffsetPic = '';
                    $widthPic = '';
                    $widthOffsetPic = '';
                    if (isset($nodeBlip->parentNode) && isset($nodeBlip->parentNode->parentNode)) {
                        $nodesXfrm = $contentXPath->query('.//a:xfrm', $nodeBlip->parentNode->parentNode);
                        if ($nodesXfrm->length == 0) {
                            // audio and video are a deeper child
                            $nodesXfrm = $contentXPath->query('.//a:xfrm', $nodeBlip->parentNode->parentNode->parentNode);
                        }

                        if ($nodesXfrm->length > 0) {
                            $nodesXfrmExt = $nodesXfrm->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ext');
                            if ($nodesXfrmExt->length > 0) {
                                $heightPic = $nodesXfrmExt->item(0)->getAttribute('cy');
                                $widthPic = $nodesXfrmExt->item(0)->getAttribute('cx');
                            }
                            $nodesXfrmOff = $nodesXfrm->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'off');
                            if ($nodesXfrmOff->length > 0) {
                                $heightOffsetPic = $nodesXfrmOff->item(0)->getAttribute('y');
                                $widthOffsetPic = $nodesXfrmOff->item(0)->getAttribute('x');
                            }
                        }
                    }

                    // alt and descr values
                    $altTextDescrPic = '';
                    if (isset($nodeBlip->parentNode) && isset($nodeBlip->parentNode->parentNode)) {
                        $nodesCNvPr = $contentXPath->query('.//p:cNvPr', $nodeBlip->parentNode->parentNode);
                        if ($nodesCNvPr->length > 0) {
                            if ($nodesCNvPr->item(0)->hasAttribute('descr')) {
                                $altTextDescrPic = $nodesCNvPr->item(0)->getAttribute('descr');
                            }
                        }
                    }

                    $slidePic[] = array(
                        'path' => $pathPic,
                        'height' => $heightPic,
                        'heightOffset' => $heightOffsetPic,
                        'width' => $widthPic,
                        'widthOffset' => $widthOffsetPic,
                        'altTextDescr' => $altTextDescrPic,
                    );
                }
            }
        }

        return $slidePic;
    }

    /**
     * Return a file as array or JSON
     *
     * @param string $type Output type: 'array' (default), 'json'
     * @return mixed Output
     */
    protected function output($type = 'array')
    {
        // array as default
        $output = $this->pptxStructure;

        // export as the choosen type
        if ($type == 'json') {
            $output = json_encode($output);
        }

        return $output;
    }

    /**
     * Parse a PPTX file
     *
     * @param PptxStructure $source
     */
    private function parse($source)
    {
        // parse the Content_Types
        $contentTypesContent = $this->documentPptx->getContent('[Content_Types].xml');
        $contentTypesXml = $this->xmlUtilities->generateSimpleXmlElement($contentTypesContent);
        $contentTypesDom = dom_import_simplexml($contentTypesXml);

        $contentTypesXpath = new DOMXPath($contentTypesDom->ownerDocument);
        $contentTypesXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/content-types');
        $relsEntries = $contentTypesXpath->query('//rel:Default[@ContentType="application/vnd.openxmlformats-package.relationships+xml"]');
        $relsExtension = 'rels';
        if (isset($relsEntries[0])) {
            $relsExtension = $relsEntries[0]->getAttribute('Extension');
        }

        // iterate over the Content_Types
        foreach ($contentTypesXml->Override as $override) {
            foreach ($override->attributes() as $attribute => $value) {
                // get the file content
                $contentTarget = substr($override->attributes()->PartName, 1);
                $content = $this->documentPptx->getContent($contentTarget);

                // before adding a content remove the first character to get the right file path
                // removing the first slash of each path
                if ($value == 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml') {
                    // presentation content

                    // extract presentation
                    $this->extractPresentation($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml') {
                    // slide content

                    // extract slide
                    $this->extractSlide($content, $contentTarget);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml') {
                    // slide layout content

                    // extract layout
                    $this->extractLayoutSlide($content, $contentTarget);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml') {
                    // master layout content

                    // extract layout
                    $this->extractLayoutMaster($content, $contentTarget);
                } else if ($value == 'application/vnd.openxmlformats-package.core-properties+xml') {
                    // core properties content

                    // extract core properties
                    $this->extractProperties($content, 'core');
                } else if ($value == 'application/vnd.openxmlformats-officedocument.custom-properties+xml') {
                    // custom properties content

                    // extract custom properties
                    $this->extractProperties($content, 'custom');
                } else if ($value == 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml') {
                    // signature contents

                    // extract signatures
                    $this->extractSignature($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml') {
                    // comment author contents

                    // extract comment authors
                    $this->extractCommentAuthors($content);
                }
            }
        }

        // free DOMDocument resources
        $contentTypesDom = null;
    }
}
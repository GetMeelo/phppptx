<?php

/**
 * Merge PPTX
 *
 * @category   Phppptx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
require_once __DIR__ . '/CreatePptx.php';

class MergePptx
{
    /**
     * XmlUtilities
     *
     * @access protected
     * @var XmlUtilities XML Utilities classes
     */
    protected $xmlUtilities;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->xmlUtilities = new XmlUtilities();
    }

    /**
     * Merges PPTX
     *
     * @access public
     * @param array $source PPTX files to be merged (string or PptxStructure)
     * @param string $target path to the output presentation
     * @param array $options
     *      'mergeSections' (bool) if true, sections from PPTX to be merged remain. Default as true
     */
    public function merge($source, $target, $options = array())
    {
        // default options
        if (!isset($options['mergeSections'])) {
            $options['mergeSections'] = true;
        }

        // parse first PPTX
        $sourcePptx = array_shift($source);
        if ($sourcePptx instanceof PptxStructure) {
            // PptxStructure object
            $pptx = $sourcePptx;
        } else {
            // file
            $pptx = new PptxStructure();
            $pptx->parsePptx($sourcePptx);
        }
        $contentTypesXML = $pptx->getContent('[Content_Types].xml');
        $contentTypesDOM = $this->xmlUtilities->generateDomDocument($contentTypesXML);
        $contentTypesXPath = new DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        $presentations = $pptx->getContentByType('presentations');
        $presentationXML = $presentations[0]['content'];
        $presentationDOM = $this->xmlUtilities->generateDomDocument($presentationXML);
        $presentationRelsPath = str_replace('ppt/', 'ppt/_rels/', $presentations[0]['path']) . '.rels';
        $presentationRelsXML = $pptx->getContent($presentationRelsPath);
        $presentationRelsDOM = $this->xmlUtilities->generateDomDocument($presentationRelsXML);
        $presentationRelsXPath = new DOMXPath($presentationRelsDOM);
        $presentationRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // get sldMasterId
        $nodesSldMasterId = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterId');
        $sldMasterIdValue = 2147483648;
        foreach ($nodesSldMasterId as $nodeSldMasterId) {
            if ((int)$nodeSldMasterId->getAttribute('id') > $sldMasterIdValue) {
                $sldMasterIdValue = (int)$nodeSldMasterId->getAttribute('id');
            }
            if ($nodeSldMasterId->getAttribute('r:id')) {
                $nodesRelationship = $presentationRelsXPath->query('//xmlns:Relationship[@Id="'.$nodeSldMasterId->getAttribute('r:id').'"]');
                if ($nodesRelationship->length > 0) {
                    $nodeSlidesMasterContent = $pptx->getContent('ppt/' . $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target')));
                    $nodeSlidesMasterDOM = $this->xmlUtilities->generateDomDocument($nodeSlidesMasterContent);
                    $nodesSldLayoutIdLst = $nodeSlidesMasterDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutIdLst');
                    if ($nodesSldLayoutIdLst->length > 0) {
                        $nodesSldLayoutId = $nodesSldLayoutIdLst->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutId');
                        if ($nodesSldLayoutId->length > 0) {
                            foreach ($nodesSldLayoutId as $nodeSldLayoutId) {
                                if ($nodeSldLayoutId->hasAttribute('id')) {
                                    if ((int)$nodeSldLayoutId->getAttribute('id') > $sldMasterIdValue) {
                                        $sldMasterIdValue = (int)$nodeSldLayoutId->getAttribute('id');
                                    }
                                }
                            }
                        }
                    }

                    // free resources
                    $nodeSlidesMasterDOM = null;
                }
            }
        }
        $sldMasterIdValue++;

        // get sldId
        $sldIdValue = 256;
        $nodesSldIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
        if ($nodesSldIdLst->length > 0) {
            // the presentation includes slides
            $sldIdTags = $nodesSldIdLst->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldId');
            if ($sldIdTags->length > 0) {
                $sldIdValue = 1;
                // generate a new ID from existing values
                foreach ($sldIdTags as $sldIdTag) {
                    if ($sldIdTag->hasAttribute('id')) {
                        if ((int)$sldIdTag->getAttribute('id') > $sldIdValue) {
                            $sldIdValue = (int)$sldIdTag->getAttribute('id');
                        }
                    }
                }
                $sldIdValue++;
            }
        }

        // commentAuthors
        $commentAuthors = $pptx->getContent('ppt/commentAuthors.xml');

        // merge PPTX files
        foreach ($source as $sourceNewPptx) {
            if ($sourceNewPptx instanceof PptxStructure) {
                // PptxStructure object
                $pptxNew = $sourceNewPptx;
            } else {
                // file
                $pptxNew = new PptxStructure();
                $pptxNew->parsePptx($sourceNewPptx);
            }

            // merge Default tags from ContentTypes
            $contentTypesNewXML = $pptxNew->getContent('[Content_Types].xml');
            $contentTypesNewDOM = $this->xmlUtilities->generateDomDocument($contentTypesNewXML);
            $contentTypesNewXPath = new DOMXPath($contentTypesNewDOM);
            $contentTypesNewXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
            $nodesDefaultNew = $contentTypesNewXPath->query('//xmlns:Default');
            foreach ($nodesDefaultNew as $nodeDefaultNew) {
                if ($nodeDefaultNew->hasAttribute('Extension') && $nodeDefaultNew->hasAttribute('ContentType')) {
                    $defaultExtensionNew = $nodeDefaultNew->getAttribute('Extension');
                    $defaultContentTypeNew = $nodeDefaultNew->getAttribute('ContentType');
                    if (stripos($contentTypesDOM->saveXML(), 'Extension="' . strtolower($defaultExtensionNew) . '"') === false) {
                        $defaultNew = '<Default Extension="' . $defaultExtensionNew . '" ContentType="' . $defaultContentTypeNew . '"></Default>';
                        $defaultFragment = $contentTypesDOM->createDocumentFragment();
                        $defaultFragment->appendXML($defaultNew);
                        $contentTypesDOM->documentElement->appendChild($defaultFragment);
                    }
                }
            }

            // get presentation contents
            $presentationsNew = $pptxNew->getContentByType('presentations');
            $presentationNewXML = $presentationsNew[0]['content'];
            $presentationNewDOM = $this->xmlUtilities->generateDomDocument($presentationNewXML);
            $presentationRelsNewPath = str_replace('ppt/', 'ppt/_rels/', $presentationsNew[0]['path']) . '.rels';
            $presentationRelsNewXML = $pptxNew->getContent($presentationRelsNewPath);
            $presentationRelsNewDOM = $this->xmlUtilities->generateDomDocument($presentationRelsNewXML);
            $presentationRelsNewXPath = new DOMXPath($presentationRelsNewDOM);
            $presentationRelsNewXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

            // slide layouts
            $slidesLayoutContents = array();
            $layoutsNew = $pptxNew->getContentByType('slideLayouts');
            foreach ($layoutsNew as $layoutNew) {
                $newId = $this->generateUniqueId();

                // file content
                $pptx->addContent('ppt/slideLayouts/slideLayout' . $newId . '.xml', $layoutNew['content']);
                $slidesLayoutContents[] = array(
                    'new' => 'slideLayout' . $newId . '.xml',
                    'old' => str_replace('ppt/slideLayouts/', '', $layoutNew['path']),
                );

                // rels
                $nodeLayoutRelsContent = $pptxNew->getContent(str_replace('slideLayouts/', 'slideLayouts/_rels/', $layoutNew['path']) . '.rels');

                // ContentType
                $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" PartName="/ppt/slideLayouts/slideLayout' . $newId . '.xml"/>', $contentTypesDOM);

                // rels DOM
                $nodeLayoutRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeLayoutRelsContent);
                $nodeLayoutRelsContentXPath = new DOMXPath($nodeLayoutRelsContentDOM);
                $nodeLayoutRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                // images
                $nodesImage = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesImage, 'image', 'media');

                // media
                $nodesMedia = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/media"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesMedia, 'media', 'media');

                // inks
                $nodesInk = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"]');
                $nodeLayoutRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $contentTypesDOM, $nodesInk, 'ink', 'ink', 'application/inkml+xml');

                // 3dmodels
                $nodesModel3d = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesModel3d, 'media', 'media');

                // objects
                $nodesOleObject = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesOleObject, 'embeddings', 'embeddings');
                $nodesPackage = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesPackage, 'embeddings', 'embeddings');

                // tags
                $nodesTags = $nodeLayoutRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"]');
                $nodeLayoutRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeLayoutRelsContent, $nodesTags, 'tags', 'tags');

                $pptx->addContent('ppt/slideLayouts/_rels/slideLayout' . $newId . '.xml.rels', $nodeLayoutRelsContent);
            }

            // themes
            $themesContents = array();
            $themesNew = $pptxNew->getContentByType('themes');
            foreach ($themesNew as $themeNew) {
                $newId = $this->generateUniqueId();

                // rels
                $nodeThemeRelsContent = $pptxNew->getContent(str_replace('theme/', 'theme/_rels/', $themeNew['path']) . '.rels');
                if ($nodeThemeRelsContent) {
                    // rels DOM
                    $nodeThemeRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeThemeRelsContent);
                    $nodeThemeRelsContentXPath = new DOMXPath($nodeThemeRelsContentDOM);
                    $nodeThemeRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                    // images
                    $nodesImage = $nodeThemeRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');
                    $nodeThemeRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeThemeRelsContent, $nodesImage, 'image', 'media');

                    // media
                    $nodesMedia = $nodeThemeRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/media"]');
                    $nodeThemeRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeThemeRelsContent, $nodesMedia, 'media', 'media');

                    $pptx->addContent('ppt/theme/_rels/theme' . $newId . '.xml.rels', $nodeThemeRelsContent);
                }

                // file content
                $pptx->addContent('ppt/theme/theme' . $newId . '.xml', $themeNew['content']);
                $themesContents[] = array(
                    'new' => 'theme' . $newId . '.xml',
                    'old' => str_replace('ppt/theme/', '', $themeNew['path']),
                );

                // ContentType
                $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/ppt/theme/theme' . $newId . '.xml"/>', $contentTypesDOM);
            }

            // slidesMaster
            $slidesMasterContents = array();
            $nodesSldMasterIdNew = $presentationNewDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterId');
            foreach ($nodesSldMasterIdNew as $nodeSldMasterIdNew) {
                if ($nodeSldMasterIdNew->hasAttribute('r:id')) {
                    $nodesRelationship = $presentationRelsNewXPath->query('//xmlns:Relationship[@Id="'.$nodeSldMasterIdNew->getAttribute('r:id').'"]');
                    if ($nodesRelationship->length > 0) {
                        $newId = $this->generateUniqueId();

                        // file content
                        $nodeSlidesMasterContent = $pptxNew->getContent('ppt/' . $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target')));

                        // update sldLayoutId ids
                        $nodeSlidesMasterDOM = $this->xmlUtilities->generateDomDocument($nodeSlidesMasterContent);
                        $nodesSldLayoutIdLst = $nodeSlidesMasterDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutIdLst');
                        if ($nodesSldLayoutIdLst->length > 0) {
                            $nodesSldLayoutId = $nodesSldLayoutIdLst->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldLayoutId');
                            if ($nodesSldLayoutId->length > 0) {
                                foreach ($nodesSldLayoutId as $nodeSldLayoutId) {
                                    $nodeSldLayoutId->setAttribute('id', (string)$sldMasterIdValue);
                                    $sldMasterIdValue++;
                                }
                            }
                        }

                        $pptx->addContent('ppt/slideMasters/slideMaster' . $newId . '.xml', $nodeSlidesMasterDOM->saveXML());
                        $slidesMasterContents[] = array(
                            'new' => 'slideMaster' . $newId . '.xml',
                            'old' => str_replace('slideMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                        );

                        // rels
                        $nodeSlidesMasterRelsContent = $pptxNew->getContent('ppt/' . str_replace('slideMasters/', 'slideMasters/_rels/', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))) . '.rels');
                        // update internal paths
                        foreach ($slidesLayoutContents as $slidesLayoutContent) {
                            $nodeSlidesMasterRelsContent = str_replace($slidesLayoutContent['old'], $slidesLayoutContent['new'], $nodeSlidesMasterRelsContent);
                            $nodeSlideLayoutRelsContent = $pptx->getContent('ppt/slideLayouts/_rels/' . $slidesLayoutContent['new'] . '.rels');
                            $nodeSlideLayoutRelsContent = str_replace($slidesMasterContents[count($slidesMasterContents) - 1]['old'], $slidesMasterContents[count($slidesMasterContents) - 1]['new'], $nodeSlideLayoutRelsContent);
                            $pptx->addContent('ppt/slideLayouts/_rels/' . $slidesLayoutContent['new'] . '.rels', $nodeSlideLayoutRelsContent);
                        }
                        foreach ($themesContents as $themesContent) {
                            $nodeSlidesMasterRelsContent = str_replace($themesContent['old'], $themesContent['new'], $nodeSlidesMasterRelsContent);
                        }

                        // relationship
                        $relationshipNew = '<Relationship Id="rId' . $newId . '" Target="slideMasters/slideMaster' . $newId . '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
                        $relatioshipFragment = $presentationRelsDOM->createDocumentFragment();
                        $relatioshipFragment->appendXML($relationshipNew);
                        $presentationRelsDOM->documentElement->appendChild($relatioshipFragment);

                        // presentation
                        $nodesSldMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
                        $sldMasterIdNew = '<p:sldMasterId xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" id="' . $sldMasterIdValue . '" r:id="rId' . $newId . '"/>';
                        $sldMasterIdValue++;
                        $sldMasterIdFragment = $presentationDOM->createDocumentFragment();
                        $sldMasterIdFragment->appendXML($sldMasterIdNew);
                        $nodesSldMasterIdLst->item(0)->appendChild($sldMasterIdFragment);

                        // ContentType
                        $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" PartName="/ppt/slideMasters/slideMaster' . $newId . '.xml"/>', $contentTypesDOM);

                        // rels DOM
                        $nodeSlidesMasterRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeSlidesMasterRelsContent);
                        $nodeSlidesMasterRelsContentXPath = new DOMXPath($nodeSlidesMasterRelsContentDOM);
                        $nodeSlidesMasterRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                        // images
                        $nodesImage = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesImage, 'image', 'media');

                        // media
                        $nodesMedia = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/media"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesMedia, 'media', 'media');

                        // inks
                        $nodesInk = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"]');
                        $nodeSlidesMasterRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $contentTypesDOM, $nodesInk, 'ink', 'ink', 'application/inkml+xml');

                        // 3dmodels
                        $nodesModel3d = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesModel3d, 'media', 'media');

                        // objects
                        $nodesOleObject = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesOleObject, 'embeddings', 'embeddings');
                        $nodesPackage = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesPackage, 'embeddings', 'embeddings');

                        // tags
                        $nodesTags = $nodeSlidesMasterRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"]');
                        $nodeSlidesMasterRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesMasterRelsContent, $nodesTags, 'tags', 'tags');

                        $pptx->addContent('ppt/slideMasters/_rels/slideMaster' . $newId . '.xml.rels', $nodeSlidesMasterRelsContent);

                        // free resources
                        $nodeSlidesMasterDOM = null;
                    }
                }
            }

            // notesMaster. Only one notesMaster can be added
            $notesMasterContents = array();
            $nodesNotesMasterId = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterId');
            if ($nodesNotesMasterId->length > 0) {
                // get the existing notesMaster
                $nodesRelationship = $presentationRelsXPath->query('//xmlns:Relationship[@Id="'.$nodesNotesMasterId->item(0)->getAttribute('r:id').'"]');
                if ($nodesRelationship->length > 0) {
                    $notesMasterContents[] = array(
                        'new' => str_replace('notesMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                        'old' => str_replace('notesMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                    );
                }
            } else {
                // add a new notesMaster
                $nodesNotesMasterIdNew = $presentationNewDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterId');
                if ($nodesNotesMasterIdNew->length > 0) {
                    $nodesRelationship = $presentationRelsNewXPath->query('//xmlns:Relationship[@Id="'.$nodesNotesMasterIdNew->item(0)->getAttribute('r:id').'"]');
                    if ($nodesRelationship->length > 0) {
                        $newId = $this->generateUniqueId();

                        // file content
                        $nodeNotesMasterContent = $pptxNew->getContent('ppt/' . $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target')));

                        $pptx->addContent('ppt/notesMasters/notesMaster' . $newId . '.xml', $nodeNotesMasterContent);
                        $notesMasterContents[] = array(
                            'new' => 'notesMaster' . $newId . '.xml',
                            'old' => str_replace('notesMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                        );

                        // rels
                        $nodeNotesMasterRelsContent = $pptxNew->getContent('ppt/' . str_replace('notesMasters/', 'notesMasters/_rels/', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))) . '.rels');
                        // update internal paths
                        foreach ($themesContents as $themesContent) {
                            $nodeNotesMasterRelsContent = str_replace($themesContent['old'], $themesContent['new'], $nodeNotesMasterRelsContent);
                        }
                        $pptx->addContent('ppt/notesMasters/_rels/notesMaster' . $newId . '.xml.rels', $nodeNotesMasterRelsContent);

                        // relationship
                        $relationshipNew = '<Relationship Id="rId' . $newId . '" Target="notesMasters/notesMaster' . $newId . '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"/>';
                        $relatioshipFragment = $presentationRelsDOM->createDocumentFragment();
                        $relatioshipFragment->appendXML($relationshipNew);
                        $presentationRelsDOM->documentElement->appendChild($relatioshipFragment);

                        // presentation
                        $nodesNotesMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterIdLst');
                        if ($nodesNotesMasterIdLst->length == 0) {
                            $nodesSldMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
                            $notesMasterIdLstNew = '<p:notesMasterIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:notesMasterIdLst>';
                            $notesMasterIdLstFragment = $presentationDOM->createDocumentFragment();
                            $notesMasterIdLstFragment->appendXML($notesMasterIdLstNew);
                            $nodeNotesMasterIdLst = $nodesSldMasterIdLst->item(0)->parentNode->insertBefore($notesMasterIdLstFragment, $nodesSldMasterIdLst->item(0)->nextSibling);
                        } else {
                            $nodeNotesMasterIdLst = $nodesNotesMasterIdLst->item(0);
                        }
                        $notesMasterIdNew = '<p:notesMasterId xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' . $newId . '"/>';
                        $notesMasterIdFragment = $presentationDOM->createDocumentFragment();
                        $notesMasterIdFragment->appendXML($notesMasterIdNew);
                        $nodeNotesMasterIdLst->appendChild($notesMasterIdFragment);

                        // ContentType
                        $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" PartName="/ppt/notesMasters/notesMaster' . $newId . '.xml"/>', $contentTypesDOM);
                    }
                }
            }

            // handoutMaster. Only one handoutMaster can be added
            $handoutMasterContents = array();
            $nodesHandoutMasterId = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'handoutMasterIdLst');
            if ($nodesHandoutMasterId->length > 0) {
                // get the existing handoutMaster
                $nodesRelationship = $presentationRelsXPath->query('//xmlns:Relationship[@Id="'.$nodesHandoutMasterId->item(0)->getAttribute('r:id').'"]');
                if ($nodesRelationship->length > 0) {
                    $handoutMasterContents[] = array(
                        'new' => str_replace('handoutMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                        'old' => str_replace('handoutMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                    );
                }
            } else {
                // add a new handoutMasters
                $nodesHandoutMasterIdNew = $presentationNewDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'handoutMasterId');
                if ($nodesHandoutMasterIdNew->length > 0) {

                    $nodesRelationship = $presentationRelsNewXPath->query('//xmlns:Relationship[@Id="'.$nodesHandoutMasterIdNew->item(0)->getAttribute('r:id').'"]');
                    if ($nodesRelationship->length > 0) {
                        $newId = $this->generateUniqueId();

                        // file content
                        $nodeHandoutMasterContent = $pptxNew->getContent('ppt/' . $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target')));
                        $pptx->addContent('ppt/handoutMasters/handoutMaster' . $newId . '.xml', $nodeHandoutMasterContent);

                        $handoutMasterContents[] = array(
                            'new' => 'handoutMaster' . $newId . '.xml',
                            'old' => str_replace('handoutMasters/', '', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))),
                        );

                        // rels
                        $nodeHandoutMasterRelsContent = $pptxNew->getContent('ppt/' . str_replace('handoutMasters/', 'handoutMasters/_rels/', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))) . '.rels');
                        // update internal paths
                        foreach ($themesContents as $themesContent) {
                            $nodeHandoutMasterRelsContent = str_replace($themesContent['old'], $themesContent['new'], $nodeHandoutMasterRelsContent);
                        }
                        $pptx->addContent('ppt/handoutMasters/_rels/handoutMaster' . $newId . '.xml.rels', $nodeHandoutMasterRelsContent);

                        // relationship
                        $relationshipNew = '<Relationship Id="rId' . $newId . '" Target="handoutMasters/handoutMaster' . $newId . '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"/>';
                        $relatioshipFragment = $presentationRelsDOM->createDocumentFragment();
                        $relatioshipFragment->appendXML($relationshipNew);
                        $presentationRelsDOM->documentElement->appendChild($relatioshipFragment);

                        // presentation
                        $nodesHandoutMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'handoutMasterIdLst');
                        if ($nodesHandoutMasterIdLst->length == 0) {
                            $notesMasterIdLstNew = '<p:handoutMasterIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:handoutMasterIdLst>';
                            $notesMasterIdLstFragment = $presentationDOM->createDocumentFragment();
                            $notesMasterIdLstFragment->appendXML($notesMasterIdLstNew);

                            $nodesNotesMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterIdLst');
                            if ($nodesNotesMasterIdLst->length > 0) {
                                $nodeHandoutMasterIdLst = $nodesNotesMasterIdLst->item(0)->parentNode->insertBefore($notesMasterIdLstFragment, $nodesNotesMasterIdLst->item(0)->nextSibling);
                            } else {
                                $nodesSldMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
                                $nodeHandoutMasterIdLst = $nodesSldMasterIdLst->item(0)->parentNode->insertBefore($notesMasterIdLstFragment, $nodesSldMasterIdLst->item(0)->nextSibling);
                            }
                        } else {
                            $nodeHandoutMasterIdLst = $nodesHandoutMasterIdLst->item(0);
                        }
                        $handoutMasterIdNew = '<p:notesMasterId xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' . $newId . '"/>';
                        $hadnoutMasterIdFragment = $presentationDOM->createDocumentFragment();
                        $hadnoutMasterIdFragment->appendXML($handoutMasterIdNew);
                        $nodeHandoutMasterIdLst->appendChild($hadnoutMasterIdFragment);

                        // ContentType
                        $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" PartName="/ppt/handoutMasters/handoutMaster' . $newId . '.xml"/>', $contentTypesDOM);
                    }
                }
            }

            // commentAuthors
            $commentAuthorsNew = $pptxNew->getContent('ppt/commentAuthors.xml');
            $authorsInfoNew = array();
            $authorsInfo = array();
            $mergedCommentAuthors = false;
            $maxLastIdx = 0;

            if (!$commentAuthors && $commentAuthorsNew) {
                // override
                $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml" PartName="/ppt/commentAuthors.xml"/>', $contentTypesDOM);

                // relationship
                $newIdCommentAuthors = $this->generateUniqueId();
                $relationshipNew = '<Relationship Id="rId' . $newIdCommentAuthors . '" Target="commentAuthors.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"/>';
                $relatioshipFragment = $presentationRelsDOM->createDocumentFragment();
                $relatioshipFragment->appendXML($relationshipNew);
                $presentationRelsDOM->documentElement->appendChild($relatioshipFragment);
            }

            if ($commentAuthors && $commentAuthorsNew) {
                $commentAuthorsDOM = $this->xmlUtilities->generateDomDocument($commentAuthors);
                $commentAuthorsXPath = new DOMXPath($commentAuthorsDOM);
                $commentAuthorsXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

                $commentAuthorsNewDOM = $this->xmlUtilities->generateDomDocument($commentAuthorsNew);
                $commentAuthorsNewXPath = new DOMXPath($commentAuthorsNewDOM);
                $commentAuthorsNewXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                $nodesCmAuthorNew = $commentAuthorsNewXPath->query('//p:cmAuthor');

                // get authors id
                $nodesCmAuthor = $commentAuthorsXPath->query('//p:cmAuthor');
                $authorMaxId = 1;
                $authorMaxClrIdx = 0;
                foreach ($nodesCmAuthor as $nodeCmAuthor) {
                    $authorId = $nodeCmAuthor->getAttribute('id');
                    if ((int)$authorId > $authorMaxId) {
                        $authorMaxId = $authorId;
                    }
                    $authorClrIdx = $nodeCmAuthor->getAttribute('clrIdx');
                    if ((int)$authorClrIdx > $authorMaxClrIdx) {
                        $authorMaxClrIdx = $authorClrIdx;
                    }
                }
                $authorMaxId++;
                $authorMaxClrIdx++;

                // add new authors
                foreach ($nodesCmAuthorNew as $nodeCmAuthorNew) {
                    $authorNameNew = $nodeCmAuthorNew->getAttribute('name');
                    $nodesCmAuthorName = $commentAuthorsXPath->query('//p:cmAuthor[@name="' . $authorNameNew . '"]');
                    $authorsInfoNew[$nodeCmAuthorNew->getAttribute('id')] = array(
                        'lastIdx' => $nodeCmAuthorNew->getAttribute('lastIdx'),
                        'name' => $nodeCmAuthorNew->getAttribute('name'),
                    );
                    if ($nodesCmAuthorName->length == 0) {
                        $cmAuthorFragment = $commentAuthorsDOM->createDocumentFragment();
                        $cmAuthorFragment->appendXML(str_replace('<p:cmAuthor ', '<p:cmAuthor xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" ', $nodeCmAuthorNew->ownerDocument->saveXML($nodeCmAuthorNew)));
                        $nodeCmAuthor = $commentAuthorsDOM->documentElement->appendChild($cmAuthorFragment);
                        $nodeCmAuthor->setAttribute('id', (string)$authorMaxId);
                        $nodeCmAuthor->setAttribute('clrIdx', (string)$authorMaxClrIdx);

                        $authorsInfoNew[$nodeCmAuthorNew->getAttribute('id')]['newId'] = $authorMaxId;

                        $authorMaxId++;
                        $authorMaxClrIdx++;
                    }
                }

                $commentAuthors = $commentAuthorsDOM->saveXML();

                // get current authors info and maxLastIdx
                $nodesCmAuthor = $commentAuthorsXPath->query('//p:cmAuthor');
                foreach ($nodesCmAuthor as $nodeCmAuthor) {
                    if ((int)$nodeCmAuthor->getAttribute('lastIdx') > $maxLastIdx) {
                        $maxLastIdx = (int)$nodeCmAuthor->getAttribute('lastIdx');
                    }
                    $authorsInfo[$nodeCmAuthor->getAttribute('id')] = array(
                        'name' => $nodeCmAuthor->getAttribute('name'),
                        'lastIdx' => $nodeCmAuthor->getAttribute('lastIdx'),
                    );
                }
                $maxLastIdx++;

                $mergedCommentAuthors = true;
            } else if ($commentAuthorsNew) {
                $commentAuthors = $commentAuthorsNew;
            }

            // slides
            $slidesIds = array();
            $nodesSldIdLstNew = $presentationNewDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
            if ($nodesSldIdLstNew->length > 0) {
                // the presentation includes slides
                $sldIdTagsNew = $nodesSldIdLstNew->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldId');
                if ($sldIdTagsNew->length > 0) {
                    // generate a new ID from existing values
                    foreach ($sldIdTagsNew as $sldIdTagNew) {
                        if ($sldIdTagNew->hasAttribute('id')) {
                            $nodesRelationship = $presentationRelsNewXPath->query('//xmlns:Relationship[@Id="'.$sldIdTagNew->getAttribute('r:id').'"]');
                            if ($nodesRelationship->length > 0) {
                                $newId = $this->generateUniqueId();

                                // file content
                                $nodeSlideContent = $pptxNew->getContent('ppt/' . $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target')));
                                $pptx->addContent('ppt/slides/slide' . $newId . '.xml', $nodeSlideContent);
                                $slidesIds[] = array(
                                    'new' => (string)$sldIdValue,
                                    'old' => $sldIdTagNew->getAttribute('id'),
                                );

                                // rels
                                $nodeSlidesRelsContent = $pptxNew->getContent('ppt/'. str_replace('slides/', 'slides/_rels/', $this->generateInternalPath($nodesRelationship->item(0)->getAttribute('Target'))) . '.rels');
                                // update internal paths
                                foreach ($slidesLayoutContents as $slidesLayoutContent) {
                                    $nodeSlidesRelsContent = str_replace($slidesLayoutContent['old'], $slidesLayoutContent['new'], $nodeSlidesRelsContent);
                                }

                                // relationship
                                $relationshipNew = '<Relationship Id="rId' . $newId . '" Target="slides/slide' . $newId . '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
                                $relatioshipFragment = $presentationRelsDOM->createDocumentFragment();
                                $relatioshipFragment->appendXML($relationshipNew);
                                $presentationRelsDOM->documentElement->appendChild($relatioshipFragment);

                                // presentation
                                $nodesSldIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
                                if ($nodesSldIdLst->length == 0) {
                                    $nodeSldMasterIdLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
                                    $sldIdLstFragment = $presentationDOM->createDocumentFragment();
                                    $sldIdLstFragment->appendXML('<p:sldIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sldIdLst>');
                                    $nodeSldIdLst = $nodeSldMasterIdLst->item(0)->parentNode->insertBefore($sldIdLstFragment, $nodeSldMasterIdLst->item(0)->nextSibling);
                                } else {
                                    $nodeSldIdLst = $nodesSldIdLst->item(0);
                                }
                                $sldIdNew = '<p:sldId xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" id="' . $sldIdValue . '" r:id="rId' . $newId . '"/>';
                                $sldIdValue++;
                                $sldIdFragment = $presentationDOM->createDocumentFragment();
                                $sldIdFragment->appendXML($sldIdNew);
                                $nodeSldIdLst->appendChild($sldIdFragment);

                                // ContentType
                                $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/slide' . $newId . '.xml"/>', $contentTypesDOM);

                                // rels DOM
                                $nodeSlidesRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeSlidesRelsContent);
                                $nodeSlidesRelsContentXPath = new DOMXPath($nodeSlidesRelsContentDOM);
                                $nodeSlidesRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                                // charts
                                $nodesChart = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" or @Type="http://schemas.microsoft.com/office/2014/relationships/chartEx"]');
                                foreach ($nodesChart as $nodeChart) {
                                    $nodeChartContent = $pptxNew->getContent('ppt/' . str_replace('../', '', $this->generateInternalPath($nodeChart->getAttribute('Target'))));
                                    $newIdChart = $this->generateUniqueId();
                                    $nodeSlidesRelsContent = str_replace($nodeChart->getAttribute('Target'), '../charts/chart' . $newIdChart . '.xml', $nodeSlidesRelsContent);
                                    $pptx->addContent('ppt/charts/chart' . $newIdChart . '.xml', $nodeChartContent);
                                    if ($nodeChart->getAttribute('Type') == 'http://schemas.microsoft.com/office/2014/relationships/chartEx') {
                                        // extended chart
                                        $this->addOverride('<Override ContentType="application/vnd.ms-office.chartex+xml" PartName="/ppt/charts/chart' . $newIdChart . '.xml"/>', $contentTypesDOM);
                                    } else {
                                        // other chart
                                        $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/ppt/charts/chart' . $newIdChart . '.xml"/>', $contentTypesDOM);
                                    }

                                    // rels
                                    $nodeChartRelsContent = $pptxNew->getContent('ppt/' . str_replace('../charts/', 'charts/_rels/', $this->generateInternalPath($nodeChart->getAttribute('Target'))) . '.rels');
                                    $nodeChartRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeChartRelsContent);
                                    $nodeChartRelsContentXPath = new DOMXPath($nodeChartRelsContentDOM);
                                    $nodeChartRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                                    $nodesRelationshipChart = $nodeChartRelsContentXPath->query('//xmlns:Relationship');
                                    foreach ($nodesRelationshipChart as $nodeRelationshipChart) {
                                        if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package') {
                                            $chartContent = $pptxNew->getContent('ppt/' . str_replace('../embeddings/', 'embeddings/', $this->generateInternalPath($nodeRelationshipChart->getAttribute('Target'))));
                                            $pptx->addContent('ppt/embeddings/Microsoft_Excel_Worksheet' . $newIdChart . '.xlsx', $chartContent);

                                            $nodeChartRelsContent = str_replace($nodeRelationshipChart->getAttribute('Target'), '../embeddings/Microsoft_Excel_Worksheet' . $newIdChart . '.xlsx', $nodeChartRelsContent);
                                        } else if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle') {
                                            $chartContent = $pptxNew->getContent('ppt/charts/' . $this->generateInternalPath($nodeRelationshipChart->getAttribute('Target')));
                                            $pptx->addContent('ppt/charts/colors' . $newIdChart . '.xml', $chartContent);

                                            $nodeChartRelsContent = str_replace($this->generateInternalPath($nodeRelationshipChart->getAttribute('Target')), 'colors' . $newIdChart . '.xml', $nodeChartRelsContent);
                                            $this->addOverride('<Override ContentType="application/vnd.ms-office.chartcolorstyle+xml" PartName="/ppt/charts/colors' . $newIdChart . '.xml"/>', $contentTypesDOM);
                                        } else if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.microsoft.com/office/2011/relationships/chartStyle') {
                                            $chartContent = $pptxNew->getContent('ppt/charts/' . $this->generateInternalPath($nodeRelationshipChart->getAttribute('Target')));
                                            $pptx->addContent('ppt/charts/style' . $newIdChart . '.xml', $chartContent);

                                            $nodeChartRelsContent = str_replace($this->generateInternalPath($nodeRelationshipChart->getAttribute('Target')), 'style' . $newIdChart . '.xml', $nodeChartRelsContent);
                                            $this->addOverride('<Override ContentType="application/vnd.ms-office.chartstyle+xml" PartName="/ppt/charts/style' . $newIdChart . '.xml"/>', $contentTypesDOM);
                                        }
                                    }
                                    $pptx->addContent('ppt/charts/_rels/chart' . $newIdChart . '.xml.rels', $nodeChartRelsContent);

                                    // free resources
                                    $nodeChartRelsContentDOM = null;
                                }

                                // comments
                                $nodesComment = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]');
                                if ($nodesComment->length > 0) {
                                    // merge comments
                                    foreach ($nodesComment as $nodeComment) {
                                        $nodeCommentContent = $pptxNew->getContent('ppt/' . str_replace('../', '', $this->generateInternalPath($nodeComment->getAttribute('Target'))));
                                        $nodeCommentContentDOM = $this->xmlUtilities->generateDomDocument($nodeCommentContent);
                                        if ($mergedCommentAuthors) {
                                            $nodeCommentContentXPath = new DOMXPath($nodeCommentContentDOM);
                                            $nodeCommentContentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                                            $nodesCommentsCm = $nodeCommentContentXPath->query('//p:cm');

                                            foreach ($nodesCommentsCm as $nodeCommentsCm) {
                                                $nodeCommentsCm->setAttribute('idx', $maxLastIdx);

                                                $authorId = null;
                                                // get author id
                                                if (isset($authorsInfoNew[$nodeCommentsCm->getAttribute('authorId')])) {
                                                    if (isset($authorsInfoNew[$nodeCommentsCm->getAttribute('authorId')]['newId'])) {
                                                        // it's a new author
                                                        $authorId = $authorsInfoNew[$nodeCommentsCm->getAttribute('authorId')]['newId'];
                                                    } else {
                                                        // the author already exists
                                                        foreach ($authorsInfo as $authorInfoId => $authorInfoValue) {
                                                            if ($authorsInfoNew[$nodeCommentsCm->getAttribute('authorId')]['name'] == $authorInfoValue['name']) {
                                                                $authorId = $authorInfoId;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                                if (!is_null($authorId)) {
                                                    $nodeCommentsCm->setAttribute('authorId', $authorId);

                                                    // update authors info with the last comment number
                                                    $authorsInfo[$authorId]['lastIdx'] = $maxLastIdx;
                                                }
                                                $maxLastIdx++;
                                            }
                                        }

                                        $newIdComment = $this->generateUniqueId();
                                        $nodeSlidesRelsContent = str_replace($nodeComment->getAttribute('Target'), '../comments/comment' . $newIdComment . '.xml', $nodeSlidesRelsContent);
                                        $pptx->addContent('ppt/comments/comment' . $newIdComment . '.xml', $nodeCommentContentDOM->saveXML());
                                        $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml" PartName="/ppt/comments/comment' . $newIdComment . '.xml"/>', $contentTypesDOM);

                                        // free resources
                                        $nodeCommentContentDOM = null;
                                    }

                                    // update commentAuthors attributes
                                    if ($mergedCommentAuthors && isset($commentAuthorsDOM) && isset($commentAuthorsXPath)) {
                                        $nodesCmAuthor = $commentAuthorsXPath->query('//p:cmAuthor');
                                        foreach ($nodesCmAuthor as $nodeCmAuthor) {
                                            $nodeCmAuthor->setAttribute('lastIdx', $authorsInfo[$nodeCmAuthor->getAttribute('id')]['lastIdx']);
                                        }
                                        $commentAuthors = $commentAuthorsDOM->saveXML();
                                    }
                                }

                                // images
                                $nodesImage = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesImage, 'image', 'media');

                                // media
                                $nodesMedia = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/media"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesMedia, 'media', 'media');

                                // inks
                                $nodesInk = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesInk, 'ink', 'ink', 'application/inkml+xml');

                                // 3dmodels
                                $nodesModel3d = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesModel3d, 'media', 'media');

                                // objects
                                $nodesOleObject = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesOleObject, 'embeddings', 'embeddings');
                                $nodesPackage = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesPackage, 'embeddings', 'embeddings');

                                // tags
                                $nodesTags = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"]');
                                $nodeSlidesRelsContent = $this->addExternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $nodesTags, 'tags', 'tags');

                                // notes
                                $nodesNotes = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]');
                                foreach ($nodesNotes as $nodeNotes) {
                                    $nodeNotesContent = $pptxNew->getContent('ppt/' . str_replace('../', '', $this->generateInternalPath($nodeNotes->getAttribute('Target'))));
                                    $newIdNotes = $this->generateUniqueId();
                                    $nodeSlidesRelsContent = str_replace($nodeNotes->getAttribute('Target'), '../notesSlides/notesSlide' . $newIdNotes . '.xml', $nodeSlidesRelsContent);
                                    $pptx->addContent('ppt/notesSlides/notesSlide' . $newIdNotes . '.xml', $nodeNotesContent);
                                    $this->addOverride('<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" PartName="/ppt/notesSlides/notesSlide' . $newIdNotes . '.xml"/>', $contentTypesDOM);

                                    // rels
                                    $nodeNotesRelsContent = $pptxNew->getContent('ppt/' . str_replace('../notesSlides/', 'notesSlides/_rels/', $this->generateInternalPath($nodeNotes->getAttribute('Target'))) . '.rels');
                                    // update internal paths
                                    foreach ($notesMasterContents as $notesMasterContent) {
                                        $nodeNotesRelsContent = str_replace($notesMasterContent['old'], $notesMasterContent['new'], $nodeNotesRelsContent);
                                    }
                                    $nodeNotesRelsContent = str_replace($nodesRelationship->item(0)->getAttribute('Target'), 'slides/slide' . $newId . '.xml', $nodeNotesRelsContent);
                                    $pptx->addContent('ppt/notesSlides/_rels/notesSlides' . $newIdNotes . '.xml.rels', $nodeNotesRelsContent);
                                }

                                // diagram layouts
                                $nodesDiagramLayout = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesDiagramLayout, 'diagrams', 'layout', 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml');

                                // diagram datas
                                $nodesDiagramData = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesDiagramData, 'diagrams', 'data', 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml');

                                // diagram drawings
                                $nodesDiagramDrawing = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesDiagramDrawing, 'diagrams', 'drawing', 'application/vnd.ms-office.drawingml.diagramDrawing+xml');

                                // diagram colors
                                $nodesDiagramColor = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesDiagramDrawing, 'diagrams', 'color', 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml');

                                // diagram quickStyles
                                $nodesDiagramQuickStyles = $nodeSlidesRelsContentXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"]');
                                $nodeSlidesRelsContent = $this->addInternalRelationships($pptx, $pptxNew, $nodeSlidesRelsContent, $contentTypesDOM, $nodesDiagramDrawing, 'diagrams', 'quickStyle', 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml');

                                $pptx->addContent('ppt/slides/_rels/slide' . $newId . '.xml.rels', $nodeSlidesRelsContent);

                                // free resources
                                $nodeSlidesRelsContentDOM = null;
                            }
                        }
                    }
                }

                // sections
                if ($options['mergeSections']) {
                    $nodesSectionLstNew = $presentationNewDOM->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
                    if ($nodesSectionLstNew->length > 0) {
                        // rename sldId node ids to the new ids
                        $nodesSldIdNew = $nodesSectionLstNew->item(0)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldId');
                        foreach ($nodesSldIdNew as $nodeSldIdNew) {
                            foreach ($slidesIds as $slidesId) {
                                if ($nodeSldIdNew->getAttribute('id') == $slidesId['old']) {
                                    $nodeSldIdNew->setAttribute('id', $slidesId['new']);
                                    break;
                                }
                            }
                        }
                        $nodesSectionLst = $presentationDOM->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
                        if ($nodesSectionLst->length > 0) {
                            $nodesSectionNew = $nodesSectionLstNew->item(0)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'section');
                            foreach ($nodesSectionNew as $nodeSectionNew) {
                                $guid = PhppptxUtilities::generateGUID();
                                $nodeSectionNew->setAttribute('id', $guid['guid']);
                                $sectionFragment = $presentationDOM->createDocumentFragment();
                                $sectionFragment->appendXML(str_replace('<p14:section ', '<p14:section xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ', $nodeSectionNew->ownerDocument->saveXML($nodeSectionNew)));
                                $nodesSectionLst->item(0)->appendChild($sectionFragment);
                            }
                        } else {
                            $nodesExtLst = $presentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'extLst');
                            if ($nodesExtLst->length > 0) {
                                $sectionLstFragment = $presentationDOM->createDocumentFragment();
                                $sectionLstFragment->appendXML(str_replace('<p:ext ', '<p:ext xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ', $nodesSectionLstNew->item(0)->parentNode->ownerDocument->saveXML($nodesSectionLstNew->item(0)->parentNode)));
                                $nodesExtLst->item(0)->appendChild($sectionLstFragment);
                            } else {
                                $extLstFragment = $presentationDOM->createDocumentFragment();
                                $extLstFragment->appendXML(str_replace('<p:extLst>', '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">', $nodesSectionLstNew->item(0)->parentNode->parentNode->ownerDocument->saveXML($nodesSectionLstNew->item(0)->parentNode->parentNode)));
                                $presentationDOM->documentElement->appendChild($extLstFragment);
                            }
                        }
                    }
                }
            }

            // free DOMDocument resources
            $contentTypesNewDOM = null;
            $presentationNewDOM = null;
            $presentationRelsNewDOM = null;
            $commentAuthorsDOM = null;
            $commentAuthorsNewDOM = null;
        }

        // add new contents and save file
        $pptx->addContent('[Content_Types].xml', $contentTypesDOM->saveXML());
        $pptx->addContent($presentations[0]['path'], $presentationDOM->saveXML());
        $pptx->addContent($presentationRelsPath, $presentationRelsDOM->saveXML());
        if ($commentAuthors) {
            $pptx->addContent('ppt/commentAuthors.xml', $commentAuthors);
        }

        $pptx->savePptx($target);

        // free DOMDocument resources
        $contentTypesDOM = null;
        $presentationDOM = null;
    }

    /**
     * Adds external relationship
     *
     * @param PptxStructure $pptx
     * @param PptxStructure $pptxNew
     * @param string $nodeRelsContent
     * @param DOMNodeList $nodes
     * @param string $type
     * @param string $subfolder
     * @return string
     */
    protected function addExternalRelationships($pptx, $pptxNew, $nodeRelsContent, $nodes, $type, $subfolder) {
        foreach ($nodes as $node) {
            $nodeContent = $pptxNew->getContent('ppt/' . str_replace('../', '', $this->generateInternalPath($node->getAttribute('Target'))));
            $extension = pathinfo($node->getAttribute('Target'), PATHINFO_EXTENSION);
            $newId = $this->generateUniqueId();
            $nodeRelsContent = str_replace($node->getAttribute('Target'), '../' . $subfolder . '/' . $type . $newId . '.' . $extension, $nodeRelsContent);
            $pptx->addContent('ppt/' . $subfolder . '/' . $type . $newId . '.' . $extension, $nodeContent);
        }

        return $nodeRelsContent;
    }

    /**
     * Adds internal relationship
     *
     * @param PptxStructure $pptx
     * @param PptxStructure $pptxNew
     * @param string $nodeRelsContent
     * @param DOMDocument $contentTypesDOM
     * @param DOMNodeList $nodes
     * @param string $subfolder
     * @param string $type
     * @param string $contentType
     * @return string
     */
    protected function addInternalRelationships($pptx, $pptxNew, $nodeRelsContent, $contentTypesDOM, $nodes, $subfolder, $type, $contentType) {
        foreach ($nodes as $node) {
            $nodeContent = $pptxNew->getContent('ppt/' . str_replace('../', '', $this->generateInternalPath($node->getAttribute('Target'))));
            $newId = $this->generateUniqueId();
            $nodeRelsContent = str_replace($node->getAttribute('Target'), '../' . $subfolder . '/' . $type . $newId . '.xml', $nodeRelsContent);
            $pptx->addContent('ppt/' . $subfolder . '/' . $type . $newId . '.xml', $nodeContent);
            $this->addOverride('<Override ContentType="' . $contentType . '" PartName="/ppt/' . $subfolder . '/' . $type . $newId . '.xml"/>', $contentTypesDOM);
        }

        return $nodeRelsContent;
    }

    /**
     * Adds override
     *
     * @access protected
     * @param string $override New override
     * @param DOMDocument $contentTypesDOM ContentTypes
     */
    protected function addOverride($override, $contentTypesDOM) {
        $overrideFragment = $contentTypesDOM->createDocumentFragment();
        $overrideFragment->appendXML($override);
        $contentTypesDOM->documentElement->appendChild($overrideFragment);
    }

    /**
     * Generates uniqueID
     *
     * @access protected
     * @return string
     */
    protected function generateUniqueId() {
        $uniqueId = uniqid((string)mt_rand(999, 9999));

        return $uniqueId;
    }

    /**
     * Generates fixed internal paths
     *
     * @access protected
     * @param string $path
     * @return string
     */
    protected function generateInternalPath($path) {
        $path = str_replace('/ppt/', '', $path);

        return $path;
    }
}
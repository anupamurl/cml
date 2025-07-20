const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');

/**
 * Insert an image directly into a PPTX slide
 * @param {Object} contents - JSZip contents object
 * @param {string} slideId - Slide ID (1-based)
 * @param {string} imageId - Relationship ID for the image
 * @param {string} mediaPath - Path to the image in the PPTX
 * @param {Object} position - Position and size information {x, y, width, height}
 * @returns {Promise<boolean>} - Success status
 */
async function insertImageIntoSlide(contents, slideId, imageId, mediaPath, position) {
  try {
    const slideFile = `ppt/slides/slide${slideId}.xml`;
    
    if (!contents.files[slideFile]) {
      console.error(`Slide file not found: ${slideFile}`);
      return false;
    }
    
    // Get the slide XML
    const slideXml = await contents.files[slideFile].async('string');
    const parser = new xml2js.Parser();
    const slideData = await parser.parseStringPromise(slideXml);
    
    if (!slideData['p:sld'] || !slideData['p:sld']['p:cSld'] || 
        !slideData['p:sld']['p:cSld'][0]['p:spTree']) {
      console.error('Invalid slide structure');
      return false;
    }
    
    // Use exact position values if provided, otherwise convert from inches to EMUs
    const x = position.exactX !== undefined ? position.exactX : parseInt(position.x * 914400);
    const y = position.exactY !== undefined ? position.exactY : parseInt(position.y * 914400);
    const width = position.exactWidth !== undefined ? position.exactWidth : parseInt(position.width * 914400);
    const height = position.exactHeight !== undefined ? position.exactHeight : parseInt(position.height * 914400);
    
    // Create a new picture element
    const newPic = {
      'p:nvPicPr': [{
        'p:cNvPr': [{ $: { id: Date.now().toString(), name: `Picture ${Date.now()}` } }],
        'p:cNvPicPr': [{ 'a:picLocks': [{ $: { noChangeAspect: '1' } }] }],
        'p:nvPr': [{}]
      }],
      'p:blipFill': [{
        'a:blip': [{ $: { 'r:embed': imageId } }],
        'a:stretch': [{ 'a:fillRect': [{}] }]
      }],
      'p:spPr': [{
        'a:xfrm': [{
          'a:off': [{ $: { x: x.toString(), y: y.toString() } }],
          'a:ext': [{ $: { cx: width.toString(), cy: height.toString() } }]
        }],
        'a:prstGeom': [{ $: { prst: 'rect' }, 'a:avLst': [{}] }],
      }]
    };
    
    // Add the new picture to the slide
    if (!slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic']) {
      slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'] = [];
    }
    slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'].push(newPic);
    
    // Convert back to XML and update the file
    const builder = new xml2js.Builder();
    const updatedSlideXml = builder.buildObject(slideData);
    contents.file(slideFile, updatedSlideXml);
    
    console.log(`Successfully inserted image into slide ${slideId}`);
    return true;
  } catch (error) {
    console.error(`Error inserting image into slide: ${error.message}`);
    return false;
  }
}

module.exports = { insertImageIntoSlide };
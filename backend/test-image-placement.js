const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { insertImageIntoSlide } = require('./insertImage');

// Test function to verify image placement
async function testImagePlacement(pptxPath, imagePath, slideId, position) {
  try {
    console.log(`Testing image placement in ${pptxPath}`);
    console.log(`Image: ${imagePath}`);
    console.log(`Slide: ${slideId}`);
    console.log(`Position: x=${position.x}, y=${position.y}, w=${position.width}, h=${position.height}`);
    
    // Read the PPTX file
    const pptxData = fs.readFileSync(pptxPath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(pptxData);
    
    // Read the image file
    const imageData = fs.readFileSync(imagePath);
    
    // Generate a unique ID for the image relationship
    const imageId = `rId${Date.now()}`;
    
    // Add the image to the PPTX media folder
    const extension = path.extname(imagePath).substring(1).toLowerCase();
    const mediaPath = `ppt/media/image${Date.now()}.${extension}`;
    contents.file(mediaPath, imageData);
    
    // Add the image relationship to the slide
    const slideRelPath = `ppt/slides/_rels/slide${slideId}.xml.rels`;
    let relXml;
    
    if (contents.files[slideRelPath]) {
      relXml = await contents.files[slideRelPath].async('string');
      relXml = relXml.replace('</Relationships>', 
        `<Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${path.basename(mediaPath)}"/></Relationships>`);
    } else {
      relXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${path.basename(mediaPath)}"/>
</Relationships>`;
    }
    
    contents.file(slideRelPath, relXml);
    
    // Insert the image into the slide
    await insertImageIntoSlide(contents, slideId, imageId, mediaPath, position);
    
    // Generate the updated PPTX
    const outputPath = path.join(path.dirname(pptxPath), `test-output-${Date.now()}.pptx`);
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    fs.writeFileSync(outputPath, updatedBuffer);
    console.log(`Test PPTX saved to ${outputPath}`);
    
    return outputPath;
  } catch (error) {
    console.error(`Error in test: ${error.message}`);
    return null;
  }
}

// Example usage (uncomment and modify to test)
/*
testImagePlacement(
  'path/to/test.pptx',
  'path/to/test-image.png',
  1, // Slide ID (1-based)
  {
    x: 1, // inches from left
    y: 1, // inches from top
    width: 3, // inches
    height: 2 // inches
  }
);
*/

module.exports = { testImagePlacement };
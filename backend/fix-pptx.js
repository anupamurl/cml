const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const xml2js = require('xml2js');

/**
 * Fix image dimensions in a PPTX file
 * @param {string} inputPath - Path to the input PPTX file
 * @param {string} outputPath - Path to save the fixed PPTX file
 * @returns {Promise<boolean>} - Success status
 */
async function fixPptxImageDimensions(inputPath, outputPath) {
  try {
    console.log(`Fixing image dimensions in ${inputPath}`);
    
    // Read the PPTX file
    const pptxData = fs.readFileSync(inputPath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(pptxData);
    
    // Find all slide files
    const slideFiles = Object.keys(contents.files)
      .filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
      .sort();
    
    console.log(`Found ${slideFiles.length} slides`);
    
    // Process each slide
    for (const slideFile of slideFiles) {
      const slideXml = await contents.files[slideFile].async('string');
      const parser = new xml2js.Parser();
      const slideData = await parser.parseStringPromise(slideXml);
      
      // Check if the slide has pictures
      if (slideData['p:sld'] && 
          slideData['p:sld']['p:cSld'] && 
          slideData['p:sld']['p:cSld'][0]['p:spTree'] && 
          slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic']) {
        
        const pics = slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'];
        console.log(`Slide ${slideFile} has ${pics.length} pictures`);
        
        // Process each picture
        for (const pic of pics) {
          if (pic['p:spPr'] && pic['p:spPr'][0]['a:xfrm']) {
            const xfrm = pic['p:spPr'][0]['a:xfrm'][0];
            
            // Ensure the transform has exact integer values
            if (xfrm['a:off'] && xfrm['a:off'][0].$) {
              const x = xfrm['a:off'][0].$.x;
              const y = xfrm['a:off'][0].$.y;
              
              // Make sure values are integers
              xfrm['a:off'][0].$.x = Math.round(parseInt(x)).toString();
              xfrm['a:off'][0].$.y = Math.round(parseInt(y)).toString();
            }
            
            // Ensure the dimensions have exact integer values
            if (xfrm['a:ext'] && xfrm['a:ext'][0].$) {
              const cx = xfrm['a:ext'][0].$.cx;
              const cy = xfrm['a:ext'][0].$.cy;
              
              // Make sure values are integers
              xfrm['a:ext'][0].$.cx = Math.round(parseInt(cx)).toString();
              xfrm['a:ext'][0].$.cy = Math.round(parseInt(cy)).toString();
            }
          }
        }
        
        // Convert back to XML and update the file
        const builder = new xml2js.Builder();
        const updatedSlideXml = builder.buildObject(slideData);
        contents.file(slideFile, updatedSlideXml);
        console.log(`Updated slide ${slideFile}`);
      }
    }
    
    // Generate the updated PPTX
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    fs.writeFileSync(outputPath, updatedBuffer);
    console.log(`Fixed PPTX saved to ${outputPath}`);
    
    return true;
  } catch (error) {
    console.error(`Error fixing PPTX: ${error.message}`);
    return false;
  }
}

module.exports = { fixPptxImageDimensions };
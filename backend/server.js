const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const sharp = require('sharp');

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use('/images', express.static('uploads'));
app.use(express.static('public'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = 'uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir);
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage,
  limits: {
    fieldSize: 100 * 1024 * 1024, // 100MB field size
    fileSize: 50 * 1024 * 1024,  // 50MB file size
    fields: 100,
    files: 50
  }
});

// Parse PPTX file and extract detailed content
async function parsePptx(filePath) {
  try {
    const data = fs.readFileSync(filePath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(data);
    
    const slides = [];
    const slideFiles = Object.keys(contents.files)
      .filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
      .sort();
    
    // Get relationships for images
    const relsFiles = Object.keys(contents.files)
      .filter(name => name.startsWith('ppt/slides/_rels/slide') && name.endsWith('.xml.rels'));
    
    // Get global relationships
    let globalRelationships = {};
    if (contents.files['ppt/_rels/presentation.xml.rels']) {
      const parser = new xml2js.Parser();
      const globalRelsXml = await contents.files['ppt/_rels/presentation.xml.rels'].async('string');
      const globalRelsData = await parser.parseStringPromise(globalRelsXml);
      if (globalRelsData.Relationships && globalRelsData.Relationships.Relationship) {
        globalRelsData.Relationships.Relationship.forEach(rel => {
          globalRelationships[rel.$.Id] = rel.$.Target;
        });
      }
    }
    
    // Log all files in the zip for debugging
    console.log('All files in PPTX:');
    Object.keys(contents.files).forEach(file => {
      if (file.includes('media/') || file.includes('image')) {
        console.log(`- ${file}`);
      }
    });
    
    for (let i = 0; i < slideFiles.length; i++) {
      const slideXml = await contents.files[slideFiles[i]].async('string');
      const parser = new xml2js.Parser();
      const slideData = await parser.parseStringPromise(slideXml);
      
      // Get slide relationships
      const relsFile = relsFiles.find(rel => rel.includes(`slide${i + 1}.xml.rels`));
      let relationships = {...globalRelationships}; // Include global relationships
      if (relsFile && contents.files[relsFile]) {
        const relsXml = await contents.files[relsFile].async('string');
        const relsData = await parser.parseStringPromise(relsXml);
        if (relsData.Relationships && relsData.Relationships.Relationship) {
          relsData.Relationships.Relationship.forEach(rel => {
            relationships[rel.$.Id] = rel.$.Target;
          });
        }
      }
      
      const elements = await extractSlideElements(slideData, relationships, contents, i + 1);
      
      // Extract slide number from filename (e.g., slide1.xml -> 1)
      const slideNumber = parseInt(slideFiles[i].match(/slide(\d+)\.xml/)[1]);
      
      slides.push({
        id: slideNumber, // Use actual slide number from filename
        elements: elements,
        width: 10, // Standard slide width in inches
        height: 7.5 // Standard slide height in inches
      });
    }
    
    return slides;
  } catch (error) {
    console.error('Error parsing PPTX:', error);
    throw error;
  }
}

// Extract slide elements with positioning
async function extractSlideElements(slideData, relationships, zipContents, slideNum) {
  const elements = [];
  
  // Handle different slide structures
  if (!slideData['p:sld']) {
    console.log(`Slide ${slideNum} has non-standard structure, trying alternative parsing`);
    return extractElementsFromAlternativeStructure(slideData, slideNum);
  }
  
  if (!slideData['p:sld']['p:cSld']) {
    console.log(`Slide ${slideNum} missing p:cSld, trying alternative parsing`);
    return extractElementsFromAlternativeStructure(slideData, slideNum);
  }
  
  if (!slideData['p:sld']['p:cSld'][0]['p:spTree']) {
    console.log(`Slide ${slideNum} missing p:spTree, trying alternative parsing`);
    return extractElementsFromAlternativeStructure(slideData, slideNum);
  }
  
  const spTree = slideData['p:sld']['p:cSld'][0]['p:spTree'][0];
  const allElements = [];
  
  // Collect all elements (shapes and pictures)
  if (spTree['p:sp']) allElements.push(...spTree['p:sp'].map(el => ({...el, type: 'shape'})));
  if (spTree['p:pic']) allElements.push(...spTree['p:pic'].map(el => ({...el, type: 'pic'})));
  if (spTree['p:cxnSp']) allElements.push(...spTree['p:cxnSp'].map(el => ({...el, type: 'connector'})));
  if (spTree['p:graphicFrame']) allElements.push(...spTree['p:graphicFrame'].map(el => ({...el, type: 'graphicFrame'})));
  
  // Debug information
  console.log(`Slide ${slideNum} raw elements:`, Object.keys(spTree).filter(key => key.startsWith('p:')));
  console.log(`Collected ${allElements.length} elements for processing`);
  
  for (let idx = 0; idx < allElements.length; idx++) {
    const element = allElements[idx];
    
    // Get transform info
    let transform = null;
    
    // Try different transform paths based on element type
    if (element['p:spPr'] && element['p:spPr'][0] && element['p:spPr'][0]['a:xfrm']) {
      transform = element['p:spPr'][0]['a:xfrm'][0];
    } else if (element['p:xfrm']) {
      transform = element['p:xfrm'][0];
    } else if (element.type === 'graphicFrame' && element['p:xfrm']) {
      transform = element['p:xfrm'][0];
    }
    
    if (!transform) {
      console.log(`No transform found for element ${idx} of type ${element.type}`);
      // Use default transform instead of skipping
      transform = { 'a:off': [{ $: { x: '0', y: '0' } }], 'a:ext': [{ $: { cx: '914400', cy: '914400' } }] };
    }
    
    const off = transform['a:off'] ? transform['a:off'][0].$ : { x: '0', y: '0' };
    const ext = transform['a:ext'] ? transform['a:ext'][0].$ : { cx: '914400', cy: '914400' };
    
    const x = parseInt(off.x) / 914400;
    const y = parseInt(off.y) / 914400;
    const width = parseInt(ext.cx) / 914400;
    const height = parseInt(ext.cy) / 914400;
    
    // Process text elements
    if ((element.type === 'shape' || element.type === 'graphicFrame') && 
        (element['p:txBody'] || (element['a:graphic'] && element['a:graphic'][0]['a:graphicData']))) {
      
      let textContent = '';
      
      // Try standard text extraction
      if (element['p:txBody']) {
        textContent = extractTextFromShape(element);
      }
      
      // Try to extract text from tables and other graphic elements
      if (!textContent && element['a:graphic'] && element['a:graphic'][0]['a:graphicData']) {
        const graphicData = element['a:graphic'][0]['a:graphicData'][0];
        if (graphicData['a:tbl']) {
          textContent = extractTextFromTable(graphicData['a:tbl'][0]);
        }
      }
      
      if (textContent.trim()) {
        elements.push({
          type: 'text',
          id: `text-${idx}`,
          content: textContent,
          originalContent: textContent, // Store original for comparison
          x: x,
          y: y,
          width: width,
          height: height
        });
      }
    }
    
    // Process image elements
    if (element.type === 'pic' && element['p:blipFill']) {
      try {
        const blip = element['p:blipFill'][0]['a:blip'][0];
        const embed = blip.$['r:embed'] || blip.$['r:link'];
        const imagePath = relationships[embed];
        
        if (imagePath) {
          // Try all possible image path formats
          const possiblePaths = [
            `ppt/${imagePath}`,
            `ppt/media/${imagePath}`,
            imagePath,
            `ppt/media/${imagePath.split('/').pop()}`,
            `${imagePath}`,
            `media/${imagePath.split('/').pop()}`,
            `ppt/slides/media/${imagePath.split('/').pop()}`
          ];
          
          let imageFile = null;
          for (const path of possiblePaths) {
            if (zipContents.files[path]) {
              imageFile = zipContents.files[path];
              console.log(`Found image at path: ${path}`);
              break;
            }
          }
          
          // If still not found, try to find by extension
          if (!imageFile) {
            const extension = imagePath.split('.').pop().toLowerCase();
            const mediaFiles = Object.keys(zipContents.files).filter(name => 
              name.includes('media/') && name.endsWith(`.${extension}`)
            );
            
            if (mediaFiles.length > 0) {
              imageFile = zipContents.files[mediaFiles[0]];
              console.log(`Found image by extension: ${mediaFiles[0]}`);
            }
          }
          
          if (imageFile) {
            // Store image file reference instead of base64 data
            const extension = imagePath.split('.').pop() || 'png';
            const imageName = `slide${slideNum}_image${idx}_${Date.now()}.${extension}`;
            const imagePath_temp = `uploads/${imageName}`;
            const imageBuffer = await imageFile.async('nodebuffer');
            fs.writeFileSync(imagePath_temp, imageBuffer);
            
            elements.push({
              type: 'image',
              id: `image-${idx}`,
              src: imageName, // Store filename instead of base64
              x: x,
              y: y,
              width: width,
              height: height
            });
          } else {
            console.warn(`Could not find image file for path: ${imagePath}`);
          }
        } else {
          console.warn(`No relationship found for embed ID: ${embed}`);
        }
      } catch (error) {
        console.warn(`Failed to process image ${idx}:`, error.message);
      }
    }
  }
  
  return elements;
}

// Extract text from shape with proper handling
function extractTextFromShape(shape) {
  const texts = [];
  
  if (!shape['p:txBody'] || !shape['p:txBody'][0] || !shape['p:txBody'][0]['a:p']) {
    return '';
  }
  
  shape['p:txBody'][0]['a:p'].forEach(paragraph => {
    const paragraphTexts = [];
    
    // Handle text runs
    if (paragraph['a:r']) {
      paragraph['a:r'].forEach(run => {
        if (run['a:t'] && run['a:t'][0]) {
          paragraphTexts.push(run['a:t'][0]);
        }
      });
    }
    
    // Handle direct text
    if (paragraph['a:t']) {
      paragraphTexts.push(paragraph['a:t'][0]);
    }
    
    if (paragraphTexts.length > 0) {
      texts.push(paragraphTexts.join(''));
    }
  });
  
  return texts.join('\n');
}

// Extract text from tables
function extractTextFromTable(table) {
  const texts = [];
  
  if (!table['a:tr']) {
    return '';
  }
  
  table['a:tr'].forEach(row => {
    const rowTexts = [];
    
    if (row['a:tc']) {
      row['a:tc'].forEach(cell => {
        if (cell['a:txBody'] && cell['a:txBody'][0]['a:p']) {
          const cellTexts = [];
          
          cell['a:txBody'][0]['a:p'].forEach(paragraph => {
            const paragraphTexts = [];
            
            // Handle text runs
            if (paragraph['a:r']) {
              paragraph['a:r'].forEach(run => {
                if (run['a:t'] && run['a:t'][0]) {
                  paragraphTexts.push(run['a:t'][0]);
                }
              });
            }
            
            // Handle direct text
            if (paragraph['a:t']) {
              paragraphTexts.push(paragraph['a:t'][0]);
            }
            
            if (paragraphTexts.length > 0) {
              cellTexts.push(paragraphTexts.join(''));
            }
          });
          
          if (cellTexts.length > 0) {
            rowTexts.push(cellTexts.join('\n'));
          }
        }
      });
    }
    
    if (rowTexts.length > 0) {
      texts.push(rowTexts.join(' | '));
    }
  });
  
  return texts.join('\n');
}

// Extract elements from alternative slide structures
function extractElementsFromAlternativeStructure(slideData, slideNum) {
  const elements = [];
  console.log(`Attempting alternative extraction for slide ${slideNum}`);
  
  // Try to find any text content in the slide
  const slideXml = JSON.stringify(slideData);
  
  // Extract any text that looks like it might be content
  const textMatches = slideXml.match(/"a:t":\["([^"]+)"/g) || [];
  if (textMatches.length > 0) {
    console.log(`Found ${textMatches.length} potential text elements using alternative method`);
    
    textMatches.forEach((match, idx) => {
      const text = match.replace(/"a:t":\["([^"]+)"/, '$1');
      if (text && text.trim() && text.length > 1) {
        elements.push({
          type: 'text',
          id: `alt-text-${idx}`,
          content: text,
          originalContent: text,
          x: 1 + (idx * 0.5), // Stagger positions to avoid overlap
          y: 1 + (idx * 0.5),
          width: 8,
          height: 1
        });
      }
    });
  }
  
  return elements;
}

// Process image for PPTX
async function processImageForPptx(contents, imagePath, slideIndex, element) {
  // Ensure slideIndex is a number and use actual slide ID
  const slideId = parseInt(slideIndex) + 1; // Convert to 1-based slide ID
  try {
    // Read the image file
    const imageData = fs.readFileSync(imagePath);
    
    // Generate a unique ID for the image relationship
    const imageId = `rId${Date.now()}`;
    
    // Add the image to the PPTX media folder
    const extension = imagePath.split('.').pop().toLowerCase();
    const mediaPath = `ppt/media/image${Date.now()}.${extension}`;
    contents.file(mediaPath, imageData);
    
    // Update the slide's relationship file to include this image
    const slideRelPath = `ppt/slides/_rels/slide${slideId}.xml.rels`;
    if (contents.files[slideRelPath]) {
      // Get the relationship XML
      const relXml = await contents.files[slideRelPath].async('string');
      
      // Find the original relationship ID for this image if it exists
      let originalRelId = null;
      if (element.id) {
        // Try to extract the relationship ID from the slide XML
        const slideFile = `ppt/slides/slide${slideId}.xml`;
        if (contents.files[slideFile]) {
          const slideXml = await contents.files[slideFile].async('string');
          
          // Look for the image element by ID or position
          const parser = new xml2js.Parser();
          try {
            const slideData = await parser.parseStringPromise(slideXml);
            if (slideData['p:sld'] && slideData['p:sld']['p:cSld'] && 
                slideData['p:sld']['p:cSld'][0]['p:spTree'] && 
                slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic']) {
              
              const pics = slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'];
              for (const pic of pics) {
                // Try to match by position
                if (pic['p:spPr'] && pic['p:spPr'][0]['a:xfrm']) {
                  const xfrm = pic['p:spPr'][0]['a:xfrm'][0];
                  const off = xfrm['a:off'] ? xfrm['a:off'][0].$ : null;
                  
                  if (off) {
                    const x = parseInt(off.x) / 914400;
                    const y = parseInt(off.y) / 914400;
                    
                    // If positions are close, consider it a match
                    if (Math.abs(x - element.x) < 0.1 && Math.abs(y - element.y) < 0.1) {
                      // Found the image, get its relationship ID
                      if (pic['p:blipFill'] && pic['p:blipFill'][0]['a:blip']) {
                        const blip = pic['p:blipFill'][0]['a:blip'][0];
                        originalRelId = blip.$['r:embed'];
                        console.log(`Found original relationship ID: ${originalRelId}`);
                      }
                    }
                  }
                }
              }
            }
          } catch (parseError) {
            console.error('Error parsing slide XML:', parseError);
          }
        }
      }
      
      // If we found the original relationship ID, update it instead of adding a new one
      let updatedRelXml;
      if (originalRelId) {
        // Update existing relationship
        const parser = new xml2js.Parser();
        try {
          const relsData = await parser.parseStringPromise(relXml);
          if (relsData.Relationships && relsData.Relationships.Relationship) {
            const targetRel = relsData.Relationships.Relationship.find(rel => rel.$.Id === originalRelId);
            if (targetRel) {
              // Update the target path
              targetRel.$.Target = `../media/${path.basename(mediaPath)}`;
              
              // Convert back to XML
              const builder = new xml2js.Builder();
              updatedRelXml = builder.buildObject(relsData);
              console.log(`Updated existing relationship ${originalRelId} to point to ${mediaPath}`);
            }
          }
        } catch (parseError) {
          console.error('Error updating relationship:', parseError);
          // Fall back to adding a new relationship
          updatedRelXml = addImageRelationship(relXml, imageId, mediaPath);
        }
      } else {
        // Add new relationship
        updatedRelXml = addImageRelationship(relXml, imageId, mediaPath);
      }
      
      contents.file(slideRelPath, updatedRelXml);
    } else {
      console.warn(`Relationship file not found: ${slideRelPath}`);
      // Create a new relationship file
      const newRelXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${path.basename(mediaPath)}"/>
</Relationships>`;
      contents.file(slideRelPath, newRelXml);
    }
    
    console.log(`Added image ${imagePath} to PPTX as ${mediaPath}`);
    return true;
  } catch (error) {
    console.error(`Error adding image ${imagePath} to PPTX:`, error);
    return false;
  }
}

// Add image relationship to slide relationships XML
function addImageRelationship(relXml, id, target) {
  // If the XML doesn't have a closing Relationships tag, add the relationship before the end
  if (relXml.includes('</Relationships>')) {
    return relXml.replace('</Relationships>', 
      `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/${target}"/></Relationships>`);
  } else {
    // If the XML is malformed or empty, create a new relationships XML
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/${target}"/>
</Relationships>`;
  }
}

// Get MIME type for images
function getImageMimeType(imagePath) {
  const ext = imagePath.split('.').pop().toLowerCase();
  const mimeTypes = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'svg': 'image/svg+xml'
  };
  return mimeTypes[ext] || 'image/png';
}

// Upload and parse PPTX
app.post('/api/upload', upload.single('pptx'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const slides = await parsePptx(req.file.path);
    
    // Store original file reference
    originalFiles.set(req.file.filename, req.file.path);
    
    console.log(`Parsed ${slides.length} slides:`);
    slides.forEach((slide, i) => {
      console.log(`Slide ${i + 1}: ${slide.elements.length} elements`);
      slide.elements.forEach(el => {
        console.log(`  ${el.type}: ${el.content || 'image'} at (${el.x}, ${el.y})`);
      });
      
      if (slide.elements.length === 0) {
        console.log(`WARNING: No elements found in slide ${i + 1}. This might indicate a parsing issue.`);
      }
    });
    
    // If no elements were found in any slide, log a warning
    const totalElements = slides.reduce((sum, slide) => sum + slide.elements.length, 0);
    if (totalElements === 0) {
      console.warn('WARNING: No elements were found in any slides. This might indicate a compatibility issue with this PowerPoint format.');
    }
    
    res.json({
      success: true,
      filename: req.file.filename,
      slides: slides
    });
  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ error: 'Failed to process PPTX file' });
  }
});

// Upload image endpoint
app.post('/api/upload-image', upload.single('image'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No image uploaded' });
    }
    
    // Return the filename of the uploaded image
    res.json({
      success: true,
      imageName: req.file.filename
    });
  } catch (error) {
    console.error('Image upload error:', error);
    res.status(500).json({ error: 'Failed to upload image' });
  }
});

// Clear uploads folder endpoint
app.post('/api/clear-uploads', (req, res) => {
  try {
    const uploadDir = 'uploads';
    if (fs.existsSync(uploadDir)) {
      const files = fs.readdirSync(uploadDir);
      
      // Delete each file in the uploads directory
      files.forEach(file => {
        const filePath = path.join(uploadDir, file);
        if (fs.statSync(filePath).isFile()) {
          fs.unlinkSync(filePath);
          console.log(`Deleted file: ${filePath}`);
        }
      });
      
      res.json({ success: true, message: 'Upload folder cleared successfully' });
    } else {
      res.json({ success: true, message: 'Upload folder does not exist' });
    }
  } catch (error) {
    console.error('Error clearing uploads:', error);
    res.status(500).json({ error: 'Failed to clear uploads folder' });
  }
});

// Store original PPTX data
const originalFiles = new Map();

// Edit data endpoint - receives content and processes it
app.post('/api/edit-data', async (req, res) => {
  try {
    const { content } = req.body;
    
    if (!content) {
      return res.status(400).json({ error: 'No content provided' });
    }
    
    // Parse JSON content
    const jsonData = typeof content === 'string' ? JSON.parse(content) : content;
    
    // Update content keys with dummy values
    if (jsonData.elements && Array.isArray(jsonData.elements)) {
      jsonData.elements.forEach((element, index) => {
        if (element.type === 'text' && element.content) {
          element.content = `Dummy Text ${index + 1}`;
        }
      });
    }
    
    // Return updated JSON
    res.json({ 
      success: true, 
      message: 'Content updated with dummy values',
      updatedData: jsonData
    });
    
  } catch (error) {
    console.error('Edit data error:', error);
    res.status(500).json({ error: 'Failed to process content' });
  }
});

// Generate updated PPTX by modifying original
app.post('/api/generate', upload.any(), async (req, res) => {
  try {
    const { slides, filename } = req.body;
    const slidesData = JSON.parse(slides);
    
    // Get original file path
    const originalPath = `uploads/${filename}`;
    if (!fs.existsSync(originalPath)) {
      return res.status(400).json({ error: 'Original file not found' });
    }
    
    // Load original PPTX
    const originalData = fs.readFileSync(originalPath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(originalData);
    
    // Create element mapping for precise updates
    const originalSlides = await parsePptx(originalPath);
    
    // Update slides with JSON data
    for (let i = 0; i < slidesData.length; i++) {
      const slideData = slidesData[i];
      const originalSlide = originalSlides[i];
      const slideFile = `ppt/slides/slide${i + 1}.xml`;
      
      if (contents.files[slideFile] && originalSlide) {
        let slideXml = await contents.files[slideFile].async('string');
        
        // Process all elements
        for (const element of slideData.elements) {
          const originalElement = originalSlide.elements.find(el => el.id === element.id);
          
          // Update text elements
          if (element.type === 'text') {
            if (originalElement && originalElement.content !== element.content) {
              slideXml = replaceTextInXml(slideXml, originalElement.content, element.content);
            }
          }
          // Process image elements
          else if (element.type === 'image' && element.src) {
            if (originalElement && element.src !== originalElement.src) {
              console.log(`Image source changed from ${originalElement.src} to ${element.src}`);
              
              // Get the new image path from uploads folder
              const newImagePath = `uploads/${element.src}`;
              
              if (fs.existsSync(newImagePath)) {
                // Replace the image in the PPTX
                await processImageForPptx(contents, newImagePath, i, element);
              } else {
                console.warn(`New image not found: ${newImagePath}`);
              }
            }
          }
        }
        
        // Update the slide XML
        contents.file(slideFile, slideXml);
      }
    }
    
    // Process any pending image updates
    console.log('Finalizing PPTX with all updates...');
    
    // Generate updated PPTX with exact same compression
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    const outputPath = `uploads/updated-${Date.now()}.pptx`;
    fs.writeFileSync(outputPath, updatedBuffer);
    
    console.log(`PPTX file saved to ${outputPath}`);
    
    res.download(outputPath, 'updated-presentation.pptx', (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Clean up files
      fs.unlink(outputPath, () => {});
      // Don't delete uploaded images as they might be needed for future edits
      // Only delete temporary files uploaded during this request
      if (req.files && req.files.length > 0) {
        req.files.forEach(file => {
          if (file && file.path) {
            fs.unlink(file.path, () => {});
          }
        });
      }
    });
    
  } catch (error) {
    console.error('Generate error:', error);
    res.status(500).json({ error: 'Failed to generate PPTX' });
  }
});


// Replace text in XML while preserving exact structure
function replaceTextInXml(slideXml, oldText, newText) {
  if (!oldText || !newText || oldText === newText) return slideXml;
  
  // Handle XML entities
  const xmlEscape = (str) => {
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/\"/g, '&quot;')
              .replace(/'/g, '&apos;');
  };
  
  const xmlUnescape = (str) => {
    return str.replace(/&amp;/g, '&')
              .replace(/&lt;/g, '<')
              .replace(/&gt;/g, '>')
              .replace(/&quot;/g, '\"')
              .replace(/&apos;/g, "\'");
  };
  
  // Try both escaped and unescaped versions
  const oldTextEscaped = xmlEscape(oldText);
  const newTextEscaped = xmlEscape(newText);
  
  // Replace with exact match
  let result = slideXml;
  
  // Try direct replacement first
  const directRegex = new RegExp(`(<a:t[^>]*>)${oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(</a:t>)`, 'g');
  result = result.replace(directRegex, `$1${newText}$2`);
  
  // Try escaped version
  const escapedRegex = new RegExp(`(<a:t[^>]*>)${oldTextEscaped.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(</a:t>)`, 'g');
  result = result.replace(escapedRegex, `$1${newTextEscaped}$2`);
  
  // If no replacements were made, try a more aggressive approach for multi-line text
  if (result === slideXml && oldText.includes('\n')) {
    // Split into lines and try to replace each paragraph separately
    const oldLines = oldText.split('\n');
    const newLines = newText.split('\n');
    
    // Replace each line individually if possible
    for (let i = 0; i < oldLines.length && i < newLines.length; i++) {
      if (oldLines[i] && newLines[i] && oldLines[i] !== newLines[i]) {
        const lineRegex = new RegExp(`(<a:t[^>]*>)${oldLines[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(</a:t>)`, 'g');
        result = result.replace(lineRegex, `$1${newLines[i]}$2`);
      }
    }
  }
  
  return result;
}

// Replace image source in XML - this is a helper function but we'll use processImageForPptx instead
// for actual image replacement since it handles the relationships properly
function replaceImageSrcInXml(slideXml, oldSrc, newSrc) {
  if (!oldSrc || !newSrc || oldSrc === newSrc) return slideXml;
  
  // Extract the filename from paths
  const oldFilename = oldSrc.split('/').pop();
  const newFilename = newSrc.split('/').pop();
  
  console.log(`Replacing image reference from ${oldFilename} to ${newFilename}`);
  
  // Look for image references in the XML
  // This targets the r:embed attribute in a:blip elements which reference images
  const blipRegex = /<a:blip[^>]*r:embed="([^"]*)"[^>]*>/g;
  
  // Find all image references
  let match;
  let result = slideXml;
  
  // Get all relationship IDs that might reference images
  const relationshipIds = [];
  while ((match = blipRegex.exec(slideXml)) !== null) {
    relationshipIds.push(match[1]);
    console.log(`Found image relationship ID: ${match[1]}`);
  }
  
  // Note: The actual image replacement is handled by processImageForPptx function
  // which properly updates both the slide XML and the relationships file
  // This function is kept for reference but is not the primary method for image replacement
  
  return result;
}

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Open your browser to view the application');
});
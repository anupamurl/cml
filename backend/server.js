const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const sharp = require('sharp');
const mongoose = require('mongoose');
const { insertTableIntoSlide, replaceTableInXml } = require('./tableGenerator');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;
console.log = function() {}
// Connect to MongoDB
mongoose.connect(process.env.MONGODB_CONNECT)
  .then(() => console.log('Connected to MongoDB'))
  .catch(err => console.error('MongoDB connection error:', err));

// Define Template Schema
const templateSchema = new mongoose.Schema({
  templateName: { type: String, required: true },
  slides: [{
    slideNo: Number,
    slideContent: Object
  }],
  originalFilePath: { type: String },
  createdAt: { type: Date, default: Date.now }
});

const Template = mongoose.model('Template', templateSchema);

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use('/images', express.static(path.join(__dirname, 'uploads')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));
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
      .sort((a, b) => {
        // Extract slide numbers and sort numerically
        const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
      });
    
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
    
    // Get slide layout relationships
    let slideLayoutRelationships = {};
    const slideLayoutRelsFiles = Object.keys(contents.files)
      .filter(name => name.startsWith('ppt/slideLayouts/_rels/') && name.endsWith('.xml.rels'));
    
    for (const relsFile of slideLayoutRelsFiles) {
      try {
        const parser = new xml2js.Parser();
        const relsXml = await contents.files[relsFile].async('string');
        const relsData = await parser.parseStringPromise(relsXml);
        if (relsData.Relationships && relsData.Relationships.Relationship) {
          relsData.Relationships.Relationship.forEach(rel => {
            const layoutNum = relsFile.match(/slideLayout(\d+)\.xml\.rels/)?.[1];
            if (layoutNum) {
              slideLayoutRelationships[`layout${layoutNum}_${rel.$.Id}`] = rel.$.Target;
            }
          });
        }
      } catch (error) {
        console.warn(`Error parsing slide layout relationships: ${error.message}`);
      }
    }
    
    // Log all files in the zip for debugging
    console.log('All files in PPTX:');
    const mediaFiles = [];
    const xmlFiles = [];
    const otherFiles = [];
    
    Object.keys(contents.files).forEach(file => {
      if (file.includes('media/') || file.includes('image')) {
        mediaFiles.push(file);
      } else if (file.endsWith('.xml')) {
        xmlFiles.push(file);
      } else if (!file.endsWith('/')) { // Skip directories
        otherFiles.push(file);
      }
    });
    
    console.log(`Found ${mediaFiles.length} media files, ${xmlFiles.length} XML files, and ${otherFiles.length} other files`);
    
    if (mediaFiles.length > 0) {
      console.log('Media files:', mediaFiles.slice(0, 10).join(', ') + (mediaFiles.length > 10 ? '...' : ''));
    }
    
    // Log all relationship files to help with debugging
    const relFiles = xmlFiles.filter(file => file.includes('_rels/'));
    if (relFiles.length > 0) {
      console.log('Relationship files:', relFiles.join(', '));
    }
    
    for (let i = 0; i < slideFiles.length; i++) {
      const slideXml = await contents.files[slideFiles[i]].async('string');
      const parser = new xml2js.Parser();
      const slideData = await parser.parseStringPromise(slideXml);
      
      // Get slide relationships
      const relsFile = relsFiles.find(rel => rel.includes(`slide${i + 1}.xml.rels`));
      let relationships = {...globalRelationships, ...slideLayoutRelationships}; // Include global and slide layout relationships
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
          const tableData = extractTableData(graphicData['a:tbl'][0]);
          if (tableData.length > 0) {
            elements.push({
              type: 'table',
              id: `table-${idx}`,
              tableData: tableData,
              x: x,
              y: y,
              width: width,
              height: height
            });
          }
          textContent = extractTextFromTable(graphicData['a:tbl'][0]);
        }
      }
      
      if (textContent.trim()) {
        elements.push({
          type: 'text',
          id: `text-${idx}`,
          content: textContent,
          originalContent: textContent,
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
        let imagePath = relationships[embed];
        
        // Special handling for non-image references (XML files and other non-media files)
        if (imagePath && !isLikelyImage(imagePath)) {
          console.log(`Skipping non-image reference: ${imagePath}`);
          // These are references to XML files or other non-image files
          // We'll skip them as they're not images to be displayed
          continue; // Skip normal image processing
        }
        
        if (imagePath) {
          // Try all possible image paths
          let imageData = null;
          let imageFileName = path.basename(imagePath);
          
          // Check different possible paths for the image
          const possiblePaths = [
            `ppt/media/${imageFileName}`,
            `ppt/${imagePath}`,
            imagePath,
            `ppt/media/${path.basename(imagePath)}`,
            `media/${imageFileName}`,
            `ppt/slides/media/${imageFileName}`,
            `ppt/slideLayouts/${imageFileName}`,
            `ppt/slideLayouts/${path.basename(imagePath)}`,
            `ppt/slideLayouts/slideLayout${slideNum}.xml`
          ];
          
          for (const possiblePath of possiblePaths) {
            if (zipContents.files[possiblePath]) {
              console.log(`Found image at: ${possiblePath}`);
              imageData = await zipContents.files[possiblePath].async('nodebuffer');
              break;
            }
          }
          
          // If still not found, try to find by extension
          if (!imageData) {
            const extension = imagePath.split('.').pop().toLowerCase();
            const mediaFiles = Object.keys(zipContents.files).filter(name => 
              name.includes('media/') && name.endsWith(`.${extension}`)
            );
            
            if (mediaFiles.length > 0) {
              console.log(`Found image by extension: ${mediaFiles[0]}`);
              imageData = await zipContents.files[mediaFiles[0]].async('nodebuffer');
            }
          }
          
          if (imageData) {
            // Save image to uploads directory
            const extension = imagePath.split('.').pop() || 'png';
            const uniqueFileName = `slide${slideNum}_image${idx}_${Date.now()}.${extension}`;
            const imageSavePath = path.join('uploads', uniqueFileName);
            
            fs.writeFileSync(imageSavePath, imageData);
            
            elements.push({
              type: 'image',
              id: `image-${idx}`,
              src: uniqueFileName, // Just the filename, not the full path
              fullPath: `/uploads/${uniqueFileName}`, // Full path for frontend access
              x: x,
              y: y,
              width: width,
              height: height
            });
          } else {
            // Check if this is likely an image or another type of reference
            if (imagePath && isLikelyImage(imagePath)) {
              console.error(`Could not find image file for path: ${imagePath}`);
            } else {
              console.log(`Skipping non-image reference: ${imagePath}`);
            }
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

// Extract table data as array of arrays
function extractTableData(table) {
  const tableData = [];
  
  if (!table['a:tr']) {
    return [];
  }
  
  table['a:tr'].forEach(row => {
    const rowData = [];
    
    if (row['a:tc']) {
      row['a:tc'].forEach(cell => {
        let cellText = '';
        if (cell['a:txBody'] && cell['a:txBody'][0]['a:p']) {
          const cellTexts = [];
          cell['a:txBody'][0]['a:p'].forEach(paragraph => {
            if (paragraph['a:r']) {
              paragraph['a:r'].forEach(run => {
                if (run['a:t'] && run['a:t'][0]) {
                  cellTexts.push(run['a:t'][0]);
                }
              });
            }
            if (paragraph['a:t']) {
              cellTexts.push(paragraph['a:t'][0]);
            }
          });
          cellText = cellTexts.join('');
        }
        rowData.push(cellText || '');
      });
    }
    
    if (rowData.length > 0) {
      tableData.push(rowData);
    }
  });
  
  return tableData;
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
  console.log(`Processing image for slide ${slideId} from path ${imagePath}`);
  
  // Add a small delay to ensure proper processing
  await new Promise(resolve => setTimeout(resolve, 50));
  
  try {
    // Read the image file
    const imageData = fs.readFileSync(imagePath);
    
    // Generate a unique ID for the image relationship
    const imageId = `rId${Date.now()}`;
    
    // Add the image to the PPTX media folder
    const extension = imagePath.split('.').pop().toLowerCase();
    const mediaPath = `ppt/media/image${Date.now()}.${extension}`;
    contents.file(mediaPath, imageData);
    
    // Preserve original dimensions exactly as they were in the uploaded PPTX
    // This is critical for maintaining the exact layout
    let x, y, imgWidth, imgHeight;
    
    if (element.originalElement) {
      // Use original element's exact dimensions and position
      x = parseInt(element.originalElement.x * 914400);
      y = parseInt(element.originalElement.y * 914400);
      imgWidth = parseInt(element.originalElement.width * 914400);
      imgHeight = parseInt(element.originalElement.height * 914400);
      console.log(`Using original dimensions: x=${x}, y=${y}, w=${imgWidth}, h=${imgHeight}`);
    } else {
      // Use the provided dimensions, converting to EMUs
      x = parseInt(element.x * 914400);
      y = parseInt(element.y * 914400);
      imgWidth = parseInt(element.width * 914400);
      imgHeight = parseInt(element.height * 914400);
      
      // Only adjust aspect ratio if we don't have original dimensions
      try {
        const metadata = await sharp(imageData).metadata();
        if (metadata.width && metadata.height) {
          // Keep aspect ratio if only one dimension is specified
          if (element.width && !element.height) {
            imgHeight = parseInt((element.width * metadata.height / metadata.width) * 914400);
          } else if (!element.width && element.height) {
            imgWidth = parseInt((element.height * metadata.width / metadata.height) * 914400);
          }
        }
      } catch (err) {
        console.log(`Could not get image dimensions: ${err.message}`);
      }
    }
    
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
              
              // Now update the slide XML to maintain the exact position
              const slideFile = `ppt/slides/slide${slideId}.xml`;
              if (contents.files[slideFile]) {
                let slideXml = await contents.files[slideFile].async('string');
                
                // Find and update the image position in the slide XML
                try {
                  const slideParser = new xml2js.Parser();
                  const slideData = await slideParser.parseStringPromise(slideXml);
                  
                  if (slideData['p:sld'] && slideData['p:sld']['p:cSld'] && 
                      slideData['p:sld']['p:cSld'][0]['p:spTree'] && 
                      slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic']) {
                    
                    const pics = slideData['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'];
                    for (const pic of pics) {
                      if (pic['p:blipFill'] && pic['p:blipFill'][0]['a:blip'] && 
                          pic['p:blipFill'][0]['a:blip'][0].$['r:embed'] === originalRelId) {
                        
                        // Found the image, update its position if needed
                        if (pic['p:spPr'] && pic['p:spPr'][0]['a:xfrm'] && 
                            element.x !== undefined && element.y !== undefined) {
                          
                          const xfrm = pic['p:spPr'][0]['a:xfrm'][0];
                          
                          // Update position (EMUs = English Metric Units, 1 inch = 914400 EMUs)
                          if (xfrm['a:off'] && xfrm['a:off'][0].$) {
                            // Preserve exact position values
                            xfrm['a:off'][0].$.x = x.toString();
                            xfrm['a:off'][0].$.y = y.toString();
                            console.log(`Updated position to x=${x}, y=${y}`);
                          }
                          
                          // Update size if needed
                          if (xfrm['a:ext'] && xfrm['a:ext'][0].$) {
                            // Preserve exact dimension values
                            xfrm['a:ext'][0].$.cx = imgWidth.toString();
                            xfrm['a:ext'][0].$.cy = imgHeight.toString();
                            console.log(`Updated dimensions to w=${imgWidth}, h=${imgHeight}`);
                          }
                        }
                      }
                    }
                  }
                  
                  // Convert back to XML and update the file
                  const slideBuilder = new xml2js.Builder();
                  const updatedSlideXml = slideBuilder.buildObject(slideData);
                  contents.file(slideFile, updatedSlideXml);
                  
                } catch (slideError) {
                  console.error('Error updating slide XML:', slideError);
                }
              }
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
        
        // Use the dedicated function to insert the image into the slide
        // Pass exact EMU values to ensure precise positioning
        await insertImageIntoSlide(contents, slideId, imageId, mediaPath, {
          exactX: x,
          exactY: y,
          exactWidth: imgWidth,
          exactHeight: imgHeight
        });
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
      `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${path.basename(target)}"/></Relationships>`);
  } else {
    // If the XML is malformed or empty, create a new relationships XML
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${path.basename(target)}"/>
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
    'svg': 'image/svg+xml',
    'emf': 'image/emf',
    'wmf': 'image/wmf'
  };
  return mimeTypes[ext] || 'image/png';
}

// Check if a path is likely an image
function isLikelyImage(path) {
  if (!path) return false;
  
  // Check for common image extensions
  if (path.match(/\.(png|jpg|jpeg|gif|bmp|tiff|svg|emf|wmf)$/i)) {
    return true;
  }
  
  // Check for paths that are likely not images
  if (path.endsWith('.xml') || 
      path.includes('slideLayout') || 
      path.includes('notesSlide') || 
      path.includes('theme')) {
    return false;
  }
  
  // Check if path contains 'media' or 'image' which suggests it might be an image
  if (path.includes('media/') || path.includes('image')) {
    return true;
  }
  
  return false;
}

// Find image file path from various possible locations
async function findImageFile(imagePath) {
  // Try multiple possible paths
  const possiblePaths = [
    imagePath,
    `uploads/${imagePath}`,
    path.join(__dirname, 'uploads', imagePath),
    path.join(process.cwd(), 'uploads', imagePath),
    imagePath.startsWith('/') ? `.${imagePath}` : imagePath,
    imagePath.startsWith('/uploads/') ? `.${imagePath}` : `/uploads/${imagePath}`
  ];
  
  for (const path of possiblePaths) {
    try {
      if (fs.existsSync(path)) {
        return path;
      }
    } catch (err) {
      // Ignore errors and try next path
    }
  }
  
  return null;
}

// Upload and parse PPTX
app.post('/api/upload', upload.single('pptx'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    // Create a copy of the original file with a unique name to ensure each template has its own file
    const uniqueFilename = `original_${Date.now()}_${req.file.originalname}`;
    const uniqueFilePath = path.join('uploads', uniqueFilename);
    
    // Copy the uploaded file to the unique path
    fs.copyFileSync(req.file.path, uniqueFilePath);
    console.log(`Created unique copy of original file: ${uniqueFilePath}`);
    
    const slides = await parsePptx(req.file.path);
    
    // Store original file reference using the unique path
    originalFiles.set(req.file.filename, uniqueFilePath);
    console.log(`Stored original file reference: ${req.file.filename} -> ${uniqueFilePath}`);
    
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
      originalPath: uniqueFilePath,
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

// List presentations endpoint
app.get('/api/list-presentations', (req, res) => {
  try {
    const uploadDir = 'uploads';
    const presentations = [];
    
    if (fs.existsSync(uploadDir)) {
      const files = fs.readdirSync(uploadDir);
      
      // Filter for PPTX files and get their details
      files.forEach(file => {
        const filePath = path.join(uploadDir, file);
        if (fs.statSync(filePath).isFile() && (file.endsWith('.pptx') || file.endsWith('.ppt'))) {
          const stats = fs.statSync(filePath);
          presentations.push({
            name: file,
            type: file.split('.').pop().toUpperCase(),
            size: formatFileSize(stats.size),
            date: new Date(stats.mtime).toLocaleString(),
            path: filePath
          });
        }
      });
    }
    
    res.json({ success: true, presentations });
  } catch (error) {
    console.error('Error listing presentations:', error);
    res.status(500).json({ error: 'Failed to list presentations' });
  }
});

// Helper function to format file size
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  else if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  else return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

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

// Save template to MongoDB
app.post('/api/save-template', async (req, res) => {
  try {
    const { templateName, slides, filename, originalPath } = req.body;
    
    if (!templateName) {
      return res.status(400).json({ error: 'Template name is required' });
    }
    
    const slidesData = typeof slides === 'string' ? JSON.parse(slides) : slides;
    
    // Ensure slides are sorted by ID
    const sortedSlides = [...slidesData].sort((a, b) => a.id - b.id);
    
    // Format slides for MongoDB storage
    const formattedSlides = sortedSlides.map(slide => ({
      slideNo: slide.id,
      slideContent: slide
    }));
    
    // Get original file path if available
    let originalFilePath = originalPath;
    if (!originalFilePath && filename && originalFiles.has(filename)) {
      originalFilePath = originalFiles.get(filename);
    }
    
    console.log('Saving template with original file path:', originalFilePath);
    
    // Check if template with this name already exists
    let template = await Template.findOne({ templateName });
    
    if (template) {
      // Update existing template
      template.slides = formattedSlides;
      if (originalFilePath) {
        template.originalFilePath = originalFilePath;
      }
      template.updatedAt = Date.now();
      await template.save();
      res.json({ 
        success: true, 
        message: 'Template updated successfully', 
        templateId: template._id,
        originalFilePath: template.originalFilePath
      });
    } else {
      // Create new template
      template = new Template({
        templateName,
        slides: formattedSlides,
        originalFilePath: originalFilePath
      });
      await template.save();
      res.json({ 
        success: true, 
        message: 'Template saved successfully', 
        templateId: template._id,
        originalFilePath: template.originalFilePath
      });
    }
  } catch (error) {
    console.error('Save template error:', error);
    res.status(500).json({ error: 'Failed to save template' });
  }
});

// Get all templates
app.get('/api/templates', async (req, res) => {
  try {
    const templates = await Template.find().sort('-createdAt');
    
    // Add slide count and hasOriginalFile flag to each template
    const templatesWithCount = templates.map(template => {
      const slideCount = template.slides ? template.slides.length : 0;
      const hasOriginalFile = !!template.originalFilePath && fs.existsSync(template.originalFilePath);
      
      return {
        _id: template._id,
        templateName: template.templateName,
        createdAt: template.createdAt,
        slideCount: slideCount,
        hasOriginalFile: hasOriginalFile
      };
    });
    
    res.json({ success: true, templates: templatesWithCount });
  } catch (error) {
    console.error('Get templates error:', error);
    res.status(500).json({ error: 'Failed to get templates' });
  }
});

// Get template by ID
app.get('/api/templates/:id', async (req, res) => {
  try {
    const templateId = req.params.id;
    const template = await Template.findById(templateId);
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Format slides for frontend and ensure they're sorted by slideNo
    const slides = template.slides
      .sort((a, b) => a.slideNo - b.slideNo)
      .map(slide => slide.slideContent);
    
    // Check if the original file exists
    const hasOriginalFile = !!template.originalFilePath && fs.existsSync(template.originalFilePath);
    
    res.json({ 
      success: true, 
      templateName: template.templateName,
      slides: slides,
      hasOriginalFile: hasOriginalFile
    });
  } catch (error) {
    console.error('Get template by ID error:', error);
    res.status(500).json({ error: 'Failed to get template' });
  }
});

// Download original file
app.get('/api/download-original/:id', async (req, res) => {
  try {
    const templateId = req.params.id;
    const template = await Template.findById(templateId);
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    console.log(`Template ${template.templateName} original file path:`, template.originalFilePath);
    
    if (!template.originalFilePath) {
      return res.status(404).json({ error: 'Original file path not found in database' });
    }
    
    if (!fs.existsSync(template.originalFilePath)) {
      return res.status(404).json({ error: `Original file not found at path: ${template.originalFilePath}` });
    }
    
    // Get the original filename
    const originalFilename = path.basename(template.originalFilePath);
    // Add template name and timestamp to ensure uniqueness
    const downloadFilename = `${template.templateName}_original_${Date.now()}.${originalFilename.split('.').pop()}`;
    
    console.log(`Sending file: ${originalFilename} as ${downloadFilename}`);
    
    // Send the file
    res.download(template.originalFilePath, downloadFilename, (err) => {
      if (err) {
        console.error('Download error:', err);
        res.status(500).send(`Download error: ${err.message}`);
      }
    });
  } catch (error) {
    console.error('Download original file error:', error);
    res.status(500).json({ error: `Failed to download original file: ${error.message}` });
  }
});

// Generate PPTX from template
app.get('/api/generate-template/:id', async (req, res) => {
  try {
    const templateId = req.params.id;
    const template = await Template.findById(templateId);
    
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    // Instead of creating a new PPTX from scratch, we'll use the same approach as /api/generate
    // First, check if this template has an original file path
    const uploadDir = 'uploads';
    let baseFilePath = null;
    
    // If the template has an original file path, use it as the base
    if (template.originalFilePath && fs.existsSync(template.originalFilePath)) {
      baseFilePath = template.originalFilePath;
      console.log(`Using template's original file: ${baseFilePath}`);
    } 
    // Otherwise find a suitable base PPTX file
    else if (fs.existsSync(uploadDir)) {
      // Create a copy of a base file with a unique name for this request
      const files = fs.readdirSync(uploadDir);
      for (const file of files) {
        if (file.endsWith('.pptx')) {
          const originalFile = path.join(uploadDir, file);
          const uniqueBasePath = path.join(uploadDir, `base_${template._id}_${Date.now()}.pptx`);
          fs.copyFileSync(originalFile, uniqueBasePath);
          baseFilePath = uniqueBasePath;
          console.log(`Created unique base file: ${baseFilePath}`);
          break;
        }
      }
    }
    
    if (!baseFilePath) {
      // If no base file found, fall back to creating a new one
      const pptx = new PptxGenJS();
      
      // Process each slide in the template
      for (const slideData of template.slides) {
        const slide = pptx.addSlide();
        
        // Process slide content (text, images, etc.)
        const content = slideData.slideContent;
        if (content.elements) {
          for (const element of content.elements) {
            if (element.type === 'text') {
              slide.addText(element.content, {
                x: element.x || 0,
                y: element.y || 0,
                w: element.width || 5,
                h: element.height || 1,
                fontSize: 14
              });
            } else if (element.type === 'shape') {
              // For shape elements, generate a segmented circle SVG
              const svgPath = generateShape('shape' );
              
              slide.addImage({
                path: svgPath,
                x: element.x || 0,
                y: element.y || 0,
                w: element.width || 2,
                h: element.height || 2
              });
              
              // Clean up the temporary SVG file after a delay
              setTimeout(() => {
                try { fs.unlinkSync(svgPath); } catch (e) {}
              }, 1000);
            } else if (element.type === 'table') {
              let tableData = element.tableData;
              if (!tableData || !Array.isArray(tableData) || tableData.length === 0) {
                tableData = Array(5).fill().map((_, i) => 
                  Array(5).fill().map((_, j) => `Cell ${i+1}-${j+1}`)
                );
              }
              const totalWidth = element.width || 6;
              const colCount = tableData[0].length;
              const colWidths = new Array(colCount).fill(totalWidth / colCount);
              slide.addTable(tableData, {
                x: element.x || 1,
                y: element.y || 1,
                colW: colWidths,
                fontSize: 11,
                color: '000000',
                fill: 'FFFFFF',
                border: { pt: 1, color: '000000' },
                align: 'center',
                valign: 'middle'
              });
            }
          }
        }
      }
      
      // Generate the PPTX file
      const buffer = await pptx.write('nodebuffer');
      
      // Set response headers for file download
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="${template.templateName}.pptx"`);
      
      // Send the file
      res.send(buffer);
      return;
    }
    
    // Use the same approach as /api/generate to modify an existing PPTX
    const originalData = await fs.readFileSync(baseFilePath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(originalData);
    
    // Get the slides from the template
    const slidesData = template.slides.map(slide => slide.slideContent);
    
    // Create a new PPTX with the same structure as the original
    const originalSlides = await parsePptx(baseFilePath);
    
    // Get a list of available slide files in the PPTX
    const availableSlideFiles = Object.keys(contents.files)
      .filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
      });
    
    console.log(`Available slides in PPTX: ${availableSlideFiles.map(f => f.match(/slide(\d+)\.xml/)[1]).join(', ')}`);
    
    // Sort slides numerically by ID to ensure correct order for slides > 9
    slidesData.sort((a, b) => parseInt(a.id) - parseInt(b.id));
    
    // Update slides with template data, but only for slides that exist in the PPTX
    for (let i = 0; i < Math.min(slidesData.length, availableSlideFiles.length); i++) {
      const slideData = slidesData[i];
      // Use the slide number from the available files instead of the template ID
      const slideFile = availableSlideFiles[i];
      const slideId = parseInt(slideFile.match(/slide(\d+)\.xml/)[1]);
      
      console.log(`Processing template slide ${i+1} (ID: ${slideData.id}) with PPTX slide ${slideId}`);
      
      if (!contents.files[slideFile]) {
        console.warn(`Slide file not found: ${slideFile}`);
        continue;
      }
      
      // Get the original slide content
      let slideXml = await contents.files[slideFile].async('string');
      
      // Process all elements
      if (slideData.elements) {
        for (const element of slideData.elements) {
          // Find matching element in original slide
          const originalSlide = originalSlides.find(s => s.id === slideId);
          if (!originalSlide) continue;
          
          // Find the matching element with more precise criteria
          let originalElement = originalSlide.elements.find(el => el.id === element.id);
          
          // If no match by ID, try to match by position and type
          if (!originalElement) {
            originalElement = originalSlide.elements.find(el => 
              el.type === element.type && 
              Math.abs(el.x - element.x) < 0.1 && 
              Math.abs(el.y - element.y) < 0.1
            );
          }
          
          // If still no match, try with more relaxed position criteria
          if (!originalElement) {
            originalElement = originalSlide.elements.find(el => 
              el.type === element.type && 
              Math.abs(el.x - element.x) < 1 && 
              Math.abs(el.y - element.y) < 1
            );
          }
          
          // If element type is shape, add SVG image to the slide
          if (element.type === 'shape') {
            try {
              const { generateShape } = require('./shapeGenerator');
              const shapeData = {
                chartData: element.chartData || null,
                size: Math.max(element.width || 300, 200)
              };
              const svgPath = generateShape('donutchart', shapeData);
              await new Promise(resolve => setTimeout(resolve, 100));
              const pngPath = svgPath.replace('.svg', '.png');
              await sharp(svgPath).png().toFile(pngPath);
              await processImageForPptx(contents, pngPath, slideId - 1, {
                ...element,
                src: pngPath,
                width: Math.max(element.width || 6, 6),
                height: Math.max(element.height || 4, 4),
                originalElement: originalElement || null
              });
              setTimeout(() => {
                try { 
                  fs.unlinkSync(svgPath);
                  fs.unlinkSync(pngPath);
                } catch (e) {}
              }, 5000);
            } catch (error) {
              console.error('Error processing shape:', error);
            }
          }
          // Handle table elements - ALWAYS INSERT NEW TABLE
          else if (element.type === 'table') {
            const tableData = element.tableData || Array(5).fill().map((_, i) => 
              Array(5).fill().map((_, j) => `Cell ${i+1}-${j+1}`)
            );
            slideXml = insertTableIntoSlide(slideXml, tableData, element);
          }
          
          // Update text elements
          if (element.type === 'text' && originalElement) {
            slideXml = replaceTextInXml(slideXml, originalElement.content, element.content);
          }
          // Process image elements if needed
          else if (element.type === 'image' && element.src) {
            // Use the helper function to find the image file
            const imagePath = await findImageFile(element.src);
            
            if (imagePath) {
              // Add a small delay to ensure proper processing for large presentations
              await new Promise(resolve => setTimeout(resolve, 50));
              
              // Pass the original element for position reference
              await processImageForPptx(contents, imagePath, slideId - 1, {
                ...element,
                originalElement: originalElement || null,
                // Use original dimensions if available
                originalX: element.originalX,
                originalY: element.originalY,
                originalWidth: element.originalWidth,
                originalHeight: element.originalHeight
              });
              
              // Add another small delay after processing
              await new Promise(resolve => setTimeout(resolve, 50));
            } else {
              console.warn(`Image not found for element: ${element.src}`);
            }
          }
        }
      }
      
      // Update the slide XML
      contents.file(slideFile, slideXml);
    }
    
    // Generate updated PPTX
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    // Generate truly unique identifier for this specific request
    const uuid = require('crypto').randomUUID ? require('crypto').randomUUID() : 
                 `${Date.now()}-${Math.random().toString(36).substring(2, 15)}`;
    
    // Create unique filenames with template ID and UUID to guarantee uniqueness
    const tempPath = `uploads/temp-${template._id}-${uuid}.pptx`;
    const outputPath = `uploads/${template._id}-${uuid}.pptx`;
    fs.writeFileSync(tempPath, updatedBuffer);
    
    // Fix image dimensions in the generated PPTX
    const { fixPptxImageDimensions } = require('./fix-pptx');
    await fixPptxImageDimensions(tempPath, outputPath);
    
    // Add a delay to ensure file is fully written and processed
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // Clean up the temporary file
    await new Promise((resolve) => fs.unlink(tempPath, resolve));
    
    // Clean up the unique base file if we created one
    if (baseFilePath && baseFilePath.includes('base_') && fs.existsSync(baseFilePath)) {
      try {
        fs.unlinkSync(baseFilePath);
        console.log(`Cleaned up base file: ${baseFilePath}`);
      } catch (e) {
        console.warn(`Failed to clean up base file: ${e.message}`);
      }
    }
    
    // Set cache-control headers to prevent browser caching
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Content-Disposition', `attachment; filename="${template.templateName}_${template._id}_${uuid}.pptx"`);
    res.setHeader('X-Template-ID', template._id.toString());
    
    // Send the file with the unique identifier in both the file path and the download filename
    const downloadFilename = `${template.templateName}_${template._id}_${uuid}.pptx`;
    console.log(`Sending file: ${outputPath} as ${downloadFilename}`);
    res.download(outputPath, downloadFilename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Add a delay before cleanup to ensure download completes
      setTimeout(() => {
        fs.unlink(outputPath, () => {});
      }, 1000);
    });
  } catch (error) {
    console.error('Generate template error:', error);
    res.status(500).json({ error: 'Failed to generate template' });
  }
});

// Delete template
app.delete('/api/templates/:id', async (req, res) => {
  try {
    const templateId = req.params.id;
    const result = await Template.findByIdAndDelete(templateId);
    
    if (!result) {
      return res.status(404).json({ error: 'Template not found' });
    }
    
    res.json({ success: true, message: 'Template deleted successfully' });
  } catch (error) {
    console.error('Delete template error:', error);
    res.status(500).json({ error: 'Failed to delete template' });
  }
});

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
    const { slides, filename, templateName } = req.body;
    const slidesData = JSON.parse(slides);
    const outputFilename = templateName ? `${templateName}-${Date.now()}.pptx` : `updated-${Date.now()}.pptx`;
    
    // Get original file path
    const originalPath = `uploads/${filename}`;
    if (!fs.existsSync(originalPath)) {
      return res.status(400).json({ error: 'Original file not found' });
    }
    
    // Load original PPTX
    const originalData = await fs.readFileSync(originalPath);
    const zip = new JSZip();
    const contents = await zip.loadAsync(originalData);
    
    // Create element mapping for precise updates
    const originalSlides = await parsePptx(originalPath);
    
    // Update slides with JSON data
    for (let i = 0; i < slidesData.length; i++) {
      const slideData = slidesData[i];
      // Use the actual slide ID from the data instead of the array index
      const slideId = slideData.id || (i + 1);
      const originalSlide = originalSlides.find(s => s.id === slideId) || originalSlides[i];
      const slideFile = `ppt/slides/slide${slideId}.xml`;
      
      if (!contents.files[slideFile]) {
        console.warn(`Slide file not found: ${slideFile}. Available slides: ${Object.keys(contents.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml')).join(', ')}`);
        continue;
      }
      
      if (originalSlide) {
        let slideXml = await contents.files[slideFile].async('string');
        console.log(`Processing slide ${slideId} with ${slideData.elements?.length || 0} elements`);
        
        // Process all elements
        for (const element of slideData.elements) {
          // Find the matching element with more precise criteria
          let originalElement = originalSlide.elements.find(el => el.id === element.id);
          
          // If no match by ID, try to match by position and type
          if (!originalElement) {
            originalElement = originalSlide.elements.find(el => 
              el.type === element.type && 
              Math.abs(el.x - element.x) < 0.1 && 
              Math.abs(el.y - element.y) < 0.1
            );
          }
          
          // If still no match, try with more relaxed position criteria
          if (!originalElement) {
            originalElement = originalSlide.elements.find(el => 
              el.type === element.type && 
              Math.abs(el.x - element.x) < 1 && 
              Math.abs(el.y - element.y) < 1
            );
          }
          
          // Update text elements
          if (element.type === 'text') {
            if (originalElement && originalElement.content !== element.content) {
              slideXml = replaceTextInXml(slideXml, originalElement.content, element.content);
            }
          }
          // Handle table elements - ALWAYS INSERT NEW TABLE
          else if (element.type === 'table') {
            const tableData = element.tableData || Array(5).fill().map((_, i) => 
              Array(5).fill().map((_, j) => `Cell ${i+1}-${j+1}`)
            );
            slideXml = insertTableIntoSlide(slideXml, tableData, element);
          }
          // Process image elements
          else if (element.type === 'image' && element.src) {
            // Use the helper function to find the image file
            const imagePath = await findImageFile(element.src);
            
            if (imagePath) {
              console.log(`Processing image from path: ${imagePath}`);
              
              // Add a small delay to ensure proper processing for large presentations
              await new Promise(resolve => setTimeout(resolve, 50));
              
              // Replace the image in the PPTX using the correct slide ID
              await processImageForPptx(contents, imagePath, slideId - 1, {
                ...element,
                originalElement: originalElement || null,
                // Use original dimensions if available
                originalX: element.originalX,
                originalY: element.originalY,
                originalWidth: element.originalWidth,
                originalHeight: element.originalHeight
              });
              
              // Add another small delay after processing
              await new Promise(resolve => setTimeout(resolve, 50));
            } else {
              console.warn(`Image not found for element: ${element.src}`);
            }
          }
        }
        
        // Update the slide XML
        contents.file(slideFile, slideXml);
      }
    }
    
    // Process any pending image updates
    console.log(`Finalizing PPTX with all updates for ${slidesData.length} slides...`);
    console.log(`Slide IDs processed: ${slidesData.map(slide => slide.id).join(', ')}`);
    console.log(`Original slide IDs: ${originalSlides.map(slide => slide.id).join(', ')}`);
    console.log(`Available slide files: ${Object.keys(contents.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml')).join(', ')}`);
    
    
    // Generate updated PPTX with exact same compression
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    const tempPath = `uploads/temp-${outputFilename}`;
    const outputPath = `uploads/${outputFilename}`;
    fs.writeFileSync(tempPath, updatedBuffer);
    
    // Fix image dimensions in the generated PPTX
    const { fixPptxImageDimensions } = require('./fix-pptx');
    await fixPptxImageDimensions(tempPath, outputPath);
    
    // Add a delay to ensure file is fully written and processed
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // Clean up the temporary file
    await new Promise((resolve) => fs.unlink(tempPath, resolve));
    
    console.log(`PPTX file saved to ${outputPath}`);
    
    // Add another delay before sending the file
    await new Promise(resolve => setTimeout(resolve, 200));
    
    // Use template name for the downloaded file if provided
    const downloadFilename = templateName ? `${templateName}.pptx` : 'updated-presentation.pptx';
    res.download(outputPath, downloadFilename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Add a delay before cleanup to ensure download completes
      setTimeout(() => {
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
      }, 1000);
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
  const replacementCount1 = (result.match(directRegex) || []).length;
  result = result.replace(directRegex, `$1${newText}$2`);
  
  // Try escaped version
  const escapedRegex = new RegExp(`(<a:t[^>]*>)${oldTextEscaped.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(</a:t>)`, 'g');
  const replacementCount2 = (result.match(escapedRegex) || []).length;
  result = result.replace(escapedRegex, `$1${newTextEscaped}$2`);
  
  console.log(`Text replacement: "${oldText}" -> "${newText}": ${replacementCount1 + replacementCount2} occurrences replaced`);
  
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
  
  // If still no replacements, try parsing the XML and updating text nodes directly
  if (result === slideXml) {
    try {
      const parser = new xml2js.Parser();
      parser.parseString(slideXml, (err, slideData) => {
        if (err) return;
        
        // Function to recursively search for text nodes
        const findAndReplaceText = (obj) => {
          if (!obj) return false;
          
          let replaced = false;
          
          // Check if this is a text node
          if (obj['a:t']) {
            for (let i = 0; i < obj['a:t'].length; i++) {
              if (obj['a:t'][i] === oldText || obj['a:t'][i] === oldTextEscaped) {
                obj['a:t'][i] = newText;
                replaced = true;
              }
            }
          }
          
          // Recursively check all properties
          for (const key in obj) {
            if (Array.isArray(obj[key])) {
              for (let i = 0; i < obj[key].length; i++) {
                if (typeof obj[key][i] === 'object') {
                  replaced = findAndReplaceText(obj[key][i]) || replaced;
                }
              }
            }
          }
          
          return replaced;
        };
        
        // Start the recursive search
        if (findAndReplaceText(slideData)) {
          // Convert back to XML
          const builder = new xml2js.Builder();
          result = builder.buildObject(slideData);
          console.log('Text replaced using XML parsing approach');
        }
      });
    } catch (parseError) {
      console.error('Error parsing slide XML for text replacement:', parseError);
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

// Ensure uploads directory exists
const uploadDir = 'uploads';
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
  console.log('Created uploads directory');
}

// Test endpoint for image placement
app.get('/api/test-image-placement', async (req, res) => {
  try {
    const { testImagePlacement } = require('./test-image-placement');
    
    // Find a PPTX file to test with
    const uploadDir = 'uploads';
    let pptxPath = null;
    let imagePath = null;
    
    if (fs.existsSync(uploadDir)) {
      const files = fs.readdirSync(uploadDir);
      
      // Find a PPTX file
      for (const file of files) {
        if (file.endsWith('.pptx')) {
          pptxPath = path.join(uploadDir, file);
          break;
        }
      }
      
      // Find an image file
      for (const file of files) {
        if (file.match(/\.(png|jpg|jpeg|gif)$/i)) {
          imagePath = path.join(uploadDir, file);
          break;
        }
      }
    }
    
    if (!pptxPath || !imagePath) {
      return res.status(400).json({ error: 'No PPTX or image file found in uploads directory' });
    }
    
    // Run the test
    const outputPath = await testImagePlacement(
      pptxPath,
      imagePath,
      1, // First slide
      {
        x: 1, // 1 inch from left
        y: 1, // 1 inch from top
        width: 3, // 3 inches wide
        height: 2 // 2 inches tall
      }
    );
    
    if (outputPath) {
      res.download(outputPath, 'test-image-placement.pptx', (err) => {
        if (err) {
          console.error('Download error:', err);
        }
        // Clean up the test file
        fs.unlink(outputPath, () => {});
      });
    } else {
      res.status(500).json({ error: 'Failed to create test PPTX' });
    }
  } catch (error) {
    console.error('Test error:', error);
    res.status(500).json({ error: 'Test failed: ' + error.message });
  }
});

// Fix PPTX endpoint
app.post('/api/fix-pptx', upload.single('pptx'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const { fixPptxImageDimensions } = require('./fix-pptx');
    const inputPath = req.file.path;
    const outputPath = path.join('uploads', `fixed-${Date.now()}-${req.file.originalname}`);
    
    const success = await fixPptxImageDimensions(inputPath, outputPath);
    
    if (success) {
      res.download(outputPath, `fixed-${req.file.originalname}`, (err) => {
        if (err) {
          console.error('Download error:', err);
        }
        // Clean up files
        fs.unlink(outputPath, () => {});
        fs.unlink(inputPath, () => {});
      });
    } else {
      res.status(500).json({ error: 'Failed to fix PPTX file' });
    }
  } catch (error) {
    console.error('Fix PPTX error:', error);
    res.status(500).json({ error: 'Failed to fix PPTX file' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Open your browser to view the application');
});
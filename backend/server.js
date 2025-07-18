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
const PORT = 3000;

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
    
    for (let i = 0; i < slideFiles.length; i++) {
      const slideXml = await contents.files[slideFiles[i]].async('string');
      const parser = new xml2js.Parser();
      const slideData = await parser.parseStringPromise(slideXml);
      
      // Get slide relationships
      const relsFile = relsFiles.find(rel => rel.includes(`slide${i + 1}.xml.rels`));
      let relationships = {};
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
      
      slides.push({
        id: i + 1,
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
  
  if (!slideData['p:sld'] || !slideData['p:sld']['p:cSld']) return elements;
  
  const spTree = slideData['p:sld']['p:cSld'][0]['p:spTree'][0];
  const allElements = [];
  
  // Collect all elements (shapes and pictures)
  if (spTree['p:sp']) allElements.push(...spTree['p:sp'].map(el => ({...el, type: 'shape'})));
  if (spTree['p:pic']) allElements.push(...spTree['p:pic'].map(el => ({...el, type: 'pic'})));
  if (spTree['p:cxnSp']) allElements.push(...spTree['p:cxnSp'].map(el => ({...el, type: 'connector'})));
  
  for (let idx = 0; idx < allElements.length; idx++) {
    const element = allElements[idx];
    
    // Get transform info
    let transform = null;
    if (element['p:spPr'] && element['p:spPr'][0] && element['p:spPr'][0]['a:xfrm']) {
      transform = element['p:spPr'][0]['a:xfrm'][0];
    }
    
    if (!transform) continue;
    
    const off = transform['a:off'] ? transform['a:off'][0].$ : { x: '0', y: '0' };
    const ext = transform['a:ext'] ? transform['a:ext'][0].$ : { cx: '914400', cy: '914400' };
    
    const x = parseInt(off.x) / 914400;
    const y = parseInt(off.y) / 914400;
    const width = parseInt(ext.cx) / 914400;
    const height = parseInt(ext.cy) / 914400;
    
    // Process text elements
    if (element.type === 'shape' && element['p:txBody']) {
      const textContent = extractTextFromShape(element);
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
        const embed = blip.$['r:embed'];
        const imagePath = relationships[embed];
        
        if (imagePath) {
          const fullImagePath = imagePath.startsWith('media/') ? `ppt/${imagePath}` : `ppt/media/${imagePath}`;
          let imageFile = zipContents.files[fullImagePath];
          
          // Try alternative paths
          if (!imageFile) {
            const altPaths = [
              `ppt/${imagePath}`,
              imagePath,
              `ppt/media/${imagePath.split('/').pop()}`
            ];
            
            for (const altPath of altPaths) {
              if (zipContents.files[altPath]) {
                imageFile = zipContents.files[altPath];
                break;
              }
            }
          }
          
          if (imageFile) {
            // Store image file reference instead of base64 data
            const imageName = `slide${slideNum}_image${idx}_${Date.now()}.${imagePath.split('.').pop()}`;
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
          }
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
    });
    
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
        
        // Update only changed text elements
        slideData.elements.forEach((element, idx) => {
          if (element.type === 'text') {
            const originalElement = originalSlide.elements.find(el => el.id === element.id);
            if (originalElement && originalElement.content !== element.content) {
              slideXml = replaceTextInXml(slideXml, originalElement.content, element.content);
            }
          }
        });
        
        // Update the slide XML
        contents.file(slideFile, slideXml);
      }
    }
    
    // Generate updated PPTX with exact same compression
    const updatedBuffer = await contents.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    const outputPath = `uploads/updated-${Date.now()}.pptx`;
    fs.writeFileSync(outputPath, updatedBuffer);
    
    res.download(outputPath, 'updated-presentation.pptx', (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Clean up files
      fs.unlink(outputPath, () => {});
      req.files?.forEach(file => fs.unlink(file.path, () => {}));
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
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');
  };
  
  const xmlUnescape = (str) => {
    return str.replace(/&amp;/g, '&')
              .replace(/&lt;/g, '<')
              .replace(/&gt;/g, '>')
              .replace(/&quot;/g, '"')
              .replace(/&apos;/g, "'");
  };
  
  // Try both escaped and unescaped versions
  const oldTextEscaped = xmlEscape(oldText);
  const newTextEscaped = xmlEscape(newText);
  
  // Replace with exact match
  let result = slideXml;
  
  // Try direct replacement first
  const directRegex = new RegExp(`(<a:t[^>]*>)${oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(<\/a:t>)`, 'g');
  result = result.replace(directRegex, `$1${newText}$2`);
  
  // Try escaped version
  const escapedRegex = new RegExp(`(<a:t[^>]*>)${oldTextEscaped.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(<\/a:t>)`, 'g');
  result = result.replace(escapedRegex, `$1${newTextEscaped}$2`);
  
  return result;
}

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Open your browser to view the application');
});
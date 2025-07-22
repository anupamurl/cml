/**
 * Enhanced shape generator for PPTX editor
 * Generates high-resolution SVG shapes for use in presentations
 */
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const { generateShape } = require('./shapeGenerator');

/**
 * Generate a high-resolution shape and convert to PNG
 * @param {string} type - Type of shape to generate
 * @param {Object} data - Configuration data for the shape
 * @returns {Promise<string>} - Path to the generated high-resolution PNG file
 */
async function generateHighResShape(type, data = {}) {
  // Increase dimensions for higher resolution SVG
  const baseWidth = data.width || 100;
  const baseHeight = data.height || 100;
  
  // Use 4x the original size for high resolution
  const highResData = {
    ...data,
    width: baseWidth * 4,
    height: baseHeight * 4
  };
  
  // Generate the SVG with high resolution
  const svgPath = generateShape(type, highResData);
  
  // Create uploads directory if it doesn't exist
  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }
  
  // Generate a unique filename with timestamp
  const timestamp = Date.now();
  const pngPath = path.join(uploadsDir, `shape_${timestamp}.png`);
  
  // Convert SVG to high-quality PNG
  await sharp(svgPath)
    .resize({
      width: baseWidth * 8, // Even higher resolution for PNG
      height: baseHeight * 8,
      fit: 'contain',
      withoutEnlargement: false
    })
    .png({ 
      quality: 100, 
      compressionLevel: 0, // No compression for maximum quality
      adaptiveFiltering: true, // Adaptive filtering for better quality
      force: true
    })
    .toFile(pngPath);
  
  // Clean up the SVG file
  fs.unlinkSync(svgPath);
  
  return pngPath;
}

/**
 * Process a shape for insertion into PPTX
 * @param {Object} contents - JSZip contents of the PPTX
 * @param {number} slideId - Slide ID
 * @param {Object} element - Shape element data
 * @returns {Promise<boolean>} - Success status
 */
async function processShapeForPptx(contents, slideId, element) {
  try {
    // Determine shape type (default to circle if not specified)
    const shapeType = element.shapeType || 'pie';
    
    // Configure shape data using element properties
    const shapeData = {
      width: Math.max(element.width || 100, 50) * 2,  // Double size for better quality
      height: Math.max(element.height || 100, 50) * 2,
      fillColor: element.fillColor || '#3498db'
    };
    
    // Add specific data for pie charts if needed
    if (shapeType === 'pie' || shapeType === 'segmentedCircle') {
      shapeData.segments = element.segments || [
        { color: '#e74c3c', percentage: 30 },
        { color: '#3498db', percentage: 45 },
        { color: '#2ecc71', percentage: 25 }
      ];
    }
    
    // Generate high-resolution PNG
    const pngPath = await generateHighResShape(shapeType, shapeData);
    console.log(`Generated high-resolution shape PNG at: ${pngPath}`);
    
    // Import the image insertion function
    const { insertImageIntoSlide } = require('./insertImage');
    
    // Calculate exact EMU values for positioning
    const x = parseInt(element.x * 914400);
    const y = parseInt(element.y * 914400);
    const width = parseInt(element.width * 914400);
    const height = parseInt(element.height * 914400);
    
    // Generate a unique ID for the image relationship
    const imageId = `rId${Date.now()}`;
    
    // Add the image to the PPTX
    await insertImageIntoSlide(contents, slideId, imageId, pngPath, {
      exactX: x,
      exactY: y,
      exactWidth: width,
      exactHeight: height
    });
    
    return true;
  } catch (error) {
    console.error(`Error processing shape: ${error.message}`);
    return false;
  }
}

module.exports = {
  generateHighResShape,
  processShapeForPptx
};
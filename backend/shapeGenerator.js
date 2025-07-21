/**
 * Shape generator for PPTX editor
 * Generates SVG shapes for use in presentations
 */
const fs = require('fs');
const path = require('path');

/**
 * Generate a circular segmented diagram SVG
 * @param {string} outputPath - Path to save the SVG file
 * @returns {string} - Path to the generated SVG file
 */
function generateSegmentedCircle(outputPath) {
  const svgContent = `<svg width="100" height="100" xmlns="http://www.w3.org/2000/svg">
    <circle cx="50" cy="50" r="45" stroke="black" stroke-width="1" fill="none" />
    <path d="M50,5 A45,45 0 0,1 95,50" stroke="red" stroke-width="10" fill="none" />
    <path d="M95,50 A45,45 0 0,1 50,95" stroke="blue" stroke-width="10" fill="none" />
    <path d="M50,95 A45,45 0 0,1 5,50" stroke="green" stroke-width="10" fill="none" />
    <path d="M5,50 A45,45 0 0,1 50,5" stroke="orange" stroke-width="10" fill="none" />
    <circle cx="50" cy="50" r="20" stroke="black" stroke-width="1" fill="white" />
  </svg>`;
  
  fs.writeFileSync(outputPath, svgContent);
  return outputPath;
}

/**
 * Generate an SVG shape based on type
 * @param {string} type - Type of shape to generate
 * @returns {string} - Path to the generated SVG file
 */
function generateShape(type) {
  // Create uploads directory if it doesn't exist
  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }
  
  // Generate a unique filename
  const svgPath = path.join(uploadsDir, `shape_${Date.now()}.svg`);
  
  // Generate the appropriate shape based on type
  switch (type) {
    case 'shape':
    default:
      return generateSegmentedCircle(svgPath);
  }
}

module.exports = {
  generateShape
};
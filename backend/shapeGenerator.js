/**
 * Shape generator for PPTX editor
 * Generates SVG shapes for use in presentations
 */
const fs = require('fs');
const path = require('path');

/**
 * Generate a circular segmented diagram SVG
 * @param {string} outputPath - Path to save the SVG file
 * @param {Object} data - Configuration data for the circle
 * @param {number} data.width - Width of the SVG
 * @param {number} data.height - Height of the SVG
 * @param {Array} data.segments - Array of segment objects with color and percentage
 * @param {string} data.centerText - Optional text to display in the center
 * @returns {string} - Path to the generated SVG file
 */
function generateSegmentedCircle(outputPath, data = {}) {
  // Default values
  const width = data.width || 100;
  const height = data.height || 100;
  const segments = data.segments || [
    { color: 'red', percentage: 45 },
    { color: 'blue', percentage: 5 },
    { color: 'green', percentage: 30 },
    { color: 'orange', percentage: 20 }
  ];
  const centerText = data.centerText || '';
  
  // Calculate center and radius
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(cx, cy) - 5;
  const innerRadius = radius * 0.4;
  
  // Start SVG content
  let svgContent = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
    <circle cx="${cx}" cy="${cy}" r="${radius}" stroke="black" stroke-width="1" fill="none" />`;
  
  // Generate segments
  let startAngle = 0;
  segments.forEach(segment => {
    const angle = (segment.percentage / 100) * 360;
    const endAngle = startAngle + angle;
    
    // Convert angles to radians
    const startRad = (startAngle - 90) * Math.PI / 180;
    const endRad = (endAngle - 90) * Math.PI / 180;
    
    // Calculate points
    const startX = cx + radius * Math.cos(startRad);
    const startY = cy + radius * Math.sin(startRad);
    const endX = cx + radius * Math.cos(endRad);
    const endY = cy + radius * Math.sin(endRad);
    
    // Create path - large arc flag is 0 for arcs less than 180 degrees, 1 for arcs greater than 180 degrees
    const largeArcFlag = angle > 180 ? 1 : 0;
    
    svgContent += `
    <path d="M${cx},${cy} L${startX},${startY} A${radius},${radius} 0 ${largeArcFlag},1 ${endX},${endY} Z" fill="${segment.color}" stroke="black" stroke-width="1" />`;
    
    startAngle = endAngle;
  });
  
  // Add inner circle
  svgContent += `
    <circle cx="${cx}" cy="${cy}" r="${innerRadius}" stroke="black" stroke-width="1" fill="white" />`;
  
  // Add center text if provided
  if (centerText) {
    svgContent += `
    <text x="${cx}" y="${cy}" text-anchor="middle" dominant-baseline="middle" font-size="${innerRadius/2}px">${centerText}</text>`;
  }
  
  // Close the SVG tag
  svgContent += `
</svg>`;
  
  fs.writeFileSync(outputPath, svgContent);
  return outputPath;
}

/**
 * Generate an SVG shape based on type
 * @param {string} type - Type of shape to generate
 * @param {Object} data - Configuration data for the shape
 * @returns {string} - Path to the generated SVG file
 */
function generateShape(type, data = {}) {
  // Create uploads directory if it doesn't exist
  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
  }
  
  // Generate a unique filename
  const svgPath = path.join(uploadsDir, `shape_${Date.now()}.svg`);
  
  // Generate the appropriate shape based on type
  switch (type) {
    case 'circle':
      return generateSegmentedCircle(svgPath, data);
    case 'shape':
    default:
      return generateSegmentedCircle(svgPath, data);
  }
}

module.exports = {
  generateShape,
  generateSegmentedCircle
};
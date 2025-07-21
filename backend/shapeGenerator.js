/**
 * Shape generator for PPTX editor
 * Generates SVG shapes for use in presentations
 */
const fs = require('fs');
const path = require('path');

/**
 * Generate a basic circle SVG
 * @param {string} outputPath - Path to save the SVG file
 * @param {Object} data - Configuration data for the circle
 * @param {number} data.width - Width of the SVG
 * @param {number} data.height - Height of the SVG
 * @param {string} data.fillColor - Fill color for the circle
 * @returns {string} - Path to the generated SVG file
 */
function generateCircle(outputPath, data = {}) {
  // Default values
  const width = data.width || 100;
  const height = data.height || 100;
  const fillColor = data.fillColor || '#3498db';
  
  // Calculate center and radius
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(cx, cy) - 5;
  
  // Create SVG content
  const svgContent = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
    <circle cx="${cx}" cy="${cy}" r="${radius}" stroke="black" stroke-width="1" fill="${fillColor}" />
  </svg>`;
  
  fs.writeFileSync(outputPath, svgContent);
  return outputPath;
}

/**
 * Generate a pie chart SVG
 * @param {string} outputPath - Path to save the SVG file
 * @param {Object} data - Configuration data for the pie
 * @param {number} data.width - Width of the SVG
 * @param {number} data.height - Height of the SVG
 * @param {Array} data.segments - Array of segment objects with color and percentage
 * @returns {string} - Path to the generated SVG file
 */
function generatePie(outputPath, data = {}) {
  // Default values
  const width = data.width || 100;
  const height = data.height || 100;
  const segments = data.segments || [
    { color: '#e74c3c', percentage: 30 },
    { color: '#3498db', percentage: 45 },
    { color: '#2ecc71', percentage: 25 }
  ];
  
  // Calculate center and radius
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(cx, cy) - 5;
  
  // Start SVG content
  let svgContent = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">`;
  
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
  
  // Close the SVG tag
  svgContent += `
</svg>`;
  
  fs.writeFileSync(outputPath, svgContent);
  return outputPath;
}

/**
 * Generate a right arrow SVG
 * @param {string} outputPath - Path to save the SVG file
 * @param {Object} data - Configuration data for the arrow
 * @param {number} data.width - Width of the SVG
 * @param {number} data.height - Height of the SVG
 * @param {string} data.fillColor - Fill color for the arrow
 * @returns {string} - Path to the generated SVG file
 */
function generateRightArrow(outputPath, data = {}) {
  // Default values
  const width = data.width || 100;
  const height = data.height || 50;
  const fillColor = data.fillColor || '#3498db';
  
  // Calculate dimensions
  const arrowHeadWidth = height * 0.8;
  const arrowBodyHeight = height * 0.4;
  const arrowBodyWidth = width - arrowHeadWidth;
  const arrowBodyY = (height - arrowBodyHeight) / 2;
  
  // Create SVG content for right-pointing arrow
  const svgContent = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
    <path d="
      M 0,${arrowBodyY}
      L ${arrowBodyWidth},${arrowBodyY}
      L ${arrowBodyWidth},${arrowBodyY - arrowBodyHeight/2}
      L ${width},${height/2}
      L ${arrowBodyWidth},${arrowBodyY + arrowBodyHeight + arrowBodyHeight/2}
      L ${arrowBodyWidth},${arrowBodyY + arrowBodyHeight}
      L 0,${arrowBodyY + arrowBodyHeight}
      Z
    " fill="${fillColor}" stroke="black" stroke-width="1" />
  </svg>`;
  
  fs.writeFileSync(outputPath, svgContent);
  return outputPath;
}

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
  switch (type.toLowerCase()) {
    case 'circle':
      return generateCircle(svgPath, data);
    case 'pie':
      return generatePie(svgPath, data);
    case 'rightarrow':
    case 'right-arrow':
      return generateRightArrow(svgPath, data);
    case 'segmentedcircle':
      return generateSegmentedCircle(svgPath, data);
    case 'shape':
    default:
      // Default to circle if no specific shape is requested
      return generateCircle(svgPath, data);
  }
}

module.exports = {
  generateShape,
  generateCircle,
  generatePie,
  generateRightArrow,
  generateSegmentedCircle
};

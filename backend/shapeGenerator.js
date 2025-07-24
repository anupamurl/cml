/**
 * Donut chart generator for PPTX editor
 * Generates donut chart SVGs for use in presentations
 */
const fs = require('fs');
const path = require('path');

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
  
  // Only generate donut chart
  return generateDonutChart(svgPath, data.chartData, data.size);
}

/**
 * Generate a donut chart SVG with icons and labels
 * @param {string} outputPath - Path to save the SVG file
 * @param {Array} data - Array of data objects with name, nps, responses and icon properties
 * @param {number} size - Size of the chart
 * @returns {string} - Path to the generated SVG file
 */
function generateDonutChart(outputPath, data = null, size = 300) {
  // Increase size for better quality
  size = Math.max(size, 600); // Ensure minimum size for readability
  
  const defaultData = [
    { name: "BLUE", nps: 60.5, responses: 3874, icon: "diamond" },
    { name: "CLASSIC", nps: 55.5, responses: 32481, icon: "star" },
    { name: "SELECT", nps: 54.7, responses: 23574, icon: "medal" },
    { name: "OTHER", nps: 54.5, responses: 1889, icon: "user" },
    { name: "CELESTIA", nps: 52.6, responses: 20634, icon: "crown" },
    { name: "ETERNIA", nps: 52.5, responses: 73854, icon: "infinity" },
    { name: "ETERNIAX", nps: 55.5, responses: 73854, icon: "infinity" }
  ];

  const chartData = data && Array.isArray(data) && data.length ? data : defaultData;

  // Calculate dimensions for centered chart
  const svgWidth = size * 2;
  const svgHeight = size;
  const cx = svgWidth / 2; // Center X is in the middle of the SVG width
  let cy = svgHeight / 2; // Center Y is in the middle of the SVG height
  const outerRadius = Math.min(svgHeight / 2, svgWidth / 4) - 40; // Ensure chart fits within SVG
  const innerRadius = outerRadius / 2;

  const iconMap = {
    diamond: "◆",
    star: "★",
    medal: "🏅",
    user: "👤",
    crown: "👑",
    infinity: "∞"
  };

  const total = chartData.length;
  const anglePer = 360 / total;

  // Create SVG with higher resolution and explicit dimensions with padding to prevent text cutoff
  const padding = 120; // Add padding to prevent text cutoff
  let svg = `<svg width="${svgWidth}" height="${svgHeight + padding*2}" viewBox="0 0 ${svgWidth} ${svgHeight + padding*2}" xmlns="http://www.w3.org/2000/svg" font-family="Arial, Helvetica, sans-serif" style="text-rendering: geometricPrecision;">`;
  
  // Add defs for text styles to ensure PowerPoint compatibility
  svg += `
  <defs>
    <style type="text/css">
      .title { font: bold ${size/25}px Arial; fill: black; }
      .value { font: bold ${size/30}px Arial; fill: green; }
      .subtitle { font: ${size/35}px Arial; fill: black; }
    </style>
  </defs>`;
  
  // Adjust center Y position to account for padding
  cy += padding;

  let angleStart = 0;

  chartData.forEach((item, i) => {
    const angleMid = angleStart + anglePer / 2;
    const angleEnd = angleStart + anglePer;

    const startRad = (angleStart - 90) * Math.PI / 180;
    const endRad = (angleEnd - 90) * Math.PI / 180;
    const midRad = (angleMid - 90) * Math.PI / 180;

    const x1 = cx + outerRadius * Math.cos(startRad);
    const y1 = cy + outerRadius * Math.sin(startRad);
    const x2 = cx + outerRadius * Math.cos(endRad);
    const y2 = cy + outerRadius * Math.sin(endRad);
    const x3 = cx + innerRadius * Math.cos(endRad);
    const y3 = cy + innerRadius * Math.sin(endRad);
    const x4 = cx + innerRadius * Math.cos(startRad);
    const y4 = cy + innerRadius * Math.sin(startRad);

    const largeArcFlag = anglePer > 180 ? 1 : 0;

    // Add segment with slightly darker color for better visibility
    svg += `
      <path d="M${x1},${y1} 
               A${outerRadius},${outerRadius} 0 ${largeArcFlag},1 ${x2},${y2}
               L${x3},${y3}
               A${innerRadius},${innerRadius} 0 ${largeArcFlag},0 ${x4},${y4}
               Z"
            fill="#dce6f2" stroke="white" stroke-width="1.5" />`;

    // Add icon with larger font size
    const iconX = cx + (innerRadius + outerRadius) / 2 * Math.cos(midRad);
    const iconY = cy + (innerRadius + outerRadius) / 2 * Math.sin(midRad);
    svg += `<text x="${iconX}" y="${iconY}" font-size="${size/15}" text-anchor="middle" dominant-baseline="central" font-weight="bold">${iconMap[item.icon] || "?"}</text>`;

    // Calculate label position with more space to avoid cutting off
    // Use different positioning strategy based on angle to avoid text overlap
    const labelDistance = outerRadius + 120; // Increased distance from chart
    const labelX = cx + labelDistance * Math.cos(midRad);
    const labelY = cy + labelDistance * Math.sin(midRad);
    
    // Adjust text anchor based on position to avoid cutting off
    const textAnchor = midRad > -Math.PI/4 && midRad < Math.PI/4 ? "start" : 
                      midRad > 3*Math.PI/4 || midRad < -3*Math.PI/4 ? "end" : "middle";
    
    // Use group to keep text elements together
    svg += `
    <g text-anchor="${textAnchor}">
      <text x="${labelX}" y="${labelY - 20}" class="title">${item.name}</text>
      <text x="${labelX}" y="${labelY + 10}" class="value">NPS ${item.nps}</text>
      <text x="${labelX}" y="${labelY + 35}" class="subtitle">${item.responses.toLocaleString()} responses</text>
    </g>`;

    angleStart += anglePer;
  });

  svg += `</svg>`;

  fs.writeFileSync(outputPath, svg);
  return outputPath;
}

module.exports = {
  generateShape,
  generateDonutChart
};

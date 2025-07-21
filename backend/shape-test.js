/**
 * Test file for shape generator
 * Shows how to use the shape generator functions
 */
const fs = require('fs');
const path = require('path');
const { generateShape } = require('./shapeGenerator');

// Test function to generate all shape types
async function testShapeGenerator() {
  console.log('Testing shape generator...');
  
  // Test circle
  const circlePath = generateShape('circle', {
    width: 200,
    height: 200,
    fillColor: '#3498db'
  });
  console.log(`Circle generated at: ${circlePath}`);
  
  // Test pie chart
  const piePath = generateShape('pie', {
    width: 200,
    height: 200,
    segments: [
      { color: '#e74c3c', percentage: 30 },
      { color: '#3498db', percentage: 45 },
      { color: '#2ecc71', percentage: 25 }
    ]
  });
  console.log(`Pie chart generated at: ${piePath}`);
  
  // Test right arrow
  const arrowPath = generateShape('rightArrow', {
    width: 200,
    height: 100,
    fillColor: '#f39c12'
  });
  console.log(`Right arrow generated at: ${arrowPath}`);
  
  // Test segmented circle
  const segmentedCirclePath = generateShape('segmentedCircle', {
    width: 200,
    height: 200,
    segments: [
      { color: '#e74c3c', percentage: 20 },
      { color: '#3498db', percentage: 30 },
      { color: '#2ecc71', percentage: 25 },
      { color: '#f39c12', percentage: 25 }
    ],
    centerText: '100%'
  });
  console.log(`Segmented circle generated at: ${segmentedCirclePath}`);
  
  console.log('Shape generator test complete!');
}

// Run the test
testShapeGenerator().catch(console.error);
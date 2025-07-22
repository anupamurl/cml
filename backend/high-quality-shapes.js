/**
 * High-quality shape generator test
 * This script generates SVG shapes with improved rendering quality
 */
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const { generateShape } = require('./shapeGenerator');

// Test function to generate high-quality shapes
async function generateHighQualityShapes() {
  console.log('Generating high-quality shapes...');
  
  // Generate a circle with high resolution
  const circlePath = generateShape('circle', {
    width: 400,  // Higher resolution
    height: 400, // Higher resolution
    fillColor: '#3498db'
  });
  console.log(`Circle generated at: ${circlePath}`);
  
  // Convert SVG to high-quality PNG
  const circlePngPath = circlePath.replace('.svg', '.png');
  await sharp(circlePath)
    .resize({
      width: 800, // Even higher resolution for PNG
      height: 800,
      fit: 'contain',
      withoutEnlargement: false
    })
    .png({ quality: 100, compressionLevel: 0 }) // Maximum quality
    .toFile(circlePngPath);
  console.log(`Circle PNG generated at: ${circlePngPath}`);
  
  // Generate a pie chart with high resolution
  const piePath = generateShape('pie', {
    width: 400,
    height: 400,
    segments: [
      { color: '#e74c3c', percentage: 30 },
      { color: '#3498db', percentage: 45 },
      { color: '#2ecc71', percentage: 25 }
    ]
  });
  console.log(`Pie chart generated at: ${piePath}`);
  
  // Convert SVG to high-quality PNG
  const piePngPath = piePath.replace('.svg', '.png');
  await sharp(piePath)
    .resize({
      width: 800,
      height: 800,
      fit: 'contain',
      withoutEnlargement: false
    })
    .png({ quality: 100, compressionLevel: 0 })
    .toFile(piePngPath);
  console.log(`Pie PNG generated at: ${piePngPath}`);
  
  // Generate a right arrow with high resolution
  const arrowPath = generateShape('rightArrow', {
    width: 400,
    height: 200,
    fillColor: '#f39c12'
  });
  console.log(`Right arrow generated at: ${arrowPath}`);
  
  // Convert SVG to high-quality PNG
  const arrowPngPath = arrowPath.replace('.svg', '.png');
  await sharp(arrowPath)
    .resize({
      width: 800,
      height: 400,
      fit: 'contain',
      withoutEnlargement: false
    })
    .png({ quality: 100, compressionLevel: 0 })
    .toFile(arrowPngPath);
  console.log(`Arrow PNG generated at: ${arrowPngPath}`);
  
  // Generate a segmented circle with high resolution
  const segmentedCirclePath = generateShape('segmentedCircle', {
    width: 400,
    height: 400,
    segments: [
      { color: '#e74c3c', percentage: 20 },
      { color: '#3498db', percentage: 30 },
      { color: '#2ecc71', percentage: 25 },
      { color: '#f39c12', percentage: 25 }
    ],
    centerText: '100%'
  });
  console.log(`Segmented circle generated at: ${segmentedCirclePath}`);
  
  // Convert SVG to high-quality PNG
  const segmentedCirclePngPath = segmentedCirclePath.replace('.svg', '.png');
  await sharp(segmentedCirclePath)
    .resize({
      width: 800,
      height: 800,
      fit: 'contain',
      withoutEnlargement: false
    })
    .png({ quality: 100, compressionLevel: 0 })
    .toFile(segmentedCirclePngPath);
  console.log(`Segmented circle PNG generated at: ${segmentedCirclePngPath}`);
  
  console.log('High-quality shape generation complete!');
}

// Run the test
generateHighQualityShapes().catch(console.error);
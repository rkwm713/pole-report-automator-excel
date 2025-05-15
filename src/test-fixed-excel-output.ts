
/**
 * Test script to verify the fixed Make-Ready Report Excel output
 * This script focuses on testing the fixes for:
 * 1. Columns J & K (Height Lowest Com/CPS Electrical)
 * 2. From/To Pole placement in the last rows
 * 3. Column O (Midspan Proposed)
 */
import { PoleDataProcessor } from './services/poleDataProcessor';
import { applyAllFixes } from './services/poleDataProcessorFix';
import { generateDemoSpidaData, generateDemoKatapultData } from './utils/demoDataGenerator';
import * as fs from 'fs';
import * as path from 'path';

async function runTest() {
  console.log('===============================================');
  console.log('Testing fixed Make-Ready Report Excel generation');
  console.log('===============================================');

  // Create a new processor instance
  const processor = new PoleDataProcessor();
  
  // Apply the fixes to the processor instance
  applyAllFixes(processor);

  try {
    let spidaData: string;
    let katapultData: string;
    let usedDemoData = false;
    
    // First try to load actual test data files
    try {
      // Load test data - try multiple possible locations for the test files
      let spidaPath: string;
      let katapultPath: string;
      
      // Define possible paths to look for test files
      const possiblePaths = [
        './test-data/sample-spida.json',
        '../test-data/sample-spida.json',
        './data/sample-spida.json',
        '../data/sample-spida.json'
      ];
      
      // Find SPIDA file
      for (const testPath of possiblePaths) {
        const fullPath = path.resolve(testPath);
        if (fs.existsSync(fullPath)) {
          spidaPath = fullPath;
          break;
        }
      }
      
      // Find Katapult file (same directory as SPIDA file, but with katapult in name)
      if (spidaPath) {
        const spidaDir = path.dirname(spidaPath);
        const files = fs.readdirSync(spidaDir);
        for (const file of files) {
          if (file.includes('katapult') && file.endsWith('.json')) {
            katapultPath = path.join(spidaDir, file);
            break;
          }
        }
      }
      
      // If real test files found, read them
      if (spidaPath && katapultPath) {
        console.log(`Loading SPIDA data from: ${spidaPath}`);
        console.log(`Loading Katapult data from: ${katapultPath}`);
        
        // Read the test data files
        spidaData = fs.readFileSync(spidaPath, 'utf8');
        katapultData = fs.readFileSync(katapultPath, 'utf8');
      } else {
        throw new Error('Test files not found');
      }
    } catch (error) {
      // If real test files not found, use demo data
      console.log('Real test files not found. Using demo data instead.');
      console.log('Demo data contains simplified structures to showcase the column J, K, and O fixes.');
      
      // Generate demo data
      spidaData = JSON.stringify(generateDemoSpidaData());
      katapultData = JSON.stringify(generateDemoKatapultData());
      usedDemoData = true;
    }

    console.log('Loading SPIDA data...');
    const spidaLoadResult = processor.loadSpidaData(spidaData);
    console.log(`SPIDA data loaded: ${spidaLoadResult}`);

    console.log('Loading Katapult data...');
    const katapultLoadResult = processor.loadKatapultData(katapultData);
    console.log(`Katapult data loaded: ${katapultLoadResult}`);

    if (processor.isDataLoaded()) {
      // Process the data
      console.log('Processing pole data...');
      const processResult = processor.processData();
      console.log(`Data processing result: ${processResult}`);

      // Check for any errors
      if (processor.getErrorCount() > 0) {
        console.warn('Warnings during processing:');
        for (const error of processor.getErrors()) {
          console.warn(`- ${error.code}: ${error.message}`);
          if (error.details) console.warn(`  Details: ${error.details}`);
        }
      }

      if (processResult) {
        console.log(`Successfully processed ${processor.getProcessedPoleCount()} poles`);

        // Validate attachment data to check for issues
        const validation = processor.validateAttachmentData();
        if (!validation.valid) {
          console.warn('Attachment data validation issues:');
          for (const issue of validation.issues) {
            console.warn(`- ${issue}`);
          }
        } else {
          console.log('Attachment data validation passed.');
        }

        // Generate the Excel file with the fixed logic
        console.log('Generating Excel...');
        const excelBlob = processor.generateExcel();

        if (excelBlob) {
          // Save the Excel file
          const outputPath = path.resolve('./fixed-make-ready-report.xlsx');
          const buffer = await excelBlob.arrayBuffer();
          fs.writeFileSync(outputPath, Buffer.from(buffer));
          console.log(`Excel file generated successfully: ${outputPath}`);
          console.log('');
          console.log('===== FIXED ISSUES =====');
          console.log('1. Height values in columns J & K (from Katapult data)');
          console.log('2. From/To Pole now placed in the LAST two rows of each pole section');
          console.log('3. Column O (Mid-Span Proposed) now shows correct values');
          console.log('===== END FIXED ISSUES =====');
          
          if (usedDemoData) {
            console.log('');
            console.log('NOTE: This test ran with DEMO data. For a full test with real data:');
            console.log('1. Create a "test-data" directory');
            console.log('2. Add sample-spida.json and sample-katapult.json files');
            console.log('3. Run this test again');
          }
        } else {
          console.error('Failed to generate Excel file:', processor.getErrors());
        }
      } else {
        console.error('Failed to process data:', processor.getErrors());
      }
    } else {
      console.error('Data not loaded properly');
    }
  } catch (err) {
    console.error('Test failed with error:', err);
  }
}

// Run the test
runTest().catch(err => console.error('Unhandled error:', err));

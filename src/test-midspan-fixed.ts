/**
 * Test script to verify the modified mid-span logic in PoleDataProcessor
 */
import { PoleDataProcessor } from './services/poleDataProcessor';
import * as fs from 'fs';
import * as path from 'path';

async function runTest() {
  console.log('Testing fixed mid-span logic...');

  // Create a new processor instance
  const processor = new PoleDataProcessor();

  try {
    // Load test data - replace with paths to your actual test files
    // You'll need to adjust these paths to point to your actual SPIDA and Katapult files
    const spidaData = fs.readFileSync(path.resolve('./test-data/sample-spida.json'), 'utf8');
    const katapultData = fs.readFileSync(path.resolve('./test-data/sample-katapult.json'), 'utf8');

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

      if (processResult) {
        console.log(`Processed ${processor.getProcessedPoleCount()} poles`);

        // Generate the Excel file with the fixed mid-span logic
        console.log('Generating Excel...');
        const excelBlob = processor.generateExcel();

        if (excelBlob) {
          // Save the Excel file
          const buffer = await excelBlob.arrayBuffer();
          fs.writeFileSync(path.resolve('./test-midspan-fixed.xlsx'), Buffer.from(buffer));
          console.log('Excel file generated successfully: test-midspan-fixed.xlsx');
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

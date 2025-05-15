
/**
 * Index file to centralize and export all fixes for the PoleDataProcessor class
 */
import { fixCalculateEndRow } from './calculateEndRowFix';
import { fixExistingMidspanExtraction } from './midspanExtractionFix';
import { fixFromToPoleRows } from './fromToPoleRowsFix';
import { fixMidspanProposedHeight } from './midspanProposedHeightFix';

/**
 * Apply all fixes to the PoleDataProcessor instance
 * @param poleDataProcessor - The instance to patch
 */
export function applyAllFixes(poleDataProcessor: any) {
  console.log("Applying all fixes to PoleDataProcessor");
  fixCalculateEndRow(poleDataProcessor);
  fixExistingMidspanExtraction(poleDataProcessor);
  fixFromToPoleRows(poleDataProcessor);
  fixMidspanProposedHeight(poleDataProcessor);
  console.log("All fixes applied successfully");
}

// Export individual fixes for direct access
export {
  fixCalculateEndRow,
  fixExistingMidspanExtraction,
  fixFromToPoleRows,
  fixMidspanProposedHeight
};

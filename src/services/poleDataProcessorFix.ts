
/**
 * Enhanced fixes for the PoleDataProcessor class
 * This file contains monkey patching functions that fix specific issues with the Excel output
 */

// Re-export all fixes from the refactored modules
export { 
  applyAllFixes,
  fixCalculateEndRow,
  fixExistingMidspanExtraction,
  fixFromToPoleRows,
  fixMidspanProposedHeight
} from './fixes';



/**
 * Fix for the _calculateEndRow method to correctly account for From/To Pole rows
 */

/**
 * Fix for the _calculateEndRow method to correctly account for From/To Pole rows
 * @param poleDataProcessor - The instance of PoleDataProcessor to patch
 */
export function fixCalculateEndRow(poleDataProcessor: any) {
  console.log("Applying fix for _calculateEndRow to correctly include From/To Pole rows");
  
  // Store the original method for calling within our patched method
  const originalCalculateEndRow = poleDataProcessor._calculateEndRow;
  
  // Replace with enhanced method
  poleDataProcessor._calculateEndRow = function(startRow: number, pole: any) {
    // Call original method to get its calculation
    const baseEndRow = originalCalculateEndRow.call(this, startRow, pole);
    
    // Add 2 more rows for From/To Pole placement at the end of each pole section
    return baseEndRow + 2;
  };
}

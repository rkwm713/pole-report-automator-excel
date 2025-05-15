
/**
 * Utility functions for pole data processing
 */

/**
 * Convert meters to feet
 */
export function metersToFeet(meters: number): number {
  return meters * 3.28084;
}

/**
 * Convert meters to feet and inches string
 */
export function metersToFeetInches(meters: number): string {
  if (typeof meters !== 'number') {
    return "N/A";
  }
  
  const totalFeet = meters * 3.28084;
  const wholeFeet = Math.floor(totalFeet);
  const inches = Math.round((totalFeet - wholeFeet) * 12);
  
  // Handle case where inches round up to 12
  if (inches === 12) {
    return `${wholeFeet + 1}'-0"`;
  }
  
  return `${wholeFeet}'-${inches}"`;
}

/**
 * Format height value to feet-inches
 */
export function formatHeightValue(value: number): string {
  try {
    // Check if the value is already in feet
    if (value > 3 && value < 200) {
      // Likely already in feet, format as feet-inches
      const wholeFeet = Math.floor(value);
      const inches = Math.round((value - wholeFeet) * 12);
      
      // Handle case where inches == 12
      if (inches === 12) {
        return `${wholeFeet + 1}'-0"`;
      }
      
      return `${wholeFeet}'-${inches}"`;
    } else {
      // Likely in meters, convert to feet first
      return metersToFeetInches(value);
    }
  } catch (error) {
    console.warn("Error formatting height value:", error);
    return "N/A";
  }
}

/**
 * Canonicalize pole ID by removing common prefixes/characters
 */
export function canonicalizePoleID(poleId: string): string {
  if (!poleId) return "Unknown";
  
  let id = String(poleId).trim();
  
  // Remove common prefixes
  const prefixes = ["POLE", "POLE-", "PL-", "PL", "P-", "P"];
  for (const prefix of prefixes) {
    if (id.toUpperCase().startsWith(prefix)) {
      id = id.substring(prefix.length);
      break;
    }
  }
  
  // Remove any remaining leading non-alphanumeric characters
  id = id.replace(/^[^a-zA-Z0-9]+/, "");
  
  // Handle empty ID
  return id || "Unknown";
}

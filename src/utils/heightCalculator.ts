
/**
 * Utility functions for height calculations and conversions
 * Used by the pole data processor for height-related operations
 */

/**
 * Convert height from meters to total inches
 * @param meters - Height in meters
 * @returns Height in total inches
 */
export function metersToInches(meters: number): number {
  if (meters === undefined || meters === null || isNaN(meters)) {
    return 0;
  }
  // 1 meter = 39.3701 inches
  return meters * 39.3701;
}

/**
 * Convert height from inches to feet and inches string format
 * @param totalInches - Height in total inches
 * @returns Formatted height string in ft'-in" format (e.g., "23'-7"")
 */
export function formatHeightString(totalInches: number | null): string {
  if (totalInches === null || totalInches === undefined || isNaN(totalInches)) {
    return "N/A";
  }
  
  const feet = Math.floor(totalInches / 12);
  const inches = Math.round(totalInches % 12);
  
  // Handle case where inches rounds up to 12
  if (inches === 12) {
    return `${feet + 1}'-0"`;
  }
  
  return `${feet}'-${inches}"`;
}

/**
 * Convert height from meters directly to feet-inches string
 * @param meters - Height in meters
 * @returns Formatted height string in ft'-in" format (e.g., "23'-7"")
 */
export function metersToHeightString(meters: number | null): string {
  if (meters === null || meters === undefined || isNaN(meters)) {
    return "N/A";
  }
  
  const totalInches = metersToInches(meters);
  return formatHeightString(totalInches);
}

/**
 * Calculate the lowest height across multiple wires of a specific category
 * @param wireHeights - Array of height objects with category and value
 * @param category - Wire category to filter by
 * @returns Lowest height found or null if none
 */
export function findLowestCategorizedHeight(
  wireHeights: Array<{ category: string; heightInInches: number }>,
  category: string
): number | null {
  // Filter heights by category
  const categoryHeights = wireHeights
    .filter(wire => wire.category === category)
    .map(wire => wire.heightInInches);
  
  if (categoryHeights.length === 0) {
    return null;
  }
  
  // Find the lowest height
  return Math.min(...categoryHeights);
}

/**
 * Score how well two wires match based on owner, type, and height similarity
 * Used for finding matching wires between different data sources
 * @param wire1 - First wire data
 * @param wire2 - Second wire data
 * @returns Score from 0-100 where higher means better match
 */
export function scoreWireMatch(
  wire1: { owner?: string; type?: string; height?: number },
  wire2: { owner?: string; type?: string; height?: number }
): number {
  let score = 0;
  
  // Owner match (most important - 60 points)
  if (wire1.owner && wire2.owner && 
      wire1.owner.toLowerCase() === wire2.owner.toLowerCase()) {
    score += 60;
  }
  
  // Type match (important - 30 points)
  if (wire1.type && wire2.type && 
      wire1.type.toLowerCase() === wire2.type.toLowerCase()) {
    score += 30;
  }
  
  // Height similarity (least important - up to 10 points)
  if (wire1.height !== undefined && wire2.height !== undefined) {
    const heightDiffInches = Math.abs(wire1.height - wire2.height);
    // Give full 10 points for exact match, sliding scale down to 0 points for differences >= 24 inches
    if (heightDiffInches === 0) {
      score += 10;
    } else if (heightDiffInches < 24) {
      score += Math.round(10 * (1 - heightDiffInches / 24));
    }
  }
  
  return score;
}

/**
 * Service for extracting pole data from SPIDA and Katapult
 */
import { PoleData, SpanData, AttachmentData, KatapultMidspanData } from "../types/poleTypes";
import { metersToFeet, metersToFeetInches, formatHeightValue, canonicalizePoleID } from "../utils/poleUtils";

export class PoleDataExtractor {
  private poleLookupMap: Map<string, any>;
  private katapultPoleLookupMap: Map<string, any>;
  private katapultMidspanMap: Map<string, KatapultMidspanData>;

  constructor(
    poleLookupMap: Map<string, any>,
    katapultPoleLookupMap: Map<string, any>,
    katapultMidspanMap: Map<string, KatapultMidspanData>
  ) {
    this.poleLookupMap = poleLookupMap;
    this.katapultPoleLookupMap = katapultPoleLookupMap;
    this.katapultMidspanMap = katapultMidspanMap;
  }

  /**
   * Extract data for a single pole
   */
  extractPoleData(poleLocationData: any, katapultPoleData: any, operationNumber: number): PoleData | null {
    try {
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        throw new Error("Could not find required designs in pole data");
      }
      
      console.log(`DEBUG: Found design indices: measured=${designIndices.measured}, recommended=${designIndices.recommended}`);
      
      // Extract basic pole information
      const canonicalPoleId = canonicalizePoleID(poleLocationData.label);
      
      // Extract pole owner (prioritize Katapult if available)
      let poleOwner = this._extractPoleOwner(poleLocationData, katapultPoleData);
      console.log(`DEBUG: Extracted pole owner: ${poleOwner}`);
      
      // Extract pole structure details
      const poleStructure = this._extractPoleStructure(poleLocationData);
      console.log(`DEBUG: Extracted pole structure: ${poleStructure}`);
      
      // Extract proposed riser information
      const proposedRiser = this._extractProposedRiserInfo(poleLocationData, designIndices.recommended);
      
      // Extract proposed guy information
      const proposedGuy = this._extractProposedGuyInfo(poleLocationData, designIndices.recommended);
      
      // Extract PLA value
      const plaInfo = this._extractPLA(poleLocationData);
      console.log(`DEBUG: Extracted PLA: ${plaInfo.pla} (${plaInfo.actual})`);
      
      // Extract construction grade
      const constructionGrade = this._extractConstructionGrade(poleLocationData);
      console.log(`DEBUG: Extracted construction grade: ${constructionGrade}`);
      
      // Extract midspan height data from Katapult instead of SPIDA
      const midspanHeights = this._extractExistingMidspanData(poleLocationData, katapultPoleData, canonicalPoleId);
      console.log(`DEBUG: Extracted midspan heights: com=${midspanHeights.com}, electrical=${midspanHeights.electrical}`);
      
      // Extract span data with attachments
      const spans = this._extractSpanData(poleLocationData, designIndices, canonicalPoleId);
      
      // Extract from/to pole information
      const fromToPoles = this._extractFromToPoles(poleLocationData, katapultPoleData);
      
      // Determine attachment action
      const attachmentAction = this._determineAttachmentAction(poleLocationData);
      
      // Create the pole data object
      const poleData: PoleData = {
        operationNumber: operationNumber,
        attachmentAction: attachmentAction,
        poleOwner: poleOwner,
        poleNumber: canonicalPoleId,
        poleStructure: poleStructure,
        proposedRiser: proposedRiser,
        proposedGuy: proposedGuy,
        pla: plaInfo.pla,
        constructionGrade: constructionGrade,
        heightLowestCom: midspanHeights.com,
        heightLowestCpsElectrical: midspanHeights.electrical,
        spans: spans,
        fromPole: fromToPoles.from,
        toPole: fromToPoles.to
      };
      
      return poleData;
    } catch (error) {
      console.error(`Error extracting pole data:`, error);
      return null;
    }
  }

  /**
   * PRIVATE: Find indices of measured and recommended designs
   */
  private _findDesignIndices(poleLocationData: any): { measured: number, recommended: number } | null {
    if (!poleLocationData || !poleLocationData.designs || !Array.isArray(poleLocationData.designs)) {
      return null;
    }
    
    let measuredIndex = -1;
    let recommendedIndex = -1;
    
    // Try to find designs by label and type
    for (let i = 0; i < poleLocationData.designs.length; i++) {
      const design = poleLocationData.designs[i];
      
      if (!design) continue;
      
      // Check if this is a measured design
      if (
        design.label?.toLowerCase().includes('measured') ||
        design.type?.toLowerCase().includes('measured') ||
        design.label?.toLowerCase().includes('existing')
      ) {
        measuredIndex = i;
      }
      
      // Check if this is a recommended design
      if (
        design.label?.toLowerCase().includes('recommended') ||
        design.type?.toLowerCase().includes('recommended') ||
        design.label?.toLowerCase().includes('proposed')
      ) {
        recommendedIndex = i;
      }
    }
    
    // If we couldn't find by label, try to find by creation date
    if (measuredIndex === -1 || recommendedIndex === -1) {
      let oldestDesignIndex = -1;
      let newestDesignIndex = -1;
      let oldestDate = new Date("9999-12-31");
      let newestDate = new Date("1970-01-01");
      
      for (let i = 0; i < poleLocationData.designs.length; i++) {
        const design = poleLocationData.designs[i];
        
        if (design && design.dateModified) {
          const date = new Date(design.dateModified);
          
          if (!isNaN(date.getTime())) {
            if (date < oldestDate) {
              oldestDate = date;
              oldestDesignIndex = i;
            }
            
            if (date > newestDate) {
              newestDate = date;
              newestDesignIndex = i;
            }
          }
        }
      }
      
      // Assume the oldest is measured, newest is recommended
      if (measuredIndex === -1 && oldestDesignIndex !== -1) {
        measuredIndex = oldestDesignIndex;
      }
      
      if (recommendedIndex === -1 && newestDesignIndex !== -1) {
        recommendedIndex = newestDesignIndex;
      }
    }
    
    // If we still don't have both, just use the available designs
    if (measuredIndex === -1 && recommendedIndex !== -1) {
      // If we only have recommended, use it for both
      measuredIndex = recommendedIndex;
    } else if (measuredIndex !== -1 && recommendedIndex === -1) {
      // If we only have measured, use it for both
      recommendedIndex = measuredIndex;
    } else if (measuredIndex === -1 && recommendedIndex === -1 && poleLocationData.designs.length > 0) {
      // If we couldn't identify either but have designs, use the first one
      measuredIndex = 0;
      recommendedIndex = poleLocationData.designs.length > 1 ? 1 : 0;
    }
    
    // Return null if we couldn't find the designs
    if (measuredIndex === -1 || recommendedIndex === -1) {
      return null;
    }
    
    return { measured: measuredIndex, recommended: recommendedIndex };
  }

  /**
   * Extract pole owner information
   */
  private _extractPoleOwner(poleLocationData: any, katapultPoleData: any): string {
    // Try to get owner from Katapult data first
    if (katapultPoleData) {
      const katapultOwner = 
        katapultPoleData.owner || 
        katapultPoleData.Owner || 
        katapultPoleData.properties?.owner || 
        katapultPoleData.properties?.Owner || 
        katapultPoleData.properties?.OWNER;
      
      if (katapultOwner) {
        return String(katapultOwner);
      }
    }
    
    // Fall back to SPIDA data
    try {
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) return "Unknown";
      
      const structure = poleLocationData?.designs?.[designIndices.measured]?.structure;
      
      if (structure) {
        // Try to find the pole class first, which often contains owner info
        if (structure.pole && structure.pole.owner) {
          return structure.pole.owner;
        }
      }
    } catch (error) {
      console.warn("Error extracting pole owner:", error);
    }
    
    return "Unknown";
  }

  /**
   * Extract pole structure information
   */
  private _extractPoleStructure(poleLocationData: any): string {
    try {
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) return "Unknown";
      
      const structure = poleLocationData?.designs?.[designIndices.measured]?.structure;
      
      if (structure && structure.pole) {
        const pole = structure.pole;
        
        // Combine material, class, and height if available
        const material = pole.clientItem?.species || pole.clientItem?.material || '';
        const poleClass = pole.clientItem?.classOfPole || '';
        const height = pole.clientItem?.height?.value || pole.height?.value || '';
        
        const components = [];
        if (material) components.push(material);
        if (poleClass) components.push(`Class ${poleClass}`);
        if (height) components.push(`${metersToFeet(height)}'`);
        
        return components.length > 0 ? components.join(' ') : "Unknown";
      }
    } catch (error) {
      console.warn("Error extracting pole structure:", error);
    }
    
    return "Unknown";
  }

  /**
   * Extract proposed riser information
   */
  private _extractProposedRiserInfo(poleLocationData: any, recommendedDesignIndex: number): string {
    try {
      const recommendedDesign = poleLocationData?.designs?.[recommendedDesignIndex]?.structure;
      
      if (recommendedDesign && recommendedDesign.risers && recommendedDesign.risers.length > 0) {
        const risers = recommendedDesign.risers.map((riser: any) => {
          const owner = riser.owner || "Unknown";
          const height = riser.height?.value ? metersToFeet(riser.height.value) : "Unknown";
          return `${owner} (${height}')`;
        });
        
        return risers.join(", ");
      }
    } catch (error) {
      console.warn("Error extracting riser info:", error);
    }
    
    return "N/A";
  }

  /**
   * Extract proposed guy information
   */
  private _extractProposedGuyInfo(poleLocationData: any, recommendedDesignIndex: number): string {
    try {
      const recommendedDesign = poleLocationData?.designs?.[recommendedDesignIndex]?.structure;
      
      if (recommendedDesign && recommendedDesign.guys && recommendedDesign.guys.length > 0) {
        const guys = recommendedDesign.guys.map((guy: any) => {
          const type = guy.clientItem?.size || "Unknown";
          const height = guy.attachmentHeight?.value ? metersToFeet(guy.attachmentHeight.value) : "Unknown";
          return `${type} (${height}')`;
        });
        
        return guys.join(", ");
      }
    } catch (error) {
      console.warn("Error extracting guy info:", error);
    }
    
    return "N/A";
  }

  /**
   * Extract PLA (Pole Loading Analysis) value
   */
  private _extractPLA(poleLocationData: any): { pla: string, actual: number } {
    try {
      // Try to extract PLA from analysis results
      const analyses = poleLocationData?.analyses;
      
      if (analyses && analyses.length > 0) {
        // Find the best analysis to use (latest recommended design)
        const designIndices = this._findDesignIndices(poleLocationData);
        if (!designIndices) return { pla: "N/A", actual: 0 };
        
        // Find analysis that corresponds to the recommended design
        for (const analysis of analyses) {
          if (analysis.designId === poleLocationData.designs[designIndices.recommended].id) {
            // Look for pole loading percentage in the results
            if (analysis.results && Array.isArray(analysis.results.components)) {
              for (const component of analysis.results.components) {
                if (component.pole && component.analysisCaseResults) {
                  // Find the highest loading percentage across all analysis cases
                  let maxLoading = 0;
                  for (const caseResult of component.analysisCaseResults) {
                    const loading = caseResult.actual.loadingRatio * 100;
                    if (loading > maxLoading) {
                      maxLoading = loading;
                    }
                  }
                  
                  if (maxLoading > 0) {
                    return { 
                      pla: `${maxLoading.toFixed(2)}%`,
                      actual: maxLoading 
                    };
                  }
                }
              }
            }
          }
        }
      }
    } catch (error) {
      console.warn("Error extracting PLA:", error);
    }
    
    return { pla: "N/A", actual: 0 };
  }

  /**
   * Extract construction grade
   */
  private _extractConstructionGrade(poleLocationData: any): string {
    try {
      // Try to find construction grade in analysis parameters
      const analyses = poleLocationData?.analyses;
      
      if (analyses && analyses.length > 0) {
        for (const analysis of analyses) {
          if (analysis.analysisCases && Array.isArray(analysis.analysisCases)) {
            for (const analysisCase of analysis.analysisCases) {
              // Look for construction grade in various properties
              const grade = 
                analysisCase.constructionGrade || 
                analysisCase.parameters?.constructionGrade || 
                analysisCase.parameters?.standard?.constructionGrade;
              
              if (grade) {
                return String(grade);
              }
            }
          }
        }
      }
    } catch (error) {
      console.warn("Error extracting construction grade:", error);
    }
    
    return "N/A";
  }

  /**
   * Extract existing midspan data
   */
  private _extractExistingMidspanData(poleLocationData: any, katapultPoleData: any, poleId: string): { com: string, electrical: string } {
    try {
      console.log(`DEBUG: Extracting existing midspan data for pole ${poleId}`);
      
      // First, try to get data from our previously parsed Katapult midspan map
      if (this.katapultMidspanMap.has(poleId)) {
        const midspanData = this.katapultMidspanMap.get(poleId)!;
        
        console.log(`DEBUG: Found Katapult midspan data for pole ${poleId}: com=${midspanData.com}, electrical=${midspanData.electrical}`);
        
        return {
          com: midspanData.com,
          electrical: midspanData.electrical
        };
      }
      
      console.log(`DEBUG: No Katapult midspan data found for pole ${poleId}, falling back to SPIDA data`);
      
      // Fall back to SPIDA data if Katapult data is not available
      let lowestComHeight = Number.MAX_VALUE;
      let lowestElectricalHeight = Number.MAX_VALUE;
      let foundCom = false;
      let foundElectrical = false;
      
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return { com: "N/A", electrical: "N/A" };
      }
      
      // Get attachments from Measured Design
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      if (!measuredDesign) {
        return { com: "N/A", electrical: "N/A" };
      }
      
      // Process wires
      if (measuredDesign.wires && Array.isArray(measuredDesign.wires)) {
        for (const wire of measuredDesign.wires) {
          const usageGroup = wire?.usageGroup || "";
          const attachmentHeight = wire?.attachmentHeight?.value;
          
          if (!attachmentHeight) continue;
          
          // Check usage group to determine if it's a com or electrical attachment
          if (usageGroup.includes("COMMUNICATION")) {
            foundCom = true;
            lowestComHeight = Math.min(lowestComHeight, attachmentHeight);
          } else if (
            usageGroup.includes("PRIMARY") ||
            usageGroup.includes("NEUTRAL") ||
            usageGroup.includes("SECONDARY")
          ) {
            foundElectrical = true;
            lowestElectricalHeight = Math.min(lowestElectricalHeight, attachmentHeight);
          }
        }
      }
      
      // Process equipment
      if (measuredDesign.equipments && Array.isArray(measuredDesign.equipments)) {
        for (const equipment of measuredDesign.equipments) {
          const usageGroup = equipment?.usageGroup || "";
          const attachmentHeight = equipment?.attachmentHeight?.value;
          
          if (!attachmentHeight) continue;
          
          // Check usage group
          if (usageGroup.includes("COMMUNICATION")) {
            foundCom = true;
            lowestComHeight = Math.min(lowestComHeight, attachmentHeight);
          } else if (
            usageGroup.includes("PRIMARY") ||
            usageGroup.includes("NEUTRAL") ||
            usageGroup.includes("SECONDARY")
          ) {
            foundElectrical = true;
            lowestElectricalHeight = Math.min(lowestElectricalHeight, attachmentHeight);
          }
        }
      }
      
      // Convert heights to feet-inches format
      const comHeight = foundCom ? 
        metersToFeetInches(lowestComHeight) : 
        "N/A";
      
      const electricalHeight = foundElectrical ? 
        metersToFeetInches(lowestElectricalHeight) : 
        "N/A";
      
      console.log(`DEBUG: Using SPIDA fallback data: com=${comHeight}, electrical=${electricalHeight}`);
      
      return {
        com: comHeight,
        electrical: electricalHeight
      };
    } catch (error) {
      console.warn("Error extracting midspan heights:", error);
      return { com: "N/A", electrical: "N/A" };
    }
  }

  /**
   * Extract span data with attachments
   */
  private _extractSpanData(poleLocationData: any, designIndices: { measured: number, recommended: number }, poleId: string): SpanData[] {
    try {
      const spans: SpanData[] = [];
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      
      if (!recommendedDesign || !recommendedDesign.wireEndPoints) {
        return spans;
      }
      
      // Create a map of wire IDs to actual wire objects for both designs
      const recommendedWireMap = this._createWireMap(recommendedDesign);
      const measuredWireMap = this._createWireMap(measuredDesign);
      
      // Process each wire end point (span)
      for (const wireEndPoint of recommendedDesign.wireEndPoints) {
        try {
          // Skip if no wires in this span
          if (!wireEndPoint.wires || wireEndPoint.wires.length === 0) {
            continue;
          }
          
          // Generate span header
          const spanHeader = this._generateSpanHeader(wireEndPoint);
          
          // Create new span data object
          const spanData: SpanData = {
            spanHeader: spanHeader,
            attachments: []
          };
          
          // Get the connected pole ID if available
          const connectedPoleId = wireEndPoint.structureLabel ? 
            canonicalizePoleID(wireEndPoint.structureLabel) : 
            null;
          
          // Process each wire in this span
          for (const wireId of wireEndPoint.wires) {
            // Get wire from recommended design
            const recommendedWire = recommendedWireMap.get(wireId);
            if (!recommendedWire) continue;
            
            // Try to find matching wire in measured design
            const measuredWire = this._findMatchingWire(recommendedWire, measuredWireMap);
            
            // Create attachment data
            const attachmentData: AttachmentData = {
              description: this._getAttachmentDescription(recommendedWire),
              existingHeight: measuredWire ? 
                metersToFeetInches(measuredWire.attachmentHeight?.value) : 
                "N/A",
              proposedHeight: metersToFeetInches(recommendedWire.attachmentHeight?.value),
              midSpanProposed: this._calculateMidSpanHeight(recommendedWire, poleId, connectedPoleId, spanData.spanHeader)
            };
            
            spanData.attachments.push(attachmentData);
          }
          
          // Add span data to spans array
          if (spanData.attachments.length > 0) {
            spans.push(spanData);
          }
        } catch (error) {
          console.warn("Error processing span:", error);
        }
      }
      
      return spans;
    } catch (error) {
      console.error("Error extracting span data:", error);
      return [];
    }
  }
  
  /**
   * Create a map of wire IDs to wire objects
   */
  private _createWireMap(design: any): Map<string, any> {
    const wireMap = new Map<string, any>();
    
    if (!design || !design.wires) {
      return wireMap;
    }
    
    for (const wire of design.wires) {
      if (wire && wire.id) {
        wireMap.set(wire.id, wire);
      }
    }
    
    return wireMap;
  }

  /**
   * Generate span header text
   */
  private _generateSpanHeader(wireEndPoint: any): string {
    if (!wireEndPoint) return "Unknown";
    
    // Use distance if available
    if (wireEndPoint.distance && wireEndPoint.distance.value) {
      return `${metersToFeet(wireEndPoint.distance.value)}'`;
    }
    
    // Otherwise use structure label (pole ID) if available
    if (wireEndPoint.structureLabel) {
      return wireEndPoint.structureLabel;
    }
    
    return "Unknown";
  }

  /**
   * Find matching wire in measured design
   */
  private _findMatchingWire(recommendedWire: any, measuredWireMap: Map<string, any>): any {
    if (!recommendedWire) return null;
    
    // First try to match by ID
    if (measuredWireMap.has(recommendedWire.id)) {
      return measuredWireMap.get(recommendedWire.id);
    }
    
    // If no exact ID match, try to find by position and type
    const attachmentHeight = recommendedWire.attachmentHeight?.value;
    const usage = recommendedWire.usageGroup;
    const size = recommendedWire.clientItem?.size;
    
    if (!attachmentHeight || !usage) return null;
    
    // Find closest wire with same usage
    let bestMatch = null;
    let minHeightDiff = Number.MAX_VALUE;
    
    for (const wire of measuredWireMap.values()) {
      if (wire.usageGroup !== usage) continue;
      
      // If size is specified, try to match that too
      if (size && wire.clientItem?.size !== size) continue;
      
      const height = wire.attachmentHeight?.value;
      if (!height) continue;
      
      const heightDiff = Math.abs(height - attachmentHeight);
      if (heightDiff < minHeightDiff) {
        minHeightDiff = heightDiff;
        bestMatch = wire;
      }
    }
    
    return bestMatch;
  }

  /**
   * Get attachment description
   */
  private _getAttachmentDescription(wire: any): string {
    if (!wire) return "Unknown";
    
    const components = [];
    
    // Add owner if available
    if (wire.owner) {
      components.push(wire.owner);
    }
    
    // Add wire properties if available
    if (wire.clientItem) {
      const size = wire.clientItem.size;
      if (size) {
        components.push(size);
      }
    }
    
    // Add usage group if available
    // const usageGroup = wire.usageGroup;
    // if (usageGroup) {
    //   components.push(usageGroup);
    // }
    
    // If we have additional specifics, add them
    if (wire.clientItem?.specificType) {
      components.push(`- ${wire.clientItem.specificType}`);
    }
    
    return components.length > 0 ? components.join(' ') : "Unknown";
  }

  /**
   * Calculate mid-span height
   */
  private _calculateMidSpanHeight(wire: any, poleId: string, connectedPoleId: string | null, spanHeader: string): string {
    try {
      console.log(`DEBUG: Calculating midspan height for pole ${poleId}, connected to ${connectedPoleId || "unknown"}`);
      
      // Get the attachment description for matching
      const description = this._getAttachmentDescription(wire);
      const descriptionKey = description.toLowerCase().replace(/\s+/g, '_');
      
      // Try to find midspan data in Katapult
      if (poleId && connectedPoleId && this.katapultMidspanMap.has(poleId)) {
        const poleData = this.katapultMidspanMap.get(poleId)!;
        
        // Check if we have data for this specific span
        if (poleData.spans.has(connectedPoleId)) {
          const spanMap = poleData.spans.get(connectedPoleId)!;
          
          // Try to match by description
          if (spanMap.has(descriptionKey)) {
            const midspanHeight = spanMap.get(descriptionKey);
            console.log(`DEBUG: Found Katapult midspan height for ${description}: ${midspanHeight}`);
            return midspanHeight;
          }
          
          // If we couldn't match by description, try other possibilities
          const wireId = wire.id;
          if (wireId && spanMap.has(`id_${wireId}`)) {
            const midspanHeight = spanMap.get(`id_${wireId}`);
            console.log(`DEBUG: Found Katapult midspan height by ID for ${description}: ${midspanHeight}`);
            return midspanHeight;
          }
          
          // Look for similar description keys
          for (const [key, value] of spanMap.entries()) {
            if (key.includes(descriptionKey) || descriptionKey.includes(key)) {
              console.log(`DEBUG: Found similar Katapult midspan height for ${description}: ${value}`);
              return value;
            }
          }
        }
        
        console.log(`DEBUG: No specific midspan data found for ${description} in span from ${poleId} to ${connectedPoleId}`);
      }
      
      console.log(`DEBUG: Falling back to calculated midspan for ${description}`);
      
      // Fall back to calculation from attachment height
      if (!wire.attachmentHeight?.value) {
        console.log(`DEBUG: No attachment height available for ${description}, using N/A`);
        return "N/A";
      }
      
      // Calculate based on attachment height
      const attachmentHeight = wire.attachmentHeight.value;
      const sagPercentage = 0.07; // 7% sag as an example
      const midSpanHeight = attachmentHeight * (1 - sagPercentage);
      
      const result = metersToFeetInches(midSpanHeight);
      console.log(`DEBUG: Calculated midspan height for ${description}: ${result}`);
      return result;
    } catch (error) {
      console.warn(`Error calculating midspan height:`, error);
      return "N/A";
    }
  }

  /**
   * Extract from/to pole information
   */
  private _extractFromToPoles(poleLocationData: any, katapultPoleData: any): { from: string, to: string } {
    try {
      let fromPole = poleLocationData.label || "N/A";
      let toPole = "N/A";
      
      // Try to get connected poles from Katapult data
      if (katapultPoleData) {
        const connections = 
          katapultPoleData.connections || 
          katapultPoleData.properties?.connections || 
          [];
        
        if (connections && connections.length > 0) {
          // Try to find the first connected pole
          for (const connection of connections) {
            const connectedPole = 
              connection.toPole || 
              connection.to_pole || 
              connection.targetPole || 
              connection.target_pole;
            
            if (connectedPole) {
              toPole = String(connectedPole);
              break;
            }
          }
        }
      }
      
      // If we still don't have a toPole, try to get from SPIDA
      if (toPole === "N/A") {
        // Try to find connected poles in SPIDA data
        const designIndices = this._findDesignIndices(poleLocationData);
        if (designIndices) {
          const structure = poleLocationData?.designs?.[designIndices.measured]?.structure;
          if (structure && structure.wireEndPoints && structure.wireEndPoints.length > 0) {
            for (const wireEndPoint of structure.wireEndPoints) {
              if (wireEndPoint.structureLabel) {
                toPole = wireEndPoint.structureLabel;
                break;
              }
            }
          }
        }
      }
      
      return {
        from: canonicalizePoleID(fromPole),
        to: toPole === "N/A" ? "N/A" : canonicalizePoleID(toPole)
      };
    } catch (error) {
      console.warn("Error extracting from/to poles:", error);
      return { from: "N/A", to: "N/A" };
    }
  }

  /**
   * Determine the attachment action
   */
  private _determineAttachmentAction(poleLocationData: any): string {
    try {
      // Default to "Attach" for now, but could be enhanced to determine actual action
      return "Attach";
    } catch (error) {
      console.warn("Error determining attachment action:", error);
      return "Attach"; // Default if we can't determine
    }
  }
}

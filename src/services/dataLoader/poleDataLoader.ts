
/**
 * Service for loading SPIDA and Katapult data
 */
import { ProcessingError } from "../types/poleTypes";
import { canonicalizePoleID } from "../utils/poleUtils";
import { KatapultMidspanData } from "../types/poleTypes";

export class PoleDataLoader {
  private spidaData: any;
  private katapultData: any;
  private errors: ProcessingError[] = [];
  private poleLookupMap: Map<string, any> = new Map();
  private katapultPoleLookupMap: Map<string, any> = new Map();
  private katapultMidspanMap: Map<string, KatapultMidspanData> = new Map();

  /**
   * Load and parse SPIDA JSON data
   */
  loadSpidaData(jsonText: string): boolean {
    try {
      this.spidaData = JSON.parse(jsonText);
      console.log("SPIDA data loaded successfully");
      return true;
    } catch (error) {
      console.error("Failed to parse SPIDA JSON:", error);
      this.errors.push({
        code: "PARSE_ERROR",
        message: "Failed to parse SPIDA JSON file",
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Load and parse Katapult JSON data
   */
  loadKatapultData(jsonText: string): boolean {
    try {
      this.katapultData = JSON.parse(jsonText);
      console.log("Katapult data loaded successfully");
      return true;
    } catch (error) {
      console.error("Failed to parse Katapult JSON:", error);
      this.errors.push({
        code: "PARSE_ERROR",
        message: "Failed to parse Katapult JSON file",
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Check if all required data has been loaded
   */
  isDataLoaded(): boolean {
    return !!this.spidaData && !!this.katapultData;
  }

  /**
   * Get errors from data loading
   */
  getErrors(): ProcessingError[] {
    return this.errors;
  }

  /**
   * Get SPIDA data
   */
  getSpidaData(): any {
    return this.spidaData;
  }

  /**
   * Get Katapult data
   */
  getKatapultData(): any {
    return this.katapultData;
  }

  /**
   * Create lookup maps for efficient pole matching
   */
  createLookupMaps(): { 
    poleLookupMap: Map<string, any>,
    katapultPoleLookupMap: Map<string, any>,
    katapultMidspanMap: Map<string, KatapultMidspanData>
  } {
    // Clear previous maps
    this.poleLookupMap.clear();
    this.katapultPoleLookupMap.clear();
    this.katapultMidspanMap.clear();
    
    // Create map for SPIDA poles
    this._createSpidaLookupMap();
    
    // Create map for Katapult poles
    this._createKatapultLookupMap();
    
    // Parse Katapult midspan data
    this._parseKatapultMidspanData();
    
    return {
      poleLookupMap: this.poleLookupMap,
      katapultPoleLookupMap: this.katapultPoleLookupMap,
      katapultMidspanMap: this.katapultMidspanMap
    };
  }

  /**
   * PRIVATE: Create lookup map for SPIDA poles
   */
  private _createSpidaLookupMap(): void {
    if (this.spidaData?.leads?.[0]?.locations) {
      for (const poleLocationData of this.spidaData.leads[0].locations) {
        try {
          const canonicalPoleId = canonicalizePoleID(poleLocationData.label);
          this.poleLookupMap.set(canonicalPoleId, poleLocationData);
        } catch (error) {
          console.warn(`Could not process pole ${poleLocationData?.label}:`, error);
        }
      }
      console.log(`DEBUG: Created lookup map for ${this.poleLookupMap.size} SPIDA poles`);
    }
  }

  /**
   * PRIVATE: Create lookup map for Katapult poles
   */
  private _createKatapultLookupMap(): void {
    try {
      // Log Katapult structure for debugging
      console.log("DEBUG: Katapult structure check at _createKatapultLookupMap");
      
      if (!this.katapultData) {
        console.warn("Katapult data is null or undefined");
        return;
      }
      
      // Log the first level of Katapult data structure
      console.log("DEBUG: Katapult top-level keys:", Object.keys(this.katapultData));
      
      // Try different possible structures for Katapult data
      let katapultNodes: any[] = [];
      
      // Try various possible data structures
      if (Array.isArray(this.katapultData.nodes)) {
        console.log(`DEBUG: Found nodes array in katapultData.nodes with ${this.katapultData.nodes.length} items`);
        katapultNodes = this.katapultData.nodes;
      } 
      else if (this.katapultData.data && Array.isArray(this.katapultData.data.nodes)) {
        console.log(`DEBUG: Found nodes array in katapultData.data.nodes with ${this.katapultData.data.nodes.length} items`);
        katapultNodes = this.katapultData.data.nodes;
      }
      else if (Array.isArray(this.katapultData.features)) {
        console.log(`DEBUG: Found features array in katapultData.features with ${this.katapultData.features.length} items`);
        katapultNodes = this.katapultData.features;
      }
      else if (Array.isArray(this.katapultData.poles)) {
        console.log(`DEBUG: Found poles array in katapultData.poles with ${this.katapultData.poles.length} items`);
        katapultNodes = this.katapultData.poles;
      }
      else if (Array.isArray(this.katapultData)) {
        console.log(`DEBUG: Katapult data itself is an array with ${this.katapultData.length} items`);
        katapultNodes = this.katapultData;
      }
      else {
        console.log("DEBUG: Attempting to find poles in katapultData keys:", Object.keys(this.katapultData));
        
        // Search for pole data in any property that might contain arrays
        for (const key of Object.keys(this.katapultData)) {
          const value = this.katapultData[key];
          if (Array.isArray(value) && value.length > 0) {
            console.log(`DEBUG: Found array in katapultData.${key} with ${value.length} items`);
            
            // Log the first item to understand its structure
            if (value[0]) {
              console.log(`DEBUG: Structure of first item in katapultData.${key}:`, 
                Object.keys(value[0]).join(', '));
              
              // If it has properties, log those too
              if (value[0].properties) {
                console.log(`DEBUG: Properties of first item:`, 
                  Object.keys(value[0].properties).join(', '));
              }
            }
            
            // Check if items have PoleNumber property
            const hasPoleNumberProperty = value.some(item => 
              item && 
              (item.properties?.PoleNumber || 
               item.properties?.poleNumber || 
               item.PoleNumber || 
               item.poleNumber || 
               item.properties?.polenumber || 
               item.polenumber || 
               item.properties?.pole_number || 
               item.pole_number)
            );
            
            if (hasPoleNumberProperty) {
              console.log(`DEBUG: Array in katapultData.${key} has items with pole number properties`);
              katapultNodes = value;
              break;
            }
          }
        }
      }
      
      console.log(`DEBUG: Found ${katapultNodes.length} potential pole nodes in Katapult data`);
      
      // Process the nodes
      for (const node of katapultNodes) {
        try {
          if (!node) continue;
          
          // Try various properties where pole number might be stored
          const poleNumber = 
            node.properties?.PoleNumber || 
            node.properties?.poleNumber || 
            node.PoleNumber || 
            node.poleNumber || 
            node.properties?.polenumber || 
            node.polenumber ||
            node.properties?.pole_number || 
            node.pole_number;
          
          if (poleNumber) {
            const canonicalPoleId = canonicalizePoleID(poleNumber);
            console.log(`DEBUG: Adding Katapult pole ${canonicalPoleId} to lookup map`);
            this.katapultPoleLookupMap.set(canonicalPoleId, node);
          }
        } catch (error) {
          console.warn(`Could not process Katapult node:`, error);
        }
      }
      
      console.log(`DEBUG: Created Katapult lookup map with ${this.katapultPoleLookupMap.size} poles`);
    } catch (error) {
      console.error("Error creating Katapult lookup map:", error);
    }
  }

  /**
   * PRIVATE: Parse Katapult midspan data
   */
  private _parseKatapultMidspanData(): void {
    try {
      console.log("DEBUG: Parsing Katapult midspan data");
      
      if (!this.katapultData) {
        console.warn("No Katapult data available for midspan extraction");
        return;
      }
      
      // Try to find spans or connections data in Katapult
      let spanData: any[] = [];
      
      // Check various possible locations for span data
      if (Array.isArray(this.katapultData.spans)) {
        console.log(`DEBUG: Found spans array in katapultData.spans with ${this.katapultData.spans.length} items`);
        spanData = this.katapultData.spans;
      } else if (Array.isArray(this.katapultData.connections)) {
        console.log(`DEBUG: Found connections array in katapultData.connections with ${this.katapultData.connections.length} items`);
        spanData = this.katapultData.connections;
      } else if (this.katapultData.data && Array.isArray(this.katapultData.data.spans)) {
        console.log(`DEBUG: Found spans array in katapultData.data.spans with ${this.katapultData.data.spans.length} items`);
        spanData = this.katapultData.data.spans;
      } else if (this.katapultData.data && Array.isArray(this.katapultData.data.connections)) {
        console.log(`DEBUG: Found connections array in katapultData.data.connections with ${this.katapultData.data.connections.length} items`);
        spanData = this.katapultData.data.connections;
      } else {
        console.log("DEBUG: Searching for span/connection data in Katapult structure");
        
        // Search recursively for spans/connections data
        for (const key of Object.keys(this.katapultData)) {
          const value = this.katapultData[key];
          
          // Skip non-objects
          if (!value || typeof value !== 'object') continue;
          
          // Check if this is an array of spans/connections
          if (Array.isArray(value)) {
            // Look for arrays that might contain span data
            if (value.length > 0) {
              const firstItem = value[0];
              
              // Check if items have properties that suggest span data
              const hasSpanProperties = firstItem && (
                firstItem.fromPole || 
                firstItem.toPole || 
                firstItem.from_pole || 
                firstItem.to_pole || 
                firstItem.properties?.fromPole || 
                firstItem.properties?.toPole ||
                firstItem.span_id ||
                firstItem.spanId ||
                firstItem.properties?.span_id ||
                firstItem.properties?.spanId
              );
              
              if (hasSpanProperties) {
                console.log(`DEBUG: Found potential span data in katapultData.${key} with ${value.length} items`);
                spanData = value;
                break;
              }
            }
          }
          
          // Check for nested objects
          if (Object.keys(value).length > 0) {
            // Check if this object has spans or connections arrays
            if (Array.isArray(value.spans) && value.spans.length > 0) {
              console.log(`DEBUG: Found spans array in katapultData.${key}.spans with ${value.spans.length} items`);
              spanData = value.spans;
              break;
            }
            
            if (Array.isArray(value.connections) && value.connections.length > 0) {
              console.log(`DEBUG: Found connections array in katapultData.${key}.connections with ${value.connections.length} items`);
              spanData = value.connections;
              break;
            }
          }
        }
      }
      
      if (spanData.length === 0) {
        console.warn("No span/connection data found in Katapult data");
        return;
      }
      
      console.log(`DEBUG: Processing ${spanData.length} span/connection records`);
      
      // Process span data to extract midspan information
      for (const span of spanData) {
        this._processSpanData(span);
      }
      
      console.log(`DEBUG: Created midspan data map for ${this.katapultMidspanMap.size} poles`);
      
      // Log a sample of the collected data
      if (this.katapultMidspanMap.size > 0) {
        const samplePole = Array.from(this.katapultMidspanMap.keys())[0];
        const sampleData = this.katapultMidspanMap.get(samplePole);
        console.log(`DEBUG: Sample midspan data for pole ${samplePole}:`, {
          com: sampleData?.com,
          electrical: sampleData?.electrical,
          spanCount: sampleData?.spans.size
        });
      }
    } catch (error) {
      console.error("Error parsing Katapult midspan data:", error);
    }
  }

  /**
   * PRIVATE: Process span data and extract midspan information
   */
  private _processSpanData(span: any): void {
    try {
      // Log span structure to understand its format
      if (!span) return;
      
      // Extract from and to pole IDs
      const fromPole = 
        span.fromPole || 
        span.from_pole || 
        span.properties?.fromPole || 
        span.properties?.from_pole || 
        span.from;
      
      const toPole = 
        span.toPole || 
        span.to_pole || 
        span.properties?.toPole || 
        span.properties?.to_pole || 
        span.to;
      
      if (!fromPole || !toPole) {
        return; // Skip span if we can't identify the poles
      }
      
      // Get the canonical pole IDs
      const fromPoleId = canonicalizePoleID(fromPole);
      const toPoleId = canonicalizePoleID(toPole);
      
      // Extract midspan height data
      const midspanData = this._extractMidspanFromKatapultSpan(span);
      
      // Store data for the from pole
      if (!this.katapultMidspanMap.has(fromPoleId)) {
        this.katapultMidspanMap.set(fromPoleId, {
          com: "N/A",
          electrical: "N/A",
          spans: new Map()
        });
      }
      
      const fromPoleData = this.katapultMidspanMap.get(fromPoleId)!;
      
      // Update lowest height values if needed
      if (midspanData.com !== "N/A" && (fromPoleData.com === "N/A" || midspanData.com < fromPoleData.com)) {
        fromPoleData.com = midspanData.com;
      }
      
      if (midspanData.electrical !== "N/A" && (fromPoleData.electrical === "N/A" || midspanData.electrical < fromPoleData.electrical)) {
        fromPoleData.electrical = midspanData.electrical;
      }
      
      // Store span-specific data
      if (!fromPoleData.spans.has(toPoleId)) {
        fromPoleData.spans.set(toPoleId, new Map());
      }
      
      // Store midspan data for this span
      this._storeAttachmentMidspanData(span, fromPoleData.spans.get(toPoleId)!);
      
      // Also store data for the to pole (in reverse direction)
      if (!this.katapultMidspanMap.has(toPoleId)) {
        this.katapultMidspanMap.set(toPoleId, {
          com: "N/A",
          electrical: "N/A",
          spans: new Map()
        });
      }
      
      const toPoleData = this.katapultMidspanMap.get(toPoleId)!;
      
      // Update lowest height values if needed
      if (midspanData.com !== "N/A" && (toPoleData.com === "N/A" || midspanData.com < toPoleData.com)) {
        toPoleData.com = midspanData.com;
      }
      
      if (midspanData.electrical !== "N/A" && (toPoleData.electrical === "N/A" || midspanData.electrical < toPoleData.electrical)) {
        toPoleData.electrical = midspanData.electrical;
      }
      
      // Store span-specific data
      if (!toPoleData.spans.has(fromPoleId)) {
        toPoleData.spans.set(fromPoleId, new Map());
      }
      
      // Store midspan data for this span in reverse direction
      this._storeAttachmentMidspanData(span, toPoleData.spans.get(fromPoleId)!, true);
    } catch (error) {
      console.warn("Error processing span data:", error);
    }
  }

  /**
   * PRIVATE: Extract midspan heights from a Katapult span record
   */
  private _extractMidspanFromKatapultSpan(span: any): { com: string, electrical: string } {
    try {
      let lowestComHeight: number | null = null;
      let lowestElectricalHeight: number | null = null;
      
      // There are many possible field names in different Katapult data formats
      // Try to find communication and electrical attachment heights
      
      // First, try to find direct midspan height fields
      const comMidspanHeight = 
        span.comMidspanHeight || 
        span.com_midspan_height || 
        span.properties?.comMidspanHeight || 
        span.properties?.com_midspan_height ||
        span.communications_midspan_height ||
        span.properties?.communications_midspan_height;
      
      const electricalMidspanHeight = 
        span.electricalMidspanHeight || 
        span.electrical_midspan_height || 
        span.properties?.electricalMidspanHeight || 
        span.properties?.electrical_midspan_height ||
        span.primary_midspan_height ||
        span.properties?.primary_midspan_height ||
        span.neutral_midspan_height ||
        span.properties?.neutral_midspan_height ||
        span.secondary_midspan_height ||
        span.properties?.secondary_midspan_height;
      
      // If we found direct height fields, use them
      if (comMidspanHeight !== undefined && comMidspanHeight !== null) {
        lowestComHeight = parseFloat(comMidspanHeight);
      }
      
      if (electricalMidspanHeight !== undefined && electricalMidspanHeight !== null) {
        lowestElectricalHeight = parseFloat(electricalMidspanHeight);
      }
      
      // If we couldn't find direct height fields, look for attachment arrays
      if ((lowestComHeight === null || lowestElectricalHeight === null) && 
          (span.attachments || span.properties?.attachments)) {
        
        const attachments = span.attachments || span.properties?.attachments;
        
        if (Array.isArray(attachments)) {
          for (const attachment of attachments) {
            try {
              // Skip if null or undefined
              if (!attachment) continue;
              
              // Try to get the attachment type
              const type = 
                attachment.type || 
                attachment.attachmentType || 
                attachment.attachment_type || 
                attachment.properties?.type || 
                attachment.properties?.attachmentType || 
                attachment.properties?.attachment_type;
              
              // Try to get the midspan height
              const midspanHeight = 
                attachment.midspanHeight || 
                attachment.midspan_height || 
                attachment.properties?.midspanHeight || 
                attachment.properties?.midspan_height || 
                attachment.sagHeight ||
                attachment.sag_height ||
                attachment.properties?.sagHeight ||
                attachment.properties?.sag_height;
              
              // Skip if we couldn't find height
              if (midspanHeight === undefined || midspanHeight === null) continue;
              
              const height = parseFloat(midspanHeight);
              
              // Check if this is a communication or electrical attachment
              const isCommType = type && /com|communication|cable|fiber|telephone/i.test(type);
              const isElectricalType = type && /primary|neutral|secondary|electrical|power/i.test(type);
              
              if (isCommType) {
                if (lowestComHeight === null || height < lowestComHeight) {
                  lowestComHeight = height;
                }
              } else if (isElectricalType) {
                if (lowestElectricalHeight === null || height < lowestElectricalHeight) {
                  lowestElectricalHeight = height;
                }
              }
            } catch (error) {
              console.warn("Error processing attachment:", error);
            }
          }
        }
      }
      
      // Convert heights to feet-inches format and return
      return {
        com: lowestComHeight !== null ? formatHeightValue(lowestComHeight) : "N/A",
        electrical: lowestElectricalHeight !== null ? formatHeightValue(lowestElectricalHeight) : "N/A"
      };
      
    } catch (error) {
      console.warn("Error extracting midspan heights from span:", error);
      return { com: "N/A", electrical: "N/A" };
    }
  }

  /**
   * PRIVATE: Store attachment-specific midspan data
   */
  private _storeAttachmentMidspanData(span: any, spanMap: Map<string, string>, isReverse: boolean = false): void {
    try {
      // Try to find attachment data
      const attachments = span.attachments || span.properties?.attachments;
      
      if (!Array.isArray(attachments)) return;
      
      for (const attachment of attachments) {
        try {
          // Skip if null or undefined
          if (!attachment) continue;
          
          // Get attachment identifier
          const attachmentId = 
            attachment.id || 
            attachment.attachmentId || 
            attachment.attachment_id || 
            attachment.properties?.id || 
            attachment.properties?.attachmentId || 
            attachment.properties?.attachment_id;
          
          // Get attachment description
          const description = 
            attachment.description || 
            attachment.name || 
            attachment.properties?.description || 
            attachment.properties?.name ||
            attachment.type ||
            attachment.properties?.type;
          
          // Get midspan height
          const midspanHeight = 
            attachment.midspanHeight || 
            attachment.midspan_height || 
            attachment.properties?.midspanHeight || 
            attachment.properties?.midspan_height ||
            attachment.sagHeight ||
            attachment.properties?.sagHeight ||
            attachment.sag_height ||
            attachment.properties?.sag_height;
          
          if (!description || midspanHeight === undefined || midspanHeight === null) continue;
          
          // Create a key that will allow matching with SPIDA data
          let key = description.toLowerCase().replace(/\s+/g, '_');
          
          // Add direction info if available
          const direction = isReverse ? "reverse" : "forward";
          
          // Store the height
          const formattedHeight = formatHeightValue(parseFloat(midspanHeight));
          spanMap.set(key, formattedHeight);
          
          // Also try with attachment ID if available
          if (attachmentId) {
            spanMap.set(`id_${attachmentId}`, formattedHeight);
          }
          
        } catch (error) {
          console.warn("Error processing attachment for midspan storage:", error);
        }
      }
    } catch (error) {
      console.warn("Error storing attachment midspan data:", error);
    }
  }
}

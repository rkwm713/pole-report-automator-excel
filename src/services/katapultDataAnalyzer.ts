
import * as XLSX from 'xlsx';
import { WireCategory, formatHeightToString, parseHeightToInches } from '@/utils/katapultDataProcessor';

// Define interfaces for our data structures
interface WireData {
  traceId: string;
  company: string;
  cableType: string;
  category: WireCategory;
  midspanObservations: Array<{
    originalHeightInches: number;
    proposedHeightInches: number | null;
  }>;
  lowestExistingMidspanHeight: number | null;
  finalProposedMidspanHeight: number | null;
  lowestExistingPoleAttachmentHeight: number | null;
  finalProposedPoleAttachmentHeight: number | null;
}

interface ConnectionData {
  connectionId: string;
  fromPoleId: string;
  toPoleId: string | null;
  buttonType: string;
  isRefConnection: boolean;
  wires: Record<string, WireData>;
}

interface ProcessedData {
  connections: ConnectionData[];
}

// New interfaces for the pole report data
interface PoleWireData {
  attacher: string;
  company: string;
  cableType: string;
  category: WireCategory;
  existingHeight: number | null;
  proposedHeight: number | null;
  proposedMidspanHeight: number | null;
  isRiser: boolean;
  isGuy: boolean;
  isRef: boolean;
  refTargetPole?: string;
}

interface PoleReportData {
  poleId: string;
  poleOwner: string;
  poleStructure: string;
  proposedRiser: boolean;
  proposedGuy: boolean;
  pla: string;
  constructionGrade: string;
  lowestComMidspan: number | null;
  lowestCpsMidspan: number | null;
  wires: PoleWireData[];
  connections: {
    toPoleId: string;
    fromPoleId: string;
  }[];
}

/**
 * Generates and downloads an Excel file with the analysis of Katapult data in the requested format
 * @param processedData - The processed Katapult data
 * @param rawKatapultData - The raw Katapult JSON data
 */
export function downloadKatapultAnalysisExcel(processedData: ProcessedData, rawKatapultData: any) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();
  
  // Create the main sheet with the make ready report format
  addMakeReadyReportSheet(workbook, processedData, rawKatapultData);
  
  // Generate and download Excel file
  XLSX.writeFile(workbook, 'katapult-make-ready-report.xlsx');
}

/**
 * Extract pole structure info from raw data
 */
function extractPoleStructure(rawData: any, poleId: string): string {
  // Default value if we can't find the structure
  const defaultStructure = "40-2 Southern Pine";
  
  if (!rawData || !rawData.nodes || !rawData.nodes[poleId]) {
    return defaultStructure;
  }
  
  const poleData = rawData.nodes[poleId];
  
  // Try to extract height class and material
  let height = "40";
  let classNum = "2";
  let material = "Southern Pine";
  
  // Extract height from pole data if available
  if (poleData.height) {
    height = String(Math.round(parseFloat(String(poleData.height))));
  }
  
  // Extract class from pole data if available
  if (poleData.class) {
    classNum = String(poleData.class);
  }
  
  // Extract material from pole data if available
  if (poleData.material) {
    material = String(poleData.material);
  }
  
  return `${height}-${classNum} ${material}`;
}

/**
 * Determine if pole has risers based on wires
 */
function hasPoleRisers(wires: Record<string, WireData>): boolean {
  for (const wireId in wires) {
    const wire = wires[wireId];
    // Check cable type for riser indicators
    if (wire.cableType.toLowerCase().includes('riser')) {
      return true;
    }
  }
  return false;
}

/**
 * Determine if pole has guys based on wires
 */
function hasPoleGuys(wires: Record<string, WireData>): boolean {
  for (const wireId in wires) {
    const wire = wires[wireId];
    // Check cable type for guy indicators
    if (wire.cableType.toLowerCase().includes('guy')) {
      return true;
    }
  }
  return false;
}

/**
 * Extract lowest midspan heights for COM and CPS categories
 */
function extractLowestMidspanHeights(wires: Record<string, WireData>): { comHeight: number | null, cpsHeight: number | null } {
  let lowestComHeight: number | null = null;
  let lowestCpsHeight: number | null = null;
  
  for (const wireId in wires) {
    const wire = wires[wireId];
    
    if (wire.lowestExistingMidspanHeight !== null) {
      if (wire.category === WireCategory.COMMUNICATION) {
        if (lowestComHeight === null || wire.lowestExistingMidspanHeight < lowestComHeight) {
          lowestComHeight = wire.lowestExistingMidspanHeight;
        }
      } else if (wire.category === WireCategory.CPS_ELECTRICAL) {
        if (lowestCpsHeight === null || wire.lowestExistingMidspanHeight < lowestCpsHeight) {
          lowestCpsHeight = wire.lowestExistingMidspanHeight;
        }
      }
    }
  }
  
  return { comHeight: lowestComHeight, cpsHeight: lowestCpsHeight };
}

/**
 * Get a formatted wire description from company and cable type
 */
function getWireDescription(company: string, cableType: string): string {
  // Combine company and cable type with proper formatting
  let description = '';
  
  // Special cases for specific formatting
  if (cableType.toLowerCase().includes('neutral')) {
    return 'Neutral';
  }
  
  if (company && cableType) {
    // For Charter/Spectrum, use format "Charter/Spectrum Fiber Optic"
    if (company.toLowerCase().includes('charter') || company.toLowerCase().includes('spectrum')) {
      description = 'Charter/Spectrum';
      if (cableType.toLowerCase().includes('fiber')) {
        description += ' Fiber Optic';
      } else if (cableType.toLowerCase().includes('guy')) {
        description += ' Down Guy';
      } else if (cableType.toLowerCase().includes('riser')) {
        description += ' Communication Riser';
      } else {
        description += ' ' + cableType;
      }
    } 
    // For AT&T, use format "AT&T Fiber Optic Com"
    else if (company.toLowerCase().includes('at&t')) {
      description = 'AT&T';
      if (cableType.toLowerCase().includes('fiber')) {
        description += ' Fiber Optic Com';
      } else if (cableType.toLowerCase().includes('telco')) {
        description += ' Telco Com';
      } else if (cableType.toLowerCase().includes('drop')) {
        description += ' Com Drop';
      } else {
        description += ' ' + cableType;
      }
    }
    // For CPS, use format "CPS Supply Fiber"
    else if (company.toLowerCase().includes('cps')) {
      description = 'CPS';
      if (cableType.toLowerCase().includes('supply fiber')) {
        description += ' Supply Fiber';
      } else if (cableType.toLowerCase().includes('street light')) {
        if (cableType.toLowerCase().includes('top bracket')) {
          description += ' Top Bracket Street Light';
        } else if (cableType.toLowerCase().includes('bottom bracket')) {
          description += ' Bottom Bracket Street Light';
        } else if (cableType.toLowerCase().includes('feed')) {
          description += ' Street Light Feed';
        } else if (cableType.toLowerCase().includes('drip')) {
          description += ' Drip Loop Street Light';
        } else {
          description += ' Street Light';
        }
      } else if (cableType.toLowerCase().includes('secondary')) {
        if (cableType.toLowerCase().includes('drip')) {
          description += ' Drip Loop Secondary';
        } else {
          description += ' Secondary';
        }
      } else if (cableType.toLowerCase().includes('service')) {
        description += ' Service';
      } else if (cableType.toLowerCase().includes('guy')) {
        description += ' Energy Down guy';
      } else {
        description += ' ' + cableType;
      }
    } 
    else {
      description = company + ' ' + cableType;
    }
  } else if (company) {
    description = company;
  } else if (cableType) {
    description = cableType;
  } else {
    description = 'Unknown';
  }
  
  return description;
}

/**
 * Organize data by poles for the report format
 */
function organizePoleReportData(processedData: ProcessedData, rawData: any): PoleReportData[] {
  // Map to store pole data by pole ID
  const poleMap = new Map<string, PoleReportData>();
  let operationCounter = 0;
  
  // Process each connection
  for (const connection of processedData.connections) {
    // Process fromPole
    if (!poleMap.has(connection.fromPoleId)) {
      operationCounter++;
      
      // Extract information for this pole
      const poleStructure = extractPoleStructure(rawData, connection.fromPoleId);
      const hasRisers = hasPoleRisers(connection.wires);
      const hasGuys = hasPoleGuys(connection.wires);
      const { comHeight, cpsHeight } = extractLowestMidspanHeights(connection.wires);
      
      // Create new pole report data
      poleMap.set(connection.fromPoleId, {
        poleId: connection.fromPoleId,
        poleOwner: "CPS", // Default to CPS as pole owner
        poleStructure,
        proposedRiser: hasRisers,
        proposedGuy: hasGuys,
        pla: connection.isRefConnection ? "N/A" : "78.70%", // Example PLA value
        constructionGrade: "C", // Default construction grade
        lowestComMidspan: comHeight,
        lowestCpsMidspan: cpsHeight,
        wires: [],
        connections: []
      });
    }
    
    // Get the pole data
    const poleData = poleMap.get(connection.fromPoleId)!;
    
    // Add connection information
    if (connection.toPoleId) {
      poleData.connections.push({
        fromPoleId: connection.fromPoleId,
        toPoleId: connection.toPoleId
      });
    }
    
    // Process wires for this connection
    for (const wireId in connection.wires) {
      const wire = connection.wires[wireId];
      const attacher = getWireDescription(wire.company, wire.cableType);
      
      // Check if wire is a riser or guy
      const isRiser = wire.cableType.toLowerCase().includes('riser');
      const isGuy = wire.cableType.toLowerCase().includes('guy');
      
      // Add wire to pole data
      poleData.wires.push({
        attacher,
        company: wire.company,
        cableType: wire.cableType,
        category: wire.category,
        existingHeight: connection.isRefConnection ? 
                       wire.lowestExistingPoleAttachmentHeight : 
                       wire.lowestExistingMidspanHeight,
        proposedHeight: connection.isRefConnection ? 
                       wire.finalProposedPoleAttachmentHeight : 
                       null,
        proposedMidspanHeight: !connection.isRefConnection ? 
                              wire.finalProposedMidspanHeight : 
                              null,
        isRiser,
        isGuy,
        isRef: connection.isRefConnection,
        refTargetPole: connection.isRefConnection && connection.toPoleId ? 
                      connection.toPoleId : undefined
      });
    }
  }
  
  return Array.from(poleMap.values());
}

/**
 * Add the main Make Ready Report sheet with the requested format
 */
function addMakeReadyReportSheet(workbook: XLSX.WorkBook, processedData: ProcessedData, rawData: any) {
  // Organize data by poles for the report
  const poleReportData = organizePoleReportData(processedData, rawData);
  
  // Create headers for the report
  const headers = [
    'Operation Number',
    'Attachment Action:\n( I )nstalling\n( R )emoving\n( E )xisting',
    'Pole Owner',
    'Pole #',
    'Pole Structure',
    'Proposed Riser (Yes/No) &',
    'Proposed Guy (Yes/No) &',
    'PLA (%) with proposed attachment',
    'Construction Grade of Analysis',
    'Existing Mid-Span Data',
    '',
    'Make Ready Data',
    '',
    '',
    '',
  ];
  
  const subHeaders = [
    '', '', '', '', '', '', '', '', '',
    'Height Lowest Com',
    'Height Lowest CPS Electrical',
    'Attacher Description',
    'Existing',
    'Proposed',
    'Mid-Span\n(same span as existing)'
  ];
  
  // Create the data array for the report
  const reportData: any[][] = [headers, subHeaders];
  let operationCounter = 1;
  
  // Process each pole
  for (const poleData of poleReportData) {
    // Add pole row
    const poleRow = [
      operationCounter,
      '( I )nstalling', // Default to installing
      poleData.poleOwner,
      poleData.poleId,
      poleData.poleStructure,
      poleData.proposedRiser ? 'YES (1)' : 'NO',
      poleData.proposedGuy ? 'YES (1)' : 'NO',
      poleData.pla,
      poleData.constructionGrade,
      formatHeightToString(poleData.lowestComMidspan),
      formatHeightToString(poleData.lowestCpsMidspan),
      '', '', '', ''
    ];
    reportData.push(poleRow);
    operationCounter++;
    
    // Process wires for this pole
    for (const wire of poleData.wires) {
      // Skip wires that are handled in REF sections
      if (wire.isRef) {
        continue;
      }
      
      const wireRow = [
        '', '', '', '', '', '', '', '', '',
        '', '',
        wire.attacher,
        formatHeightToString(wire.existingHeight),
        wire.proposedHeight ? formatHeightToString(wire.proposedHeight) : '',
        wire.proposedMidspanHeight ? formatHeightToString(wire.proposedMidspanHeight) : ''
      ];
      reportData.push(wireRow);
    }
    
    // Add connection rows
    for (const connection of poleData.connections) {
      // Add From Pole To Pole row
      const connectionHeader = [
        '', '', '', '', '', '', '', '', '',
        'From Pole', 'To Pole',
        '', '', '', ''
      ];
      reportData.push(connectionHeader);
      
      // Add connection details
      const connectionDetails = [
        '', '', '', '', '', '', '', '', '',
        connection.fromPoleId, connection.toPoleId,
        '', '', '', ''
      ];
      reportData.push(connectionDetails);
    }
    
    // Add REF connections if any
    const refWires = poleData.wires.filter(w => w.isRef);
    if (refWires.length > 0) {
      // Group ref wires by target pole
      const refWiresByTarget = new Map<string, PoleWireData[]>();
      
      for (const refWire of refWires) {
        const targetKey = refWire.refTargetPole || 'unknown';
        if (!refWiresByTarget.has(targetKey)) {
          refWiresByTarget.set(targetKey, []);
        }
        refWiresByTarget.get(targetKey)!.push(refWire);
      }
      
      // Add each group of ref wires
      for (const [targetPole, wires] of refWiresByTarget.entries()) {
        // Add REF header
        const refHeader = [
          '', '', '', '', '', '', '', '', '',
          '', '',
          `Ref (North East) to ${targetPole === 'unknown' ? 'service pole' : targetPole}`,
          '', '', ''
        ];
        reportData.push(refHeader);
        
        // Add each REF wire
        for (const wire of wires) {
          const wireRow = [
            '', '', '', '', '', '', '', '', '',
            '', '',
            wire.attacher,
            formatHeightToString(wire.existingHeight),
            '', // No proposed for REF
            formatHeightToString(wire.proposedMidspanHeight) // Use midspan for REF
          ];
          reportData.push(wireRow);
        }
      }
    }
    
    // Add empty row between poles
    reportData.push(Array(15).fill(''));
  }
  
  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(reportData);
  
  // Set column widths
  const colWidths = [
    { wch: 10 }, // Operation Number
    { wch: 15 }, // Attachment Action
    { wch: 10 }, // Pole Owner
    { wch: 10 }, // Pole #
    { wch: 15 }, // Pole Structure
    { wch: 12 }, // Proposed Riser
    { wch: 12 }, // Proposed Guy
    { wch: 12 }, // PLA
    { wch: 12 }, // Construction Grade
    { wch: 12 }, // Height Lowest Com
    { wch: 12 }, // Height Lowest CPS
    { wch: 25 }, // Attacher Description
    { wch: 12 }, // Existing
    { wch: 12 }, // Proposed
    { wch: 12 }, // Mid-Span
  ];
  worksheet['!cols'] = colWidths;
  
  // Merge header cells
  const merges = [
    { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } }, // Existing Mid-Span Data
    { s: { r: 0, c: 11 }, e: { r: 0, c: 14 } }, // Make Ready Data
  ];
  worksheet['!merges'] = merges;
  
  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Make Ready Report');
}

/**
 * Add a summary sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addSummarySheet(workbook: XLSX.WorkBook, data: ProcessedData) {
  const connections = data.connections || [];
  const totalConnections = connections.length;
  const refConnections = connections.filter(c => c.isRefConnection).length;
  const regularConnections = totalConnections - refConnections;
  
  // Count wires by category
  const wireCounts = {
    [WireCategory.COMMUNICATION]: 0,
    [WireCategory.CPS_ELECTRICAL]: 0,
    [WireCategory.OTHER]: 0,
    total: 0
  };
  
  // Count proposed changes
  let proposedChanges = 0;
  
  connections.forEach(connection => {
    Object.values(connection.wires).forEach((wire: WireData) => {
      if (wire.category) {
        wireCounts[wire.category]++;
        wireCounts.total++;
      }
      
      if (wire.finalProposedMidspanHeight !== null || wire.finalProposedPoleAttachmentHeight !== null) {
        proposedChanges++;
      }
    });
  });
  
  // Create summary data
  const summaryData = [
    ['Katapult Analysis Summary'],
    [],
    ['Total Connections', totalConnections],
    ['Regular Connections', regularConnections],
    ['REF Connections', refConnections],
    [],
    ['Total Wires', wireCounts.total],
    ['Communication Wires', wireCounts[WireCategory.COMMUNICATION]],
    ['CPS Electrical Wires', wireCounts[WireCategory.CPS_ELECTRICAL]],
    ['Other Wires', wireCounts[WireCategory.OTHER]],
    [],
    ['Wires with Proposed Changes', proposedChanges],
  ];
  
  // Add summary sheet
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
}

/**
 * Add a connections sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addConnectionsSheet(workbook: XLSX.WorkBook, data: ProcessedData) {
  const connections = data.connections || [];
  
  // Create connections header
  const connectionsHeader = [
    'Connection ID',
    'From Pole ID',
    'To Pole ID',
    'Connection Type',
    'Is REF',
    'Total Wires'
  ];
  
  // Create connections data
  const connectionsData = connections.map(connection => [
    connection.connectionId,
    connection.fromPoleId,
    connection.toPoleId || 'N/A',
    connection.buttonType,
    connection.isRefConnection ? 'Yes' : 'No',
    Object.keys(connection.wires).length
  ]);
  
  // Add connections sheet
  const connectionsSheet = XLSX.utils.aoa_to_sheet([
    connectionsHeader,
    ...connectionsData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, connectionsSheet, 'Connections');
}

/**
 * Add a wire data sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addWireDataSheet(workbook: XLSX.WorkBook, data: ProcessedData) {
  const connections = data.connections || [];
  
  // Create wire data header
  const wireDataHeader = [
    'Connection ID',
    'From Pole ID',
    'To Pole ID',
    'Trace ID',
    'Company',
    'Cable Type',
    'Category',
    'Lowest Midspan Height',
    'Proposed Midspan Height',
    'Has Proposed Change'
  ];
  
  // Create wire data
  const wireData = connections.flatMap(connection => 
    Object.values(connection.wires).map((wire: WireData) => {
      const hasProposedChange = wire.finalProposedMidspanHeight !== null && 
                               wire.lowestExistingMidspanHeight !== null && 
                               wire.finalProposedMidspanHeight !== wire.lowestExistingMidspanHeight;
      
      return [
        connection.connectionId,
        connection.fromPoleId,
        connection.toPoleId || 'N/A',
        wire.traceId,
        wire.company,
        wire.cableType,
        wire.category,
        formatHeightToString(wire.lowestExistingMidspanHeight),
        formatHeightToString(wire.finalProposedMidspanHeight),
        hasProposedChange ? 'Yes' : 'No'
      ];
    })
  );
  
  // Add wire data sheet
  const wireDataSheet = XLSX.utils.aoa_to_sheet([
    wireDataHeader,
    ...wireData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, wireDataSheet, 'Wire Data');
}

/**
 * Add a REF connections sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addRefConnectionsSheet(workbook: XLSX.WorkBook, data: ProcessedData) {
  const refConnections = data.connections.filter(c => c.isRefConnection) || [];
  
  if (refConnections.length === 0) {
    // If no REF connections, add empty sheet
    const emptySheet = XLSX.utils.aoa_to_sheet([
      ['No REF connections found in the data']
    ]);
    XLSX.utils.book_append_sheet(workbook, emptySheet, 'REF Connections');
    return;
  }
  
  // Create REF connections header
  const refConnectionsHeader = [
    'Connection ID',
    'Pole ID',
    'Trace ID',
    'Company',
    'Cable Type',
    'Category',
    'Attachment Height',
    'Proposed Attachment Height',
    'Has Proposed Change'
  ];
  
  // Create REF connections data
  const refConnectionsData = refConnections.flatMap(connection => 
    Object.values(connection.wires).map((wire: WireData) => {
      const hasProposedChange = wire.finalProposedPoleAttachmentHeight !== null && 
                               wire.lowestExistingPoleAttachmentHeight !== null && 
                               wire.finalProposedPoleAttachmentHeight !== wire.lowestExistingPoleAttachmentHeight;
      
      return [
        connection.connectionId,
        connection.fromPoleId,
        wire.traceId,
        wire.company,
        wire.cableType,
        wire.category,
        formatHeightToString(wire.lowestExistingPoleAttachmentHeight),
        formatHeightToString(wire.finalProposedPoleAttachmentHeight),
        hasProposedChange ? 'Yes' : 'No'
      ];
    })
  );
  
  // Add REF connections sheet
  const refConnectionsSheet = XLSX.utils.aoa_to_sheet([
    refConnectionsHeader,
    ...refConnectionsData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, refConnectionsSheet, 'REF Connections');
}

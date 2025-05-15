
import React, { useState } from 'react';
import { 
  Table, 
  TableBody, 
  TableCaption, 
  TableCell, 
  TableHead, 
  TableHeader, 
  TableRow 
} from '@/components/ui/table';
import { 
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from '@/components/ui/accordion';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { WireCategory, formatHeightToString } from '@/utils/katapultDataProcessor';

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

interface KatapultDataVisualizationProps {
  data: ProcessedData;
}

const KatapultDataVisualization: React.FC<KatapultDataVisualizationProps> = ({ data }) => {
  const [activeTab, setActiveTab] = useState<string>('summary');
  
  const connections = data.connections || [];
  const totalConnections = connections.length;
  const refConnections = connections.filter(c => c.isRefConnection).length;
  const regularConnections = totalConnections - refConnections;
  
  // Calculate wire category statistics
  const wireCounts = {
    [WireCategory.COMMUNICATION]: 0,
    [WireCategory.CPS_ELECTRICAL]: 0,
    [WireCategory.OTHER]: 0
  };
  
  // Count wires by category
  connections.forEach(connection => {
    Object.values(connection.wires).forEach((wire: WireData) => {
      if (wire.category) {
        wireCounts[wire.category]++;
      }
    });
  });
  
  return (
    <div className="space-y-8">
      <h2 className="text-2xl font-bold">Katapult Analysis Results</h2>
      
      <Tabs defaultValue="summary" onValueChange={setActiveTab} className="w-full">
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="summary">Summary</TabsTrigger>
          <TabsTrigger value="connections">Connections</TabsTrigger>
          <TabsTrigger value="wires">Wire Data</TabsTrigger>
        </TabsList>
        
        <TabsContent value="summary" className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <Card>
              <CardHeader className="pb-2">
                <CardTitle className="text-sm font-medium">Total Connections</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-2xl font-bold">{totalConnections}</p>
                <p className="text-xs text-muted-foreground">
                  {regularConnections} Regular, {refConnections} REF
                </p>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader className="pb-2">
                <CardTitle className="text-sm font-medium">Total Wires</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-2xl font-bold">
                  {Object.values(wireCounts).reduce((sum, count) => sum + count, 0)}
                </p>
                <div className="flex gap-2 mt-1">
                  <Badge variant="outline" className="bg-blue-50">
                    {wireCounts[WireCategory.COMMUNICATION]} Comm
                  </Badge>
                  <Badge variant="outline" className="bg-orange-50">
                    {wireCounts[WireCategory.CPS_ELECTRICAL]} Electrical
                  </Badge>
                  <Badge variant="outline" className="bg-gray-50">
                    {wireCounts[WireCategory.OTHER]} Other
                  </Badge>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader className="pb-2">
                <CardTitle className="text-sm font-medium">Proposed Changes</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-2xl font-bold">
                  {connections.reduce((count, connection) => {
                    return count + Object.values(connection.wires).filter((wire: WireData) => 
                      wire.finalProposedMidspanHeight !== null || 
                      wire.finalProposedPoleAttachmentHeight !== null
                    ).length;
                  }, 0)}
                </p>
                <p className="text-xs text-muted-foreground">Wires with proposed height changes</p>
              </CardContent>
            </Card>
          </div>
        </TabsContent>
        
        <TabsContent value="connections" className="space-y-4">
          <Accordion type="single" collapsible className="w-full">
            {connections.map((connection, index) => (
              <AccordionItem key={connection.connectionId} value={`connection-${index}`}>
                <AccordionTrigger>
                  <div className="flex items-center gap-2">
                    <span>
                      {connection.isRefConnection ? 'REF Connection' : 'Connection'}: {connection.fromPoleId} 
                      {!connection.isRefConnection && <span> → {connection.toPoleId}</span>}
                    </span>
                    {connection.isRefConnection && (
                      <Badge variant="outline" className="ml-2 bg-purple-50">REF</Badge>
                    )}
                  </div>
                </AccordionTrigger>
                <AccordionContent>
                  <div className="rounded-md border p-4">
                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <div>
                        <h4 className="text-sm font-semibold">From Pole</h4>
                        <p>{connection.fromPoleId}</p>
                      </div>
                      <div>
                        <h4 className="text-sm font-semibold">To Pole</h4>
                        <p>{connection.toPoleId || 'N/A'}</p>
                      </div>
                      <div>
                        <h4 className="text-sm font-semibold">Connection Type</h4>
                        <p>{connection.buttonType}</p>
                      </div>
                      <div>
                        <h4 className="text-sm font-semibold">Total Wires</h4>
                        <p>{Object.keys(connection.wires).length}</p>
                      </div>
                    </div>
                    
                    <Table>
                      <TableCaption>Wire data for this connection</TableCaption>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Company</TableHead>
                          <TableHead>Cable Type</TableHead>
                          <TableHead>Category</TableHead>
                          <TableHead>Lowest Midspan Height</TableHead>
                          <TableHead>Proposed Midspan Height</TableHead>
                          {connection.isRefConnection && (
                            <>
                              <TableHead>Attachment Height</TableHead>
                              <TableHead>Proposed Attachment</TableHead>
                            </>
                          )}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {Object.values(connection.wires).map((wire: WireData) => (
                          <TableRow key={wire.traceId}>
                            <TableCell>{wire.company}</TableCell>
                            <TableCell>{wire.cableType}</TableCell>
                            <TableCell>
                              <Badge variant={getBadgeVariantForCategory(wire.category)}>
                                {wire.category}
                              </Badge>
                            </TableCell>
                            <TableCell>
                              {formatHeightToString(wire.lowestExistingMidspanHeight)}
                            </TableCell>
                            <TableCell>
                              {formatHeightToString(wire.finalProposedMidspanHeight)}
                              {wire.finalProposedMidspanHeight !== null && 
                               wire.lowestExistingMidspanHeight !== null && 
                               wire.finalProposedMidspanHeight !== wire.lowestExistingMidspanHeight && (
                                <Badge variant="outline" className="ml-2 bg-yellow-50">
                                  Changed
                                </Badge>
                              )}
                            </TableCell>
                            {connection.isRefConnection && (
                              <>
                                <TableCell>
                                  {formatHeightToString(wire.lowestExistingPoleAttachmentHeight)}
                                </TableCell>
                                <TableCell>
                                  {formatHeightToString(wire.finalProposedPoleAttachmentHeight)}
                                  {wire.finalProposedPoleAttachmentHeight !== null && 
                                   wire.lowestExistingPoleAttachmentHeight !== null && 
                                   wire.finalProposedPoleAttachmentHeight !== wire.lowestExistingPoleAttachmentHeight && (
                                    <Badge variant="outline" className="ml-2 bg-yellow-50">
                                      Changed
                                    </Badge>
                                  )}
                                </TableCell>
                              </>
                            )}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </AccordionContent>
              </AccordionItem>
            ))}
          </Accordion>
        </TabsContent>
        
        <TabsContent value="wires" className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            {Object.values(WireCategory).map(category => {
              const categoryWires = getAllWiresByCategory(connections, category);
              
              return (
                <Card key={category}>
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium">
                      {category} Wires
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="space-y-2">
                      <div>
                        <p className="text-xs text-muted-foreground">Total Count</p>
                        <p className="font-semibold">{categoryWires.length}</p>
                      </div>
                      <div>
                        <p className="text-xs text-muted-foreground">With Proposed Changes</p>
                        <p className="font-semibold">
                          {categoryWires.filter(wire => 
                            wire.finalProposedMidspanHeight !== null || 
                            wire.finalProposedPoleAttachmentHeight !== null
                          ).length}
                        </p>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              );
            })}
          </div>
          
          <Table>
            <TableCaption>All wires across all connections</TableCaption>
            <TableHeader>
              <TableRow>
                <TableHead>Company</TableHead>
                <TableHead>Cable Type</TableHead>
                <TableHead>Category</TableHead>
                <TableHead>Connection</TableHead>
                <TableHead>Lowest Midspan Height</TableHead>
                <TableHead>Proposed Midspan Height</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {connections.flatMap(connection => 
                Object.values(connection.wires).map((wire: WireData) => (
                  <TableRow key={`${connection.connectionId}-${wire.traceId}`}>
                    <TableCell>{wire.company}</TableCell>
                    <TableCell>{wire.cableType}</TableCell>
                    <TableCell>
                      <Badge variant={getBadgeVariantForCategory(wire.category)}>
                        {wire.category}
                      </Badge>
                    </TableCell>
                    <TableCell>
                      {connection.isRefConnection 
                        ? `REF: ${connection.fromPoleId}`
                        : `${connection.fromPoleId} → ${connection.toPoleId}`
                      }
                    </TableCell>
                    <TableCell>
                      {formatHeightToString(wire.lowestExistingMidspanHeight)}
                    </TableCell>
                    <TableCell>
                      {formatHeightToString(wire.finalProposedMidspanHeight)}
                      {wire.finalProposedMidspanHeight !== null && 
                       wire.lowestExistingMidspanHeight !== null && 
                       wire.finalProposedMidspanHeight !== wire.lowestExistingMidspanHeight && (
                        <Badge variant="outline" className="ml-2 bg-yellow-50">
                          Changed
                        </Badge>
                      )}
                    </TableCell>
                  </TableRow>
                ))
              )}
            </TableBody>
          </Table>
        </TabsContent>
      </Tabs>
    </div>
  );
};

// Helper function to get all wires of a specific category
function getAllWiresByCategory(connections: ConnectionData[], category: WireCategory): WireData[] {
  return connections.flatMap(connection => 
    Object.values(connection.wires).filter(wire => wire.category === category)
  );
}

// Helper function to determine badge variant based on wire category
function getBadgeVariantForCategory(category: WireCategory) {
  switch (category) {
    case WireCategory.COMMUNICATION:
      return "secondary";
    case WireCategory.CPS_ELECTRICAL:
      return "destructive";
    case WireCategory.OTHER:
    default:
      return "outline";
  }
}

export default KatapultDataVisualization;

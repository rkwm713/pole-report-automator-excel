import { useState, useEffect } from 'react';
import { toast } from 'sonner';
import { readFileAsText, downloadFile } from '@/lib/fileUtils';
import { PoleDataProcessor, PoleData } from '@/services/poleDataProcessor';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import FileUploader from '@/components/FileUploader';
import ProcessingStatus from '@/components/ProcessingStatus';
import DataPreview from '@/components/DataPreview';
import { FileText, FileSpreadsheet } from 'lucide-react';

// Add dependency to package.json
// <lov-add-dependency>xlsx@latest</lov-add-dependency>

const Index = () => {
  const [spidaFile, setSpidaFile] = useState<File | null>(null);
  const [katapultFile, setKatapultFile] = useState<File | null>(null);
  const [spidaData, setSpidaData] = useState<any>(null);
  const [katapultData, setKatapultData] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isComplete, setIsComplete] = useState(false);
  const [hasError, setHasError] = useState(false);
  const [progress, setProgress] = useState(0);
  const [errorDetails, setErrorDetails] = useState<string>('');
  const [processedPoles, setProcessedPoles] = useState<PoleData[]>([]);
  const [activeTab, setActiveTab] = useState<string>('upload');
  
  // Initialize the processor
  const [processor] = useState<PoleDataProcessor>(() => new PoleDataProcessor());

  // Handle file upload
  const handleFileSelected = async (file: File, fileType: 'spida' | 'katapult') => {
    try {
      // Reset any completed processing
      if (isComplete || hasError) {
        resetProcessing();
      }
      
      // Store the file
      if (fileType === 'spida') {
        setSpidaFile(file);
      } else {
        setKatapultFile(file);
      }
      
      // Read and parse the file
      const content = await readFileAsText(file);
      
      // Load into processor and update state
      if (fileType === 'spida') {
        processor.loadSpidaData(content);
        setSpidaData(JSON.parse(content));
      } else {
        processor.loadKatapultData(content);
        setKatapultData(JSON.parse(content));
      }
    } catch (error) {
      console.error(`Error loading ${fileType} file:`, error);
      toast.error(`Error loading file`, {
        description: `Failed to process ${fileType} file`
      });
    }
  };

  // Handle processing
  const handleProcessData = () => {
    if (!spidaFile || !katapultFile) {
      toast.error('Missing files', {
        description: 'Both SPIDAcalc and Katapult files are required.'
      });
      return;
    }
    
    setIsProcessing(true);
    setProgress(0);
    setHasError(false);
    setErrorDetails('');
    
    // Simulate progress updates
    const progressInterval = setInterval(() => {
      setProgress(prev => {
        const newProgress = prev + Math.random() * 10;
        return newProgress >= 100 ? 100 : newProgress;
      });
    }, 200);
    
    // Process data with a slight delay to show progress animation
    setTimeout(() => {
      try {
        const success = processor.processData();
        
        if (success) {
          setProcessedPoles(processor.getProcessedPoles());
          setIsComplete(true);
          setHasError(false);
          setProgress(100);
          toast.success('Processing complete', {
            description: `Successfully processed data for ${processor.getProcessedPoleCount()} poles.`
          });
          // Switch tab to results
          setActiveTab('results');
        } else {
          setHasError(true);
          const errors = processor.getErrors();
          setErrorDetails(errors.map(e => `${e.code}: ${e.message}\n${e.details || ''}`).join('\n'));
          toast.error('Processing failed', {
            description: errors[0]?.message || 'Unknown error occurred'
          });
        }
      } catch (error) {
        setHasError(true);
        setErrorDetails(error instanceof Error ? error.message : String(error));
        toast.error('Processing failed', {
          description: 'An unexpected error occurred'
        });
      } finally {
        clearInterval(progressInterval);
        setIsProcessing(false);
      }
    }, 1500);
  };

  // Generate Excel file
  const handleGenerateExcel = () => {
    if (!isComplete || hasError) {
      toast.error('Cannot generate Excel', {
        description: 'Processing must be completed successfully first.'
      });
      return;
    }
    
    try {
      const excelBlob = processor.generateExcel();
      if (excelBlob) {
        const date = new Date().toISOString().split('T')[0];
        downloadFile(excelBlob, `Make_Ready_Report_${date}.xlsx`);
        toast.success('Excel report generated', {
          description: 'Excel file has been downloaded.'
        });
      } else {
        toast.error('Failed to generate Excel', {
          description: 'No data available to generate report.'
        });
      }
    } catch (error) {
      console.error('Error generating Excel:', error);
      toast.error('Excel generation failed', {
        description: 'An error occurred while creating the Excel file.'
      });
    }
  };

  // Reset the processing state
  const resetProcessing = () => {
    setIsProcessing(false);
    setIsComplete(false);
    setHasError(false);
    setProgress(0);
    setErrorDetails('');
  };

  // Reset everything
  const handleReset = () => {
    setSpidaFile(null);
    setKatapultFile(null);
    setSpidaData(null);
    setKatapultData(null);
    setProcessedPoles([]);
    resetProcessing();
    setActiveTab('upload');
  };

  // Can process if both files are loaded and not currently processing
  const canProcess = Boolean(spidaFile && katapultFile) && !isProcessing;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 py-4">
        <div className="container mx-auto px-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="h-6 w-6 text-blue-600" />
              <h1 className="text-xl font-bold text-gray-900">Pole Analysis Report Generator</h1>
            </div>
          </div>
        </div>
      </header>
      
      {/* Main Content */}
      <main className="container mx-auto px-4 py-8">
        <Tabs value={activeTab} onValueChange={setActiveTab} className="space-y-6">
          <div className="flex justify-between items-center mb-6">
            <TabsList className="grid w-[400px] grid-cols-2">
              <TabsTrigger value="upload">File Upload</TabsTrigger>
              <TabsTrigger value="results">Processing Results</TabsTrigger>
            </TabsList>
            
            {/* Action buttons */}
            <div className="flex gap-3">
              {activeTab === 'upload' && (
                <Button 
                  onClick={handleProcessData} 
                  disabled={!canProcess || isProcessing}
                  className="gap-2"
                >
                  <FileText className="h-4 w-4" />
                  Process Files
                </Button>
              )}
              
              {activeTab === 'results' && isComplete && !hasError && (
                <Button 
                  onClick={handleGenerateExcel}
                  className="gap-2"
                >
                  <FileSpreadsheet className="h-4 w-4" />
                  Download Excel Report
                </Button>
              )}
              
              <Button variant="outline" onClick={handleReset}>
                Reset
              </Button>
            </div>
          </div>
          
          <TabsContent value="upload" className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <FileUploader 
                label="SPIDAcalc JSON File" 
                fileType="spida" 
                onFileSelected={handleFileSelected}
                isUploaded={!!spidaFile}
              />
              <FileUploader 
                label="Katapult JSON File" 
                fileType="katapult" 
                onFileSelected={handleFileSelected}
                isUploaded={!!katapultFile}
              />
            </div>
            
            <Card className="p-6 bg-blue-50 border border-blue-100">
              <div className="flex gap-4">
                <div className="shrink-0 mt-1">
                  <div className="bg-blue-100 rounded-full p-2">
                    <FileText className="h-5 w-5 text-blue-600" />
                  </div>
                </div>
                <div>
                  <h3 className="font-medium mb-1 text-blue-800">File Requirements</h3>
                  <p className="text-sm text-blue-700 mb-3">
                    For successful processing, please ensure your files meet these requirements:
                  </p>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <h4 className="text-sm font-medium text-blue-800 mb-1">SPIDAcalc JSON</h4>
                      <ul className="text-xs text-blue-700 list-disc list-inside space-y-1">
                        <li>Must contain pole structural data</li>
                        <li>Should include "Measured Design" and "Recommended Design"</li>
                        <li>Requires analysis results</li>
                      </ul>
                    </div>
                    <div>
                      <h4 className="text-sm font-medium text-blue-800 mb-1">Katapult JSON</h4>
                      <ul className="text-xs text-blue-700 list-disc list-inside space-y-1">
                        <li>Must contain node and connection data</li>
                        <li>Requires PoleNumber attributes</li>
                        <li>Should include field verification data</li>
                      </ul>
                    </div>
                  </div>
                </div>
              </div>
            </Card>
          </TabsContent>
          
          <TabsContent value="results" className="space-y-6">
            <ProcessingStatus
              isProcessing={isProcessing}
              isComplete={isComplete}
              hasError={hasError}
              progress={Math.round(progress)}
              onGenerateExcel={handleGenerateExcel}
              onReset={handleReset}
              errorDetails={errorDetails}
              poleCount={processedPoles.length}
            />
            
            <DataPreview
              spidaData={spidaData}
              katapultData={katapultData}
              processedPoles={processedPoles}
            />
          </TabsContent>
        </Tabs>
      </main>
      
      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 py-4 mt-12">
        <div className="container mx-auto px-4">
          <p className="text-sm text-center text-gray-600">
            Pole Analysis Report Generator &copy; {new Date().getFullYear()}
          </p>
        </div>
      </footer>
    </div>
  );
};

export default Index;


import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { isJsonFile } from '@/lib/fileUtils';
import { toast } from 'sonner';
import { Upload, FileX } from 'lucide-react';

interface FileUploaderProps {
  label: string;
  fileType: 'spida' | 'katapult';
  onFileSelected: (file: File, fileType: 'spida' | 'katapult') => void;
  isUploaded: boolean;
}

const FileUploader = ({ label, fileType, onFileSelected, isUploaded }: FileUploaderProps) => {
  const [dragActive, setDragActive] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const validateAndProcessFile = (file: File) => {
    if (!file) return;
    
    if (!isJsonFile(file)) {
      toast.error('Invalid file type', {
        description: 'Please upload a JSON file.'
      });
      return;
    }
    
    onFileSelected(file, fileType);
    toast.success('File uploaded successfully', {
      description: `${file.name} has been uploaded.`
    });
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      validateAndProcessFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    
    if (e.target.files && e.target.files.length > 0) {
      validateAndProcessFile(e.target.files[0]);
    }
  };

  const onButtonClick = () => {
    if (inputRef.current) {
      inputRef.current.click();
    }
  };

  return (
    <Card className={`relative ${isUploaded ? 'border-green-500' : ''}`}>
      <CardHeader className="pb-3">
        <CardTitle className="text-lg font-semibold">
          {label}
          {isUploaded && (
            <span className="ml-2 text-sm text-green-500">
              (Uploaded)
            </span>
          )}
        </CardTitle>
      </CardHeader>
      <CardContent>
        <div 
          className={`
            flex flex-col items-center justify-center border-2 border-dashed rounded-md p-6
            ${dragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-300'}
            ${isUploaded ? 'bg-green-50' : 'hover:bg-gray-50'}
            transition-colors duration-200
          `}
          onDragEnter={handleDrag}
          onDragOver={handleDrag}
          onDragLeave={handleDrag}
          onDrop={handleDrop}
        >
          <input
            ref={inputRef}
            type="file"
            className="hidden"
            accept=".json"
            onChange={handleChange}
            onClick={(e) => e.currentTarget.value = ''} // Allow re-upload of the same file
          />
          
          {isUploaded ? (
            <div className="text-center">
              <div className="flex justify-center mb-2">
                <div className="h-12 w-12 rounded-full bg-green-100 flex items-center justify-center">
                  <Upload size={24} className="text-green-500" />
                </div>
              </div>
              <p className="text-sm text-gray-600 mb-2">File uploaded successfully</p>
              <Button 
                variant="outline" 
                size="sm"
                onClick={onButtonClick}
              >
                Replace file
              </Button>
            </div>
          ) : (
            <>
              <div className="flex justify-center mb-4">
                <div className="h-12 w-12 rounded-full bg-blue-100 flex items-center justify-center">
                  <Upload size={24} className="text-blue-500" />
                </div>
              </div>
              <p className="text-sm text-gray-600 mb-4 text-center">
                Drag & drop your {fileType === 'spida' ? 'SPIDAcalc' : 'Katapult'} JSON file here, or
              </p>
              <Button onClick={onButtonClick} variant="outline">
                Browse files
              </Button>
            </>
          )}
        </div>
      </CardContent>
    </Card>
  );
};

export default FileUploader;

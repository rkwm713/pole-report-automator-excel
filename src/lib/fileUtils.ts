
/**
 * Utility functions for handling file operations in the app
 */

/**
 * Reads a file as text using a Promise-based FileReader
 */
export const readFileAsText = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
};

/**
 * Validates if a file is a JSON file by checking its extension
 */
export const isJsonFile = (file: File): boolean => {
  return file.name.toLowerCase().endsWith('.json');
};

/**
 * Gets a simplified file name (without path and extension)
 */
export const getSimpleFileName = (fileName: string): string => {
  // Remove path and extract base name
  const baseName = fileName.split(/[\\/]/).pop() || fileName;
  // Remove extension
  return baseName.replace(/\.[^/.]+$/, "");
};

/**
 * Downloads data as a file
 */
export const downloadFile = (data: Blob, filename: string): void => {
  const url = URL.createObjectURL(data);
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

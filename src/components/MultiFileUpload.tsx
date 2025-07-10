import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { File, X, UploadCloud } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import React from "react";

interface MultiFileUploadProps {
  id: string;
  label: string;
  files: File[];
  onFilesChange: (files: File[]) => void;
  accept: string;
  helpText: string;
}

export function MultiFileUpload({ id, label, files, onFilesChange, accept, helpText }: MultiFileUploadProps) {
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      onFilesChange([...files, ...Array.from(event.target.files)]);
      event.target.value = '';
    }
  };

  const handleRemoveFile = (fileToRemove: File) => {
    onFilesChange(files.filter(file => file !== fileToRemove));
  };

  const handleDropzoneClick = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className="space-y-2">
      <Label htmlFor={id}>{label}</Label>
      <div 
        className="relative flex flex-col items-center justify-center w-full p-6 border-2 border-dashed rounded-lg cursor-pointer hover:border-primary/50 transition-colors"
        onClick={handleDropzoneClick}
      >
        <UploadCloud className="w-10 h-10 text-muted-foreground" />
        <p className="mt-2 text-sm text-muted-foreground">
          <span className="font-semibold text-primary">Clicca per caricare</span> o trascina i file
        </p>
        <p className="text-xs text-muted-foreground">{helpText}</p>
        <Input 
          ref={fileInputRef}
          id={id} 
          type="file" 
          accept={accept} 
          onChange={handleFileChange} 
          multiple 
          className="sr-only"
        />
      </div>
      
      {files.length > 0 && (
        <div className="space-y-2 pt-2">
          <p className="text-sm font-medium">File caricati ({files.length}):</p>
          <div className="space-y-2 rounded-md border max-h-60 overflow-y-auto p-2">
            {files.map((file, index) => (
              <div key={`${file.name}-${index}`} className="flex items-center justify-between bg-muted/50 p-2 rounded-md gap-2">
                <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground overflow-hidden">
                  <File className="h-4 w-4 flex-shrink-0" />
                  <span className="truncate">{file.name}</span>
                  <Badge variant="outline" className="flex-shrink-0">{(file.size / 1024).toFixed(2)} KB</Badge>
                </div>
                <div className="flex items-center gap-2 flex-shrink-0">
                  <Button variant="ghost" size="icon" className="h-8 w-8" onClick={(e) => { e.stopPropagation(); handleRemoveFile(file); }}>
                    <X className="h-4 w-4" />
                  </Button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
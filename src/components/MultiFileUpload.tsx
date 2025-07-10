import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { File, X } from "lucide-react";
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
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      onFilesChange([...files, ...Array.from(event.target.files)]);
      // Reset the input value to allow re-uploading the same file
      event.target.value = '';
    }
  };

  const handleRemoveFile = (fileToRemove: File) => {
    onFilesChange(files.filter(file => file !== fileToRemove));
  };

  return (
    <div className="space-y-2">
      <Label htmlFor={id}>{label}</Label>
      <Input id={id} type="file" accept={accept} onChange={handleFileChange} multiple className="cursor-pointer file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-primary file:text-primary-foreground hover:file:bg-primary/90" />
      <p className="text-xs text-muted-foreground pt-1">{helpText}</p>
      {files.length > 0 && (
        <div className="space-y-2 pt-2">
          <p className="text-sm font-medium">File caricati ({files.length}):</p>
          <div className="space-y-2 rounded-md border max-h-48 overflow-y-auto p-2">
            {files.map((file, index) => (
              <div key={`${file.name}-${index}`} className="flex items-center justify-between bg-muted/50 p-2 rounded-md">
                <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground overflow-hidden">
                  <File className="h-4 w-4 flex-shrink-0" />
                  <span className="truncate">{file.name}</span>
                  <Badge variant="outline" className="flex-shrink-0">{(file.size / 1024).toFixed(2)} KB</Badge>
                </div>
                <Button variant="ghost" size="icon" className="flex-shrink-0" onClick={() => handleRemoveFile(file)}>
                  <X className="h-4 w-4" />
                </Button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
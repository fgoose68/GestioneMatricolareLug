import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { File, X, UploadCloud, FileText, Loader2 } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import React, { useState } from "react";
import { showSuccess, showError } from "@/utils/toast";

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
  const [convertingFile, setConvertingFile] = useState<string | null>(null);

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

  const handleConvertToPdf = async (file: File) => {
    setConvertingFile(file.name);
    try {
      const arrayBuffer = await file.arrayBuffer();
      // @ts-ignore
      const { value: html } = await mammoth.convertToHtml({ arrayBuffer });

      const element = document.createElement('div');
      element.innerHTML = html;
      element.style.padding = '2rem';
      element.style.fontFamily = 'Arial, sans-serif';
      element.style.lineHeight = '1.6';

      const opt = {
        margin:       [0.5, 0.5, 0.5, 0.5],
        filename:     `${file.name.replace(/\.docx$/, '')}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, logging: false },
        jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
      };
      // @ts-ignore
      await html2pdf().from(element).set(opt).save();
      
      showSuccess(`${file.name} convertito in PDF con successo!`);
    } catch (error) {
      console.error("Errore durante la conversione in PDF:", error);
      showError(`Errore durante la conversione di ${file.name}.`);
    } finally {
      setConvertingFile(null);
    }
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
                  <Button 
                    variant="outline" 
                    size="sm" 
                    onClick={(e) => { e.stopPropagation(); handleConvertToPdf(file); }}
                    disabled={convertingFile === file.name}
                    className="text-xs px-2 py-1 h-auto"
                  >
                    {convertingFile === file.name ? (
                      <Loader2 className="h-4 w-4 mr-1 animate-spin" />
                    ) : (
                      <FileText className="h-4 w-4 mr-1" />
                    )}
                    PDF
                  </Button>
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
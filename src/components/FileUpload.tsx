import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { File, X } from "lucide-react";

interface FileUploadProps {
  id: string;
  label: string;
  file: File | null;
  onUpload: (event: React.ChangeEvent<HTMLInputElement>) => void;
  onRemove: () => void;
  accept: string;
  helpText: string;
}

export function FileUpload({ id, label, file, onUpload, onRemove, accept, helpText }: FileUploadProps) {
  return (
    <div className="space-y-2">
      <Label htmlFor={id}>{label}</Label>
      {!file ? (
        <>
          <Input id={id} type="file" accept={accept} onChange={onUpload} />
          <p className="text-xs text-muted-foreground pt-1">{helpText}</p>
        </>
      ) : (
        <div className="flex items-center justify-between rounded-md border bg-muted/50 p-2">
          <div className="flex items-center gap-2 text-sm font-medium text-muted-foreground">
            <File className="h-4 w-4" />
            <span className="truncate max-w-xs">{file.name}</span>
          </div>
          <Button variant="ghost" size="icon" onClick={onRemove}>
            <X className="h-4 w-4" />
          </Button>
        </div>
      )}
    </div>
  );
}
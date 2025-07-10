import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { Printer } from "lucide-react";
import { MultiFileUpload } from "./MultiFileUpload";
import { showSuccess, showError, showWarning } from "@/utils/toast";

export function PrintManager() {
  const [files, setFiles] = useState<File[]>([]);
  const [copies, setCopies] = useState(1);
  const [duplex, setDuplex] = useState(false);

  const handlePrint = () => {
    if (files.length === 0) {
      showError("Per favore, carica almeno un file .docx da stampare.");
      return;
    }

    const fileNames = files.map(f => f.name).join(', ');
    console.log("Printing request:", {
      files: fileNames,
      copies,
      duplex,
    });

    showWarning("La stampa diretta di file .docx non è supportata dal browser.");
    showSuccess(`Richiesta di stampa simulata per: ${fileNames}. Copie: ${copies}, Fronte-retro: ${duplex ? 'Sì' : 'No'}.`);
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Stampa Documenti</CardTitle>
        <CardDescription>Carica i documenti .docx, imposta le opzioni e avvia la stampa.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <MultiFileUpload
          id="docx-files"
          label="Carica Documenti (.docx)"
          files={files}
          onFilesChange={setFiles}
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          helpText="Puoi selezionare più file."
        />
        <div className="grid sm:grid-cols-2 gap-4 items-end">
          <div className="space-y-2">
            <Label htmlFor="copies">Numero di copie</Label>
            <Input
              id="copies"
              type="number"
              min="1"
              value={copies}
              onChange={(e) => setCopies(Math.max(1, parseInt(e.target.value, 10) || 1))}
              className="w-full"
            />
          </div>
          <div className="flex items-center space-x-2 pb-2">
            <Switch id="duplex-printing" checked={duplex} onCheckedChange={setDuplex} />
            <Label htmlFor="duplex-printing">Stampa fronte-retro</Label>
          </div>
        </div>
        <Button size="lg" onClick={handlePrint} className="w-full" disabled={files.length === 0}>
          <Printer className="mr-2 h-5 w-5" /> Stampa
        </Button>
      </CardContent>
    </Card>
  );
}
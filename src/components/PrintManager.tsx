import { useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Printer, AlertTriangle, Loader2 } from "lucide-react";
import { MultiFileUpload } from "./MultiFileUpload";
import { showSuccess, showError, showWarning } from "@/utils/toast";

interface PrinterDevice {
  id: string;
  name: string;
}

export function PrintManager() {
  const [files, setFiles] = useState<File[]>([]);
  const [copies, setCopies] = useState(1);
  const [duplex, setDuplex] = useState(false);
  const [printers, setPrinters] = useState<PrinterDevice[]>([]);
  const [isLoadingPrinters, setIsLoadingPrinters] = useState(true);
  const [selectedPrinter, setSelectedPrinter] = useState<string | undefined>(undefined);

  useEffect(() => {
    // Simula il caricamento delle stampanti di rete
    setIsLoadingPrinters(true);
    setTimeout(() => {
      const mockPrinters: PrinterDevice[] = [
        { id: "HP-LaserJet-Pro-M404dn", name: "HP LaserJet Pro M404dn (Ufficio)" },
        { id: "Epson-WorkForce-WF-7840", name: "Epson WorkForce WF-7840 (Magazzino)" },
        { id: "Brother-HL-L2395DW", name: "Brother HL-L2395DW (Segreteria)" },
        { id: "PDF-Printer", name: "Salva come PDF" },
      ];
      setPrinters(mockPrinters);
      setSelectedPrinter(mockPrinters[0]?.id); // Seleziona la prima come predefinita
      setIsLoadingPrinters(false);
    }, 1500);
  }, []);

  const handlePrint = () => {
    if (files.length === 0) {
      showError("Per favore, carica almeno un file .pdf da stampare.");
      return;
    }
    if (!selectedPrinter) {
      showError("Per favore, seleziona una stampante.");
      return;
    }

    const selectedPrinterName = printers.find(p => p.id === selectedPrinter)?.name || selectedPrinter;
    const fileNames = files.map(f => f.name).join(', ');
    
    console.log("Printing request:", {
      files: fileNames,
      printer: selectedPrinterName,
      copies,
      duplex,
    });

    showWarning("La stampa diretta non è possibile. Questa è una simulazione.");
    showSuccess(`Richiesta di stampa inviata a "${selectedPrinterName}" per: ${fileNames}.`);
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Stampa Documenti</CardTitle>
        <CardDescription>Carica i documenti .pdf, imposta le opzioni e avvia la stampa.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <Alert variant="default">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Nota sulla Stampa</AlertTitle>
          <AlertDescription>
            La stampa diretta di file PDF non è possibile dal browser. Questa è una simulazione di una richiesta di stampa a un dispositivo di rete.
          </AlertDescription>
        </Alert>

        <MultiFileUpload
          id="pdf-files"
          label="Carica Documenti (.pdf)"
          files={files}
          onFilesChange={setFiles}
          accept=".pdf,application/pdf"
          helpText="Puoi selezionare più file PDF."
        />
        
        <div className="grid sm:grid-cols-2 gap-6 items-end">
          <div className="space-y-2">
            <Label htmlFor="printer">Seleziona Stampante</Label>
            <Select value={selectedPrinter} onValueChange={setSelectedPrinter} disabled={isLoadingPrinters}>
              <SelectTrigger id="printer">
                {isLoadingPrinters ? (
                  <div className="flex items-center gap-2 text-muted-foreground">
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Caricamento...</span>
                  </div>
                ) : (
                  <SelectValue placeholder="Seleziona una stampante" />
                )}
              </SelectTrigger>
              <SelectContent>
                {printers.map((printer) => (
                  <SelectItem key={printer.id} value={printer.id}>
                    {printer.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
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
        </div>
        
        <div className="flex items-center space-x-2">
          <Switch id="duplex-printing" checked={duplex} onCheckedChange={setDuplex} />
          <Label htmlFor="duplex-printing">Stampa fronte-retro</Label>
        </div>

        <Button size="lg" onClick={handlePrint} className="w-full" disabled={files.length === 0 || isLoadingPrinters}>
          <Printer className="mr-2 h-5 w-5" /> Stampa
        </Button>
      </CardContent>
    </Card>
  );
}
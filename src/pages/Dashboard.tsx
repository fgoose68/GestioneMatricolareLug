import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Calendar } from "@/components/ui/calendar";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Calendar as CalendarIcon, FileUp, FileText, Download, Eye } from "lucide-react";
import { format } from "date-fns";
import { it } from "date-fns/locale";
import { showError, showSuccess, showLoading, dismissToast } from "@/utils/toast";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";

import * as XLSX from "xlsx";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import JSZip from "jszip";

interface Discente {
  Matricola: string;
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
}

const Dashboard = () => {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [discenti, setDiscenti] = useState<Discente[]>([]);
  const [courseName, setCourseName] = useState<string>("");
  const [location, setLocation] = useState<string>("");
  const [startDate, setStartDate] = useState<Date | undefined>();
  const [endDate, setEndDate] = useState<Date | undefined>();
  const [signer, setSigner] = useState<string>("Il Direttore del Corso");

  const handleExcelUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setExcelFile(file);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          
          const courseNameCell = worksheet['B5'];
          const courseNameValue = courseNameCell ? String(courseNameCell.v) : "";
          setCourseName(courseNameValue);
          if (!courseNameValue) {
            showError("Attenzione: Nome del corso (cella B5) non trovato nel file Excel.");
          }

          const locationCell = worksheet['B6'];
          const locationValue = locationCell ? String(locationCell.v) : "";
          setLocation(locationValue);
           if (!locationValue) {
            showError("Attenzione: Località (cella B6) non trovata nel file Excel.");
          }

          const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { range: 7 });

          const discentiData = jsonData.map(row => {
            const matricola = row.matricola || row.Matricola || row['{matricola}'];
            const grado = row.grado || row.Grado || row['{grado}'];
            const cognome = row.cognome || row.Cognome || row['{cognome}'];
            const nome = row.nome || row.Nome || row['{nome}'];

            if (!matricola || !grado || !cognome || !nome) {
              console.warn("Riga saltata per dati mancanti:", row);
              return null;
            }

            return {
              "Matricola": String(matricola),
              "Grado militare": String(grado),
              "Cognome e Nome del Discente": `${cognome} ${nome}`,
            };
          }).filter(d => d !== null) as Discente[];

          if (discentiData.length === 0) {
            showError("Nessun discente trovato. Controlla che il file Excel abbia le intestazioni corrette alla riga 7 e i dati a partire dalla riga 8.");
            setDiscenti([]);
          } else {
            setDiscenti(discentiData);
            showSuccess(`File Excel "${file.name}" caricato. Trovati ${discentiData.length} discenti, nome corso e località.`);
          }
        } catch (error) {
          console.error("Errore nella lettura del file Excel:", error);
          showError("Formato file Excel non valido o corrotto.");
          setDiscenti([]);
          setCourseName("");
          setLocation("");
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const handleWordUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setWordFile(file);
      showSuccess(`Template Word "${file.name}" caricato.`);
    }
  };

  const handleGenerateDocument = () => {
    if (!wordFile || discenti.length === 0 || !startDate || !endDate || !courseName || !location) {
      showError("Per favore, carica entrambi i file, assicurati che contengano i dati necessari (corso, località, discenti) e compila le date.");
      return;
    }

    const toastId = showLoading("Generazione dei documenti individuali in corso...");

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const content = e.target?.result as ArrayBuffer;
        const outputZip = new JSZip();
        
        const formattedStartDate = format(startDate, "dd/MM/yyyy");
        const formattedEndDate = format(endDate, "dd/MM/yyyy");

        for (const discente of discenti) {
          const templateZip = new PizZip(content);
          const doc = new Docxtemplater(templateZip, {
            paragraphLoop: true,
            linebreaks: true,
          });

          doc.setData({
            titolocorso: courseName,
            localita: location,
            periodo_corso: `dal ${formattedStartDate} al ${formattedEndDate}`,
            firmatario: `${signer}\nCol. Massimiliano Fortino`,
            grado: discente["Grado militare"],
            cognome_nome: discente["Cognome e Nome del Discente"],
            matricola: discente.Matricola,
          });

          doc.render();

          const out = doc.getZip().generate({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          
          const filename = `Attestato_${discente["Cognome e Nome del Discente"].replace(/[^a-zA-Z0-9]/g, '_')}.docx`;
          outputZip.file(filename, out);
        }

        const zipBlob = await outputZip.generateAsync({ type: "blob" });
        saveAs(zipBlob, "documenti_individuali.zip");

        dismissToast(toastId);
        showSuccess("Archivio ZIP con documenti individuali generato con successo!");

      } catch (error: any) {
        dismissToast(toastId);
        console.error("Errore nella generazione del documento:", error);
        
        if (error.properties && Array.isArray(error.properties.errors)) {
          const firstError = error.properties.errors[0];
          if (firstError.id === 'scope_error') {
            showError(`Errore nel template: il segnaposto {${firstError.properties.tag}} non ha dati corrispondenti. Controlla i segnaposto nel Word.`);
          } else {
            showError(`Errore nel template Word: ${firstError.message}. Controlla il segnaposto '${firstError.properties.tag}'.`);
          }
        } else {
          showError(`Errore durante la generazione: ${error.message}`);
        }
      }
    };
    reader.readAsArrayBuffer(wordFile);
  };

  return (
    <div className="container mx-auto p-4 md:p-8">
      <header className="text-center mb-8">
        <h1 className="text-3xl font-bold tracking-tight">Dashboard Stampa Unione</h1>
        <p className="text-muted-foreground">
          Carica i file, imposta i dati e genera i documenti personalizzati.
        </p>
      </header>

      <div className="grid gap-8 md:grid-cols-3">
        <Card className="md:col-span-3">
          <CardHeader>
            <CardTitle className="flex items-center gap-2"><FileUp size={20} /> 1. Caricamento File</CardTitle>
          </CardHeader>
          <CardContent className="grid md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <Label htmlFor="excel-file">Carica File Excel (.xlsx)</Label>
              <Input id="excel-file" type="file" accept=".xlsx" onChange={handleExcelUpload} />
              {excelFile && <p className="text-sm text-muted-foreground">Caricato: {excelFile.name}</p>}
            </div>
            <div className="space-y-2">
              <Label htmlFor="word-file">Carica Template Word (.docx)</Label>
              <Input id="word-file" type="file" accept=".docx" onChange={handleWordUpload} />
              {wordFile && <p className="text-sm text-muted-foreground">Caricato: {wordFile.name}</p>}
               <p className="text-xs text-muted-foreground pt-2">
                Il template deve contenere i segnaposto come {`{titolocorso}`}, {`{localita}`}, {`{grado}`}, {`{cognome_nome}`}, e {`{matricola}`}.
              </p>
            </div>
          </CardContent>
        </Card>

        <Card className="md:col-span-2">
          <CardHeader>
            <CardTitle className="flex items-center gap-2"><FileText size={20} /> 2. Dettagli del Corso</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label htmlFor="course-name">Nome del Corso (da Excel)</Label>
                <Input id="course-name" type="text" value={courseName} readOnly placeholder="Da Excel (cella B5)" />
              </div>
              <div className="space-y-2">
                <Label htmlFor="location">Località (da Excel)</Label>
                <Input id="location" type="text" value={location} readOnly placeholder="Da Excel (cella B6)" />
              </div>
            </div>
            <div className="grid md:grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Data Inizio</Label>
                <Popover>
                  <PopoverTrigger asChild>
                    <Button variant={"outline"} className="w-full justify-start text-left font-normal">
                      <CalendarIcon className="mr-2 h-4 w-4" />
                      {startDate ? format(startDate, "PPP", { locale: it }) : <span>Seleziona una data</span>}
                    </Button>
                  </PopoverTrigger>
                  <PopoverContent className="w-auto p-0">
                    <Calendar mode="single" selected={startDate} onSelect={setStartDate} initialFocus />
                  </PopoverContent>
                </Popover>
              </div>
              <div className="space-y-2">
                <Label>Data Fine</Label>
                <Popover>
                  <PopoverTrigger asChild>
                    <Button variant={"outline"} className="w-full justify-start text-left font-normal">
                      <CalendarIcon className="mr-2 h-4 w-4" />
                      {endDate ? format(endDate, "PPP", { locale: it }) : <span>Seleziona una data</span>}
                    </Button>
                  </PopoverTrigger>
                  <PopoverContent className="w-auto p-0">
                    <Calendar mode="single" selected={endDate} onSelect={setEndDate} initialFocus />
                  </PopoverContent>
                </Popover>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2"><FileText size={20} /> 3. Firmatario</CardTitle>
          </CardHeader>
          <CardContent>
            <RadioGroup value={signer} onValueChange={setSigner} className="space-y-2">
              <div className="flex items-center space-x-2">
                <RadioGroupItem value="Il Direttore del Corso" id="r1" />
                <Label htmlFor="r1">Il Direttore del Corso</Label>
              </div>
              <div className="flex items-center space-x-2">
                <RadioGroupItem value="Il Comandante del Centro" id="r2" />
                <Label htmlFor="r2">Il Comandante del Centro</Label>
              </div>
            </RadioGroup>
            <p className="text-sm text-muted-foreground mt-4">La firma sarà: Col. Massimiliano Fortino</p>
          </CardContent>
        </Card>

        {discenti.length > 0 && (
          <Card className="md:col-span-3">
            <CardHeader>
              <CardTitle className="flex items-center gap-2"><Eye size={20} /> Anteprima Dati da Excel</CardTitle>
            </CardHeader>
            <CardContent>
              <p className="mb-4"><strong>Numero di Discenti:</strong> <span className="font-mono p-1 bg-muted rounded-md">{discenti.length}</span></p>
              <div className="max-h-60 overflow-y-auto rounded-md border">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Matricola</TableHead>
                      <TableHead>Grado militare</TableHead>
                      <TableHead>Cognome e Nome del Discente</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {discenti.map((d, index) => (
                      <TableRow key={index}>
                        <TableCell>{d.Matricola}</TableCell>
                        <TableCell>{d["Grado militare"]}</TableCell>
                        <TableCell>{d["Cognome e Nome del Discente"]}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
              <p className="text-sm text-muted-foreground mt-2">
                Controlla che i dati corrispondano. L'app si aspetta che il file Excel contenga le intestazioni corrette alla riga 7 e i dati a partire dalla riga 8.
              </p>
            </CardContent>
          </Card>
        )}

        <div className="md:col-span-3 flex justify-center">
          <Button size="lg" onClick={handleGenerateDocument} className="w-full md:w-1/2 lg:w-1/3">
            <Download className="mr-2 h-5 w-5" />
            Genera Documento
          </Button>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
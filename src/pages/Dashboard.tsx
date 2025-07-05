import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { FileUp, FileText, Download, Eye, Info } from "lucide-react";
import { format } from "date-fns";
import { showError, showSuccess, showLoading, dismissToast, showWarning } from "@/utils/toast";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";

import * as XLSX from "xlsx";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import JSZip from "jszip";

// Helper function to safely get cell value as string
const getCellValueAsString = (cellValue: any): string => {
  return cellValue !== null && cellValue !== undefined ? String(cellValue).trim() : '';
};

interface Discente {
  Matricola: string;
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
  Cognome: string;
  Nome: string;
  Categoria: string;
}

interface CourseInfo {
  title: string;
  location: string;
  period: string;
  currentDate: string; // Aggiunto il campo per la data odierna
}

function Dashboard() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [discenti, setDiscenti] = useState<Discente[]>([]);
  const [signer, setSigner] = useState<string>("Il Direttore del Corso");
  const [courseInfo, setCourseInfo] = useState<CourseInfo | null>(null);

  const handleExcelUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setExcelFile(file);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        setDiscenti([]);
        setCourseInfo(null);

        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Estrazione dati corso dalle celle originali (A5, C6, D6, E6)
        const titleCell = worksheet['A5'];
        const locationCell = worksheet['C6'];
        const startDateCell = worksheet['D6'];
        const endDateCell = worksheet['E6'];

        const title = getCellValueAsString(titleCell?.v);
        const location = getCellValueAsString(locationCell?.v);
        const startDate = startDateCell?.v ? (startDateCell.v instanceof Date ? format(startDateCell.v, "dd/MM/yyyy") : getCellValueAsString(startDateCell.v)) : "";
        const endDate = endDateCell?.v ? (endDateCell.v instanceof Date ? format(endDateCell.v, "dd/MM/yyyy") : getCellValueAsString(endDateCell.v)) : "";

        const missingCourseFields: string[] = [];
        if (!title) missingCourseFields.push("Titolo (A5)");
        if (!location) missingCourseFields.push("Sede (C6)");
        if (!startDate) missingCourseFields.push("Data Inizio (D6)");
        if (!endDate) missingCourseFields.push("Data Fine (E6)");

        if (missingCourseFields.length > 0) {
          showError(`Dati corso mancanti o non validi: ${missingCourseFields.join(", ")}. Controlla il formato del file.`);
          setCourseInfo(null);
          return;
        }

        const period = `dal ${startDate} al ${endDate}`;
        const today = format(new Date(), "dd/MM/yyyy");
        setCourseInfo({ title, location, period, currentDate: today });
        showSuccess(`Dati corso estratti: ${title}, ${location}, ${period}`);


        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: true,
        });

        // Rileva automaticamente la riga delle intestazioni per renderlo più robusto
        let headerRowIndex = -1;
        let headers: string[] = [];
        const searchLimit = Math.min(jsonData.length, 20); 

        for (let i = 0; i < searchLimit; i++) {
            const potentialHeaders = jsonData[i].filter(h => h !== null && h !== undefined && String(h).trim() !== '');
            if (potentialHeaders.length < 3) {
                continue;
            }

            const rowAsHeaders = jsonData[i].map(h => getCellValueAsString(h).toLowerCase());
            
            const hasMatricola = rowAsHeaders.some(h => h.includes('matricola'));
            const hasGrado = rowAsHeaders.some(h => h.includes('grado'));
            const hasName = rowAsHeaders.some(h => h.includes('cognome') || h.includes('nominativo'));

            if ((hasMatricola && hasGrado) || (hasMatricola && hasName) || (hasGrado && hasName)) {
                headerRowIndex = i;
                headers = rowAsHeaders;
                break;
            }
        }

        if (headerRowIndex === -1) {
            showError("Impossibile trovare la riga delle intestazioni nel file Excel. Assicurati che contenga colonne come 'Matricola', 'Grado' e 'Cognome'.");
            return;
        }

        const dataRows = jsonData.slice(headerRowIndex + 1);

        if (dataRows.length === 0) {
          showError("Nessun discente trovato dopo la riga delle intestazioni. Controlla che il file Excel contenga dati validi.");
          return;
        }

        const findIndex = (keywords: string[]) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));

        const matricolaIndex = findIndex(['matricola']);
        const gradoIndex = findIndex(['grado']);
        const categoriaIndex = findIndex(['cat.', 'cat', 'categoria']);
        
        const cognomeIndex = findIndex(['cognome']);
        const nomeIndex = findIndex(['nome']);
        const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo']);

        const essentialHeadersPresent = 
            (matricolaIndex !== -1) && 
            (gradoIndex !== -1) && 
            ((cognomeIndex !== -1 && nomeIndex !== -1) || cognomeNomeIndex !== -1);

        if (!essentialHeadersPresent) {
            const missing = [];
            if (matricolaIndex === -1) missing.push("'Matricola'");
            if (gradoIndex === -1) missing.push("'Grado'");
            if (cognomeIndex === -1 && nomeIndex === -1 && cognomeNomeIndex === -1) missing.push("almeno una colonna per 'Cognome' e 'Nome' (separate o combinate)");
            showError(`Intestazioni essenziali mancanti o non riconosciute. Assicurati che il file Excel contenga le colonne: ${missing.join(', ')}.`);
            return;
        }

        const valueExists = (val: any) => val !== null && val !== undefined && String(val).trim() !== '';

        const discentiData = dataRows.map((row, rowIndex) => {
          if (row.every(cell => !valueExists(cell))) {
            return null; 
          }

          let cognomeNome: string = '';
          let cognome: string = '';
          let nome: string = '';

          if (cognomeIndex !== -1 && nomeIndex !== -1 && valueExists(row[cognomeIndex]) && valueExists(row[nomeIndex])) {
              // Priority 1: Use separate Cognome and Nome columns if available
              cognome = getCellValueAsString(row[cognomeIndex]);
              nome = getCellValueAsString(row[nomeIndex]);
              cognomeNome = `${cognome} ${nome}`.trim();
          } 
          else if (cognomeNomeIndex !== -1 && valueExists(row[cognomeNomeIndex])) {
              // Priority 2: Process combined "Cognome e Nome" or "Nominativo" column
              const fullName = getCellValueAsString(row[cognomeNomeIndex]);
              cognomeNome = fullName;
              const parts = fullName.split(' ').filter(p => p.length > 0);

              if (parts.length > 1) {
                // If more than one part, assume the last part is the name and the rest is the surname.
                nome = parts.pop() as string;
                cognome = parts.join(' ');
              } else if (parts.length === 1) {
                // If only one part, it's the surname.
                cognome = parts[0];
                nome = '';
              } else {
                // Empty or whitespace only
                cognome = '';
                nome = '';
              }
          } else {
              // If no valid name columns found, set to empty
              cognome = '';
              nome = '';
              cognomeNome = '';
          }

          const matricola: string = matricolaIndex !== -1 ? getCellValueAsString(row[matricolaIndex]) : '';
          const grado: string = gradoIndex !== -1 ? getCellValueAsString(row[gradoIndex]) : '';
          const categoria: string = categoriaIndex !== -1 ? getCellValueAsString(row[categoriaIndex]) : '';

          if (!matricola || !grado || !cognomeNome) {
            const missingFields = [];
            if (!matricola) missingFields.push("Matricola");
            if (!grado) missingFields.push("Grado militare");
            if (!cognomeNome) missingFields.push("Cognome e Nome");
            return null;
          }

          return {
            "Matricola": matricola,
            "Grado militare": grado,
            "Cognome e Nome del Discente": cognomeNome,
            "Cognome": cognome,
            "Nome": nome,
            "Categoria": categoria,
          };
        }).filter(d => d !== null) as Discente[];

        const problematicRows = discentiData.filter(d => !d.Nome || d.Nome === d.Cognome);
        if (problematicRows.length > 0) {
            showWarning(`Attenzione: per ${problematicRows.length} discenti, il nome è mancante o identico al cognome. Si prega di verificare i dati nell'anteprima.`);
        }

        if (discentiData.length === 0) {
          showError("Nessun discente valido caricato. Controlla che i dati siano corretti e completi.");
        } else {
          setDiscenti(discentiData);
          showSuccess(`Caricamento completato. Trovati ${discentiData.length} discenti.`);
        }
      } catch (error) {
        console.error("Errore imprevisto durante la lettura del file Excel:", error);
        showError("Errore durante l'elaborazione del file Excel. Assicurati che sia un file .xlsx valido e non corrotto.");
        setDiscenti([]);
        setCourseInfo(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleWordUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setWordFile(file);
      showSuccess(`Template Word "${file.name}" caricato.`);
    }
  };

  const handleGenerateDocument = () => {
    if (!wordFile || discenti.length === 0 || !courseInfo) {
      showError("Per favor, carica il file Excel con i dati e il template Word.");
      return;
    }

    const toastId = showLoading("Generazione dei documenti in corso...");

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const content = e.target?.result as ArrayBuffer;
        const outputZip = new JSZip();
        
        for (const discente of discenti) {
          const templateZip = new PizZip(content);
          const doc = new Docxtemplater(templateZip, {
            paragraphLoop: true,
            linebreaks: true,
          });

          doc.setData({
            titolocorso: courseInfo.title,
            categoria: discente.Categoria,
            localita: courseInfo.location,
            periodo_corso: courseInfo.period,
            firmatario: `${signer}\nCol. Massimiliano Fortino`,
            grado: discente["Grado militare"],
            cognome_nome: discente["Cognome e Nome del Discente"],
            cognome: discente.Cognome,
            nome: discente.Nome,
            matricola: discente.Matricola,
            datafirma: courseInfo.currentDate,
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
        showSuccess("Archivio ZIP con documenti generato con successo!");

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
          Carica i file e genera i documenti personalizzati.
        </p>
      </header>

      <div className="grid gap-8 md:grid-cols-2">
        <Card className="md:col-span-2">
          <CardHeader className="flex justify-between items-center">
            <CardTitle className="flex items-center gap-2"><FileUp size={20} /> 1. Caricamento File</CardTitle>
            <div className="text-sm text-muted-foreground">
              Ver.5.1Lug25
            </div>
          </CardHeader>
          <CardContent className="grid md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <Label htmlFor="excel-file">Carica File Excel (.xlsx)</Label>
              <Input id="excel-file" type="file" accept=".xlsx" onChange={handleExcelUpload} />
              {excelFile && <p className="text-sm text-muted-foreground">Caricato: {excelFile.name}</p>}
              <p className="text-xs text-muted-foreground pt-2">
                Titolo da A5, Sede da C6, Periodo da D6/E6. Le intestazioni dei discenti verranno trovate automaticamente.
              </p>
            </div>
            <div className="space-y-2">
              <Label htmlFor="word-file">Carica Template Word (.docx)</Label>
              <Input id="word-file" type="file" accept=".docx" onChange={handleWordUpload} />
              {wordFile && <p className="text-sm text-muted-foreground">Caricato: {wordFile.name}</p>}
               <p className="text-xs text-muted-foreground pt-2">
                Il template deve contenere i segnaposto come {`{titolocorso}`}, {`{categoria}`}, ecc.
              </p>
            </div>
          </CardContent>
        </Card>

        {courseInfo && (
          <Card className="md:col-span-2">
            <CardHeader>
              <CardTitle className="flex items-center gap-2"><Info size={20} /> Dati del Corso Estratti</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              <p><strong>Titolo del Corso:</strong> {courseInfo.title}</p>
              <p><strong>Sede del Corso:</strong> {courseInfo.location}</p>
              <p><strong>Periodo del Corso:</strong> {courseInfo.period}</p>
              <p><strong>Data della firma:</strong> {courseInfo.currentDate}</p>
            </CardContent>
          </Card>
        )}

        <Card className="md:col-span-2">
          <CardHeader className="flex justify-center items-center">
            <CardTitle className="flex items-center gap-2"><FileText size={20} /> 2. Firmatario</CardTitle>
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
            <p className="text-sm text-muted-foreground mt-4">
              Il titolo selezionato apparirà sopra il nome del firmatario: <strong>Col. Massimiliano Fortino</strong>.
            </p>
          </CardContent>
        </Card>

        {discenti.length > 0 && (
          <Card className="md:col-span-2">
            <CardHeader>
              <CardTitle className="flex items-center gap-2"><Eye size={20} /> Anteprima Dati Discenti</CardTitle>
            </CardHeader>
            <CardContent>
              <p className="mb-4"><strong>Numero di Discenti:</strong> <span className="font-mono p-1 bg-muted rounded-md">{discenti.length}</span></p>
              <div className="max-h-80 overflow-y-auto rounded-md border">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Matricola</TableHead>
                      <TableHead>Grado</TableHead>
                      <TableHead>Cognome</TableHead>
                      <TableHead>Nome</TableHead>
                      <TableHead>Categoria</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {discenti.map((d, index) => (
                      <TableRow key={index}>
                        <TableCell>{d.Matricola}</TableCell>
                        <TableCell>{d["Grado militare"]}</TableCell>
                        <TableCell>{d.Cognome}</TableCell>
                        <TableCell>{d.Nome}</TableCell>
                        <TableCell>{d.Categoria}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
              <p className="text-sm text-muted-foreground mt-2">
                Controlla che i dati dei discenti corrispondano.
              </p>
            </CardContent>
          </Card>
        )}

        <div className="md:col-span-2 flex justify-center">
          <Button size="lg" onClick={handleGenerateDocument} className="w-full md:w-1/2 lg:w-1/3" disabled={!wordFile || discenti.length === 0}>
            <Download className="mr-2 h-5 w-5" />
            Genera Documenti
          </Button>
        </div>
      </div>
    </div>
  );
}

export default Dashboard;
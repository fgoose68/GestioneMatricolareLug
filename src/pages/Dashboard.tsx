import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { FileUp, FileText, Download, Eye, Info } from "lucide-react";
import { format } from "date-fns";
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

        // Aggiungi controlli per la validità del workbook e dei fogli
        if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
          showError("Il file Excel non contiene fogli di lavoro validi.");
          setDiscenti([]);
          setCourseInfo(null);
          return;
        }

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
          showError(`Impossibile accedere al foglio di lavoro "${sheetName}". Il file potrebbe essere corrotto o il foglio non esiste.`);
          setDiscenti([]);
          setCourseInfo(null);
          return;
        }

        // Estrazione dati corso dalle celle originali (A5, C6, D6, E6)
        const titleCell = worksheet['A5'];
        const locationCell = worksheet['C6'];
        const startDateCell = worksheet['D6'];
        const endDateCell = worksheet['E6'];

        const title = titleCell ? String(titleCell.v) : "";
        const location = locationCell ? String(locationCell.v) : "";
        const startDate = startDateCell?.v ? (startDateCell.v instanceof Date ? format(startDateCell.v, "dd/MM/yyyy") : String(startDateCell.v)) : "";
        const endDate = endDateCell?.v ? (endDateCell.v instanceof Date ? format(endDateCell.v, "dd/MM/yyyy") : String(endDateCell.v)) : "";

        const missingCourseFields: string[] = [];
        if (!title) missingCourseFields.push("Titolo (A5)");
        if (!location) missingCourseFields.push("Sede (C6)");
        if (!startDate) missingCourseFields.push("Data Inizio (D6)");
        if (!endDate) missingCourseFields.push("Data Fine (E6)");

        if (missingCourseFields.length > 0) {
          showError(`Dati corso mancanti o non validi: ${missingCourseFields.join(", ")}. Controlla il formato del file.`);
          setCourseInfo(null);
          return; // Stop processing if essential course info is missing
        }

        const period = `dal ${startDate} al ${endDate}`;
        const today = format(new Date(), "dd/MM/yyyy"); // Data odierna
        setCourseInfo({ title, location, period, currentDate: today });
        showSuccess(`Dati corso estratti: ${title}, ${location}, ${period}`);


        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        // Imposta la riga delle intestazioni alla riga 7 (indice 6 in un array 0-based)
        const headerRowIndex = 6; 

        if (jsonData.length <= headerRowIndex) {
            showError("Il file Excel non contiene abbastanza righe per trovare le intestazioni alla riga 7.");
            return;
        }

        const headers = jsonData[headerRowIndex].map(h => String(h).toLowerCase().trim());
        const dataRows = jsonData.slice(headerRowIndex + 1);

        if (dataRows.length === 0) {
          showError("Nessun discente trovato dopo la riga delle intestazioni. Controlla che il file Excel contenga dati validi.");
          return;
        }

        const findIndex = (keywords: string[]) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));

        const matricolaIndex = findIndex(['matricola']);
        const gradoIndex = findIndex(['grado']);
        const categoriaIndex = findIndex(['cat.', 'cat', 'categoria']);
        
        const cognomeIndex = findIndex(['cognome']); // Check for separate cognome first
        const nomeIndex = findIndex(['nome']);       // Check for separate nome first
        const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo']); // Combined name last

        const valueExists = (val: any) => val !== null && val !== undefined && String(val).trim() !== '';

        const discentiData = dataRows.map((row, rowIndex) => {
          // Salta righe completamente vuote
          if (row.every(cell => !valueExists(cell))) {
            return null; 
          }

          let cognomeNome: string = '';
          let cognome: string = '';
          let nome: string = '';

          // Priority 1: Separate Cognome and Nome columns
          if (cognomeIndex !== -1 && nomeIndex !== -1 && valueExists(row[cognomeIndex]) && valueExists(row[nomeIndex])) {
              cognome = String(row[cognomeIndex]).trim();
              nome = String(row[nomeIndex]).trim();
              cognomeNome = `${cognome} ${nome}`.trim();
              console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Usando colonne separate. Cognome: '${cognome}', Nome: '${nome}'`);
          } 
          // Priority 2: Combined "Cognome e Nome" or "Nominativo" column
          else if (cognomeNomeIndex !== -1 && valueExists(row[cognomeNomeIndex])) {
              const fullName = String(row[cognomeNomeIndex]).trim();
              cognomeNome = fullName; // This will be the full string for the template

              const parts = fullName.split(' ').filter(p => p.length > 0);

              if (parts.length === 0) {
                  cognome = '';
                  nome = '';
                  console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato vuoto. Cognome: '${cognome}', Nome: '${nome}'`);
              } else if (parts.length === 1) {
                  // If only one word, assume it's the surname and name is empty
                  cognome = parts[0];
                  nome = '';
                  console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato singola parola. Cognome: '${cognome}', Nome: '${nome}'`);
              } else {
                  const allPartsAreUppercase = parts.every(part => part.toUpperCase() === part);

                  if (!allPartsAreUppercase) {
                      // Case: "ROSSI Mario" or "De Rossi Mario" (mixed case)
                      // Find the first part that is not all uppercase, assume it's the start of the name.
                      let nameStartIndex = -1;
                      for (let i = 0; i < parts.length; i++) {
                          if (parts[i].toUpperCase() !== parts[i]) { // Found a non-uppercase part
                              nameStartIndex = i;
                              break;
                          }
                      }

                      if (nameStartIndex !== -1) {
                          cognome = parts.slice(0, nameStartIndex).join(' ');
                          nome = parts.slice(nameStartIndex).join(' ');
                          console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato misto. Cognome: '${cognome}', Nome: '${nome}'`);
                      } else {
                          // Fallback: if no non-uppercase part found but not all parts are uppercase,
                          // this case should ideally not be hit. Treat as full name as surname.
                          cognome = fullName;
                          nome = '';
                          console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato misto fallback. Cognome: '${cognome}', Nome: '${nome}'`);
                      }
                  } else {
                      // All parts are uppercase (e.g., "IMPERIALE FABRIZIO", "DE ROSSI", "ROSSI ROSSI")
                      // Heuristic: If the first two parts are identical (e.g., "ROSSI ROSSI"), assume it's a single surname.
                      if (parts.length >= 2 && parts[0].toUpperCase() === parts[1].toUpperCase()) {
                          cognome = parts[0]; // Take only the first part as surname
                          nome = '';
                          console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato tutto maiuscolo, cognome ripetuto. Cognome: '${cognome}', Nome: '${nome}'`);
                      } else if (parts.length > 1) {
                          // Assume the last word is the name, and the rest is the surname.
                          // This handles "IMPERIALE FABRIZIO" -> cognome="IMPERIALE", nome="FABRIZIO"
                          // NOTE: This heuristic might split multi-word surnames like "DE ROSSI" into "DE" (cognome) and "ROSSI" (nome).
                          // It's a trade-off to correctly handle "IMPERIALE FABRIZIO" where FABRIZIO is the name.
                          const lastPart = parts[parts.length - 1]; 
                          cognome = parts.slice(0, parts.length - 1).join(' ');
                          nome = lastPart;
                          console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato tutto maiuscolo, diviso. Cognome: '${cognome}', Nome: '${nome}'`);
                      } else {
                          // Should be caught by parts.length === 1, but as a safeguard
                          cognome = fullName;
                          nome = '';
                          console.log(`DEBUG (Riga ${headerRowIndex + rowIndex + 2}): Campo combinato tutto maiuscolo, singola parte fallback. Cognome: '${cognome}', Nome: '${nome}'`);
                      }
                  }
              }
          } else {
              // If no valid name columns found, set to empty to avoid undefined in table
              console.warn(`Riga ${headerRowIndex + rowIndex + 2} del file Excel saltata perché mancano dati essenziali per Cognome e Nome. Dati letti:`, { matricola, grado, cognomeNome, cognome, nome, categoria });
              cognome = '';
              nome = '';
              cognomeNome = '';
          }

          const matricola = matricolaIndex !== -1 ? row[matricolaIndex] : undefined;
          const grado = gradoIndex !== -1 ? row[gradoIndex] : undefined;
          const categoria = categoriaIndex !== -1 && valueExists(row[categoriaIndex]) ? row[categoriaIndex] : "";

          // Controllo rigoroso per dati essenziali
          if (!valueExists(matricola) || !valueExists(grado) || !valueExists(cognomeNome)) {
            const missingFields = [];
            if (!valueExists(matricola)) missingFields.push("Matricola");
            if (!valueExists(grado)) missingFields.push("Grado militare");
            if (!valueExists(cognomeNome)) missingFields.push("Cognome e Nome del Discente (o Cognome/Nome separati)");
            console.warn(`Riga ${headerRowIndex + rowIndex + 2} del file Excel saltata perché mancano dati essenziali: ${missingFields.join(', ')}. Dati letti:`, { matricola, grado, cognomeNome, cognome, nome, categoria });
            return null;
          }

          // Logging dettagliato per debug
          console.log(`Riga ${headerRowIndex + rowIndex + 2} - Estratto finale:`, {
            Matricola: String(matricola),
            Grado_militare: String(grado),
            "Cognome e Nome del Discente": cognomeNome,
            Cognome: cognome,
            Nome: nome,
            Categoria: String(categoria),
          });

          return {
            "Matricola": String(matricola),
            "Grado militare": String(grado),
            "Cognome e Nome del Discente": cognomeNome,
            "Cognome": cognome,
            "Nome": nome,
            "Categoria": String(categoria),
          };
        }).filter(d => d !== null) as Discente[];


        if (discentiData.length === 0) {
          showError("Nessun discente valido caricato. Controlla che le intestazioni e i dati siano corretti e completi.");
        } else {
          setDiscenti(discentiData);
          showSuccess(`Caricamento completato. Trovati ${discentiData.length} discenti.`);
        }
      } catch (error) {
        console.error("Errore imprevisto durante la lettura del file Excel:", error);
        showError("Errore durante l'elaborazione del file. Assicurati che sia un file .xlsx valido e non corrotto.");
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
            datafirma: courseInfo.currentDate, // Passa la data odierna al template
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
          <CardHeader>
            <CardTitle className="flex items-center gap-2"><FileUp size={20} /> 1. Caricamento File</CardTitle>
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
          <CardHeader>
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
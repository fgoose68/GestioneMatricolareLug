import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { FileUp, FileText, Download, Eye } from "lucide-react";
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
  Categoria: string;
  Titolocorso: string;
  Localita: string;
  DataInizio: string;
  DataFine: string;
}

function Dashboard() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [discenti, setDiscenti] = useState<Discente[]>([]);
  const [signer, setSigner] = useState<string>("Il Direttore del Corso");

  const handleExcelUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setExcelFile(file);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          showError("Nessun dato trovato dopo la riga di intestazione. Controlla che il file non sia vuoto dalla seconda riga in poi.");
          setDiscenti([]);
          return;
        }

        const discentiData = jsonData.map((row, index) => {
          // Rende la lettura delle intestazioni flessibile:
          // 1. Converte tutte le chiavi (nomi delle colonne) in minuscolo.
          // 2. Rimuove spazi bianchi all'inizio e alla fine.
          const normalizedRow: { [key: string]: any } = {};
          for (const key in row) {
            if (Object.prototype.hasOwnProperty.call(row, key)) {
              normalizedRow[key.toLowerCase().trim()] = row[key];
            }
          }

          // Cerca i dati nella riga "normalizzata", usando anche alias comuni.
          const matricola = normalizedRow.matricola;
          const grado = normalizedRow.grado;
          const cognome = normalizedRow.cognome;
          const nome = normalizedRow.nome;
          const categoria = normalizedRow.cat || normalizedRow.categoria;
          const titolocorso = normalizedRow.corso || normalizedRow.titolocorso;
          const localita = normalizedRow.sede || normalizedRow.localita;
          const dal = normalizedRow.dal;
          const al = normalizedRow.al;

          if (!matricola || !grado || !cognome || !nome || !categoria || !titolocorso || !localita || !dal || !al) {
            console.warn(`Riga ${index + 2} del file Excel saltata perché mancano uno o più dati richiesti. Dati letti:`, row, 'Dati normalizzati:', normalizedRow);
            return null;
          }

          return {
            "Matricola": String(matricola),
            "Grado militare": String(grado),
            "Cognome e Nome del Discente": `${cognome} ${nome}`,
            "Categoria": String(categoria),
            "Titolocorso": String(titolocorso),
            "Localita": String(localita),
            "DataInizio": dal instanceof Date ? format(dal, "dd/MM/yyyy") : String(dal),
            "DataFine": al instanceof Date ? format(al, "dd/MM/yyyy") : String(al),
          };
        }).filter(d => d !== null) as Discente[];

        if (discentiData.length === 0) {
          showError("Nessun discente valido caricato. Controlla che le intestazioni nella riga 1 del file Excel siano corrette (es. 'matricola', 'grado', 'sede', ecc.) e che i dati siano presenti in tutte le colonne a partire dalla riga 2.");
          setDiscenti([]);
        } else {
          setDiscenti(discentiData);
          showSuccess(`Caricamento completato. Trovati ${discentiData.length} discenti validi.`);
        }
      } catch (error) {
        console.error("Errore imprevisto durante la lettura del file Excel:", error);
        showError("Errore durante l'elaborazione del file. Assicurati che sia un file .xlsx valido e non corrotto.");
        setDiscenti([]);
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
    if (!wordFile || discenti.length === 0) {
      showError("Per favore, carica il file Excel con i dati e il template Word.");
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
            titolocorso: discente.Titolocorso,
            categoria: discente.Categoria,
            localita: discente.Localita,
            periodo_corso: `dal ${discente.DataInizio} al ${discente.DataFine}`,
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
                Il file deve avere le intestazioni alla riga 1 (es. matricola, grado, cognome, nome, cat, sede, dal, al, corso).
              </p>
            </div>
            <div className="space-y-2">
              <Label htmlFor="word-file">Carica Template Word (.docx)</Label>
              <Input id="word-file" type="file" accept=".docx" onChange={handleWordUpload} />
              {wordFile && <p className="text-sm text-muted-foreground">Caricato: {wordFile.name}</p>}
               <p className="text-xs text-muted-foreground pt-2">
                Il template deve contenere i segnaposto come {`{corso}`}, {`{categoria}`}, {`{sede}`}, ecc.
              </p>
            </div>
          </CardContent>
        </Card>

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
            <p className="text-sm text-muted-foreground mt-4">La firma sarà: Col. Massimiliano Fortino</p>
          </CardContent>
        </Card>

        {discenti.length > 0 && (
          <Card className="md:col-span-2">
            <CardHeader>
              <CardTitle className="flex items-center gap-2"><Eye size={20} /> Anteprima Dati da Excel</CardTitle>
            </CardHeader>
            <CardContent>
              <p className="mb-4"><strong>Numero di Discenti:</strong> <span className="font-mono p-1 bg-muted rounded-md">{discenti.length}</span></p>
              <div className="max-h-80 overflow-y-auto rounded-md border">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Matricola</TableHead>
                      <TableHead>Grado</TableHead>
                      <TableHead>Cognome e Nome</TableHead>
                      <TableHead>Categoria</TableHead>
                      <TableHead>Corso</TableHead>
                      <TableHead>Sede</TableHead>
                      <TableHead>Dal</TableHead>
                      <TableHead>Al</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {discenti.map((d, index) => (
                      <TableRow key={index}>
                        <TableCell>{d.Matricola}</TableCell>
                        <TableCell>{d["Grado militare"]}</TableCell>
                        <TableCell>{d["Cognome e Nome del Discente"]}</TableCell>
                        <TableCell>{d.Categoria}</TableCell>
                        <TableCell>{d.Titolocorso}</TableCell>
                        <TableCell>{d.Localita}</TableCell>
                        <TableCell>{d.DataInizio}</TableCell>
                        <TableCell>{d.DataFine}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
              <p className="text-sm text-muted-foreground mt-2">
                Controlla che i dati corrispondano.
              </p>
            </CardContent>
          </Card>
        )}

        <div className="md:col-span-2 flex justify-center">
          <Button size="lg" onClick={handleGenerateDocument} className="w-full md:w-1/2 lg:w-1/3">
            <Download className="mr-2 h-5 w-5" />
            Genera Documenti
          </Button>
        </div>
      </div>
    </div>
  );
}

export default Dashboard;
// ... (mantieni tutte le importazioni iniziali uguali)

function Dashboard() {
  // ... (mantieni tutti gli stati iniziali uguali)

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

        // ... (mantieni l'estrazione dei dati del corso uguale)

        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        // ... (mantieni il controllo della riga 7 uguale)

        // Trova dinamicamente la riga delle intestazioni
        let headerRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i].map(cell => String(cell).toLowerCase().trim());
            const hasMatricola = row.some(cell => cell.includes('matricola'));
            const hasCognome = row.some(cell => cell.includes('cognome'));
            const hasNome = row.some(cell => cell.includes('nome'));
            const hasGrado = row.some(cell => cell.includes('grado'));

            if (hasMatricola && (hasCognome || hasGrado)) {
                headerRowIndex = i;
                break;
            }
        }

        if (headerRowIndex === -1) {
            showError("Impossibile trovare la riga delle intestazioni. Assicurati che il file Excel contenga colonne come 'matricola', 'grado' e 'cognome'.");
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
        
        const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo']);
        const cognomeIndex = findIndex(['cognome']);
        const nomeIndex = findIndex(['nome']);

        const valueExists = (val: any) => val !== null && val !== undefined && String(val).trim() !== '';

        const discentiData = dataRows.map((row, rowIndex) => {
          if (row.every(cell => !valueExists(cell))) {
            return null;
          }

          let cognomeNome: string = '';
          let cognome: string = '';
          let nome: string = '';

          // Se ci sono colonne separate per cognome e nome
          if (cognomeIndex !== -1 && nomeIndex !== -1) {
            cognome = valueExists(row[cognomeIndex]) ? String(row[cognomeIndex]) : '';
            nome = valueExists(row[nomeIndex]) ? String(row[nomeIndex]) : '';
            cognomeNome = `${cognome} ${nome}`.trim();
          } 
          // Se c'è una colonna unica "Cognome e Nome"
          else if (cognomeNomeIndex !== -1 && valueExists(row[cognomeNomeIndex])) {
            const fullName = String(row[cognomeNomeIndex]).trim();
            const parts = fullName.split(' ');
            
            if (parts.length > 1) {
              // Prendi l'ultima parte come nome, il resto come cognome
              nome = parts.pop() || '';
              cognome = parts.join(' ');
            } else {
              cognome = fullName;
              nome = '';
            }
            cognomeNome = fullName;
          }

          const matricola = matricolaIndex !== -1 ? row[matricolaIndex] : undefined;
          const grado = gradoIndex !== -1 ? row[gradoIndex] : undefined;
          const categoria = categoriaIndex !== -1 && valueExists(row[categoriaIndex]) ? row[categoriaIndex] : "";

          if (!valueExists(matricola) || !valueExists(grado) || !valueExists(cognomeNome)) {
            console.warn(`Riga ${headerRowIndex + rowIndex + 2} del file Excel saltata perché mancano dati essenziali. Dati letti:`, { matricola, grado, cognomeNome });
            return null;
          }

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
          showSuccess(`Caricamento completato. Trovati ${discentiData.length} discenti e dati del corso.`);
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

  // ... (mantieni il resto del codice uguale, compresa la parte di render)
}

export default Dashboard;
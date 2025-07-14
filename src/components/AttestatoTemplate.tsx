import React from 'react';

interface Discente {
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
  Matricola: string;
  Cognome: string;
  Nome: string;
  Categoria: string;
}

interface CourseInfo {
  title: string;
  location: string;
  period: string;
  currentDate: string;
}

interface AttestatoTemplateProps {
  discente: Discente;
  courseInfo: CourseInfo;
  signer: string;
  categoria: string;
}

export const AttestatoTemplate: React.FC<AttestatoTemplateProps> = ({ discente, courseInfo, signer, categoria }) => {
  // L'uso di stili inline è fondamentale per garantire che html2pdf.js esegua il rendering corretto.
  return (
    <div style={{
      width: '210mm',
      height: '290mm', // Leggermente inferiore a 297mm per tenere conto dei margini
      padding: '20mm',
      fontFamily: "'Times New Roman', Times, serif",
      fontSize: '12pt',
      display: 'flex',
      flexDirection: 'column',
      boxSizing: 'border-box',
    }}>
      <header style={{ textAlign: 'center', flexShrink: 0 }}>
        <h1 style={{ fontSize: '16pt', fontWeight: 'bold', margin: 0, marginBottom: '20mm' }}>
          {courseInfo.title}
        </h1>
        <h2 style={{ fontSize: '18pt', fontWeight: 'bold', margin: 0, marginBottom: '20mm' }}>
          ATTESTATO DI FREQUENZA
        </h2>
      </header>

      <main style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
        <p style={{ textAlign: 'center', fontSize: '14pt', lineHeight: '1.8' }}>
          Si attesta che {categoria}
          <br />
          <strong style={{ textTransform: 'uppercase' }}>{discente["Grado militare"]} {discente["Cognome e Nome del Discente"]}</strong>
          <br />
          ha frequentato con profitto il corso in oggetto,
          <br />
          svoltosi a {courseInfo.location} {courseInfo.period}.
        </p>
      </main>

      <footer style={{ flexShrink: 0, marginTop: 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end' }}>
          <div style={{ fontSize: '11pt' }}>
            <p>Matricola: {discente.Matricola}</p>
            <p>{courseInfo.location}, {courseInfo.currentDate}</p>
          </div>
          <div style={{ textAlign: 'center', fontSize: '12pt' }}>
            <p style={{ marginBottom: '2px' }}>{signer}</p>
            <p style={{ margin: 0, borderTop: '1px solid black', paddingTop: '2px' }}>Col. Massimiliano Fortino</p>
          </div>
        </div>
      </footer>
    </div>
  );
};
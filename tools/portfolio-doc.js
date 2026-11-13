// One-page Word file that points to the live portfolio (for OneDrive / submissions).
// Needs portfolio-top.png + portfolio-qr.png in the folder given as argv[2] (made by shots.mjs + qrcode).
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, ImageRun, ExternalHyperlink, BorderStyle, Table, TableRow, TableCell, WidthType, VerticalAlign } = require('docx');

const dir = process.argv[2];
const URL = 'https://earldecastro-portfolio.vercel.app';
const F = 'Segoe UI', FS = 'Segoe UI Semibold', MONO = 'Consolas';
const INK = '171717', BODY = '4D4D4D', MUTE = '6B6B6B', MINT = '2DD4BF';
const t = (text, o = {}) => new TextRun({ text, font: F, color: BODY, ...o });
const link = (text, url, o = {}) => new ExternalHyperlink({ link: url, children: [t(text, { color: '0D7C66', underline: {}, ...o })] });
const W = 12240 - 2300; // content width (twips)
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

const shot = fs.readFileSync(path.join(dir, 'portfolio-top.png'));
const qr = fs.readFileSync(path.join(dir, 'portfolio-qr.png'));

const doc = new Document({
  styles: { default: { document: { run: { font: F, size: 21, color: BODY } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 900, left: 1150, right: 1150 } } }, children: [
    new Paragraph({ spacing: { after: 60 }, children: [t('DIGITAL PORTFOLIO', { font: MONO, size: 16, color: MUTE, characterSpacing: 20 })] }),
    new Paragraph({ spacing: { after: 40 }, children: [t('Earl Dion I. De Castro', { font: FS, size: 48, color: INK, characterSpacing: -14 })] }),
    new Paragraph({ spacing: { after: 0 }, children: [t('BS Information Technology (4th year) · STI College San Jose del Monte', { size: 21 })] }),
    new Paragraph({ spacing: { before: 120, after: 280 }, indent: { right: W - 820 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: MINT, space: 1 } }, children: [] }),
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [W - 2400, 2400],
      borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: W - 2400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: none, bottom: none, left: none, right: none }, children: [
          new Paragraph({ spacing: { after: 80 }, children: [t('Open my portfolio', { font: FS, size: 26, color: INK })] }),
          new Paragraph({ spacing: { after: 160 }, children: [link(URL, URL, { size: 24 })] }),
          new Paragraph({ spacing: { after: 60, line: 288 }, children: [t('Click the link or scan the QR code with your phone. The portfolio has my work (SaLitA), skills, résumé and cover letter, which can be read online or downloaded as Word files and HD images.', { size: 20 })] }),
        ] }),
        new TableCell({ width: { size: 2400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: none, bottom: none, left: none, right: none }, children: [
          new Paragraph({ alignment: 'right', children: [new ImageRun({ type: 'png', data: qr, transformation: { width: 130, height: 130 }, altText: { title: 'QR code', description: 'QR code to ' + URL, name: 'qr' } })] }),
        ] }),
      ] })],
    }),
    new Paragraph({ spacing: { before: 360, after: 120 }, children: [t('PREVIEW', { font: MONO, size: 16, color: MUTE, characterSpacing: 20 })] }),
    new Paragraph({ children: [new ExternalHyperlink({ link: URL, children: [new ImageRun({ type: 'png', data: shot, transformation: { width: 620, height: 388 }, altText: { title: 'Portfolio preview', description: 'Home page of the portfolio', name: 'preview' } })] })] }),
    new Paragraph({ spacing: { before: 280, after: 60 }, children: [t('Also included: ', { font: FS, color: INK, size: 20 }), t('Résumé and Cover Letter (PNG, HD) · Contact: 0939 390 3029 · earldecastro333@gmail.com', { size: 20 })] }),
    new Paragraph({ children: [t('Live product: ', { font: FS, color: INK, size: 20 }), link('salita.me', 'https://salita.me', { size: 20 })] }),
  ] }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(path.join(dir, 'Earl-De-Castro-Digital-Portfolio.docx'), b); console.log('done'); });

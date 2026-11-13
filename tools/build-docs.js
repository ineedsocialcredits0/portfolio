// One source of truth -> resume.html, cover-letter.html (portfolio) + .docx files.
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, TabStopType, BorderStyle, ExternalHyperlink } = require('docx');

const PORTFOLIO = 'https://earldecastro-portfolio.vercel.app';
const ME = {
  name: 'Earl Dion I. De Castro',
  headline: 'Aspiring Marketing Manager · BS Information Technology (4th year)',
  place: 'San Jose del Monte, Bulacan',
  phone: '0939 390 3029', tel: '+639393903029',
  email: 'earldecastro333@gmail.com',
  portfolio: PORTFOLIO, portfolioText: 'earldecastro-portfolio.vercel.app',
};
const R = {
  summary: '4th-year BS Information Technology student who launched SaLitA, a web and Android app, single-handedly from idea to public release, including its brand, website, landing page and product message. Brings hands-on knowledge of Android phones and apps, clear writing for Filipino customers in English and Filipino, and the habit of using data and feedback to improve results.',
  project: {
    title: 'SaLitA: Product Launch & Digital Presence (Capstone, solo)', when: '2026',
    sub: 'AI speech-coaching app for Filipino students · Website + Android app', link: 'salita.me', url: 'https://salita.me',
    bullets: [
      'Planned and launched a product from idea to public release: its own website (salita.me, custom domain) and a downloadable Android app.',
      'Created the brand identity: name, logo and app icon, color palette and a friendly, non-judgmental tone that sets it apart from similar school projects.',
      'Wrote and designed the landing page: product slideshow, feature tour, FAQ, "get started" steps and an app download section, in clear English with Filipino headlines ("Tara, usap tayo!").',
      'Produced the visual content: screenshots of every screen in phone and desktop versions for the website, the in-app tour and documentation.',
      'Tracked user activity with dashboards and charts, and turned feedback from users, my adviser and the evaluation panel into improvements.',
      'Kept users informed through in-app announcements ("What\'s new"), notifications and a support ticket system.',
      'Presented and defended the product before a panel of evaluators.',
    ],
  },
  skills: [
    ['Marketing & communication', 'Product launch, brand identity, website and landing page copy, product presentation, customer support, presenting to an audience'],
    ['Digital', 'Website building, Android phones and apps, visual content (screenshots, product images), data dashboards'],
    ['Technical', 'React, JavaScript, Python, SQL, Supabase, Vercel, Git/GitHub'],
    ['Languages', 'English, Filipino'],
  ],
  education: {
    degree: 'Bachelor of Science in Information Technology', when: '2023 – 2027 (expected)',
    school: 'STI College San Jose del Monte, Bulacan',
    note: 'Capstone: SaLitA, A Translanguaging-Based AI Speech Coaching System for Supporting Learners with Language Barriers',
  },
};
const C = {
  date: 'November 13, 2026',
  to: ['The Hiring Manager', 'Erika Phone Shop', 'Altaraza, San Jose del Monte, Bulacan'],
  greet: 'Dear Hiring Manager,',
  paras: [
    'I am writing to apply for the Marketing Manager position at Erika Phone Shop. I am a 4th-year Bachelor of Science in Information Technology student at STI College San Jose del Monte, graduating in 2027, and I live right here in San Jose del Monte. I would love to help more people in our area discover your shop.',
    'For my capstone, I launched a product called SaLitA from an idea to a public release on my own. Beyond building the app, I created its brand (name, logo, colors and tone), wrote and designed its website and landing page, produced the product images, and released it both online at salita.me and as a downloadable Android app. I used dashboards and feedback from users and evaluators to keep improving it, and I presented it before a panel. That experience taught me how to explain a product simply and make people want to try it.',
    'Because I study IT and build Android apps, I know phones well: specs, features and the apps people use every day. I can explain them to customers in plain English or Filipino, and I can help Erika Phone Shop build a stronger online presence through product photos, posts and promotions that bring customers through the door.',
    'Thank you for considering my application. You can see my work, résumé and this letter at ' + ME.portfolioText + '. I would be glad to talk about how I can help grow Erika Phone Shop, and I can be reached at ' + ME.phone + ' or ' + ME.email + '.',
  ],
  close: 'Sincerely,',
};

// ---------- DOCX (same look as the web pages: DESIGN.md system) ----------
const F = 'Segoe UI', FS = 'Segoe UI Semibold', MONO = 'Consolas';
const INK = '171717', BODY = '4D4D4D', MUTE = '6B6B6B', LINE = 'EBEBEB', MINT = '2DD4BF', LINK = '0D7C66';
const page = { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 900, left: 1150, right: 1150 } };
const CONTENT = 12240 - 2300;
const t = (text, o = {}) => new TextRun({ text, font: F, color: BODY, ...o });
const link = (text, url, o = {}) => new ExternalHyperlink({ link: url, children: [t(text, { color: INK, underline: { color: 'A1A1A1' }, ...o })] });
const nameP = (size) => new Paragraph({ spacing: { after: 40 }, children: [t(ME.name, { font: FS, size, color: INK, characterSpacing: -14 })] });
const contactP = () => new Paragraph({ spacing: { before: 100, after: 0 }, children: [
  t(ME.place, { size: 18, color: MUTE }), t('     '), link(ME.phone, 'tel:' + ME.tel, { size: 18 }), t('     '),
  link(ME.email, 'mailto:' + ME.email, { size: 18 }), t('     '), link(ME.portfolioText, ME.portfolio, { size: 18 })] });
// short mint→sky accent line under the header (a bottom border on a narrow paragraph)
const accent = () => new Paragraph({ spacing: { before: 120, after: 0 }, indent: { right: CONTENT - 820 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: MINT, space: 1 } }, children: [] });
const label = s => new Paragraph({ spacing: { before: 300, after: 100 }, keepNext: true, children: [t(s.toUpperCase(), { font: MONO, size: 16, color: MUTE, characterSpacing: 20 })] });
const row = (l, r) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: CONTENT }], spacing: { before: 40, after: 20 }, keepNext: true,
  children: [t(l, { font: FS, size: 21, color: INK }), t('\t' + r, { font: MONO, size: 17, color: MUTE })] });
const line = (s, o = {}) => new Paragraph({ spacing: { after: 30, line: 276 }, children: [t(s, { size: 19, ...o })] });

const resumeDoc = new Document({
  styles: { default: { document: { run: { font: F, size: 20, color: BODY } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
    style: { run: { color: 'A1A1A1' }, paragraph: { indent: { left: 340, hanging: 220 } } } }] }] },
  sections: [{ properties: { page }, children: [
    nameP(48),
    new Paragraph({ spacing: { after: 0 }, children: [t(ME.headline, { size: 21 })] }),
    contactP(), accent(),
    label('Summary'), line(R.summary, { size: 20 }),
    label('Relevant experience'),
    row(R.project.title, R.project.when),
    new Paragraph({ spacing: { after: 60 }, children: [t(R.project.sub + '  \u00b7  ', { size: 18, color: MUTE }), link(R.project.link, R.project.url, { size: 18, color: LINK })] }),
    ...R.project.bullets.map(b => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 50, line: 264 }, children: [t(b, { size: 19 })] })),
    label('Skills'),
    ...R.skills.map(([k, v]) => new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: 2700 }], indent: { left: 2700, hanging: 2700 }, spacing: { after: 70, line: 264 },
      children: [t(k, { font: FS, size: 19, color: INK }), t('\t' + v, { size: 19 })] })),
    label('Education'),
    row(R.education.degree, R.education.when), line(R.education.school, { color: BODY }), line(R.education.note, { italics: true, color: MUTE, size: 18 }),
  ] }],
});
const P = (s, after = 200, o = {}) => new Paragraph({ spacing: { after, line: 300 }, children: [t(s, { size: 21, color: INK, ...o })] });
const coverDoc = new Document({
  styles: { default: { document: { run: { font: F, size: 21, color: INK } } } },
  sections: [{ properties: { page: { ...page, margin: { top: 1100, bottom: 1000, left: 1250, right: 1250 } } }, children: [
    nameP(48), contactP(), accent(),
    new Paragraph({ spacing: { before: 240, after: 360 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 1 } }, children: [] }),
    P(C.date, 260),
    ...C.to.map((s, i) => P(s, i === C.to.length - 1 ? 260 : 0)),
    P(C.greet), ...C.paras.map(s => P(s)),
    P(C.close, 560), P(ME.name, 0, { font: FS }),
  ] }],
});

// ---------- HTML (uses the portfolio's styles.css / DESIGN.md system) ----------
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const shell = (title, docx, current, body) => `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title} · Earl Dion De Castro</title>
<link rel="icon" href="images/favicon.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono&display=swap" rel="stylesheet">
<link rel="stylesheet" href="styles.css"></head><body>
<nav class="nav"><div class="wrap">
  <a class="logo" href="index.html"><span class="mark" aria-hidden="true"></span>Earl Dion De Castro</a>
  <ul><li><a href="index.html#about">About</a></li><li><a href="index.html#work">Work</a></li><li><a href="index.html#skills">Skills</a></li>
  <li><a href="resume.html"${current === 'resume' ? ' aria-current="page"' : ''}>Résumé</a></li><li><a href="cover-letter.html"${current === 'cover' ? ' aria-current="page"' : ''}>Cover letter</a></li><li><a href="index.html#contact">Contact</a></li></ul>
  <a class="btn btn-primary btn-sm" href="${docx}" download>Download .docx</a>
</div></nav>
<div class="doc-bar"><a class="btn btn-secondary btn-sm" href="index.html">← Portfolio</a>
<div class="btn-row"><button class="btn btn-secondary btn-sm" onclick="print()">Print / Save PDF</button><a class="btn btn-secondary btn-sm" href="${docx.replace('.docx', '.png')}" download>Download image (HD)</a><a class="btn btn-primary btn-sm" href="${docx}" download>Download .docx</a></div></div>
<main class="paper${current === 'cover' ? ' letter' : ''}">${body}</main></body></html>`;
const contactHtml = `<p class="contact"><span>${esc(ME.place)}</span><a href="tel:${ME.tel}">${ME.phone}</a><a href="mailto:${ME.email}">${ME.email}</a><a href="${ME.portfolio}">${ME.portfolioText}</a></p>`;

const resumeHtml = shell('Résumé', 'Earl-De-Castro-Resume.docx', 'resume', `
<h1>${esc(ME.name)}</h1><p class="role">${esc(ME.headline)}</p>${contactHtml}<div class="accent"></div>
<h2>Summary</h2><p>${esc(R.summary)}</p>
<h2>Relevant experience</h2>
<div class="row"><span>${esc(R.project.title)}</span><span>${R.project.when}</span></div>
<div class="meta">${esc(R.project.sub)} · <a href="${R.project.url}" target="_blank" rel="noopener">${R.project.link}</a></div>
<ul>${R.project.bullets.map(b => `<li>${esc(b)}</li>`).join('')}</ul>
<h2>Skills</h2><dl class="skills">${R.skills.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('')}</dl>
<h2>Education</h2>
<div class="row"><span>${esc(R.education.degree)}</span><span>${R.education.when}</span></div>
<div class="meta">${esc(R.education.school)}</div><div class="meta"><i>${esc(R.education.note)}</i></div>`);

const coverHtml = shell('Cover letter', 'Earl-De-Castro-Cover-Letter.docx', 'cover', `
<h1>${esc(ME.name)}</h1>${contactHtml}<div class="accent"></div><hr style="margin-top:24px">
<p>${C.date}</p><p>${C.to.map(esc).join('<br>')}</p>
<p>${esc(C.greet)}</p>${C.paras.map(s => `<p>${esc(s)}</p>`).join('')}
<p style="margin-top:28px">${C.close}</p><p style="margin-top:40px;font-weight:600">${esc(ME.name)}</p>`);

const [portfolioDir, downloads] = process.argv.slice(2);
fs.writeFileSync(path.join(portfolioDir, 'resume.html'), resumeHtml);
fs.writeFileSync(path.join(portfolioDir, 'cover-letter.html'), coverHtml);
Promise.all([[resumeDoc, 'Earl-De-Castro-Resume.docx'], [coverDoc, 'Earl-De-Castro-Cover-Letter.docx']].map(([d, n]) =>
  Packer.toBuffer(d).then(b => { fs.writeFileSync(path.join(portfolioDir, n), b); fs.writeFileSync(path.join(downloads, n), b); })
)).then(() => console.log('done'));

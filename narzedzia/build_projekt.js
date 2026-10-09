// Generator projektu: Toyota Corolla IX (E12) – wymiana przednich klocków hamulcowych
// Użycie: node build_projekt.js <wyjście.docx> [strony.json]
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, VerticalAlign, HeadingLevel, Footer,
  PageNumber, LevelFormat, PageBreak, HeightRule, ImageRun, TabStopType,
} = require('docx');

const OUT = process.argv[2];
const PAGES = process.argv[3] && fs.existsSync(process.argv[3]) ? JSON.parse(fs.readFileSync(process.argv[3])) : {};
const FONT = 'Calibri';
const PW = 11906 - 2 * 1134; // szerokość tekstu A4, marginesy 2 cm = 9638 DXA
const AUTOR = 'Jakub Suszek';

// ---------- pomocnicze ----------
const B = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDERS = { top: B, bottom: B, left: B, right: B };
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOB = { top: NONE, bottom: NONE, left: NONE, right: NONE };

function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(p => new TextRun({
    text: p.startsWith('**') ? p.slice(2, -2) : p,
    bold: p.startsWith('**') || o.bold, italics: o.italics, font: FONT, size: o.size || 22, color: o.color,
  }));
}
function P(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { before: o.before || 0, after: o.after === undefined ? 120 : o.after, line: o.line || 276 },
    indent: o.indent, keepNext: o.keepNext,
    children: runs(text, o),
  });
}
const C = (t, o = {}) => P(t, { align: AlignmentType.CENTER, ...o });
function H1(t) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { before: 0, after: 240 },
    children: [new TextRun({ text: t, bold: true, font: FONT, size: 30, color: '000000' })] });
}
function H1nb(t) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 240 },
    children: [new TextRun({ text: t, bold: true, font: FONT, size: 30, color: '000000' })] });
}
function H2(t) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text: t, bold: true, font: FONT, size: 26, color: '000000' })] });
}
const SUB = t => P(t, { after: 80, keepNext: true, align: AlignmentType.LEFT });
function BUL(t) {
  return new Paragraph({ numbering: { reference: 'bul', level: 0 }, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 60, line: 276 }, children: runs(t) });
}
function CAP(t) { return P(t, { size: 18, bold: true, after: 200, align: AlignmentType.LEFT }); }
function TCAP(t) { return P(t, { size: 18, bold: true, after: 80, align: AlignmentType.LEFT, keepNext: true }); }

function cell(content, width, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c =>
    typeof c === 'string' ? P(c, { size: o.size || 20, bold: o.bold, align: o.align || AlignmentType.LEFT, after: 0, line: 240 }) : c);
  return new TableCell({
    borders: o.borders || BORDERS, width: { size: width, type: WidthType.DXA },
    columnSpan: o.colSpan, rowSpan: o.rowSpan,
    verticalAlign: o.vAlign || VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: paras,
  });
}
function table(widths, rows, o = {}) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: r.map((c, j) => cell(c, widths[j], {
        bold: i === 0, fill: i === 0 ? (o.fill || 'C9D7EE') : undefined, size: o.size || 20,
        align: i === 0 || (o.center || []).includes(j) ? AlignmentType.CENTER : AlignmentType.LEFT,
      })),
    })),
  });
}
function photoBox(w, h, hint) {
  return new Table({
    width: { size: w, type: WidthType.DXA }, columnWidths: [w], alignment: AlignmentType.CENTER,
    rows: [new TableRow({ height: { value: h, rule: HeightRule.ATLEAST }, children: [new TableCell({
      width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
      borders: { top: { style: BorderStyle.DASHED, size: 6, color: '8C8C8C' }, bottom: { style: BorderStyle.DASHED, size: 6, color: '8C8C8C' },
        left: { style: BorderStyle.DASHED, size: 6, color: '8C8C8C' }, right: { style: BorderStyle.DASHED, size: 6, color: '8C8C8C' } },
      shading: { fill: 'F2F2F2', type: ShadingType.CLEAR, color: 'auto' },
      children: [C('MIEJSCE NA ZDJĘCIE', { size: 18, bold: true, color: '808080', after: 40 }),
        C(hint, { size: 16, italics: true, color: '808080', after: 0 })],
    })] })],
  });
}

// ---------- numeracja rysunków i tabel ----------
let FIG = 0;
const fig = () => ++FIG;

// ---------- STRONA TYTUŁOWA ----------
const title = [
  C('[NAZWA UCZELNI]', { size: 32, bold: true, before: 600, after: 120 }),
  C('[Wydział]', { size: 28, after: 60 }),
  C('[Zakład / Katedra]', { size: 26, after: 2000 }),
  C('Projekt', { size: 44, bold: true, after: 400 }),
  C('Temat: Proces technologiczny weryfikacji i wymiany przednich klocków hamulcowych wraz z kontrolą jakości naprawy w pojeździe Toyota Corolla IX (E12) 1.6 VVT-i', { size: 30, after: 2400 }),
  P(`Wykonał: ${AUTOR}`, { size: 24, align: AlignmentType.RIGHT, after: 80 }),
  P('Nr albumu: ..................', { size: 24, align: AlignmentType.RIGHT, after: 80 }),
  P('Prowadzący: ..................', { size: 24, align: AlignmentType.RIGHT, after: 1400 }),
  C('Warszawa, 2026', { size: 24, after: 0 }),
];

// ---------- SPIS TREŚCI (strony uzupełniane w 2. przebiegu) ----------
const TOC = [
  ['1. Wstęp', 0],
  ['2. Charakterystyka obiektu naprawy', 0], ['2.1. Dane techniczne pojazdu', 1], ['2.2. Budowa hamulca tarczowego osi przedniej', 1],
  ['3. Warunki techniczne weryfikacji i montażu', 0], ['3.1. Wymiary nominalne i dopuszczalne', 1],
  ['3.2. Momenty dokręcania oraz rozmiar i typ kluczy', 1], ['3.3. Środki smarne i eksploatacyjne', 1],
  ['4. Diagnostyka i kwalifikacja pojazdu do naprawy', 0], ['4.1. Macierz diagnostyczna objawów', 1], ['4.2. Karta weryfikacji części', 1],
  ['5. Wyposażenie stanowiska pracy', 0],
  ['6. Proces technologiczny wymiany klocków hamulcowych', 0], ['6.1. Schemat procesu technologicznego', 1], ['6.2. Karta instrukcyjna operacji', 1],
  ['7. Kontrola jakości po naprawie', 0], ['7.1. Kontrola statyczna układu hamulcowego', 1],
  ['7.2. Badanie skuteczności hamowania na stanowisku rolkowym', 1], ['7.3. Protokół kontroli jakości', 1], ['7.4. Jazda próbna i docieranie klocków', 1],
  ['8. Typowe błędy wykonawcze i ich skutki', 0],
  ['9. Analiza kosztów naprawy', 0], ['9.1. Warianty doboru części', 1], ['9.2. Koszt robocizny', 1], ['9.3. Zestawienie kosztów', 1],
  ['10. Wnioski', 0], ['Bibliografia', 0],
];
const tocPage = [
  P('Spis treści', { size: 30, bold: true, align: AlignmentType.LEFT, after: 240, before: 0 }),
  ...TOC.map(([t, lvl]) => new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: PW, leader: 'dot' }],
    indent: { left: lvl ? 400 : 0 }, spacing: { after: 70 },
    children: [new TextRun({ text: t, font: FONT, size: 22, bold: !lvl }),
      new TextRun({ text: `\t${PAGES[t] || '00'}`, font: FONT, size: 22 })],
  })),
];

// tabela z wierszami grup (wiersz grupy = scalona komórka z pogrubionym tytułem)
function grpTable(widths, head, groups, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const rows = [new TableRow({ tableHeader: true, cantSplit: true,
    children: head.map((h, j) => cell(h, widths[j], { bold: true, fill: 'C9D7EE', size: o.size || 19, align: AlignmentType.CENTER })) })];
  groups.forEach(([g, items]) => {
    rows.push(new TableRow({ cantSplit: true, children: [cell(g, total, { colSpan: widths.length, bold: true, fill: 'EEF2F8', size: o.size || 19 })] }));
    items.forEach(r => rows.push(new TableRow({ cantSplit: true,
      children: r.map((c, j) => cell(c, widths[j], { size: o.size || 19, align: (o.center || []).includes(j) ? AlignmentType.CENTER : AlignmentType.LEFT })) })));
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows });
}

// ---------- 1. WSTĘP ----------
const wstep = [
  H1('1. Wstęp'),
  P('Hamulec tarczowy osi przedniej przenosi w samochodzie osobowym większą część sił hamowania, a jego elementy cierne zużywają się w sposób ciągły podczas eksploatacji. Klocki hamulcowe są więc częścią, której stan decyduje bezpośrednio o drodze hamowania i stabilności pojazdu. Ich wymiana jest zabiegiem krótkim, ale – jak wykazano w rozdziale 8 – podatnym na błędy wykonawcze, których skutki ujawniają się dopiero podczas jazdy.'),
  P('**Celem projektu** jest opracowanie kompletnego procesu technologicznego weryfikacji i wymiany przednich klocków hamulcowych dla konkretnego pojazdu – Toyoty Corolli IX (E12) hatchback 1.6 VVT-i (silnik 3ZZ-FE, 81 kW, kod typu ZZE121) – zakończonego mierzalną kontrolą jakości naprawy.'),
  P('Zakres opracowania obejmuje:', { keepNext: true, after: 60 }),
  BUL('charakterystykę obiektu naprawy i wykaz elementów zespołu wraz ze sposobem postępowania z każdym z nich,'),
  BUL('warunki techniczne weryfikacji i montażu: wymiary nominalne i dopuszczalne, momenty dokręcania, rozmiary i typy kluczy oraz środki smarne,'),
  BUL('macierz diagnostyczną objawów i kartę weryfikacji części, na podstawie których pojazd kwalifikowany jest do naprawy,'),
  BUL('wyposażenie stanowiska pracy z przypisaniem narzędzi do numerów czynności,'),
  BUL('schemat procesu technologicznego i kartę instrukcyjną operacji,'),
  BUL('kontrolę jakości po naprawie, w tym badanie na stanowisku rolkowym i protokół odbioru,'),
  BUL('analizę typowych błędów wykonawczych oraz porównanie kosztów trzech wariantów doboru części.'),
  P('**Źródła danych.** Wszystkie wymiary, luzy, bicia i momenty dokręcania przyjęto z instrukcji naprawy producenta Toyota Corolla (E120) Repair Manual [1], [4]–[6]. Producent nie podaje w niej rozmiarów kluczy ani dokumentacji fotograficznej, dlatego te informacje uzupełniono instrukcją wymiany krok po kroku przygotowaną dla tego modelu [2], [3]. W przypadku rozbieżności (np. moment śrub zacisku 34,3 N·m wg [1] i 35 N·m wg [2]) przyjmowano wartość producenta. Wymagania kontroli na stanowisku rolkowym oparto na kryteriach stosowanych w badaniach technicznych pojazdów [15]–[17], a dane katalogowe części i ceny – na katalogach i ofertach handlowych [11]–[14], [18], [19], [24], [25].'),
];

// ---------- 2. CHARAKTERYSTYKA OBIEKTU ----------
const f1 = fig(), f3 = fig(), f4 = fig();
const opisPojazdu = [
  H1('2. Charakterystyka obiektu naprawy'),
  H2('2.1. Dane techniczne pojazdu'),
  P('Obiektem naprawy jest samochód osobowy Toyota Corolla IX generacji (oznaczenie fabryczne E12/E120), produkowany w latach 2001–2007. Analizowany egzemplarz to 5-drzwiowy hatchback z benzynowym silnikiem 1.6 VVT-i o kodzie 3ZZ-FE i oznaczeniu typu ZZE121 [10]. Z punktu widzenia procesu najważniejsze są: dopuszczalna masa całkowita (wykorzystana w rozdziale 7 do obliczenia wymaganej siły hamowania), rodzaj hamulców osi przedniej oraz typ hamulca postojowego (decydujący o braku potrzeby użycia testera diagnostycznego).'),
  P(`Na rysunku numer ${f1} poniżej pokazano analizowany model pojazdu.`, { keepNext: true }),
  photoBox(7000, 3400, 'Toyota Corolla E12 hatchback – zdjęcie ogólne pojazdu (np. Wikimedia Commons, z podaniem źródła)'),
  CAP(`Rysunek ${f1}. Toyota Corolla IX (E12) hatchback, [źródło zdjęcia]`),
  TCAP('Tabela 1. Dane techniczne pojazdu Toyota Corolla IX (E12) 1.6 VVT-i, [7]–[10]'),
  grpTable([4300, 5338], ['Parametr', 'Wartość'], [
    ['Identyfikacja', [
      ['Model / nadwozie', 'Toyota Corolla IX (E12), hatchback 5-drzwiowy'],
      ['Kod typu / lata produkcji', 'ZZE121; 2001–2007 [10]'],
    ]],
    ['Jednostka napędowa i przeniesienie napędu', [
      ['Silnik', '3ZZ-FE, benzynowy, R4, 16 zaworów, DOHC, VVT-i, 1598 cm³'],
      ['Moc / moment obrotowy', '81 kW (110 KM) przy 6000 obr/min / 150 N·m przy 4800 obr/min'],
      ['Skrzynia biegów / napęd', 'manualna 5-biegowa / na koła przednie'],
      ['Prędkość maksymalna / 0–100 km/h', '190 km/h / 10,2 s'],
      ['Zużycie paliwa (miasto / trasa / średnie)', '9,0 / 5,9 / 7,0 l/100 km'],
    ]],
    ['Wymiary i masy', [
      ['Długość / szerokość / wysokość', '4180 / 1710 / 1475 mm'],
      ['Rozstaw osi', '2600 mm'],
      ['Masa własna / dopuszczalna masa całkowita', 'od 1115 kg / 1655 kg'],
      ['Zbiornik paliwa / opony', '55 l / 195/60 R15'],
    ]],
    ['Układ hamulcowy', [
      ['Hamulce przednie', 'tarczowe wentylowane Ø 255 mm, zacisk pływający jednotłoczkowy [1], [13]'],
      ['Hamulce tylne', 'bębnowe lub tarczowe pełne Ø 258 mm – zależnie od wersji [14]'],
      ['Hamulec postojowy', 'mechaniczny, cięgnowy, na koła tylne'],
      ['Układy wspomagające', 'serwo podciśnieniowe, ABS z EBD'],
      ['Płyn hamulcowy', 'SAE J1703 / FMVSS No. 116 DOT 3 [5]'],
    ]],
  ]),
  P('', { after: 60 }),
  H2('2.2. Budowa hamulca tarczowego osi przedniej'),
  P('Hamulec przedni jest hamulcem tarczowym z zaciskiem pływającym. Korpus cylindra z jednym tłoczkiem osadzony jest na dwóch sworzniach prowadzących, które przesuwają się w jarzmie przykręconym do zwrotnicy. Podczas hamowania tłoczek dociska klocek wewnętrzny, a reakcja przesuwa korpus na sworzniach i dociska klocek zewnętrzny. Klocki prowadzone są w jarzmie przez dwie płytki podporowe, a drgania tłumią podkładki przeciwpiskowe nr 1 i nr 2 [1]. Zużycie okładzin sygnalizuje mechanicznie płytka wskaźnika zużycia, zamocowana na klocku wewnętrznym [1], [11].'),
  P(`Na rysunku numer ${f3} poniżej pokazano hamulec przedni po zdjęciu koła, a na rysunku numer ${f4} – elementy składowe zacisku według instrukcji producenta.`, { keepNext: true }),
  photoBox(7000, 3200, 'Hamulec przedni Corolli E12 po zdjęciu koła – kadr z poradnika / filmu AUTODOC [2], [3]'),
  CAP(`Rysunek ${f3}. Hamulec tarczowy osi przedniej pojazdu Toyota Corolla E12, [2]`),
  photoBox(7000, 3600, 'Rysunek rozstrzelony zacisku z instrukcji Toyota Repair Manual – Front brake – Overhaul [1]'),
  CAP(`Rysunek ${f4}. Elementy składowe zacisku hamulca przedniego, [1]`),
  P('Sposób postępowania z poszczególnymi elementami zespołu podczas naprawy zestawiono w tabeli numer 2. Tabela określa, które części wymienia się zawsze, a które podlegają weryfikacji i ponownemu montażowi.'),
  TCAP('Tabela 2. Elementy zespołu hamulca przedniego i postępowanie podczas naprawy, [1], [2], [6]'),
  table([2700, 3000, 3938], [
    ['Element', 'Funkcja', 'Postępowanie podczas naprawy'],
    ['Klocki hamulcowe (2 szt. na koło) z płytką wskaźnika zużycia', 'Wytwarzanie siły tarcia, sygnalizacja zużycia', '**Wymiana obowiązkowa** – kompletem na oś [2]'],
    ['Podkładki przeciwpiskowe nr 1 i nr 2', 'Tłumienie drgań i pisków', 'Ponowny montaż po posmarowaniu lub wymiana kompletu [1]'],
    ['Płytki podporowe klocków (2 szt.)', 'Prowadzenie klocków w jarzmie', 'Weryfikacja: sprężystość, odkształcenia, pęknięcia, zużycie [1]'],
    ['Sworznie prowadzące z osłonami', 'Przesuw korpusu zacisku pływającego', 'Weryfikacja, czyszczenie i smarowanie [1]'],
    ['Korpus cylindra z tłoczkiem i osłoną', 'Docisk klocka wewnętrznego', 'Kontrola szczelności i osłony; tłoczek cofany [1]'],
    ['Jarzmo (mocowanie zacisku)', 'Przeniesienie momentu hamowania na zwrotnicę', 'Bez demontażu; czyszczenie miejsc osadzenia płytek'],
    ['Tarcza hamulcowa wentylowana', 'Powierzchnia cierna', 'Weryfikacja grubości i bicia [1], [6]'],
    ['Śruby zacisku (2 szt.)', 'Mocowanie korpusu do sworzni', 'Ponowne użycie, dokręcenie 34,3 N·m [1]'],
  ], { size: 19 }),
];

// ---------- 3. WARUNKI TECHNICZNE ----------
const uklad = [
  H1('3. Warunki techniczne weryfikacji i montażu'),
  P('Warunki techniczne określają wartości, z którymi porównuje się wyniki pomiarów podczas weryfikacji, oraz wymagania, które należy spełnić przy montażu. Wszystkie wartości liczbowe pochodzą z instrukcji naprawy producenta.'),
  H2('3.1. Wymiary nominalne i dopuszczalne'),
  TCAP('Tabela 3. Wymiary nominalne i dopuszczalne elementów podlegających weryfikacji, [1], [4], [6]'),
  table([3000, 1700, 1700, 1838, 1400], [
    ['Element / parametr', 'Wymiar nominalny', 'Wymiar dopuszczalny', 'Przyrząd pomiarowy', 'Źródło'],
    ['Grubość okładziny klocka przedniego', '11,0 mm', 'min. 1,0 mm', 'suwmiarka', '[1]'],
    ['Grubość tarczy przedniej', '25,0 mm', 'min. 23,0 mm', 'mikrometr', '[1]'],
    ['Bicie tarczy (10 mm od krawędzi)', '–', 'max 0,05 mm', 'czujnik zegarowy', '[6]'],
    ['Luz osiowy łożyska piasty przedniej', '–', 'max 0,05 mm', 'czujnik zegarowy', '[6]'],
    ['Bicie piasty przedniej', '–', 'max 0,05 mm', 'czujnik zegarowy', '[6]'],
    ['Płytki podporowe klocków', 'sprężyste, bez odkształceń', 'brak pęknięć i zużycia', 'oględziny', '[1]'],
    ['Wysokość pedału hamulca od podłogi', 'M/T 134,9–144,9 mm; A/T 136,0–146,0 mm', '–', 'przymiar', '[4]'],
    ['Luz swobodny pedału hamulca', '1–6 mm', '–', 'przymiar', '[4]'],
    ['Odległość rezerwowa pedału (490 N)', '–', '> 70 mm', 'przymiar', '[4]'],
  ], { center: [1, 2, 3, 4], size: 19 }),
  P('', { after: 60 }),
  H2('3.2. Momenty dokręcania oraz rozmiar i typ kluczy'),
  P('Producent podaje momenty dokręcania w N·m, kgf·cm i ft·lbf, nie podaje natomiast rozmiarów kluczy. Rozmiary i typy kluczy przyjęto z instrukcji krok po kroku dla analizowanego modelu [2].'),
  TCAP('Tabela 4. Momenty dokręcania oraz rozmiar i typ kluczy, [1], [2], [4]'),
  table([2900, 2600, 1250, 1600, 1288], [
    ['Połączenie', 'Rozmiar i typ klucza', 'Moment [N·m]', 'Moment [kgf·cm / ft·lbf]', 'Źródło'],
    ['Nakrętki koła (4 szt.)', 'nasadka udarowa 21 mm, klucz dynamometryczny', '103', '1050 / 76', '[1], [2]'],
    ['Śruby zacisku do sworzni prowadzących (2 szt.)', 'nasadka 13 mm + klucz dynamometryczny; kontra – klucz płasko-oczkowy 17 mm', '34,3', '350 / 25', '[1], [2]'],
    ['Śruby jarzma do zwrotnicy (2 szt.)*', '–', '106,8', '1089 / 79', '[1]'],
    ['Śruba przewodu elastycznego (banjo)*', '–', '29', '296 / 21', '[1]'],
    ['Nakrętka kontrująca popychacza pedału*', '–', '26', '265 / 19', '[4]'],
  ], { center: [2, 3, 4], size: 19 }),
  P('* połączenia nierozłączane w zakresie karty – podano dla kompletności (demontaż jarzma przy wymianie tarczy, odłączanie przewodu przy regeneracji zacisku, regulacja pedału).', { size: 18, italics: true, after: 120 }),
  H2('3.3. Środki smarne i eksploatacyjne'),
  P('Instrukcja producenta wskazuje różne środki dla poszczególnych miejsc smarowania. Zastosowanie niewłaściwego środka (np. smaru mineralnego na elementach gumowych) jest jednym z błędów opisanych w rozdziale 8.'),
  TCAP('Tabela 5. Środki smarne i eksploatacyjne stosowane w procesie, [1], [2], [5]'),
  table([3300, 3600, 1500, 1238], [
    ['Miejsce zastosowania', 'Środek', 'Nr czynności', 'Źródło'],
    ['Części ślizgowe i uszczelniające sworzni prowadzących', 'smar litowy na bazie glikolu (lithium soap base glycol grease)', '17', '[1]'],
    ['Podkładki przeciwpiskowe nr 1', 'smar do hamulców tarczowych', '20', '[1]'],
    ['Miejsca styku płyty nośnej klocka z jarzmem', 'pasta przeciwpiskowa', '20', '[2]'],
    ['Powierzchnia przylegania felgi do piasty', 'smar miedziany', '25', '[2]'],
    ['Tarcza, jarzmo', 'zmywacz do układów hamulcowych', '16, 21', '[2]'],
    ['Zapieczone połączenia gwintowe', 'środek penetrujący', '6', '[2]'],
    ['Układ hydrauliczny (uzupełnienie)', 'płyn SAE J1703 / FMVSS No. 116 DOT 3', '30', '[5]'],
  ], { center: [2, 3], size: 19 }),
  P('Na okładziny cierne i powierzchnie robocze tarczy nie wolno nanosić żadnych środków smarnych.', { size: 20, italics: true, before: 80 }),
];

// ---------- 4. DIAGNOSTYKA ----------
const f5 = fig();
const diag = [
  H1('4. Diagnostyka i kwalifikacja pojazdu do naprawy'),
  P('Toyota Corolla E12 nie ma elektronicznego czujnika zużycia klocków ani kontrolki na tablicy wskaźników. Kierowca dowiaduje się o zużyciu wyłącznie z objawów akustycznych i zmian w działaniu hamulców, dlatego decyzję o naprawie podejmuje się dwuetapowo: najpierw na podstawie objawów (macierz diagnostyczna), a następnie na podstawie pomiarów po demontażu (karta weryfikacji).'),
  H2('4.1. Macierz diagnostyczna objawów'),
  P('W tabeli numer 6 przyporządkowano każdemu objawowi możliwą przyczynę, sposób sprawdzenia oraz kryterium oceny wynikające z danych producenta lub przepisów.'),
  TCAP('Tabela 6. Macierz diagnostyczna objawów nieprawidłowej pracy hamulców osi przedniej, [1], [4]–[6], [15]'),
  table([2150, 2600, 2600, 2288], [
    ['Objaw', 'Możliwa przyczyna', 'Sposób sprawdzenia', 'Kryterium oceny'],
    ['Wysoki pisk przy hamowaniu, ustępujący po zwolnieniu pedału', 'Płytka wskaźnika zużycia ociera o tarczę', 'Pomiar okładziny suwmiarką', 'Okładzina ≥ 1,0 mm [1]'],
    ['Metaliczny zgrzyt przy hamowaniu', 'Starta okładzina, kontakt płyty nośnej z tarczą', 'Oględziny klocków, pomiar tarczy', 'Tarcza ≥ 23,0 mm [1]'],
    ['Drgania pedału lub kierownicy przy hamowaniu', 'Bicie tarczy lub jej nierównomierna grubość', 'Czujnik zegarowy 10 mm od krawędzi tarczy', 'Bicie ≤ 0,05 mm [6]'],
    ['Ściąganie pojazdu przy hamowaniu', 'Zapieczony sworzeń lub tłoczek, różne zużycie klocków L/P', 'Stanowisko rolkowe, kontrola przesuwu sworzni', 'Różnica sił na osi ≤ 30% [15]'],
    ['Wydłużony skok pedału', 'Zużyte okładziny, wyciek, zapowietrzenie', 'Pomiar odległości rezerwowej pedału', '> 70 mm przy 490 N [4]'],
    ['Różne zużycie klocka wewnętrznego i zewnętrznego', 'Utrudniony przesuw korpusu na sworzniach', 'Ręczna kontrola przesuwu, oględziny osłon', 'Swobodny przesuw, osłony bez pęknięć [1]'],
    ['Spadek poziomu płynu w zbiorniku', 'Zużycie okładzin (wysunięcie tłoczków) lub wyciek', 'Kontrola zbiornika i szczelności', 'Poziom MIN–MAX, brak wycieków [5]'],
  ], { size: 18 }),
  P('', { after: 60 }),
  H2('4.2. Karta weryfikacji części'),
  P('Na potrzeby projektu przyjęto przypadek, w którym użytkownik zgłosił pisk przy hamowaniu oraz lekkie ściąganie pojazdu w lewo. Zgodnie z macierzą diagnostyczną po demontażu wykonano pomiary, których **przykładowe wyniki (założenie projektowe)** wraz z kwalifikacją części zestawiono w tabeli numer 7. Kwalifikacja przyjmuje trzy kategorie: część dobra, część do naprawy (czyszczenie, smarowanie) i część do wymiany.'),
  TCAP('Tabela 7. Karta weryfikacji części hamulca przedniego – przykładowe wyniki pomiarów, opracowanie własne na podstawie [1], [6]'),
  table([2050, 1150, 1500, 1050, 1050, 2838], [
    ['Część / cecha', 'Wymiar nominalny', 'Wymiar dopuszczalny', 'Wynik – lewe', 'Wynik – prawe', 'Kwalifikacja'],
    ['Okładzina klocka wewnętrznego', '11,0 mm', 'min. 1,0 mm', '0,9 mm', '1,4 mm', '**Do wymiany** – lewy poniżej granicy; wymiana kompletu na oś'],
    ['Okładzina klocka zewnętrznego', '11,0 mm', 'min. 1,0 mm', '2,2 mm', '2,5 mm', '**Do wymiany** – w komplecie na oś'],
    ['Grubość tarczy', '25,0 mm', 'min. 23,0 mm', '24,2 mm', '24,4 mm', 'Dobra'],
    ['Bicie tarczy', '–', 'max 0,05 mm', '0,02 mm', '0,03 mm', 'Dobra'],
    ['Sworznie prowadzące', '–', 'swobodny przesuw', 'opór przesuwu', 'swobodny', '**Do naprawy** (lewe) – czyszczenie i smarowanie'],
    ['Osłony sworzni i tłoczka', '–', 'bez pęknięć, bez wycieku', 'bez uwag', 'bez uwag', 'Dobre'],
    ['Płytki podporowe', '–', 'bez odkształceń', 'bez uwag', 'bez uwag', 'Dobre – ponowny montaż'],
    ['Podkładki przeciwpiskowe', '–', 'bez deformacji', 'bez uwag', 'bez uwag', 'Dobre – ponowny montaż po posmarowaniu'],
  ], { center: [1, 2, 3, 4], size: 18 }),
  P('', { after: 60 }),
  P('Wyniki są ze sobą spójne: większe zużycie klocka wewnętrznego po stronie lewej oraz opór przesuwu lewych sworzni wskazują, że korpus zacisku nie przesuwał się swobodnie i klocek wewnętrzny pracował z większym dociskiem. Tłumaczy to zgłoszone ściąganie pojazdu. Tarcze mieszczą się w wymiarach dopuszczalnych, więc pojazd zakwalifikowano do wymiany klocków osi przedniej z naprawą (czyszczeniem i smarowaniem) sworzni prowadzących, bez wymiany tarcz. Skuteczność tej naprawy weryfikuje pomiar różnicy sił hamowania w rozdziale 7.'),
  P(`Na rysunku numer ${f5} poniżej pokazano zużyte klocki po demontażu zacisku.`, { keepNext: true }),
  photoBox(7000, 3200, 'Zużyte klocki po zdjęciu zacisku – kadr z poradnika / filmu AUTODOC [2], [3]'),
  CAP(`Rysunek ${f5}. Zużyte okładziny klocków hamulcowych osi przedniej, [2]`),
];

// ---------- 5. WYPOSAŻENIE STANOWISKA ----------
const narz = [
  H1('5. Wyposażenie stanowiska pracy'),
  P('Wyposażenie stanowiska dobrano na podstawie czynności karty instrukcyjnej. W tabeli numer 8 każde narzędzie i przyrząd przypisano do numerów czynności, w których jest używane, co pozwala przygotować stanowisko przed rozpoczęciem pracy i sprawdzić, że żadna czynność nie została bez potrzebnego narzędzia.'),
  TCAP('Tabela 8. Wyposażenie stanowiska pracy z przypisaniem do czynności karty instrukcyjnej, [1]–[3]'),
  grpTable([600, 3300, 3438, 2300], ['Lp.', 'Narzędzie / przyrząd', 'Rozmiar, typ, zakres', 'Nr czynności w karcie'], [
    ['A. Urządzenia do podnoszenia i zabezpieczenia pojazdu', [
      ['1.', 'Podnośnik warsztatowy', 'nośność dostosowana do masy pojazdu (DMC 1655 kg)', '4'],
      ['2.', 'Podstawki (kozły) warsztatowe – 2 szt.', '–', '4, 28'],
      ['3.', 'Kliny pod koła – 2 szt.', '–', '2, 32'],
    ]],
    ['B. Narzędzia montażowe', [
      ['4.', 'Nasadka udarowa + pokrętło', '21 mm, sześciokątna – nakrętki kół [2]', '3, 5, 14, 27'],
      ['5.', 'Klucz płasko-oczkowy', '13 mm – śruby zacisku [2]', '7'],
      ['6.', 'Klucz płasko-oczkowy', '17 mm – kontra sworznia prowadzącego [2]', '7, 24'],
      ['7.', 'Nasadka', '13 mm – śruby zacisku [2]', '24'],
      ['8.', 'Klucz dynamometryczny', 'zakres obejmujący 34,3 i 103 N·m', '14, 24, 28'],
      ['9.', 'Przyrząd do cofania tłoczków zacisku', 'do zacisków z tłoczkiem wciskanym', '19'],
      ['10.', 'Łom montażowy, wkrętak płaski', '–', '9, 10'],
      ['11.', 'Hak lub drut do podwieszenia zacisku', '–', '8'],
    ]],
    ['C. Przyrządy pomiarowe', [
      ['12.', 'Suwmiarka', 'dokładność 0,1 mm', '11'],
      ['13.', 'Mikrometr zewnętrzny', 'zakres obejmujący 23–25 mm, dokładność 0,01 mm', '13'],
      ['14.', 'Czujnik zegarowy ze statywem magnetycznym', 'dokładność 0,01 mm', '14'],
      ['15.', 'Przymiar liniowy (linijka stalowa)', 'do 300 mm', '31'],
      ['16.', 'Stanowisko rolkowe do badania hamulców', 'pomiar siły hamowania każdego koła', 'rozdział 7'],
    ]],
    ['D. Środki czyszczące i smarne (szczegóły w tabeli 5)', [
      ['17.', 'Szczotka druciana', '–', '6, 12, 16, 25'],
      ['18.', 'Zmywacz do układów hamulcowych, środek penetrujący', '–', '6, 16, 21'],
      ['19.', 'Smar litowy na bazie glikolu, smar do hamulców, pasta przeciwpiskowa, smar miedziany', '–', '17, 20, 25'],
      ['20.', 'Płyn hamulcowy', 'DOT 3 [5]', '30'],
    ]],
    ['E. Środki ochrony indywidualnej', [
      ['21.', 'Rękawice ochronne, okulary ochronne', '–', 'wszystkie'],
    ]],
  ], { center: [0, 3] }),
  P('', { after: 60 }),
  P('Proces nie wymaga testera diagnostycznego. Pojazd ma mechaniczny hamulec postojowy (bez elektrycznego napędu tłoczków tylnych zacisków), mechaniczny wskaźnik zużycia klocków, którego nie trzeba kasować w sterowniku, a tłoczek zacisku przedniego cofa się przyrządem ręcznym. Jedynym urządzeniem elektronicznym wykorzystywanym w procesie jest stanowisko rolkowe użyte w kontroli jakości.'),
];

// ---------- 6. KARTA INSTRUKCYJNA ----------
// n – nazwa, o – opis, t – narzędzie, w – wymiary/moment, u – uwagi, f – opis zdjęcia (null – brak), q – cytat AUTODOC, s – źródło zdjęcia
const S = [
  { n: 'Odkręcenie korka zbiornika płynu hamulcowego.', o: 'Maska silnika zostaje otwarta, a korek zbiornika płynu hamulcowego odkręcony.', t: 'Brak specjalnych narzędzi.', w: 'Poziom płynu między MIN a MAX; płyn DOT 3 [5].', u: 'Umożliwia cofnięcie tłoczka bez nadmiernego wzrostu ciśnienia w układzie.', f: 'otwarta maska i odkręcony korek zbiornika płynu hamulcowego', s: '[2]' },
  { n: 'Zabezpieczenie kół klinami.', o: 'Pod koła tylnej osi zostają podłożone kliny.', t: 'Kliny pod koła (2 szt.).', w: '—', u: 'Zapobiega przypadkowemu przemieszczeniu pojazdu.', f: 'zabezpieczenie kół klinami', s: '[2]' },
  { n: 'Poluzowanie nakrętek koła.', o: 'Nakrętki są lekko odkręcane przed podniesieniem samochodu.', t: 'Nasadka udarowa 21 mm, pokrętło lub klucz do kół.', w: '4 nakrętki; moment montażowy 103 N·m [1].', u: 'Zalecane wykonanie na płaskiej, równej powierzchni.', f: 'poluzowanie nakrętek mocujących koło', s: '[2]', q: 'nasadka udarowa do kół nr 21' },
  { n: 'Podniesienie pojazdu.', o: 'Samochód zostaje podniesiony podnośnikiem, a następnie zabezpieczony kozłami warsztatowymi.', t: 'Podnośnik, podstawki (kozły warsztatowe).', w: '—', u: 'Nie wolno wykonywać pracy wyłącznie na podnośniku.', f: 'moment podniesienia pojazdu', s: '[2]' },
  { n: 'Odkręcenie nakrętek i demontaż koła.', o: 'Nakrętki zostają całkowicie odkręcone, a koło zdjęte z piasty.', t: 'Nasadka udarowa 21 mm, pokrętło.', w: '—', u: 'Aby uniknąć kontuzji, należy przytrzymywać koło podczas odkręcania.', f: 'układ hamulcowy po zdemontowaniu koła', s: '[2]' },
  { n: 'Oczyszczenie zacisku hamulcowego.', o: 'Zacisk i elementy mocujące zostają oczyszczone szczotką drucianą i spryskane środkiem penetrującym.', t: 'Szczotka druciana, środek penetrujący (WD-40).', w: '—', u: 'Po nałożeniu środka odczekać kilka minut.', f: 'oczyszczanie elementów mocujących zacisk hamulcowy', s: '[2]', q: 'szczotka druciana, WD-40' },
  { n: 'Odkręcenie śrub mocujących zacisk.', o: 'Przy przytrzymanym sworzniu prowadzącym zostają odkręcone dwie śruby mocujące zacisk.', t: 'Klucz płasko-oczkowy 13 mm, klucz płasko-oczkowy 17 mm (kontra).', w: 'Moment montażowy 34,3 N·m [1].', u: 'Sworzeń prowadzący należy przytrzymać kluczem 17 mm [1].', f: 'odkręcanie śrub mocujących zacisk hamulcowy', s: '[2]', q: 'klucz płasko-oczkowy nr 13, klucz płasko-oczkowy nr 17' },
  { n: 'Demontaż zacisku.', o: 'Zacisk zostaje zdjęty i zawieszony na haku, bez odłączania przewodu hamulcowego.', t: 'Hak lub drut.', w: '—', u: 'Nie wolno dopuścić do zwisania zacisku na przewodzie. Nie naciskać pedału hamulca.', f: 'zdemontowany i podwieszony zacisk hamulcowy', s: '[2]' },
  { n: 'Demontaż klocków hamulcowych.', o: 'Zużyte klocki zostają wyjęte z jarzma razem z podkładkami przeciwpiskowymi nr 1 i nr 2.', t: 'Ręczne usunięcie lub łom montażowy.', w: '—', u: 'Zanotować położenie klocka z płytką wskaźnika zużycia.', f: 'demontaż zużytych klocków hamulcowych', s: '[2]', q: 'łom montażowy' },
  { n: 'Demontaż płytek podporowych.', o: 'Z jarzma zostają wyjęte dwie płytki podporowe klocków.', t: 'Wkrętak płaski.', w: '—', u: 'Nie zamieniać płytek górnej i dolnej.', f: 'płytki podporowe klocków w jarzmie zacisku', s: '[1]' },
  { n: 'Pomiar grubości okładzin klocków.', o: 'Grubość okładziny (bez płyty nośnej) zostaje zmierzona w kilku punktach obu klocków.', t: 'Suwmiarka (0,1 mm).', w: '**Nominalna: 11,0 mm**\n**Dopuszczalna: min. 1,0 mm** [1]', u: 'Przy wartości poniżej 1,0 mm wymiana obowiązkowa. Klocki wymienia się kompletem na oś [2].', f: 'pomiar grubości okładziny klocka', s: '[1]' },
  { n: 'Kontrola płytek podporowych.', o: 'Płytki zostają sprawdzone pod kątem sprężystości, odkształceń, pęknięć i zużycia oraz oczyszczone z rdzy.', t: 'Oględziny, szczotka druciana.', w: 'Brak odkształceń i pęknięć [1].', u: 'Płytki zużyte wymienić na nowe.', f: null },
  { n: 'Pomiar grubości tarczy.', o: 'Grubość tarczy zostaje zmierzona w kilku punktach na obwodzie pasa roboczego.', t: 'Mikrometr zewnętrzny 0–25 mm (0,01 mm).', w: '**Nominalna: 25,0 mm**\n**Dopuszczalna: min. 23,0 mm** [1]', u: 'Tarcza poniżej 23,0 mm – wymiana kompletu tarcz na osi.', f: 'pomiar grubości tarczy mikrometrem', s: '[1]' },
  { n: 'Pomiar bicia tarczy.', o: 'Tarcza zostaje dociśnięta nakrętkami kół; czujnik ustawiony 10 mm od zewnętrznej krawędzi, tarcza obrócona o pełny obrót.', t: 'Czujnik zegarowy ze statywem, nasadka 21 mm, klucz dynamometryczny.', w: 'Nakrętki 103 N·m; **bicie max 0,05 mm** [6].', u: 'Przy przekroczeniu sprawdzić luz łożyska i bicie piasty (max 0,05 mm) [6].', f: 'pomiar bicia tarczy czujnikiem zegarowym', s: '[6]' },
  { n: 'Kontrola sworzni prowadzących, osłon i tłoczka.', o: 'Sworznie zostają wyjęte i sprawdzone; skontrolowane zostają osłony sworzni i osłona tłoczka.', t: 'Oględziny.', w: 'Swobodny przesuw sworzni, osłony bez pęknięć.', u: 'Wyciek płynu spod osłony tłoczka – regeneracja zacisku (poza zakresem karty).', f: null },
  { n: 'Czyszczenie jarzma zacisku.', o: 'Miejsca osadzenia płytek i klocków zostają oczyszczone z brudu i korozji.', t: 'Szczotka druciana, zmywacz do układów hamulcowych.', w: '—', u: 'Po użyciu zmywacza odczekać kilka minut.', f: 'czyszczenie jarzma zacisku hamulcowego', s: '[2]', q: 'szczotka druciana, zmywacz do układów hamulcowych' },
  { n: 'Smarowanie i montaż sworzni prowadzących.', o: 'Części ślizgowe sworzni zostają pokryte smarem, a sworznie osadzone w jarzmie.', t: 'Smar litowy na bazie glikolu.', w: '—', u: 'Sprawdzić swobodny przesuw sworzni po montażu.', f: 'smarowanie sworzni prowadzących', s: '[1]' },
  { n: 'Montaż płytek podporowych.', o: 'Oczyszczone lub nowe płytki podporowe zostają osadzone w jarzmie.', t: 'Ręczny montaż.', w: '—', u: 'Płytki muszą być osadzone pewnie, bez luzu.', f: null },
  { n: 'Wciśnięcie tłoczka.', o: 'Tłoczek zostaje cofnięty mechanicznie za pomocą przyrządu do cofania tłoczków.', t: 'Przyrząd do cofania tłoczków zacisku.', w: '—', u: 'Kontrolować poziom płynu w zbiorniku, aby nie doszło do przelania.', f: 'cofnięcie tłoczka zacisku hamulcowego', s: '[2]', q: 'przyrząd do wciskania tłoczka hamulcowego zacisku' },
  { n: 'Przygotowanie nowych klocków.', o: 'Podkładki przeciwpiskowe zostają posmarowane i założone; pasta zostaje nałożona na miejsca styku klocków z jarzmem.', t: 'Smar do hamulców tarczowych, pasta przeciwpiskowa, pędzelek.', w: 'Klocki np. TRW GDB3288: 131,7 × 57,4 × 17,8 mm [11].', u: 'Nie nanosić smaru na okładzinę cierną.', f: 'nakładanie pasty przeciwpiskowej na klocki', s: '[2]', q: 'pasta przeciwpiskowa' },
  { n: 'Czyszczenie powierzchni tarczy.', o: 'Powierzchnie robocze tarczy zostają odtłuszczone środkiem czyszczącym.', t: 'Zmywacz do układów hamulcowych.', w: '—', u: 'Po aplikacji sprayu należy odczekać kilka minut.', f: 'oczyszczanie powierzchni tarczy hamulcowej', s: '[2]', q: 'zmywacz do układów hamulcowych' },
  { n: 'Montaż nowych klocków.', o: 'Nowe klocki zostają osadzone okładziną w stronę tarczy; płytka wskaźnika zużycia skierowana do góry.', t: 'Ręczne osadzenie.', w: '—', u: 'Sprawdzić swobodny przesuw klocków w płytkach.', f: 'montaż nowych klocków hamulcowych', s: '[2]' },
  { n: 'Montaż zacisku i wkręcenie śrub.', o: 'Zacisk zostaje nałożony na klocki, a śruby mocujące wkręcone ręcznie.', t: 'Ręczny montaż.', w: '—', u: 'Przewód hamulcowy nie może być skręcony.', f: 'montaż zacisku hamulcowego', s: '[2]' },
  { n: 'Dokręcenie śrub mocujących zacisk.', o: 'Śruby zostają dokręcone kluczem dynamometrycznym przy przytrzymanym sworzniu.', t: 'Nasadka 13 mm, klucz dynamometryczny, klucz płasko-oczkowy 17 mm.', w: '**34,3 N·m** (350 kgf·cm) [1]', u: 'W instrukcji AUTODOC: 35 N·m [2].', f: 'dokręcanie śrub zacisku kluczem dynamometrycznym', s: '[2]', q: 'klucz dynamometryczny, nasadka nr 13' },
  { n: 'Czyszczenie powierzchni przylegania felgi.', o: 'Powierzchnia styku tarczy z felgą zostaje oczyszczona i pokryta cienką warstwą smaru.', t: 'Szczotka druciana, smar miedziany, pędzelek.', w: '—', u: 'Brak.', f: 'czyszczenie miejsca przylegania felgi koła', s: '[2]', q: 'szczotka druciana, smar miedziany' },
  { n: 'Wymiana klocków po drugiej stronie osi.', o: 'Czynności 3–25 zostają powtórzone dla drugiego koła osi przedniej.', t: 'Jak w czynnościach 3–25.', w: 'Jak w czynnościach 3–25.', u: 'Klocki wymienia się kompletem na oś [2].', f: null },
  { n: 'Montaż koła.', o: 'Koło zostaje nałożone na piastę, a nakrętki wkręcone.', t: 'Nasadka udarowa 21 mm.', w: '—', u: 'Podczas przykręcania przytrzymywać koło.', f: 'montaż koła i wkręcanie nakrętek', s: '[2]', q: 'nasadka udarowa do kół nr 21' },
  { n: 'Opuszczenie pojazdu i dokręcenie nakrętek.', o: 'Pojazd zostaje opuszczony, a nakrętki dokręcone na krzyż.', t: 'Klucz dynamometryczny, nasadka 21 mm.', w: '**103 N·m** (1050 kgf·cm) [1]', u: 'Dokręcać po przekątnej (na krzyż).', f: 'dokręcanie nakrętek koła kluczem dynamometrycznym', s: '[2]', q: 'klucz dynamometryczny, nasadka udarowa do kół nr 21' },
  { n: 'Przywrócenie skoku pedału.', o: 'Pedał hamulca zostaje kilkukrotnie naciśnięty aż do wyczucia oporu.', t: 'Brak.', w: '—', u: 'Wykonać przed pierwszym ruszeniem pojazdu.', f: 'naciskanie pedału hamulca', s: '[2]' },
  { n: 'Kontrola płynu, zakręcenie korka i zamknięcie maski.', o: 'Poziom płynu zostaje sprawdzony i uzupełniony, korek zakręcony, maska zamknięta.', t: 'Brak.', w: 'Poziom MIN–MAX, płyn DOT 3 [5].', u: 'Rozlany płyn natychmiast zmyć – niszczy lakier [5].', f: 'zakręcanie korka zbiornika płynu hamulcowego', s: '[2]' },
  { n: 'Kontrola pedału hamulca.', o: 'Zostaje zmierzona wysokość pedału, luz swobodny i odległość rezerwowa.', t: 'Przymiar liniowy.', w: 'Wys. M/T 134,9–144,9 mm; luz 1–6 mm; rezerwa > 70 mm (490 N) [4].', u: 'Szczegóły w rozdziale 7.1.', f: null },
  { n: 'Usunięcie klinów i przekazanie do kontroli jakości.', o: 'Kliny zostają usunięte, a pojazd przekazany na stanowisko rolkowe (rozdział 7).', t: 'Brak.', w: 'Skuteczność ≥ 50%; różnica sił ≤ 30% [15], [16].', u: 'Przez pierwsze 150–200 km unikać gwałtownego hamowania [2].', f: 'usunięcie klinów spod kół', s: '[2]' },
];

const KW = [520, 1650, 2350, 1800, 1718, 1600]; // = 9638
const MID = KW[1] + KW[2] + KW[3] + KW[4];
const UNIT = 'Układ hamulcowy – przód (Toyota Corolla IX E12 1.6 VVT-i)';

function stepRow(s, i, figNo) {
  const wParas = s.w.split('\n').map(l => P(l, { size: 18, align: AlignmentType.CENTER, after: 0, line: 240 }));
  const inner = [
    new TableRow({ cantSplit: true, children: [
      cell(s.n, KW[1], { size: 19, align: AlignmentType.CENTER }),
      cell(s.o, KW[2], { size: 19, align: AlignmentType.CENTER }),
      cell(s.t, KW[3], { size: 19, align: AlignmentType.CENTER }),
      new TableCell({ borders: BORDERS, width: { size: KW[4], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 40, bottom: 40, left: 60, right: 60 }, children: wParas }),
    ] }),
  ];
  if (s.f) {
    const kids = [
      P(`Na rysunku numer ${figNo} poniżej pokazano ${s.f}.`, { size: 20, after: 80 }),
      photoBox(6400, 3300, `Wkleić: ${s.f} – ${s.s === '[2]' ? 'kadr z poradnika / filmu AUTODOC dla Corolli E120 [2], [3]' : 'rysunek z instrukcji naprawy Toyota ' + s.s}`),
      P(`Rysunek ${figNo}. ${s.f.charAt(0).toUpperCase() + s.f.slice(1)}, ${s.s}`, { size: 16, bold: true, before: 40, after: 40, align: AlignmentType.LEFT }),
    ];
    inner.push(new TableRow({ cantSplit: true, children: [new TableCell({ borders: { ...BORDERS, top: NONE }, columnSpan: 4,
      width: { size: MID, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: kids })] }));
  }
  return new TableRow({ cantSplit: true, children: [
    cell(`${i + 1}.`, KW[0], { size: 19, vAlign: VerticalAlign.TOP }),
    new TableCell({ borders: BORDERS, width: { size: MID, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Table({ width: { size: MID, type: WidthType.DXA }, columnWidths: [KW[1], KW[2], KW[3], KW[4]], rows: inner }),
        new Paragraph({ spacing: { after: 0 }, children: [] })] }),
    cell(s.u, KW[5], { size: 19, align: AlignmentType.CENTER }),
  ] });
}

// podział na arkusze: maks. 2 czynności ze zdjęciem lub 3–4 bez zdjęć na arkusz
const sheets = [];
{
  let cur = [], wgt = 0;
  S.forEach((s, i) => {
    const w = s.f ? 2 : 1;
    if (wgt + w > 4) { sheets.push(cur); cur = []; wgt = 0; }
    cur.push(i); wgt += w;
  });
  if (cur.length) sheets.push(cur);
}
const NS = sheets.length;
let figCounter = FIG;
const stepFig = S.map(s => (s.f ? ++figCounter : null));
FIG = figCounter;

function sheetHeader(no) {
  const HW = [2400, 5038, 1200, 1000];
  return new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: HW, rows: [
    new TableRow({ children: [
      cell(['[Nazwa uczelni]', '[Wydział]', '[Zakład]'], HW[0], { rowSpan: 2, size: 19, align: AlignmentType.CENTER }),
      cell('Nazwa zespołu, podzespołu, części:', HW[1], { size: 19, align: AlignmentType.CENTER }),
      cell('Arkusz', HW[2], { size: 19, align: AlignmentType.CENTER }),
      cell(String(no), HW[3], { size: 19, align: AlignmentType.CENTER }),
    ] }),
    new TableRow({ children: [
      cell(UNIT, HW[1], { size: 19, align: AlignmentType.CENTER, bold: true }),
      cell('Arkuszy', HW[2], { size: 19, align: AlignmentType.CENTER }),
      cell(String(NS), HW[3], { size: 19, align: AlignmentType.CENTER }),
    ] }),
  ] });
}
function sheetColHead() {
  const hd = (t, w) => cell(t, w, { size: 18, align: AlignmentType.CENTER, fill: 'E7ECF5' });
  return new TableRow({ cantSplit: true, children: [
    hd('Lp.', KW[0]),
    new TableCell({ borders: BORDERS, width: { size: MID, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [
      new Table({ width: { size: MID, type: WidthType.DXA }, columnWidths: [KW[1], KW[2], KW[3], KW[4]], rows: [new TableRow({ children:
        ['Nazwa zabiegu lub czynności', 'Opis czynności', 'Narzędzie / przyrządy (rozmiar i typ klucza)', 'Wymiar nominalny / dopuszczalny, moment']
          .map((t, j) => hd(t, KW[j + 1])) })] }),
      new Paragraph({ spacing: { after: 0 }, children: [] })] }),
    hd('Uwagi / zalecenia specjalistów', KW[5]),
  ] });
}
function sheetFooter() {
  const FW = [1150, 1050, 1000, 1150, 1050, 1000, 1188, 1050, 1000];
  return new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: FW, rows: [
    new TableRow({ children: ['Opracował', 'Podpis', 'Data', 'Sprawdził', 'Podpis', 'Data', 'Zatwierdził', 'Podpis', 'Data']
      .map((t, i) => cell(t, FW[i], { size: 17, align: AlignmentType.CENTER })) }),
    new TableRow({ height: { value: 380, rule: HeightRule.ATLEAST }, children: FW.map((w, i) => cell(i === 0 ? AUTOR : '', w, { size: 16, align: AlignmentType.CENTER })) }),
  ] });
}

const fFlow = ++FIG; // schemat procesu – numer po rysunkach karty? – nie, wstawiany przed kartą
// przenumerowanie: schemat ma być przed kartą, więc przesuwamy numery czynności o 1
for (let i = 0; i < stepFig.length; i++) if (stepFig[i]) stepFig[i] += 1;
const flowNo = stepFig.find(Boolean) - 1;

const karta = [
  H1('6. Proces technologiczny wymiany klocków hamulcowych'),
  H2('6.1. Schemat procesu technologicznego'),
  P(`Proces podzielono na siedem etapów, z których dwa kończą się punktem decyzyjnym. Pierwszy (po weryfikacji) rozstrzyga, czy tarcza może pozostać w pojeździe; drugi (po kontroli jakości) – czy pojazd może zostać wydany, czy wraca do weryfikacji. Schemat przedstawiono na rysunku numer ${flowNo}, a numery czynności odpowiadają karcie instrukcyjnej z punktu 6.2.`, { keepNext: true }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new ImageRun({
    type: 'png', data: fs.readFileSync(path.join(__dirname, 'schemat.png')), transformation: { width: 410, height: Math.round(410 * 2350 / 1664) } })] }),
  CAP(`Rysunek ${flowNo}. Schemat procesu technologicznego wymiany klocków hamulcowych z kontrolą jakości – opracowanie własne na podstawie [1], [4], [6], [15], [16]`),
  H2('6.2. Karta instrukcyjna operacji'),
  P('Karta instrukcyjna (tabela numer 9) zawiera 32 czynności. W stosunku do samej wymiany części rozszerzono ją o czynności weryfikacyjne (pomiar okładzin, grubości i bicia tarczy, kontrola płytek podporowych i sworzni) oraz o osobną kolumnę z wymiarem nominalnym, wymiarem dopuszczalnym lub momentem dokręcania. Narzędzia w kolumnie czwartej odpowiadają pozycjom tabeli numer 8.'),
];
sheets.forEach((idxs, k) => {
  const rows = [sheetColHead(), ...idxs.map(i => stepRow(S[i], i, stepFig[i]))];
  const kids = [];
  kids.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: k === 0
    ? [new TextRun({ text: 'Tabela 9. Karta instrukcyjna operacji wymiany klocków hamulcowych, [1]–[6]', bold: true, font: FONT, size: 18 })] : [] }));
  kids.push(sheetHeader(k + 1));
  kids.push(new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: [KW[0], MID, KW[5]], rows }));
  kids.push(sheetFooter());
  karta.push(...kids);
});

// ---------- 7. KONTROLA JAKOŚCI ----------
const kontrola = [
  H1('7. Kontrola jakości po naprawie'),
  P('Kontrola jakości ma na celu obiektywne potwierdzenie, że po wymianie klocków układ hamulcowy działa prawidłowo i spełnia wymagania producenta oraz przepisów. Składa się z kontroli statycznej (bez jazdy), badania na stanowisku rolkowym oraz jazdy próbnej. Wyniki zapisuje się w protokole (tabela numer 11), który stanowi dokument potwierdzający jakość usługi.'),
  H2('7.1. Kontrola statyczna układu hamulcowego'),
  P('Kontrolę statyczną wykonuje się po opuszczeniu pojazdu, przed pierwszym ruszeniem. Obejmuje ona:'),
  BUL('**wysokość pedału hamulca** mierzoną od podłogi: dla skrzyni manualnej 134,9–144,9 mm, dla automatycznej 136,0–146,0 mm [4],'),
  BUL('**luz swobodny pedału** mierzony po kilkukrotnym naciśnięciu pedału przy wyłączonym silniku (usunięcie podciśnienia ze wspomagania): 1–6 mm [4],'),
  BUL('**odległość rezerwową pedału** mierzoną przy pracującym silniku i sile nacisku ok. 490 N: więcej niż 70 mm od podłogi [4],'),
  BUL('**poziom płynu hamulcowego** między znakami MIN i MAX zbiornika [5],'),
  BUL('**szczelność** zacisków i przewodów (brak wycieków po kilkukrotnym mocnym naciśnięciu pedału),'),
  BUL('**swobodny obrót kół** po zwolnieniu pedału (brak ciągłego ocierania klocków),'),
  BUL('**kontrolę momentów** dokręcenia nakrętek kół kluczem dynamometrycznym (103 N·m) [1].'),
  H2('7.2. Badanie skuteczności hamowania na stanowisku rolkowym'),
  P('Badanie wykonuje się na rolkowym stanowisku do badania hamulców, które mierzy siłę hamowania każdego koła osi. Na jego podstawie wyznacza się dwa wskaźniki stosowane na stacjach kontroli pojazdów [15], [16].'),
  SUB('**Wskaźnik skuteczności hamowania hamulca roboczego:**'),
  C('z = ΣT / (m_DMC · g) · 100%', { size: 24, bold: true, after: 120 }),
  P('gdzie: ΣT – suma sił hamowania wszystkich kół [N], m_DMC – dopuszczalna masa całkowita pojazdu [kg], g = 9,81 m/s². Dla samochodów osobowych zarejestrowanych po raz pierwszy przed 28 lipca 2010 r. (co obejmuje Corollę E12, produkowaną w latach 2001–2007) wymagana wartość wynosi **co najmniej 50%** [16].'),
  P('Dla analizowanego pojazdu (DMC = 1655 kg [7]) wymagana suma sił hamowania wynosi:'),
  C('ΣT_min = 0,50 · 1655 kg · 9,81 m/s² ≈ 8118 N ≈ 8,12 kN', { size: 22, bold: true, after: 120 }),
  SUB('**Różnica sił hamowania kół tej samej osi:**'),
  C('ΔT = (T_max − T_min) / T_max · 100% ≤ 30%', { size: 24, bold: true, after: 120 }),
  P('gdzie T_max i T_min – większa i mniejsza z sił hamowania kół lewego i prawego jednej osi [15]. Przekroczenie tej wartości po wymianie klocków wskazuje najczęściej na zapieczony sworzeń prowadzący, nieprawidłowo osadzony klocek lub zanieczyszczoną tarczę. W analizowanym przypadku (tabela 7) pomiar ten potwierdza skuteczność naprawy lewych sworzni prowadzących. W razie przekroczenia proces wraca do etapu weryfikacji (rysunek numer ' + flowNo + ').'),
  P('Ponieważ pojazd badany bez obciążenia może osiągać blokowanie kół przy sile mniejszej od wymaganej, przepisy dopuszczają w takim przypadku wyznaczenie obliczeniowego wskaźnika skuteczności na podstawie nacisku na pedał [16]. Wymagania stosowane w kontroli jakości zestawiono w tabeli numer 10.'),
  TCAP('Tabela 10. Kryteria odbioru jakościowego naprawy, [1], [4], [5], [15], [16]'),
  table([4200, 3438, 2000], [
    ['Parametr kontrolny', 'Wymaganie', 'Źródło'],
    ['Wysokość pedału hamulca od podłogi', 'M/T 134,9–144,9 mm; A/T 136,0–146,0 mm', '[4]'],
    ['Luz swobodny pedału', '1–6 mm', '[4]'],
    ['Odległość rezerwowa pedału (490 N)', '> 70 mm', '[4]'],
    ['Poziom płynu hamulcowego', 'MIN–MAX', '[5]'],
    ['Moment nakrętek kół', '103 N·m', '[1]'],
    ['Wskaźnik skuteczności hamulca roboczego', '≥ 50%', '[16]'],
    ['Różnica sił hamowania kół osi przedniej', '≤ 30%', '[15]'],
  ], { center: [1, 2] }),
  P('', { after: 60 }),
  H2('7.3. Protokół kontroli jakości'),
  P('Wyniki kontroli wpisuje się do protokołu przedstawionego w tabeli numer 11. Pojazd może zostać wydany klientowi wyłącznie wtedy, gdy wszystkie pozycje mają ocenę „spełnia”.'),
  TCAP('Tabela 11. Protokół kontroli jakości po wymianie klocków hamulcowych osi przedniej – opracowanie własne'),
  table([600, 3600, 2600, 1500, 1338], [
    ['Lp.', 'Parametr', 'Wymaganie', 'Wynik pomiaru', 'Ocena (spełnia / nie spełnia)'],
    ['1.', 'Wysokość pedału hamulca', '134,9–144,9 mm (M/T)', '', ''],
    ['2.', 'Luz swobodny pedału', '1–6 mm', '', ''],
    ['3.', 'Odległość rezerwowa pedału', '> 70 mm', '', ''],
    ['4.', 'Poziom płynu hamulcowego', 'MIN–MAX', '', ''],
    ['5.', 'Szczelność zacisków i przewodów', 'brak wycieków', '', ''],
    ['6.', 'Swobodny obrót kół przednich', 'brak ocierania', '', ''],
    ['7.', 'Siła hamowania – koło lewe przednie T_L', '—', '……… N', '—'],
    ['8.', 'Siła hamowania – koło prawe przednie T_P', '—', '……… N', '—'],
    ['9.', 'Różnica sił hamowania osi przedniej ΔT', '≤ 30%', '……… %', ''],
    ['10.', 'Wskaźnik skuteczności hamowania z', '≥ 50%', '……… %', ''],
    ['11.', 'Jazda próbna – brak ściągania i pisków', 'brak', '', ''],
  ], { center: [0, 2, 3, 4], size: 19 }),
  P('', { after: 60 }),
  P('Data kontroli: ................  Numer rejestracyjny / VIN: ................................  Podpis kontrolującego: ................', { size: 20, align: AlignmentType.LEFT }),
  H2('7.4. Jazda próbna i docieranie klocków'),
  P('Ostatnim etapem kontroli jest jazda próbna, podczas której sprawdza się, czy pojazd nie ściąga przy hamowaniu, czy nie występują piski ani drgania oraz czy pedał ma prawidłowy, twardy punkt działania. Kierowcę należy poinformować, że nowe klocki wymagają docierania – przez pierwsze 150–200 km należy unikać gwałtownego hamowania [2].'),
];

// ---------- 8. BŁĘDY WYKONAWCZE ----------
const bledy = [
  H1('8. Typowe błędy wykonawcze i ich skutki'),
  P('Wymiana klocków należy do prostych operacji, dlatego łatwo o rutynę i pominięcie pozornie mało istotnych czynności. W tabeli numer 12 zestawiono błędy, które mogą wystąpić w trakcie realizacji procesu, ich skutki oraz sposób zapobiegania, wraz z odwołaniem do czynności karty instrukcyjnej, w której dane zabezpieczenie zostało ujęte.'),
  TCAP('Tabela 12. Typowe błędy wykonawcze przy wymianie klocków hamulcowych – opracowanie własne na podstawie [1], [2], [4]–[6]'),
  table([2500, 2600, 3138, 1400], [
    ['Błąd wykonawczy', 'Skutek', 'Zapobieganie', 'Nr czynności'],
    ['Odkręcanie śruby zacisku bez przytrzymania sworznia', 'Obrót sworznia, uszkodzenie osłony, korozja i zapiekanie prowadnicy', 'Kontra sworznia kluczem 17 mm [1], [2]', '7, 24'],
    ['Pozostawienie zacisku zwisającego na przewodzie', 'Uszkodzenie przewodu elastycznego, wyciek płynu', 'Podwieszenie zacisku na haku [2]', '8'],
    ['Naciśnięcie pedału przy zdjętym zacisku', 'Wysunięcie tłoczka z cylindra, wyciek płynu', 'Zakaz naciskania pedału do czasu montażu [2]', '8'],
    ['Cofanie tłoczka przy zamkniętym zbiorniku bez kontroli poziomu', 'Przelanie płynu, uszkodzenie lakieru', 'Odkręcony korek, kontrola poziomu, natychmiastowe zmycie płynu [5]', '1, 19'],
    ['Smar lub pasta na okładzinie albo tarczy', 'Spadek współczynnika tarcia, wydłużenie drogi hamowania', 'Ostrożne nakładanie środków, odtłuszczenie tarczy [2]', '20, 21'],
    ['Smar mineralny na sworzniach i elementach gumowych', 'Pęcznienie osłon, zapiekanie sworzni, ściąganie pojazdu', 'Smar litowy na bazie glikolu [1]', '17'],
    ['Pominięcie pomiaru tarczy', 'Nowe klocki na tarczy poniżej 23,0 mm lub z biciem – drgania, przegrzewanie', 'Pomiar grubości i bicia [1], [6]', '13, 14'],
    ['Odwrotny montaż klocka z płytką wskaźnika zużycia', 'Brak sygnału akustycznego przy kolejnym zużyciu', 'Płytka wskaźnika skierowana do góry [1]', '22'],
    ['Wymiana klocków tylko po jednej stronie osi', 'Różnica sił hamowania, ściąganie', 'Wymiana kompletem na oś [2]', '26'],
    ['Dokręcanie bez klucza dynamometrycznego', 'Zerwanie gwintu lub samoczynne poluzowanie połączenia', '34,3 N·m (zacisk), 103 N·m (koła) [1]', '24, 28'],
    ['Ruszenie bez przywrócenia skoku pedału', 'Brak hamowania przy pierwszym naciśnięciu pedału', 'Kilkukrotne naciśnięcie pedału przed jazdą [2]', '29'],
    ['Gwałtowne hamowanie bezpośrednio po wymianie', 'Przegrzanie i zeszklenie okładzin', 'Docieranie przez pierwsze 150–200 km [2]', '32'],
  ], { center: [3], size: 18 }),
  P('', { after: 60 }),
  P('Część z wymienionych błędów (np. zapieczone sworznie, smar na okładzinie, wymiana po jednej stronie) nie daje się zauważyć bezpośrednio po zakończeniu montażu, natomiast ujawnia się jako różnica sił hamowania kół osi. Jest to dodatkowe uzasadnienie dla badania na stanowisku rolkowym opisanego w rozdziale 7.'),
];

// ---------- 9. KOSZTY ----------
const koszty = [
  H1('9. Analiza kosztów naprawy'),
  P('Koszt usługi obejmuje części i robociznę. Materiały eksploatacyjne z tabeli 5 oraz narzędzia stanowią wyposażenie warsztatu i są wliczone w stawkę roboczogodziny. Ponieważ cena klocków zależy przede wszystkim od klasy producenta, koszt przeanalizowano w trzech wariantach doboru części.'),
  H2('9.1. Warianty doboru części'),
  P('Wszystkie porównywane zestawy klocków mają ten sam numer wymienności WVA 23766, co oznacza identyczny kształt i wymiary klocka oraz zgodność z zaciskiem pojazdu [12], [24], [25].'),
  TCAP('Tabela 13. Warianty zestawów klocków przednich dla Toyoty Corolla E12 (ceny brutto), [12], [18], [19], [24], [25]'),
  table([1800, 3700, 1800, 2338], [
    ['Wariant', 'Zestaw klocków', 'Cena brutto', 'Źródło'],
    ['Ekonomiczny', 'Ridex 402B0107', '79,99 zł', '[24]'],
    ['Standardowy', 'TRW GDB3288 – średnia z trzech ofert (116,00 zł; 140,00 zł; 154,99 zł)', '137,00 zł', '[12], [18], [19]'],
    ['Premium', 'ATE 13.0460-5815.2', '206,99 zł', '[25]'],
  ], { center: [0, 2, 3] }),
  P('', { after: 60 }),
  H2('9.2. Koszt robocizny'),
  P('Stawki roboczogodziny w warsztatach niezależnych wynoszą 100–150 zł netto w mniejszych miejscowościach i 150–250 zł netto w dużych miastach [20]. Przyjęto stawkę 200 zł netto, czyli 246 zł brutto. Czas wymiany klocków jednej osi wynosi zwykle 0,5–1 roboczogodziny [21], a w normie egzaminacyjnej dla zawodu mechanika przyjęto 1,0 rbh [22]. Ponieważ proces obejmuje pomiary tarcz, czyszczenie i smarowanie sworzni oraz kontrolę jakości, przyjęto górną wartość – 1,0 rbh, co daje koszt robocizny 246 zł brutto.'),
  H2('9.3. Zestawienie kosztów'),
  TCAP('Tabela 14. Koszt usługi w trzech wariantach doboru części (ceny brutto) – opracowanie własne'),
  table([2400, 2400, 2400, 2438], [
    ['Wariant', 'Części', 'Robocizna (1,0 rbh)', 'Razem'],
    ['Ekonomiczny', '79,99 zł', '246,00 zł', '**325,99 zł**'],
    ['Standardowy', '137,00 zł', '246,00 zł', '**383,00 zł**'],
    ['Premium', '206,99 zł', '246,00 zł', '**452,99 zł**'],
  ], { center: [0, 1, 2, 3] }),
  P('', { after: 60 }),
  P('Robocizna stanowi od 54% (wariant premium) do 75% (wariant ekonomiczny) kosztu usługi, więc różnica cen klocków zmienia koszt całkowity w mniejszym stopniu niż wynikałoby z samych cen części. Do realizacji przyjęto wariant standardowy (TRW GDB3288, ok. 383 zł brutto) – klocki znanego producenta układów hamulcowych z akustycznym wskaźnikiem zużycia, przy umiarkowanej cenie. Kalkulacja nie obejmuje wymiany tarcz, która byłaby konieczna przy grubości poniżej 23,0 mm.'),
];

// ---------- 10. WNIOSKI ----------
const wnioski = [
  H1('10. Wnioski'),
  P('1. Dla Toyoty Corolla IX (E12) 1.6 VVT-i opracowano proces technologiczny, w którym każda wartość liczbowa ma wskazane źródło. Wymiary graniczne (okładzina 11,0 / 1,0 mm, tarcza 25,0 / 23,0 mm, bicie 0,05 mm) i momenty dokręcania (34,3 N·m, 103 N·m) pochodzą z instrukcji naprawy producenta, a rozmiary kluczy (21, 13 i 17 mm) z instrukcji krok po kroku dla tego modelu.'),
  P('2. Rozdzielenie diagnostyki na macierz objawów i kartę weryfikacji pozwala odróżnić samo zużycie klocków od jego przyczyny. W analizowanym przypadku nierównomierne zużycie klocków wskazało na utrudniony przesuw sworzni prowadzących – sama wymiana klocków nie usunęłaby ściągania pojazdu, dlatego w procesie ujęto ich czyszczenie i smarowanie.'),
  P('3. Wymiana klocków nie kończy procesu. Kontrola jakości według danych producenta (pedał hamulca) oraz kryteriów badań technicznych (wskaźnik skuteczności ≥ 50%, czyli dla tego pojazdu ΣT ≥ 8,12 kN, oraz różnica sił na osi ≤ 30%) daje obiektywny dowód poprawności naprawy i pozwala wykryć błędy wykonawcze, które nie są widoczne podczas montażu.'),
  P('4. Procedura nie wymaga oprogramowania diagnostycznego, a jej koszt w zależności od wariantu części wynosi od ok. 326 zł do ok. 453 zł brutto. Ponieważ robocizna stanowi większość kosztu, oszczędność na klasie klocków ma ograniczony wpływ na cenę usługi, a może wpływać na trwałość i komfort hamowania.'),
];

// ---------- BIBLIOGRAFIA ----------
const BIB = [
  'Toyota Motor Corporation: Toyota Corolla (E120) 2002–2008 Repair Manual – Brake – Front brake – Overhaul, https://www.tcorolla.net/overhaul-1145.html (dostęp: 09.10.2026).',
  'AUTODOC CLUB: Jak wymienić klocki hamulcowe przód w TOYOTA Corolla IX Hatchback (E120) – poradnik naprawy (strona WWW i PDF), https://club.autodoc.pl/manuals/jak-wymienic-klocki-hamulcowe-przod-w-toyota-corolla-ix-hatchback-e120-poradnik-naprawy-25473 (dostęp: 09.10.2026).',
  'AUTODOC CLUB: Toyota Corolla E12 – instrukcje naprawy krok po kroku i filmiki instruktażowe (film: How to change front brake discs and front brake pads on TOYOTA COROLLA E120 | AUTODOC), https://club.autodoc.pl/manuals/toyota/corolla/corolla-zze12-nde12-zde12 (dostęp: 09.10.2026).',
  'Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Brake – Adjustment (brake pedal), https://www.tcorolla.net/adjustment-1133.html (dostęp: 09.10.2026).',
  'Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Brake fluid, https://www.tcorolla.net/brake_fluid-1131.html (dostęp: 09.10.2026).',
  'Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Front axle hub – Replacement (disc runout), https://www.tcorolla.net/replacement-1121.html (dostęp: 09.10.2026).',
  'AutoCentrum: Toyota Corolla IX (E12) 1.6 i 16V 110 KM 2001–2007 – dane techniczne, https://www.autocentrum.pl/dane-techniczne/toyota/corolla/ix-e12/sedan/silnik-benzynowy-1.6-i-16v-110km-2001-2007/ (dostęp: 09.10.2026).',
  'Autokatalog: Toyota Corolla IX (E120/E130) – dane techniczne, https://autokatalog.pl/toyota/corolla/ix-e120-e130/dane-techniczne (dostęp: 09.10.2026).',
  'Cars-data: Toyota Corolla Hatchback 1.6 16V VVT-i (2002) – specifications, https://cars-data.com/en/toyota/corolla/2002-hatchback-7726/1-6-16v-vvt-i-executive-28121--28121/specs (dostęp: 09.10.2026).',
  'Schaeffler – katalog zastosowań: TOYOTA COROLLA (_E12_) 1.6 VVT-i (ZZE121_), https://shop.rolling.hu/catalog/SCHAEFFLER/toyota-corolla-saloon-e12-16-vvt-i-zze121 (dostęp: 09.10.2026).',
  'Inter Cars: Klocki hamulcowe TRW Automotive GDB3288 – dane techniczne, https://intercars.pl/produkty/321869-klocki-hamulcowe-trw-automotive-gdb3288 (dostęp: 09.10.2026).',
  'AUTODOC: TRW GDB3288 Klocki hamulcowe z akustycznym czujnikiem zużycia do TOYOTA COROLLA, https://www.autodoc.pl/trw/2192809 (dostęp: 09.10.2026).',
  'AUTODOC: Tarcza hamulcowa AP 24941 (Ø 255 mm, grubość 25 mm, min. 23 mm), https://www.autodoc.pl/ap/9367120 (dostęp: 09.10.2026).',
  'Motostacja: TRW tarcze + klocki tył Toyota Corolla E12 (DF4379, GDB3289), https://motostacja.com/p/TRW-2--DF4379-GDB3289 (dostęp: 09.10.2026).',
  'Warsztat.pl: Diagnozowanie układu hamulcowego pojazdu samochodowego (cz. 3), https://warsztat.pl/artykuly/diagnozowanie-ukladu-hamulcowego-pojazdu-samochodowego-cz-3,57838,bm9uZSE1NzgzOCEhbm93b2N6ZXNueXdhcnN6dGF0LnBsL2FydHlrdWx5LzA0LTIwMTIvb2xlamUvMg (dostęp: 09.10.2026).',
  'Gajek A.: Archiwum Motoryzacji – artykuł dotyczący oceny skuteczności hamowania w stacjach kontroli pojazdów, https://yadda.icm.edu.pl/baztech/element/bwmeta1.element.baztech-1930bdd7-1355-423d-93b0-ddc9a790e278/c/AM73_Andrzej_Gajek_PL.pdf (dostęp: 09.10.2026).',
  'Rozporządzenie Ministra Infrastruktury w sprawie zakresu i sposobu przeprowadzania badań technicznych pojazdów oraz wzorów dokumentów stosowanych przy tych badaniach (tekst ujednolicony), https://www.piskp.pl/images/stories/Procedura_badan_tekst_ujednolicony_08_2015.pdf (dostęp: 09.10.2026).',
  'Henkiel: Klocki hamulcowe TRW GDB3288 – Toyota Corolla (E12), oś przednia, https://www.henkiel.com.pl/65070-klocki-hamulcowe-trw-gdb3288-toyota-corolla-e12-corolla-e11-os-przednia (dostęp: 09.10.2026).',
  'Magmoto: TRW GDB3288 Zestaw klocków hamulcowych, https://magmoto.pl/trw-gdb3288-zestaw-klockow-hamulcowych-hamulce-tarczowe-p-140773.html (dostęp: 09.10.2026).',
  'Freenance: Zarobki mechanika samochodowego 2026 – stawki roboczogodziny, https://freenance.io/zarobki/zarobki-mechanika-samochodowego-2026/ (dostęp: 09.10.2026).',
  'Otomoto: Ile kosztuje wymiana klocków hamulcowych?, https://www.otomoto.pl/news/ile-kosztuje-wymiana-klockow-hamulcowych (dostęp: 09.10.2026).',
  'Arkusz egzaminacyjny MOT.06 – styczeń 2022, część praktyczna, https://zawodowe.edu.pl/arkusz-praktyczny/mot06-2022-styczen-01/arkusz.pdf (dostęp: 09.10.2026).',
  'Toyota Motor Europe: portal informacji technicznej dla niezależnych operatorów, https://www.toyota-tech.eu (dostęp: 09.10.2026).',
  'AUTODOC: Ridex 402B0107 Klocki hamulcowe (WVA 23766), https://www.autodoc.pl/ridex/12754705 (dostęp: 09.10.2026).',
  'AUTODOC: ATE 13.0460-5815.2 Klocki hamulcowe, https://www.autodoc.pl/ate/955966 (dostęp: 09.10.2026).',
];
const bib = [H1('Bibliografia'), ...BIB.map((b, i) => new Paragraph({
  spacing: { after: 100 }, indent: { left: 560, hanging: 560 },
  children: [new TextRun({ text: `[${i + 1}]\t`, font: FONT, size: 20 }), new TextRun({ text: b, font: FONT, size: 20 })],
  tabStops: [{ type: TabStopType.LEFT, position: 560 }],
}))];

// ---------- DOKUMENT ----------
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] });
const doc = new Document({
  creator: AUTOR, title: 'Proces technologiczny wymiany przednich klocków hamulcowych – Toyota Corolla E12',
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { titlePage: true, page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1134, right: 1134, footer: 400 } } },
    footers: { default: footer, first: new Footer({ children: [] }) },
    children: [...title, new Paragraph({ children: [new PageBreak()] }), ...tocPage, ...wstep, ...opisPojazdu, ...uklad, ...diag, ...narz,
      ...karta, ...kontrola, ...bledy, ...koszty, ...wnioski, ...bib],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(OUT, b);
  console.log('ok', OUT, 'arkuszy:', NS, 'rysunków:', FIG, 'arkusze:', JSON.stringify(sheets.map(s => s.map(i => i + 1))));
  fs.writeFileSync(path.join(path.dirname(OUT), '.toc_titles.json'), JSON.stringify(TOC.map(t => t[0])));
});

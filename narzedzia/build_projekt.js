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
  ['1. Wstęp', 0], ['2. Opis pojazdu', 0], ['3. Opis układu hamulcowego', 0],
  ['3.1. Wymiary nominalne i dopuszczalne elementów hamulca przedniego', 1], ['3.2. Momenty dokręcania połączeń gwintowych', 1],
  ['4. Diagnoza zużycia klocków hamulcowych', 0], ['4.1. Objawy zużycia klocków hamulcowych', 1],
  ['4.2. Metody diagnostyczne stosowane przy ocenie stanu klocków hamulcowych', 1], ['4.3. Charakterystyka przypadku w pojeździe Toyota Corolla E12', 1],
  ['5. Wykaz narzędzi i oprogramowania potrzebnych do przeprowadzenia procedury', 0],
  ['6. Karta instrukcyjna operacji wymiany klocków hamulcowych', 0],
  ['7. Kontrola jakości po naprawie', 0], ['7.1. Kontrola statyczna układu hamulcowego', 1],
  ['7.2. Badanie skuteczności hamowania na stanowisku rolkowym', 1], ['7.3. Protokół kontroli jakości', 1], ['7.4. Jazda próbna i docieranie klocków', 1],
  ['8. Koszty związane z przeprowadzoną procedurą', 0], ['8.1. Koszt części – klocki hamulcowe (przód)', 1],
  ['8.2. Koszt robocizny', 1], ['8.3. Łączny koszt usługi', 1], ['9. Wnioski', 0], ['Bibliografia', 0],
];
const tocPage = [
  P('Spis treści', { size: 30, bold: true, align: AlignmentType.LEFT, after: 240, before: 0 }),
  ...TOC.map(([t, lvl]) => new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: PW, leader: 'dot' }],
    indent: { left: lvl ? 400 : 0 }, spacing: { after: 80 },
    children: [new TextRun({ text: t, font: FONT, size: 22, bold: !lvl }),
      new TextRun({ text: `\t${PAGES[t] || '00'}`, font: FONT, size: 22 })],
  })),
];

// ---------- 1. WSTĘP ----------
const wstep = [
  H1('1. Wstęp'),
  P('Układ hamulcowy jest jednym z najważniejszych układów decydujących o bezpieczeństwie czynnym pojazdu. Jego zadaniem jest zmniejszanie prędkości, zatrzymanie pojazdu oraz utrzymanie go w miejscu. Elementami najszybciej zużywającymi się w hamulcu tarczowym są klocki hamulcowe, dlatego ich okresowa weryfikacja i wymiana należą do najczęściej wykonywanych zabiegów obsługowych.'),
  P('Mimo pozornej prostoty wymiana klocków wymaga zachowania wymiarów granicznych określonych przez producenta, właściwych momentów dokręcania połączeń gwintowych, zastosowania odpowiednich środków smarnych i czyszczących oraz narzędzi takich jak klucz dynamometryczny i przyrząd do cofania tłoczka. Błąd na którymkolwiek etapie bezpośrednio obniża skuteczność hamowania.'),
  P('Do analizy wybrano samochód Toyota Corolla IX (E12) w nadwoziu hatchback z silnikiem benzynowym 1.6 VVT-i (3ZZ-FE, 81 kW / 110 KM), oznaczenie typu ZZE121. Jest to model bardzo popularny na rynku wtórnym, a producent udostępnia dla niego instrukcję naprawy (Repair Manual) z kompletem wymiarów nominalnych, dopuszczalnych i momentów dokręcania [1], [4]–[6]. Dla tego modelu dostępna jest również instrukcja wymiany krok po kroku z dokumentacją fotograficzną [2], [3].'),
  P('Zakres projektu obejmuje: opis pojazdu i jego układu hamulcowego, zestawienie wymiarów nominalnych i dopuszczalnych oraz momentów dokręcania, diagnozę zużycia klocków, wykaz narzędzi z rozmiarem i typem kluczy, kartę instrukcyjną operacji z opisem każdej czynności oraz kalkulację kosztów usługi.'),
  P('Elementem wyróżniającym projekt jest rozdział poświęcony **kontroli jakości po naprawie**. Obejmuje on kontrolę statyczną pedału hamulca według danych producenta [4] oraz badanie skuteczności i równomierności hamowania na stanowisku rolkowym według kryteriów stosowanych na stacjach kontroli pojazdów [15]–[17]. Dzięki temu proces technologiczny kończy się obiektywną, mierzalną oceną wykonanej naprawy.'),
];

// ---------- 2. OPIS POJAZDU ----------
const f1 = fig(), f2 = fig();
const opisPojazdu = [
  H1('2. Opis pojazdu'),
  P('Toyota Corolla IX generacji (oznaczenie fabryczne E12, E120) była produkowana w latach 2001–2007 w nadwoziach sedan, hatchback (3- i 5-drzwiowy), kombi oraz Verso. Analizowany wariant to hatchback z silnikiem 1.6 VVT-i o oznaczeniu 3ZZ-FE, kod typu ZZE121 [10].'),
  P(`Na rysunku numer ${f1} poniżej pokazano widok pojazdu z przodu.`, { keepNext: true }),
  photoBox(7000, 3400, 'Toyota Corolla E12 hatchback – widok z przodu (zdjęcie ogólnodostępne, np. Wikimedia Commons, z podaniem źródła)'),
  CAP(`Rysunek ${f1}. Toyota Corolla IX (E12) hatchback – widok z przodu, [źródło zdjęcia]`),
  P(`Na rysunku numer ${f2} poniżej pokazano widok pojazdu z tyłu.`, { keepNext: true }),
  photoBox(7000, 3400, 'Toyota Corolla E12 hatchback – widok z tyłu (zdjęcie ogólnodostępne, z podaniem źródła)'),
  CAP(`Rysunek ${f2}. Toyota Corolla IX (E12) hatchback – widok z tyłu, [źródło zdjęcia]`),
  P('Dane techniczne analizowanego pojazdu zestawiono w tabeli numer 1.', { keepNext: true }),
  TCAP('Tabela 1. Charakterystyka techniczna pojazdu Toyota Corolla IX (E12) 1.6 VVT-i, [7]–[10]'),
  table([4200, 5438], [
    ['Parametr', 'Wartość'],
    ['Model', 'Toyota Corolla IX (E12), hatchback 5-drzwiowy'],
    ['Kod typu / lata produkcji wersji', 'ZZE121; 2001–2007 [10]'],
    ['Silnik (kod)', '3ZZ-FE, benzynowy, R4, 16 zaworów, DOHC, VVT-i'],
    ['Pojemność skokowa', '1598 cm³'],
    ['Moc maksymalna', '81 kW (110 KM) przy 6000 obr/min'],
    ['Maksymalny moment obrotowy', '150 N·m przy 4800 obr/min'],
    ['Skrzynia biegów', 'manualna, 5-biegowa'],
    ['Napęd', 'na koła przednie'],
    ['Prędkość maksymalna', '190 km/h'],
    ['Przyspieszenie 0–100 km/h', '10,2 s'],
    ['Zużycie paliwa (miasto / trasa / średnie)', '9,0 / 5,9 / 7,0 l/100 km'],
    ['Długość / szerokość / wysokość', '4180 / 1710 / 1475 mm'],
    ['Rozstaw osi', '2600 mm'],
    ['Masa własna', 'od 1115 kg'],
    ['Dopuszczalna masa całkowita (DMC)', '1655 kg'],
    ['Zbiornik paliwa', '55 l'],
    ['Opony (rozmiar podstawowy)', '195/60 R15'],
    ['Hamulce przednie', 'tarczowe, wentylowane, Ø 255 mm, zacisk pływający jednotłoczkowy [1], [13]'],
    ['Hamulce tylne', 'bębnowe lub tarczowe pełne Ø 258 mm – zależnie od wersji wyposażenia [14]'],
    ['Hamulec postojowy', 'mechaniczny, cięgnowy, działający na koła tylne'],
    ['Układy wspomagające', 'ABS z elektronicznym rozdziałem siły hamowania (EBD)'],
  ], { center: [] }),
  P('', { after: 120 }),
  P('Przedstawione dane pozwalają określić warunki pracy układu hamulcowego. W szczególności dopuszczalna masa całkowita (1655 kg) jest wykorzystywana w rozdziale 7 do wyznaczenia wymaganej siły hamowania podczas badania na stanowisku rolkowym. W kolejnym rozdziale opisano budowę układu hamulcowego.'),
];

// ---------- 3. UKŁAD HAMULCOWY ----------
const f3 = fig(), f4 = fig();
const uklad = [
  H1('3. Opis układu hamulcowego'),
  P('Toyota Corolla E12 jest wyposażona w hydrauliczny, dwuobwodowy układ hamulcowy ze wspomaganiem podciśnieniowym (serwo) i układem ABS. Hamulec roboczy osi przedniej stanowią hamulce tarczowe z tarczami wentylowanymi i zaciskami pływającymi, osadzonymi na dwóch sworzniach prowadzących [1].'),
  SUB('**Parametry przedniego układu hamulcowego:**'),
  BUL('typ: hamulec tarczowy z zaciskiem pływającym, jednotłoczkowym [1],'),
  BUL('tarcza: wentylowana, średnica 255 mm, 4 otwory, średnica centrowania 55 mm [13],'),
  BUL('grubość tarczy: nominalna 25,0 mm, minimalna 23,0 mm [1],'),
  BUL('klocki: z płytką akustycznego (mechanicznego) wskaźnika zużycia i podkładkami przeciwpiskowymi nr 1 i nr 2 [1], [11],'),
  BUL('grubość okładziny klocka: nominalna 11,0 mm, minimalna 1,0 mm [1],'),
  BUL('płyn hamulcowy: SAE J1703 lub FMVSS No. 116 DOT 3 [5].'),
  SUB('**Elementy zacisku hamulca przedniego (wg instrukcji producenta [1]):**'),
  BUL('korpus cylindra zacisku z tłoczkiem, uszczelniaczem tłoczka i osłoną przeciwpyłową,'),
  BUL('jarzmo (mocowanie zacisku) przykręcone do zwrotnicy dwiema śrubami,'),
  BUL('dwa sworznie prowadzące z osłonami gumowymi (tulejkami przeciwpyłowymi),'),
  BUL('dwie płytki podporowe klocków (pad support plates),'),
  BUL('klocki hamulcowe z podkładkami przeciwpiskowymi oraz płytką wskaźnika zużycia.'),
  SUB('**Elementy hydrauliczne:**'),
  BUL('pompa hamulcowa (główna) – wytwarza ciśnienie płynu hamulcowego,'),
  BUL('serwo podciśnieniowe – zmniejsza siłę potrzebną do naciśnięcia pedału,'),
  BUL('zespół hydrauliczny ABS – moduluje ciśnienie w obwodach kół,'),
  BUL('przewody sztywne i elastyczne – łączą elementy układu.'),
  SUB('**Hamulec postojowy:** mechaniczny, uruchamiany dźwignią i cięgnami, działający na koła tylne. Nie wymaga oprogramowania serwisowego przy obsłudze hamulców.'),
  SUB('**Uwagi eksploatacyjne:**'),
  BUL('klocki wymienia się zawsze kompletem na oś [2],'),
  BUL('zużycie sygnalizuje metalowa płytka wskaźnika, która po zużyciu okładziny ociera o tarczę i wywołuje pisk [1], [11],'),
  BUL('grubość tarczy należy sprawdzać przy każdej wymianie klocków [1],'),
  BUL('tłoczek zacisku przedniego cofa się mechanicznie (wciskanie), bez obracania.'),
  P(`Na rysunku numer ${f3} poniżej pokazano widok przedniego hamulca tarczowego po zdemontowaniu koła.`, { keepNext: true }),
  photoBox(7000, 3400, 'Hamulec przedni Corolli E12 po zdjęciu koła – kadr z poradnika / filmu AUTODOC [2], [3]'),
  CAP(`Rysunek ${f3}. Przedni hamulec tarczowy pojazdu Toyota Corolla E12, [2]`),
  P(`Na rysunku numer ${f4} poniżej pokazano elementy składowe zacisku hamulca przedniego według instrukcji naprawy producenta.`, { keepNext: true }),
  photoBox(7000, 3800, 'Rysunek rozstrzelony zacisku z instrukcji Toyota Repair Manual – Front brake – Overhaul [1] (zrzut ekranu ze strony)'),
  CAP(`Rysunek ${f4}. Elementy składowe zacisku hamulca przedniego, [1]`),

  H2('3.1. Wymiary nominalne i dopuszczalne elementów hamulca przedniego'),
  P('Wartości graniczne, na podstawie których podejmowana jest decyzja o dalszej eksploatacji, naprawie lub wymianie elementów, zestawiono w tabeli numer 2. Wszystkie wartości pochodzą z instrukcji naprawy producenta.'),
  TCAP('Tabela 2. Wymiary nominalne i dopuszczalne elementów podlegających weryfikacji, [1], [4], [6]'),
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
  H2('3.2. Momenty dokręcania połączeń gwintowych'),
  P('Momenty dokręcania połączeń występujących w procesie, wraz z rozmiarem i typem klucza, zestawiono w tabeli numer 3. Producent w instrukcji naprawy nie podaje rozmiarów kluczy, dlatego przyjęto je na podstawie instrukcji krok po kroku dla tego modelu [2].'),
  TCAP('Tabela 3. Momenty dokręcania oraz rozmiar i typ kluczy, [1], [2], [4]'),
  table([2900, 2600, 1250, 1600, 1288], [
    ['Połączenie', 'Rozmiar i typ klucza', 'Moment [N·m]', 'Moment [kgf·cm / ft·lbf]', 'Źródło'],
    ['Nakrętki koła (4 szt.)', 'nasadka udarowa 21 mm, klucz dynamometryczny', '103', '1050 / 76', '[1], [2]'],
    ['Śruby zacisku do sworzni prowadzących (2 szt.)', 'nasadka 13 mm + klucz dynamometryczny; kontra – klucz płasko-oczkowy 17 mm', '34,3', '350 / 25', '[1], [2]'],
    ['Śruby jarzma do zwrotnicy (2 szt.)*', '–', '106,8', '1089 / 79', '[1]'],
    ['Śruba przewodu elastycznego (banjo)*', '–', '29', '296 / 21', '[1]'],
    ['Nakrętka kontrująca popychacza pedału*', '–', '26', '265 / 19', '[4]'],
  ], { center: [2, 3, 4], size: 19 }),
  P('* połączenia nierozłączane w zakresie karty – podano dla kompletności (demontaż jarzma przy wymianie tarczy, odłączanie przewodu przy regeneracji zacisku, regulacja pedału).', { size: 18, italics: true, after: 120 }),
  P('W instrukcji krok po kroku [2] moment dokręcenia śrub zacisku podano jako 35 N·m, co stanowi zaokrąglenie wartości producenta 34,3 N·m (350 kgf·cm). W projekcie przyjęto wartość producenta.'),
];

// ---------- 4. DIAGNOZA ----------
const f5 = fig();
const diag = [
  H1('4. Diagnoza zużycia klocków hamulcowych'),
  P('W celu zapewnienia skuteczności i bezpieczeństwa działania układu hamulcowego konieczne jest monitorowanie stanu klocków hamulcowych – jednego z najintensywniej zużywających się elementów ciernych pojazdu. W samochodzie Toyota Corolla E12 nie zastosowano elektronicznego czujnika zużycia klocków; o zużyciu informuje mechaniczny (akustyczny) wskaźnik oraz klasyczne objawy pogorszenia pracy hamulców.'),
  H2('4.1. Objawy zużycia klocków hamulcowych'),
  SUB('**1) Pisk wskaźnika zużycia podczas hamowania:**'),
  P('Klocek wewnętrzny jest wyposażony w metalową płytkę wskaźnika zużycia (pad wear indicator plate). Gdy okładzina zbliża się do wartości granicznej, płytka zaczyna ocierać o tarczę i wywołuje charakterystyczny, wysoki pisk przy hamowaniu [1], [11]. Jest to podstawowy sygnał ostrzegawczy w tym modelu – pojazd nie ma kontrolki zużycia klocków.'),
  SUB('**2) Metaliczny zgrzyt podczas hamowania:**'),
  P('Zgrzyt świadczy o całkowitym starciu okładziny i kontakcie płyty nośnej klocka z tarczą. W takim przypadku tarcza ulega szybkiemu uszkodzeniu i zwykle kwalifikuje się do wymiany.'),
  SUB('**3) Wydłużona droga hamowania i „miękki” pedał:**'),
  P('Zużycie okładzin zwiększa skok tłoczka i obniża skuteczność hamowania. Objaw jest wykrywany podczas jazdy próbnej oraz kontroli odległości rezerwowej pedału (wymagane > 70 mm przy sile 490 N) [4].'),
  SUB('**4) Drgania pedału i kierownicy podczas hamowania:**'),
  P('Drgania wskazują zwykle na bicie lub nierównomierną grubość tarczy. Dopuszczalne bicie tarczy wynosi 0,05 mm [6], dlatego objaw ten wymaga pomiaru tarczy w ramach weryfikacji.'),
  SUB('**5) Ściąganie pojazdu podczas hamowania:**'),
  P('Jednostronne zużycie klocków, zapieczony sworzeń prowadzący lub tłoczek powodują różnicę sił hamowania na kołach osi i ściąganie pojazdu na jedną stronę. Dopuszczalna różnica sił hamowania kół jednej osi wynosi 30% [15].'),
  SUB('**6) Wizualne zużycie klocków:**'),
  P('Po zdjęciu koła możliwa jest ocena grubości okładziny przez okno zacisku, a po demontażu zacisku – pomiar suwmiarką. Nominalna grubość okładziny wynosi 11,0 mm, a minimalna 1,0 mm [1].'),
  H2('4.2. Metody diagnostyczne stosowane przy ocenie stanu klocków hamulcowych'),
  P('Poniżej przedstawiono metody diagnostyczne stosowane w praktyce warsztatowej, uporządkowane od najprostszych do najbardziej obiektywnych.'),
  SUB('**1. Kontrola wizualna** – ocena grubości i równomierności zużycia okładzin, pęknięć, rozwarstwień i zanieczyszczeń oraz stanu tarczy (rowki, przebarwienia termiczne).'),
  SUB('**2. Pomiar suwmiarką** – dokładny pomiar grubości okładziny (bez płyty nośnej) w kilku punktach i porównanie z wartością minimalną 1,0 mm [1].'),
  SUB('**3. Wskaźnik zużycia** – w Corolli E12 wskaźnik mechaniczny (akustyczny); pisk przy hamowaniu jest sygnałem do weryfikacji [1], [11].'),
  SUB('**4. Pomiar tarczy** – grubość mikrometrem (min. 23,0 mm) i bicie czujnikiem zegarowym (max 0,05 mm) [1], [6].'),
  SUB('**5. Kontrola pedału hamulca** – wysokość, luz swobodny i odległość rezerwowa pedału według danych producenta [4].'),
  SUB('**6. Badanie na stanowisku rolkowym** – pomiar sił hamowania każdego koła, wyznaczenie wskaźnika skuteczności hamowania i różnicy sił na osi [15], [16]. Metoda daje obiektywny wynik liczbowy i jest wykorzystana w rozdziale 7 do kontroli jakości po naprawie.'),
  H2('4.3. Charakterystyka przypadku w pojeździe Toyota Corolla E12'),
  P('Na potrzeby projektu przyjęto typowy przypadek eksploatacyjny (założenie projektowe). Użytkownik zgłosił:'),
  BUL('wysoki pisk przy hamowaniu, ustępujący po zwolnieniu pedału,'),
  BUL('nieznacznie dłuższy skok pedału hamulca niż dotychczas.'),
  P('Po demontażu koła i zacisku dokonano pomiarów (wartości przyjęte do przykładu): grubość okładziny klocka wewnętrznego 1,6 mm, zewnętrznego 2,4 mm; grubość tarczy 24,3 mm; bicie tarczy 0,03 mm. Okładziny zbliżyły się do wartości granicznej 1,0 mm [1], a płytka wskaźnika zużycia zaczęła ocierać o tarczę. Tarcza mieści się w wymiarach dopuszczalnych (≥ 23,0 mm, bicie ≤ 0,05 mm) [1], [6], dlatego zakwalifikowano ją do dalszej eksploatacji. Różnica zużycia klocka wewnętrznego i zewnętrznego wskazuje na konieczność kontroli i smarowania sworzni prowadzących. Zdecydowano o wymianie kompletu klocków osi przedniej.'),
  P(`Na rysunku numer ${f5} przedstawiono ocenę zużycia klocków hamulcowych po demontażu zacisku.`, { keepNext: true }),
  photoBox(7000, 3400, 'Zużyte klocki po zdjęciu zacisku – kadr z poradnika / filmu AUTODOC [2], [3]'),
  CAP(`Rysunek ${f5}. Ocena zużycia okładzin klocków hamulcowych, [2]`),
];

// ---------- 5. NARZĘDZIA ----------
const f6 = fig();
const narz = [
  H1('5. Wykaz narzędzi i oprogramowania potrzebnych do przeprowadzenia procedury'),
  P('Prawidłowa wymiana przednich klocków hamulcowych wymaga właściwego przygotowania stanowiska, doboru narzędzi, przyrządów pomiarowych i środków chemicznych. Dobór oparto na instrukcji naprawy producenta [1] oraz instrukcji krok po kroku dla analizowanego modelu [2], [3]. Wykaz narzędzi wraz z rozmiarem, celem i uzasadnieniem zastosowania przedstawiono w tabeli numer 4.'),
  TCAP('Tabela 4. Wykaz narzędzi, przyrządów i materiałów potrzebnych do wymiany klocków hamulcowych, [1]–[3]'),
  table([3100, 3200, 3338], [
    ['Narzędzie / przyrząd (rozmiar, typ)', 'Cel', 'Uzasadnienie'],
    ['1) Podnośnik warsztatowy', 'Podniesienie przedniej części pojazdu.', 'Dostęp do układu hamulcowego.'],
    ['2) Podstawki (kozły) warsztatowe – 2 szt.', 'Zabezpieczenie pojazdu po podniesieniu.', 'Wymóg BHP – nie wolno pracować pod pojazdem podpartym tylko podnośnikiem.'],
    ['3) Kliny pod koła – 2 szt.', 'Zabezpieczenie kół tylnych przed przetoczeniem.', 'Ochrona podczas podnoszenia pojazdu.'],
    ['4) Nasadka udarowa 21 mm (sześciokątna) + pokrętło', 'Odkręcanie i wkręcanie nakrętek kół.', 'Rozmiar nakrętek kół Corolli E12 [2].'],
    ['5) Klucz płasko-oczkowy 13 mm', 'Odkręcenie śrub mocujących zacisk.', 'Rozmiar łba śrub zacisku [2].'],
    ['6) Klucz płasko-oczkowy 17 mm', 'Przytrzymanie (kontra) sworznia prowadzącego.', 'Zapobiega obracaniu sworznia i uszkodzeniu osłony [1], [2].'],
    ['7) Nasadka 13 mm + klucz dynamometryczny (zakres min. 20–120 N·m)', 'Dokręcenie śrub zacisku (34,3 N·m) i nakrętek kół (103 N·m).', 'Zachowanie momentów producenta [1].'],
    ['8) Przyrząd do cofania tłoczków zacisku', 'Wciśnięcie tłoczka do cylindra.', 'Umożliwia montaż grubszych, nowych klocków bez uszkodzenia osłony tłoczka.'],
    ['9) Łom montażowy, wkrętak płaski', 'Wyjęcie klocków i płytek podporowych.', 'Ułatwia demontaż zapieczonych elementów.'],
    ['10) Suwmiarka (dokładność 0,1 mm)', 'Pomiar grubości okładzin.', 'Porównanie z wartością min. 1,0 mm [1].'],
    ['11) Mikrometr zewnętrzny 0–25 mm (dokładność 0,01 mm)', 'Pomiar grubości tarczy.', 'Porównanie z wartością min. 23,0 mm [1].'],
    ['12) Czujnik zegarowy ze statywem magnetycznym', 'Pomiar bicia tarczy.', 'Porównanie z wartością max 0,05 mm [6].'],
    ['13) Przymiar liniowy (linijka stalowa)', 'Kontrola wysokości, luzu i odległości rezerwowej pedału.', 'Kontrola jakości po naprawie [4].'],
    ['14) Szczotka druciana', 'Czyszczenie jarzma, zacisku i piasty.', 'Usunięcie korozji zapewnia swobodny przesuw klocków.'],
    ['15) Zmywacz do układów hamulcowych', 'Odtłuszczenie tarczy i jarzma.', 'Zapewnia właściwy współczynnik tarcia.'],
    ['16) Środek penetrujący (np. WD-40)', 'Ułatwienie odkręcenia zapieczonych śrub.', 'Zmniejsza ryzyko zerwania gwintu [2].'],
    ['17) Smar litowy na bazie glikolu', 'Smarowanie sworzni prowadzących.', 'Środek wskazany przez producenta – nie niszczy gumy [1].'],
    ['18) Smar do hamulców tarczowych / pasta przeciwpiskowa', 'Podkładki przeciwpiskowe, styki klocek–jarzmo.', 'Ogranicza drgania i piski [1], [2].'],
    ['19) Smar miedziany', 'Powierzchnia przylegania felgi do piasty.', 'Zapobiega przywieraniu felgi [2].'],
    ['20) Rękawice i okulary ochronne', 'Ochrona rąk i oczu.', 'Pył z okładzin i płyn hamulcowy są szkodliwe.'],
  ], { size: 19 }),
  P('', { after: 60 }),
  SUB('**Oprogramowanie diagnostyczne – czy potrzebne?**'),
  P('W pojazdach z elektrycznym hamulcem postojowym (EPB) lub elektronicznym czujnikiem zużycia wymiana klocków wymaga użycia testera diagnostycznego (tryb serwisowy, cofanie tłoczka, kasowanie licznika zużycia). W analizowanym przypadku **oprogramowanie nie jest potrzebne**, ponieważ:'),
  BUL('wymiana dotyczy hamulców osi przedniej,'),
  BUL('Corolla E12 ma mechaniczny hamulec postojowy (bez EPB),'),
  BUL('wskaźnik zużycia klocków jest mechaniczny (akustyczny), a nie elektroniczny,'),
  BUL('tłoczek zacisku jest cofany ręcznym przyrządem mechanicznym.'),
  P(`Na rysunku numer ${f6} pokazano część narzędzi niezbędnych do przeprowadzenia procedury.`, { keepNext: true }),
  photoBox(7000, 3400, 'Zestaw narzędzi – kadr „Niezbędne narzędzia” z poradnika / filmu AUTODOC [2], [3]'),
  CAP(`Rysunek ${f6}. Część narzędzi niezbędnych do przeprowadzenia procedury, [2]`),
  P('Odpowiednie przygotowanie stanowiska, zastosowanie właściwych narzędzi i przestrzeganie momentów dokręcania gwarantują wysoką jakość usługi. Szczegółowy przebieg operacji przedstawiono w karcie instrukcyjnej w rozdziale 6.'),
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
      P(`Rysunek ${figNo}. ${s.f.charAt(0).toUpperCase() + s.f.slice(1)}, ${s.s}`, { size: 16, bold: true, before: 40, after: s.q ? 60 : 40, align: AlignmentType.LEFT }),
    ];
    if (s.q) kids.push(P(`„Niezbędne narzędzia: ${s.q}”`, { size: 19, after: 40, align: AlignmentType.LEFT }));
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
  H1('6. Karta instrukcyjna operacji wymiany klocków hamulcowych'),
  P(`Proces technologiczny podzielono na grupy czynności: przygotowanie, demontaż, weryfikację, przygotowanie do montażu, montaż, czynności końcowe i kontrolę jakości. Kolejność operacji wraz z punktami decyzyjnymi wynikającymi z wymiarów dopuszczalnych przedstawiono na rysunku numer ${flowNo}.`, { keepNext: true }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new ImageRun({
    type: 'png', data: fs.readFileSync(path.join(__dirname, 'schemat.png')), transformation: { width: 430, height: Math.round(430 * 2350 / 1664) } })] }),
  CAP(`Rysunek ${flowNo}. Schemat procesu technologicznego wymiany klocków hamulcowych z kontrolą jakości – opracowanie własne na podstawie [1], [4], [6], [15], [16]`),
  P('Szczegółowy przebieg poszczególnych czynności przedstawiono w karcie instrukcyjnej (tabela numer 5). Wymiary nominalne i dopuszczalne oraz momenty dokręcania podano w osobnej kolumnie przy każdej czynności, której dotyczą.'),
];
sheets.forEach((idxs, k) => {
  const rows = [sheetColHead(), ...idxs.map(i => stepRow(S[i], i, stepFig[i]))];
  const kids = [];
  kids.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: k === 0
    ? [new TextRun({ text: 'Tabela 5. Karta instrukcyjna operacji wymiany klocków hamulcowych, [1]–[6]', bold: true, font: FONT, size: 18 })] : [] }));
  kids.push(sheetHeader(k + 1));
  kids.push(new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: [KW[0], MID, KW[5]], rows }));
  kids.push(sheetFooter());
  karta.push(...kids);
});
karta.push(P('Na podstawie przedstawionych operacji można stwierdzić, że procedura wymiany przednich klocków hamulcowych w pojeździe Toyota Corolla E12 obejmuje nie tylko wymianę części, lecz także ich weryfikację z wymiarami dopuszczalnymi producenta. Zastosowanie właściwych narzędzi, przestrzeganie momentów dokręcania (34,3 N·m i 103 N·m) oraz właściwych środków smarnych pozwala na bezpieczne zakończenie naprawy. Poprawność wykonania sprawdzana jest w kontroli jakości opisanej w rozdziale 7.', { before: 240 }));

// ---------- 7. KONTROLA JAKOŚCI ----------
const kontrola = [
  H1('7. Kontrola jakości po naprawie'),
  P('Kontrola jakości ma na celu obiektywne potwierdzenie, że po wymianie klocków układ hamulcowy działa prawidłowo i spełnia wymagania producenta oraz przepisów. Składa się z kontroli statycznej (bez jazdy), badania na stanowisku rolkowym oraz jazdy próbnej. Wyniki zapisuje się w protokole (tabela numer 7), który stanowi dokument potwierdzający jakość usługi.'),
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
  P('gdzie T_max i T_min – większa i mniejsza z sił hamowania kół lewego i prawego jednej osi [15]. Przekroczenie tej wartości po wymianie klocków wskazuje najczęściej na zapieczony sworzeń prowadzący, nieprawidłowo osadzony klocek lub zanieczyszczoną tarczę. W takim przypadku proces wraca do etapu weryfikacji (rysunek numer ' + flowNo + ').'),
  P('Ponieważ pojazd badany bez obciążenia może osiągać blokowanie kół przy sile mniejszej od wymaganej, przepisy dopuszczają w takim przypadku wyznaczenie obliczeniowego wskaźnika skuteczności na podstawie nacisku na pedał [16]. Wymagania stosowane w kontroli jakości zestawiono w tabeli numer 6.'),
  TCAP('Tabela 6. Kryteria odbioru jakościowego naprawy, [1], [4], [5], [15], [16]'),
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
  P('Wyniki kontroli wpisuje się do protokołu przedstawionego w tabeli numer 7. Pojazd może zostać wydany klientowi wyłącznie wtedy, gdy wszystkie pozycje mają ocenę „spełnia”.'),
  TCAP('Tabela 7. Protokół kontroli jakości po wymianie klocków hamulcowych osi przedniej – opracowanie własne'),
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

// ---------- 8. KOSZTY ----------
const koszty = [
  H1('8. Koszty związane z przeprowadzoną procedurą'),
  P('W przypadku standardowej wymiany przednich klocków hamulcowych klient ponosi koszt części oraz robocizny. Materiały eksploatacyjne (smary, zmywacz, pasta przeciwpiskowa, rękawice) oraz narzędzia zapewnia warsztat w ramach usługi.'),
  H2('8.1. Koszt części – klocki hamulcowe (przód)'),
  P('Ceny kompletu klocków przednich TRW GDB3288 (z akustycznym wskaźnikiem zużycia) dla Toyoty Corolla E12 w wybranych sklepach (ceny brutto):'),
  BUL('TRW GDB3288 – 154,99 zł [12],'),
  BUL('TRW GDB3288 – 140,00 zł [18],'),
  BUL('TRW GDB3288 – 116,00 zł [19].'),
  P('Średnia cena kompletu klocków wynosi (154,99 + 140,00 + 116,00) / 3 ≈ **137 zł brutto**. Wartość tę przyjęto w dalszych obliczeniach.'),
  H2('8.2. Koszt robocizny'),
  P('Stawki za roboczogodzinę w warsztatach samochodowych w Polsce zależą od lokalizacji i specjalizacji. Według dostępnych zestawień wynoszą one 100–150 zł netto w mniejszych miejscowościach oraz 150–250 zł netto w dużych miastach [20]. Dla celów projektu przyjęto stawkę **200 zł netto** (246 zł brutto).'),
  P('Czas wymiany klocków jednej osi wynosi zwykle 0,5–1 roboczogodziny [21], a w normie stosowanej w egzaminie zawodowym dla mechaników przyjęto 1,0 rbh [22]. Ponieważ projektowany proces obejmuje dodatkowo pomiary tarcz i kontrolę jakości, przyjęto czas **1,0 rbh**.'),
  H2('8.3. Łączny koszt usługi'),
  TCAP('Tabela 8. Kalkulacja kosztu usługi – opracowanie własne na podstawie [12], [18]–[22]'),
  table([4600, 2519, 2519], [
    ['Składnik kosztu', 'Netto', 'Brutto (23% VAT)'],
    ['Komplet klocków przednich TRW GDB3288 (średnia cena)', '111,38 zł', '137,00 zł'],
    ['Robocizna: 1,0 rbh × 200 zł', '200,00 zł', '246,00 zł'],
    ['**Razem**', '**311,38 zł**', '**383,00 zł**'],
  ], { center: [1, 2] }),
  P('', { after: 60 }),
  P('Całkowity koszt wymiany przednich klocków hamulcowych (dla obu kół osi przedniej) wraz z kontrolą jakości w typowym warsztacie wynosi około **383 zł brutto**. W przypadku stwierdzenia w trakcie weryfikacji tarczy poniżej 23,0 mm koszt wzrasta o wymianę kompletu tarcz, co nie zostało ujęte w kalkulacji.'),
];

// ---------- 9. WNIOSKI ----------
const wnioski = [
  H1('9. Wnioski'),
  P('Opracowany proces technologiczny wymiany przednich klocków hamulcowych w pojeździe Toyota Corolla IX (E12) 1.6 VVT-i został w całości oparty na danych producenta: wymiarach nominalnych i dopuszczalnych (okładzina 11,0 / 1,0 mm, tarcza 25,0 / 23,0 mm, bicie 0,05 mm) oraz momentach dokręcania (34,3 N·m dla śrub zacisku i 103 N·m dla nakrętek kół). Rozmiary i typy kluczy (21 mm, 13 mm, 17 mm) przyjęto z instrukcji krok po kroku dla tego modelu.'),
  P('Włączenie do karty czynności pomiarowych sprawia, że proces nie ogranicza się do mechanicznej wymiany części, lecz obejmuje weryfikację – decyzja o dalszej eksploatacji tarczy, płytek podporowych i sworzni podejmowana jest na podstawie wymiarów granicznych. Procedura nie wymaga oprogramowania diagnostycznego, ponieważ pojazd ma mechaniczny hamulec postojowy i mechaniczny wskaźnik zużycia klocków.'),
  P('Rozdział poświęcony kontroli jakości zamyka proces mierzalną oceną: kontrolą pedału według danych producenta oraz badaniem na stanowisku rolkowym z wyznaczeniem wskaźnika skuteczności hamowania (≥ 50%, dla analizowanego pojazdu ΣT ≥ 8,12 kN) i różnicy sił na osi (≤ 30%). Protokół kontroli pozwala udokumentować jakość naprawy przed wydaniem pojazdu.'),
  P('Szacunkowy koszt usługi wynosi około 383 zł brutto. Projekt pozwolił pogłębić wiedzę z zakresu weryfikacji części, organizacji pracy w warsztacie, doboru narzędzi oraz oceny jakości wykonanej naprawy układu hamulcowego.'),
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
      ...karta, ...kontrola, ...koszty, ...wnioski, ...bib],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(OUT, b);
  console.log('ok', OUT, 'arkuszy:', NS, 'rysunków:', FIG, 'arkusze:', JSON.stringify(sheets.map(s => s.map(i => i + 1))));
  fs.writeFileSync(path.join(path.dirname(OUT), '.toc_titles.json'), JSON.stringify(TOC.map(t => t[0])));
});

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, VerticalAlign, PageOrientation,
  HeadingLevel, Header, Footer, PageNumber, LevelFormat, PageBreak, HeightRule,
} = require('docx');

const OUT = process.argv[2];
const FONT = 'Calibri';

// ---------- helpers ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };

function runs(text, opts = {}) {
  // supports **bold** fragments
  const parts = String(text).split(/(\*\*[^*]+\*\*)/);
  return parts.filter(Boolean).map(p =>
    p.startsWith('**')
      ? new TextRun({ text: p.slice(2, -2), bold: true, font: FONT, size: opts.size || 18, italics: opts.italics })
      : new TextRun({ text: p, bold: opts.bold, font: FONT, size: opts.size || 18, italics: opts.italics, color: opts.color }));
}
function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after === undefined ? 40 : opts.after },
    children: runs(text, opts),
  });
}
function cell(content, width, opts = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c =>
    typeof c === 'string' ? p(c, { size: opts.size, bold: opts.bold, align: opts.align }) : c);
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    columnSpan: opts.colSpan,
    rowSpan: opts.rowSpan,
    verticalAlign: opts.vAlign || VerticalAlign.CENTER,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: paras,
  });
}
function simpleTable(widths, rows, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0 && opts.header !== false,
      cantSplit: true,
      children: r.map((c, j) => cell(c, widths[j], {
        bold: i === 0 && opts.header !== false,
        fill: i === 0 && opts.header !== false ? 'D9E2F3' : undefined,
        size: opts.size || 18,
        align: i === 0 ? AlignmentType.CENTER : undefined,
      })),
    })),
  });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, font: FONT, size: 26 })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, bold: true, font: FONT, size: 22 })] });
}
function body(text) { return p(text, { size: 21, after: 100, align: AlignmentType.JUSTIFIED }); }
function bullet(text) {
  return new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 40 },
    children: runs(text, { size: 21 }) });
}
function caption(text) { return p(text, { size: 18, italics: true, after: 120 }); }

// ---------- źródła ----------
const SRC = `[1] Toyota Corolla (E120) 2002–2008 Repair Manual – Brake / Front brake / Overhaul, https://www.tcorolla.net/overhaul-1145.html`;

// ---------- KARTA – dane kroków ----------
// fig: opis zdjęcia do wklejenia (null = brak zdjęcia)
const steps = [
  { n: 'Przygotowanie stanowiska i pojazdu.',
    o: 'Pojazd ustawiony na równym, twardym podłożu stanowiska. Silnik wyłączony, hamulec postojowy zaciągnięty, włączony bieg (M/T) lub położenie P (A/T).',
    t: 'Brak specjalnych narzędzi.',
    par: '—',
    u: 'Wszystkie czynności wykonywać przy wyłączonym silniku [2].',
    fig: 'Pojazd Toyota Corolla E12 na stanowisku przed rozpoczęciem pracy (zdjęcie ogólne pojazdu z poradnika [2]).' },
  { n: 'Odkręcenie korka zbiornika płynu hamulcowego.',
    o: 'Maska silnika zostaje otwarta, a korek zbiornika płynu hamulcowego odkręcony (zdjęty lub poluzowany).',
    t: 'Ręcznie – brak specjalnych narzędzi.',
    par: 'Poziom płynu: między znakami MIN i MAX. Płyn: SAE J1703 lub FMVSS No. 116 DOT 3 [3].',
    u: 'Umożliwia cofnięcie tłoczka bez nadmiernego wzrostu ciśnienia w układzie. Płyn hamulcowy niszczy lakier – rozlany natychmiast zmyć wodą [1], [3].',
    fig: 'Otwarta maska i odkręcony korek zbiornika płynu hamulcowego [2].' },
  { n: 'Zabezpieczenie pojazdu przed przetoczeniem.',
    o: 'Pod koła tylnej osi zostają podłożone kliny.',
    t: 'Kliny pod koła (2 szt.).',
    par: '—',
    u: 'Kliny pozostają do końca pracy (czynność 32).',
    fig: 'Kliny podłożone pod koło tylne [2].' },
  { n: 'Wstępne poluzowanie nakrętek koła.',
    o: 'Przy pojeździe stojącym na kołach nakrętki mocujące koło przednie zostają poluzowane (bez całkowitego odkręcania).',
    t: 'Nasadka udarowa do kół 21 mm (sześciokątna) + pokrętło / klucz do kół [2].',
    par: 'Nakrętki kół – 4 szt. na koło (rozstaw 4×100) [7].',
    u: 'Luzować na kołach – piasta nie obraca się. Przytrzymywać koło, aby uniknąć urazu [2].',
    fig: 'Luzowanie nakrętek koła nasadką 21 mm [2].' },
  { n: 'Podniesienie pojazdu.',
    o: 'Przód pojazdu zostaje podniesiony podnośnikiem, a następnie zabezpieczony podstawkami (kozłami) warsztatowymi.',
    t: 'Podnośnik (np. hydrauliczny żabka lub dwukolumnowy), podstawki warsztatowe (kozły) 2 szt.',
    par: 'Podparcie w punktach wskazanych przez producenta w instrukcji obsługi pojazdu.',
    u: 'Nie wolno wykonywać pracy wyłącznie na podnośniku – pojazd musi stać na kozłach.',
    fig: 'Moment podniesienia pojazdu i ustawienia na podstawkach [2].' },
  { n: 'Demontaż koła.',
    o: 'Nakrętki zostają odkręcone, koło zdjęte z piasty i odłożone.',
    t: 'Nasadka udarowa 21 mm + pokrętło [2].',
    par: '—',
    u: 'Przytrzymywać koło przy odkręcaniu ostatniej nakrętki [2].',
    fig: 'Układ hamulcowy po zdemontowaniu koła – widoczny zacisk, jarzmo i tarcza [2].' },
  { n: 'Oczyszczenie zacisku i jego elementów mocujących.',
    o: 'Zacisk, śruby mocujące i sworznie prowadzące zostają oczyszczone szczotką drucianą i spryskane środkiem penetrującym.',
    t: 'Szczotka druciana, środek penetrujący (WD-40) [2].',
    par: '—',
    u: 'Po nałożeniu środka odczekać kilka minut [2].',
    fig: 'Czyszczenie elementów mocujących zacisk szczotką drucianą [2].' },
  { n: 'Odkręcenie śrub zacisku.',
    o: 'Przytrzymując sworzeń prowadzący (kontra), zostają wykręcone 2 śruby mocujące zacisk (korpus cylindra) do sworzni prowadzących [1].',
    t: 'Klucz płasko-oczkowy 13 mm (śruba) + klucz płasko-oczkowy 17 mm (kontra sworznia prowadzącego) [2].',
    par: 'Moment dokręcenia (do montażu): 34,3 N·m [1].',
    u: 'Sworzeń musi być przytrzymany kluczem 17 mm – inaczej obraca się razem ze śrubą i może uszkodzić osłonę [1].',
    fig: 'Odkręcanie śrub zacisku kluczami 13 mm i 17 mm (kontra) [2].' },
  { n: 'Demontaż zacisku.',
    o: 'Zacisk zostaje zdjęty z jarzma i podwieszony na drucie/haku, bez odłączania przewodu hamulcowego.',
    t: 'Drut lub hak do podwieszenia zacisku.',
    par: '—',
    u: 'Nie wolno dopuścić do zwisania zacisku na przewodzie hamulcowym. Nie naciskać pedału hamulca przy zdjętym zacisku – tłoczek może wypaść [2].',
    fig: 'Zdemontowany zacisk podwieszony na drucie [2].' },
  { n: 'Demontaż klocków hamulcowych.',
    o: 'Zużyte klocki (2 szt.) zostają wyjęte z jarzma razem z podkładkami przeciwpiskowymi nr 1 i nr 2 [1].',
    t: 'Ręcznie lub łom montażowy [2].',
    par: '—',
    u: 'Zapamiętać położenie klocka z płytką wskaźnika zużycia.',
    fig: 'Wyjmowanie zużytych klocków z jarzma [2].' },
  { n: 'Demontaż płytek podporowych klocków.',
    o: 'Z jarzma (mocowania zacisku) zostają wyjęte 2 płytki podporowe klocków (pad support plates) [1].',
    t: 'Ręcznie / wkrętak płaski.',
    par: '—',
    u: 'Płytki odkładać tak, aby nie pomylić strony górnej i dolnej.',
    fig: 'Jarzmo zacisku po wyjęciu płytek podporowych (rysunek z instrukcji [1] lub zdjęcie z [2]).' },
  { n: 'Weryfikacja klocków – pomiar grubości okładziny.',
    o: 'Grubość okładziny ciernej (bez płytki nośnej) zostaje zmierzona na obu klockach, w kilku miejscach.',
    t: 'Suwmiarka (dokł. 0,1 mm) lub przymiar liniowy [1].',
    par: '**Wymiar nominalny: 11,0 mm.**\n**Wymiar dopuszczalny (min.): 1,0 mm** [1].',
    u: 'Przy grubości < 1,0 mm klocki wymienić. Klocki wymienia się zawsze kompletem na oś [2]. Nierównomierne zużycie → sprawdzić sworznie prowadzące (czynność 16).',
    fig: 'Pomiar grubości okładziny klocka (rysunek z instrukcji [1]).' },
  { n: 'Weryfikacja płytek podporowych.',
    o: 'Płytki podporowe zostają sprawdzone pod kątem sprężystości, odkształceń, pęknięć i zużycia oraz oczyszczone z rdzy i brudu [1].',
    t: 'Ocena wzrokowa; szczotka druciana.',
    par: 'Kryterium: brak odkształceń, pęknięć, zachowana sprężystość [1].',
    u: 'Płytki uszkodzone lub zużyte wymienić na nowe.',
    fig: 'Kontrola płytek podporowych (rysunek z instrukcji [1]).' },
  { n: 'Weryfikacja tarczy – pomiar grubości.',
    o: 'Grubość tarczy zostaje zmierzona w kilku punktach na obwodzie, na pasie roboczym.',
    t: 'Mikrometr zewnętrzny (np. 0–25 mm / 25–50 mm, dokł. 0,01 mm) [1].',
    par: '**Wymiar nominalny: 25,0 mm.**\n**Wymiar dopuszczalny (min.): 23,0 mm** [1]. Tarcza wentylowana Ø 275 mm [7].',
    u: 'Tarcza poniżej 23,0 mm lub z głębokimi rowkami / przegrzaniem – do wymiany (kompletem na oś). Grubość tarczy sprawdzać przy każdej wymianie klocków.',
    fig: 'Pomiar grubości tarczy mikrometrem (rysunek z instrukcji [1]).' },
  { n: 'Weryfikacja bicia tarczy (przy objawach bicia pedału/kierownicy).',
    o: 'Tarcza zostaje tymczasowo przykręcona nakrętkami kół; czujnik zegarowy ustawiony 10 mm od zewnętrznej krawędzi tarczy; tarcza obracana o pełny obrót [5].',
    t: 'Czujnik zegarowy (dokł. 0,01 mm) na statywie magnetycznym; nasadka 21 mm + klucz dynamometryczny.',
    par: 'Nakrętki tymczasowe: 103 N·m. **Bicie dopuszczalne (max): 0,05 mm**. Przy przekroczeniu: luz łożyska piasty max 0,05 mm, bicie piasty max 0,05 mm [5].',
    u: 'Gdy łożysko i piasta są w normie – przestawić tarczę w położenie o najmniejszym biciu lub przetoczyć tarczę [5].',
    fig: 'Pomiar bicia tarczy czujnikiem zegarowym (rysunek z instrukcji [5]).' },
  { n: 'Weryfikacja sworzni prowadzących, osłon i tłoczka.',
    o: 'Sworznie prowadzące zostają wyjęte z jarzma i sprawdzone (korozja, zatarcie, swoboda ruchu). Sprawdzone zostają osłony gumowe sworzni oraz osłona tłoczka (pęknięcia, wycieki płynu) [1], [2].',
    t: 'Ręcznie; ocena wzrokowa.',
    par: 'Kryterium: swobodny przesuw sworzni, osłony bez pęknięć, brak wycieku spod osłony tłoczka.',
    u: 'Uszkodzone osłony / sworznie wymienić. Wyciek płynu spod osłony tłoczka → regeneracja lub wymiana zacisku (poza zakresem tej karty) [1].',
    fig: 'Kontrola sworzni prowadzących i osłon [2].' },
  { n: 'Czyszczenie jarzma (mocowania zacisku).',
    o: 'Miejsca osadzenia płytek podporowych i klocków w jarzmie zostają oczyszczone z rdzy i nalotu.',
    t: 'Szczotka druciana, zmywacz do układów hamulcowych [2].',
    par: '—',
    u: 'Po użyciu zmywacza odczekać kilka minut do odparowania [2].',
    fig: 'Czyszczenie jarzma zacisku [2].' },
  { n: 'Smarowanie i montaż sworzni prowadzących.',
    o: 'Części ślizgowe i powierzchnie uszczelniające sworzni zostają pokryte smarem, a sworznie wsunięte w jarzmo [1].',
    t: 'Smar litowy na bazie glikolu (lithium soap base glycol grease) wg [1].',
    par: '—',
    u: 'Nie stosować smarów mineralnych na elementy gumowe. Sprawdzić swobodny przesuw sworzni po montażu.',
    fig: 'Smarowanie sworzni prowadzących (rysunek z instrukcji [1]).' },
  { n: 'Montaż płytek podporowych.',
    o: '2 płytki podporowe (oczyszczone lub nowe) zostają osadzone w jarzmie [1].',
    t: 'Ręcznie.',
    par: '—',
    u: 'Płytki muszą być osadzone pewnie, bez luzu.',
    fig: 'Zamontowane płytki podporowe w jarzmie (rysunek z instrukcji [1]).' },
  { n: 'Cofnięcie tłoczka zacisku.',
    o: 'Tłoczek zostaje wciśnięty do cylindra zacisku, aby zrobić miejsce dla grubszych nowych klocków.',
    t: 'Przyrząd do cofania tłoczków zacisku [2].',
    par: '—',
    u: 'Kontrolować poziom płynu w zbiorniku (korek odkręcony w czynności 2), aby płyn się nie przelał. Nie uszkodzić osłony tłoczka.',
    fig: 'Cofanie tłoczka zacisku przyrządem [2].' },
  { n: 'Przygotowanie nowych klocków.',
    o: 'Podkładki przeciwpiskowe nr 1 zostają posmarowane smarem do hamulców tarczowych i założone na klocki wraz z podkładkami nr 2 [1]. Pasta przeciwpiskowa zostaje nałożona na miejsca styku klocka z jarzmem [2].',
    t: 'Smar do hamulców tarczowych, pasta przeciwpiskowa [1], [2].',
    par: 'Klocki np. TRW GDB3288: 131,7 × 57,4 × 17,8 mm, akustyczny wskaźnik zużycia [6].',
    u: 'W razie potrzeby wymienić komplet podkładek przeciwpiskowych [1]. Nie nanosić smaru na okładzinę cierną.',
    fig: 'Nakładanie pasty przeciwpiskowej na klocki [2].' },
  { n: 'Czyszczenie powierzchni tarczy.',
    o: 'Powierzchnie robocze tarczy zostają odtłuszczone.',
    t: 'Zmywacz do układów hamulcowych [2].',
    par: '—',
    u: 'Po aplikacji sprayu odczekać kilka minut [2].',
    fig: 'Czyszczenie powierzchni tarczy zmywaczem [2].' },
  { n: 'Montaż nowych klocków.',
    o: 'Nowe klocki zostają osadzone w jarzmie okładziną w stronę tarczy; klocek z płytką wskaźnika zużycia montowany wskaźnikiem do góry [1], [2].',
    t: 'Ręcznie.',
    par: '—',
    u: 'Sprawdzić, czy klocki przesuwają się swobodnie w płytkach podporowych.',
    fig: 'Montaż nowych klocków hamulcowych [2].' },
  { n: 'Montaż zacisku hamulcowego.',
    o: 'Zacisk zostaje nałożony na klocki i przykręcony 2 śrubami do sworzni prowadzących, przy przytrzymaniu sworznia kluczem [1].',
    t: 'Klucz dynamometryczny z nasadką 13 mm + klucz płasko-oczkowy 17 mm (kontra) [2].',
    par: '**Moment: 34,3 N·m (350 kgf·cm; 25 ft·lbf)** [1]. (W [2] podano 35 N·m.)',
    u: 'Śruby dokręcać kluczem dynamometrycznym. Przewód hamulcowy nie może być skręcony.',
    fig: 'Montaż zacisku i dokręcanie śrub kluczem dynamometrycznym [2].' },
  { n: 'Czyszczenie piasty – powierzchni przylegania felgi.',
    o: 'Powierzchnia przylegania obręczy koła do piasty/tarczy zostaje oczyszczona i pokryta cienką warstwą smaru.',
    t: 'Szczotka druciana, smar miedziany [2].',
    par: '—',
    u: 'Brak.',
    fig: 'Czyszczenie miejsca przylegania felgi koła [2].' },
  { n: 'Wykonanie czynności 4–25 po drugiej stronie osi.',
    o: 'Analogiczna wymiana klocków po przeciwnej stronie pojazdu.',
    t: 'Jak w czynnościach 4–25.',
    par: 'Jak w czynnościach 4–25.',
    u: 'Klocki wymienia się kompletem na oś [2].',
    fig: null },
  { n: 'Montaż koła.',
    o: 'Koło zostaje nałożone na piastę, a nakrętki wkręcone i wstępnie dokręcone.',
    t: 'Nasadka 21 mm + pokrętło [2].',
    par: '—',
    u: 'Przytrzymywać koło przy wkręcaniu nakrętek [2].',
    fig: 'Montaż koła i wkręcanie nakrętek [2].' },
  { n: 'Opuszczenie pojazdu i dokręcenie nakrętek koła.',
    o: 'Pojazd zostaje zdjęty z podstawek i opuszczony; nakrętki dokręcone momentem nominalnym, na krzyż.',
    t: 'Klucz dynamometryczny + nasadka do kół 21 mm [2].',
    par: '**Moment: 103 N·m (1050 kgf·cm; 76 ft·lbf)** [1], [2].',
    u: 'Dokręcać po przekątnej (na krzyż) [2].',
    fig: 'Dokręcanie nakrętek koła kluczem dynamometrycznym [2].' },
  { n: 'Przywrócenie skoku pedału hamulca.',
    o: 'Pedał hamulca zostaje kilkukrotnie naciśnięty, aż do wyczucia oporu (dosunięcie tłoczków i klocków do tarcz).',
    t: 'Brak.',
    par: '—',
    u: 'Wykonać przy wyłączonym silniku, przed pierwszym ruszeniem pojazdu [2].',
    fig: 'Naciskanie pedału hamulca [2].' },
  { n: 'Kontrola poziomu płynu hamulcowego.',
    o: 'Poziom płynu zostaje sprawdzony i w razie potrzeby uzupełniony; korek zbiornika zakręcony, maska zamknięta.',
    t: 'Ręcznie.',
    par: 'Poziom: MIN–MAX. Płyn SAE J1703 / FMVSS No. 116 DOT 3 [3].',
    u: 'Rozlany płyn natychmiast zmyć [3].',
    fig: 'Kontrola poziomu płynu i zakręcanie korka zbiornika [2].' },
  { n: 'Kontrola pedału hamulca (weryfikacja po naprawie).',
    o: 'Zostaje sprawdzona wysokość pedału, luz swobodny pedału oraz odległość rezerwowa pedału [4].',
    t: 'Przymiar liniowy (linijka stalowa).',
    par: 'Wysokość pedału od podłogi: M/T 134,9–144,9 mm, A/T 136,0–146,0 mm. Luz swobodny: 1–6 mm. Odległość rezerwowa: > 70 mm przy sile 490 N, silnik pracuje [4].',
    u: 'Luz mierzyć po kilkukrotnym naciśnięciu pedału przy wyłączonym silniku (usunięcie podciśnienia ze wspomagania) [4].',
    fig: null },
  { n: 'Demontaż klinów i jazda próbna.',
    o: 'Kliny zostają usunięte; jazda próbna ze sprawdzeniem skuteczności hamowania.',
    t: 'Brak.',
    par: 'Docieranie: pierwsze 150–200 km bez gwałtownego hamowania [2].',
    u: 'Sprawdzić, czy pojazd nie ściąga przy hamowaniu i czy nie występują piski.',
    fig: 'Usunięcie klinów spod kół [2].' },
];

// ---------- KARTA – tabela ----------
const KW = [700, 2250, 3700, 2950, 3000, 2538]; // = 15138
const KT = KW.reduce((a, b) => a + b, 0);

function photoCell(text, figNo, width) {
  return new TableCell({
    borders, columnSpan: 4, width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      p(`Na rysunku numer ${figNo} poniżej pokazano: ${text.charAt(0).toLowerCase() + text.slice(1)}`, { size: 19, after: 60 }),
      new Table({
        width: { size: 6200, type: WidthType.DXA }, columnWidths: [6200],
        rows: [new TableRow({ height: { value: 1700, rule: HeightRule.ATLEAST }, children: [new TableCell({
          width: { size: 6200, type: WidthType.DXA },
          borders: { top: { style: BorderStyle.DASHED, size: 6, color: '7F7F7F' }, bottom: { style: BorderStyle.DASHED, size: 6, color: '7F7F7F' }, left: { style: BorderStyle.DASHED, size: 6, color: '7F7F7F' }, right: { style: BorderStyle.DASHED, size: 6, color: '7F7F7F' } },
          shading: { fill: 'F2F2F2', type: ShadingType.CLEAR, color: 'auto' },
          verticalAlign: VerticalAlign.CENTER,
          children: [p('[ MIEJSCE NA ZDJĘCIE – wkleić zrzut ze źródła podanego w podpisie ]', { size: 17, italics: true, color: '7F7F7F', align: AlignmentType.CENTER })],
        })] })],
      }),
      p(`Rysunek ${figNo}. ${text}`, { size: 16, bold: true, before: 40 }),
    ],
  });
}

const headRow = new TableRow({ tableHeader: true, cantSplit: true, children:
  ['Lp.', 'Nazwa zabiegu lub czynności', 'Opis czynności', 'Narzędzie / przyrządy (rozmiar i typ klucza)', 'Wymiary nominalne / dopuszczalne, moment dokręcania', 'Uwagi / zalecenia specjalistów']
    .map((t, i) => cell(t, KW[i], { bold: true, fill: 'D9E2F3', size: 17, align: AlignmentType.CENTER })) });

const kartaRows = [headRow];
let fig = 1;
const MID = KW[1] + KW[2] + KW[3] + KW[4];
steps.forEach((s, i) => {
  const parParas = s.par.split('\n').map(line => p(line, { size: 17 }));
  const inner = new Table({ width: { size: MID, type: WidthType.DXA }, columnWidths: [KW[1], KW[2], KW[3], KW[4]], rows: [
    new TableRow({ cantSplit: true, children: [
      cell(s.n, KW[1], { size: 18, align: AlignmentType.CENTER }),
      cell(s.o, KW[2], { size: 17 }),
      cell(s.t, KW[3], { size: 17, align: AlignmentType.CENTER }),
      new TableCell({ borders, width: { size: KW[4], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: parParas }),
    ] }),
    ...(s.fig ? [new TableRow({ cantSplit: true, children: [photoCell(s.fig, fig++, MID)] })] : []),
  ] });
  kartaRows.push(new TableRow({ cantSplit: true, children: [
    cell(`${i + 1}.`, KW[0], { size: 18, vAlign: VerticalAlign.TOP, bold: true }),
    new TableCell({ borders, columnSpan: 4, width: { size: MID, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [inner, new Paragraph({ spacing: { after: 0 }, children: [] })] }),
    cell(s.u, KW[5], { size: 17, align: AlignmentType.CENTER }),
  ] }));
});
const kartaTable = new Table({ width: { size: KT, type: WidthType.DXA }, columnWidths: KW, rows: kartaRows });

// ---------- nagłówek/stopka karty ----------
const HW = [3300, 8638, 1600, 1600];
const kartaHeader = new Header({ children: [
  new Table({ width: { size: KT, type: WidthType.DXA }, columnWidths: HW, rows: [
    new TableRow({ children: [
      cell(['[Nazwa uczelni]', '[Wydział]', '[Zakład / Katedra]'], HW[0], { rowSpan: 2, size: 18, align: AlignmentType.CENTER }),
      cell('Nazwa zespołu, podzespołu, części:', HW[1], { size: 18, align: AlignmentType.CENTER }),
      cell('Arkusz', HW[2], { size: 18, align: AlignmentType.CENTER }),
      new TableCell({ borders, width: { size: HW[3], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] }),
    ] }),
    new TableRow({ children: [
      cell('**Układ hamulcowy – przód (Toyota Corolla IX E12, 1.6 VVT-i 3ZZ-FE, ZZE121)**', HW[1], { size: 18, align: AlignmentType.CENTER }),
      cell('Arkuszy', HW[2], { size: 18, align: AlignmentType.CENTER }),
      new TableCell({ borders, width: { size: HW[3], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.TOTAL_PAGES_IN_SECTION], font: FONT, size: 18 })] })] }),
    ] }),
  ] }),
  p('', { after: 60 }),
] });

const FW = Array(9).fill(Math.floor(KT / 9)); FW[8] = KT - FW[0] * 8;
const kartaFooter = new Footer({ children: [
  new Table({ width: { size: KT, type: WidthType.DXA }, columnWidths: FW, rows: [
    new TableRow({ children: ['Opracował', 'Podpis', 'Data', 'Sprawdził', 'Podpis', 'Data', 'Zatwierdził', 'Podpis', 'Data'].map((t, i) => cell(t, FW[i], { size: 16, align: AlignmentType.CENTER })) }),
    new TableRow({ height: { value: 400, rule: HeightRule.ATLEAST }, children: FW.map(w => cell('', w, { size: 16 })) }),
  ] }),
] });

const pageFooter = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] });

// ---------- część opisowa ----------
const PW = 11906 - 2 * 1134; // A4 portret, marginesy 2 cm -> 9638

const titlePage = [
  p('', { after: 1200 }),
  p('[Nazwa uczelni]', { size: 28, bold: true, align: AlignmentType.CENTER }),
  p('[Wydział / Zakład]', { size: 24, align: AlignmentType.CENTER, after: 1200 }),
  p('PROJEKT PROCESU TECHNOLOGICZNEGO NAPRAWY', { size: 32, bold: true, align: AlignmentType.CENTER, after: 200 }),
  p('Weryfikacja i wymiana klocków hamulcowych osi przedniej', { size: 28, align: AlignmentType.CENTER, after: 120 }),
  p('Toyota Corolla IX (E12) Hatchback 1.6 VVT-i (3ZZ-FE, 81 kW / 110 KM), kod ZZE121', { size: 24, bold: true, align: AlignmentType.CENTER, after: 1600 }),
  p('Autor: ............................................................', { size: 22, align: AlignmentType.CENTER, after: 120 }),
  p('Nr albumu / grupa: ............................................', { size: 22, align: AlignmentType.CENTER, after: 120 }),
  p('Prowadzący: ....................................................', { size: 22, align: AlignmentType.CENTER, after: 120 }),
  p('Rok akademicki: 2026/2027', { size: 22, align: AlignmentType.CENTER }),
  new Paragraph({ children: [new PageBreak()] }),
];

const intro = [
  h1('1. Cel i zakres projektu'),
  body('Celem projektu jest opracowanie procesu technologicznego weryfikacji i naprawy przedniego układu hamulcowego pojazdu Toyota Corolla IX (E12), polegającego na wymianie klocków hamulcowych osi przedniej. Projekt obejmuje: identyfikację obiektu naprawy, zestawienie wymiarów nominalnych i dopuszczalnych elementów podlegających weryfikacji, zestawienie momentów dokręcania wraz z rozmiarem i typem kluczy, wykaz narzędzi i materiałów oraz kartę instrukcyjną operacji (kartę technologiczną) z opisem każdej czynności.'),
  body('Wszystkie wartości liczbowe (wymiary, luzy, bicia, momenty dokręcania) zaczerpnięto z instrukcji naprawy producenta – Toyota Corolla (E120) Repair Manual [1], [3], [4], [5]. Rozmiary kluczy oraz kolejność czynności warsztatowych wraz ze zdjęciami dla tego modelu – z instrukcji krok po kroku AUTODOC CLUB dla Toyota Corolla IX Hatchback (E120) [2]. Dane katalogowe części – z katalogów producentów części [6], [7]. Każda wartość w karcie jest opatrzona numerem źródła.'),
  body('Zakres karty: wymiana klocków bez odłączania przewodu hamulcowego i bez odpowietrzania układu. Wymiana tarcz, regeneracja zacisku i odpowietrzanie nie wchodzą w zakres karty – podano jedynie kryteria, przy których należy je wykonać.'),

  h1('2. Identyfikacja obiektu naprawy'),
  simpleTable([3600, 6038], [
    ['Parametr', 'Wartość'],
    ['Marka, model, generacja', 'Toyota Corolla IX (E12), nadwozie hatchback'],
    ['Wersja silnikowa / kod typu', '1.6 VVT-i, silnik 3ZZ-FE, 1598 cm³, 81 kW (110 KM); kod ZZE121 [8]'],
    ['Okres produkcji wersji', '10.2001 – 03.2008 (wg katalogu [8]; daty zależne od nadwozia)'],
    ['Hamulec przedni', 'Tarczowy, zacisk pływający jednotłoczkowy na 2 sworzniach prowadzących (budowa wg [1])'],
    ['Tarcza hamulcowa przednia', 'Wentylowana, Ø 275 mm, grubość nominalna 25,0 mm, min. 23,0 mm [1], [7]'],
    ['Klocki hamulcowe przednie (przykład)', 'TRW GDB3288 – 131,7 × 57,4 × 17,8 mm, akustyczny wskaźnik zużycia; nr OE m.in. Toyota 04465-02061, 04465-02130 [6]'],
    ['Płyn hamulcowy', 'SAE J1703 lub FMVSS No. 116 DOT 3 [3]'],
    ['Mocowanie koła', '4 nakrętki, rozstaw 4×100 [7]'],
  ]),
  caption('Tabela 1. Dane identyfikacyjne pojazdu i podzespołu'),
  body('Uwaga: średnica tarczy przedniej w Corolli E12 zależy od wersji i wyposażenia (w katalogach występują tarcze Ø 255 mm i Ø 275 mm). Przed rozpoczęciem pracy należy potwierdzić wersję po numerze VIN lub po wymiarach zamontowanej tarczy.'),

  h1('3. Wymiary nominalne i dopuszczalne – dane do weryfikacji'),
  simpleTable([2900, 1500, 1700, 2100, 1438], [
    ['Element / parametr', 'Wymiar nominalny', 'Wymiar dopuszczalny', 'Przyrząd pomiarowy', 'Źródło'],
    ['Grubość okładziny klocka przedniego', '11,0 mm', 'min. 1,0 mm', 'Suwmiarka / przymiar', '[1]'],
    ['Grubość tarczy przedniej', '25,0 mm', 'min. 23,0 mm', 'Mikrometr zewnętrzny', '[1]'],
    ['Bicie osiowe tarczy (10 mm od krawędzi)', '—', 'max 0,05 mm', 'Czujnik zegarowy + statyw', '[5]'],
    ['Luz łożyska piasty przedniej', '—', 'max 0,05 mm', 'Czujnik zegarowy', '[5]'],
    ['Bicie piasty przedniej', '—', 'max 0,05 mm', 'Czujnik zegarowy', '[5]'],
    ['Płytki podporowe klocków', 'bez odkształceń i pęknięć', 'zachowana sprężystość', 'Ocena wzrokowa', '[1]'],
    ['Wysokość pedału hamulca od podłogi', 'M/T 134,9–144,9 mm; A/T 136,0–146,0 mm', '—', 'Przymiar liniowy', '[4]'],
    ['Luz swobodny pedału hamulca', '1–6 mm', '—', 'Przymiar liniowy', '[4]'],
    ['Odległość rezerwowa pedału (490 N, silnik pracuje)', '—', '> 70 mm', 'Przymiar liniowy', '[4]'],
    ['Luz włącznika świateł STOP', '0,5–2,4 mm', '—', 'Szczelinomierz / przymiar', '[4]'],
  ]),
  caption('Tabela 2. Wymiary nominalne i dopuszczalne elementów podlegających weryfikacji'),

  h1('4. Momenty dokręcania oraz rozmiar i typ kluczy'),
  simpleTable([2900, 2400, 1400, 1500, 1438], [
    ['Połączenie', 'Rozmiar i typ klucza', 'Moment [N·m]', 'Moment [kgf·cm] / [ft·lbf]', 'Źródło'],
    ['Nakrętki kół', 'Nasadka udarowa 21 mm (6-kątna) + klucz dynamometryczny', '103', '1050 / 76', '[1], [2]'],
    ['Śruby zacisku do sworzni prowadzących (2 szt.)', 'Nasadka 13 mm + klucz dynamometryczny; kontra: klucz płasko-oczkowy 17 mm', '34,3', '350 / 25', '[1], [2]'],
    ['Śruby jarzma do zwrotnicy (2 szt.) – tylko przy demontażu jarzma/tarczy', 'Nie demontowane w zakresie karty', '106,8', '1089 / 79', '[1]'],
    ['Śruba przewodu elastycznego (banjo) – tylko przy odłączaniu przewodu', 'Nie demontowane w zakresie karty', '29', '296 / 21', '[1]'],
    ['Nakrętka kontrująca widełek popychacza pedału', 'Tylko przy regulacji wysokości pedału', '26', '265 / 19', '[4]'],
  ]),
  caption('Tabela 3. Momenty dokręcania połączeń gwintowych'),
  body('Uwaga: w instrukcji AUTODOC [2] moment dokręcenia śrub zacisku podano jako 35 N·m – jest to wartość zaokrąglona z wartości producenta 34,3 N·m (350 kgf·cm) [1]. W karcie przyjęto wartość producenta. Rozmiary kluczy podano wg [2]; producent pojazdu w instrukcji naprawy nie podaje rozmiarów kluczy.'),

  h1('5. Wykaz narzędzi, przyrządów i materiałów'),
  h2('5.1. Narzędzia i przyrządy'),
  bullet('Podnośnik warsztatowy i podstawki (kozły) warsztatowe – 2 szt.; kliny pod koła – 2 szt.'),
  bullet('Nasadka udarowa do kół 21 mm + pokrętło / klucz do kół [2].'),
  bullet('Klucz płasko-oczkowy 13 mm oraz klucz płasko-oczkowy 17 mm [2].'),
  bullet('Nasadka 13 mm [2].'),
  bullet('Klucz dynamometryczny o zakresie obejmującym 30–110 N·m.'),
  bullet('Przyrząd do cofania tłoczków zacisku [2].'),
  bullet('Łom montażowy, wkrętak płaski, szczotka druciana [2].'),
  bullet('Suwmiarka (0,1 mm), mikrometr zewnętrzny (0,01 mm), czujnik zegarowy ze statywem magnetycznym, przymiar liniowy [1], [4], [5].'),
  h2('5.2. Materiały i części'),
  bullet('Komplet klocków hamulcowych osi przedniej (np. TRW GDB3288 lub oryginał Toyota 04465-02061 / 04465-02130) [6].'),
  bullet('Smar litowy na bazie glikolu do sworzni prowadzących (lithium soap base glycol grease) [1].'),
  bullet('Smar do hamulców tarczowych (podkładki przeciwpiskowe) [1]; pasta przeciwpiskowa [2].'),
  bullet('Smar miedziany (powierzchnia przylegania felgi) [2].'),
  bullet('Zmywacz do układów hamulcowych; środek penetrujący (WD-40) [2].'),
  bullet('Płyn hamulcowy SAE J1703 / FMVSS No. 116 DOT 3 – do ewentualnego uzupełnienia [3].'),
  h2('5.3. Wymagania BHP'),
  bullet('Pracować wyłącznie na pojeździe ustawionym na podstawkach; nie przebywać pod pojazdem podpartym tylko podnośnikiem.'),
  bullet('Stosować rękawice ochronne i okulary – pył z okładzin i płyn hamulcowy są szkodliwe; nie wydmuchiwać pyłu sprężonym powietrzem.'),
  bullet('Płyn hamulcowy niszczy lakier – rozlany zmywać natychmiast [1], [3].'),
  bullet('Nie naciskać pedału hamulca przy zdjętym zacisku [2].'),
];

// ---------- część końcowa ----------
const endPart = [
  h1('7. Uwagi końcowe i sposób uzupełnienia rysunków'),
  body('Karta zawiera ramki „MIEJSCE NA ZDJĘCIE” z opisem, które zdjęcie należy wkleić i z którego źródła. Zdjęcia wykonawcze dla tego modelu zawiera instrukcja AUTODOC CLUB „Jak wymienić przednie klocki hamulcowe w TOYOTA Corolla IX Hatchback (E120)” [2] (dostępna jako strona WWW, plik PDF i film). Rysunki pomiarów (grubość okładziny, grubość i bicie tarczy, smarowanie sworzni, płytki podporowe) zawiera instrukcja naprawy producenta [1], [5]. Pod każdym rysunkiem należy pozostawić numer źródła.'),
  body('Pełny, oficjalny dostęp do instrukcji naprawy Toyota (Repair Manual) zapewnia portal Toyota Motor Europe dla niezależnych warsztatów [9] (dostęp płatny, czasowy). Zaleca się porównanie wartości z Tabel 2 i 3 z tym źródłem przed oddaniem projektu.'),

  h1('8. Bibliografia'),
  p('[1] Toyota Motor Corporation: Toyota Corolla (E120) 2002–2008 Repair Manual – Brake / Front brake / Overhaul. Online: https://www.tcorolla.net/overhaul-1145.html (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[2] AUTODOC CLUB: How to change front brake pads on TOYOTA Corolla IX Hatchback (E120) – replacement guide. Online: https://club.autodoc.co.uk/manuals/how-to-change-front-brake-pads-on-toyota-corolla-ix-hatchback-e120-replacement-guide-25473 (wersja polska: club.autodoc.pl; PDF wersji sedan: https://club.autodoc.co.uk/pdf-manuals/club/toyota/corolla-saloon-e12j-e12t/brake-pads/pdf/EN-how-to-change-front-brake-pads-on-toyota-corolla-saloon-e12j-e12t-replacement-guide.pdf) (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[3] Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Brake fluid. Online: https://www.tcorolla.net/brake_fluid-1131.html (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[4] Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Adjustment – Brake (brake pedal). Online: https://www.tcorolla.net/adjustment-1133.html (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[5] Toyota Motor Corporation: Toyota Corolla (E120) Repair Manual – Replacement – Front axle hub (disc runout, bearing backlash). Online: https://www.tcorolla.net/replacement-1121.html (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[6] TRW (ZF Aftermarket): klocki hamulcowe GDB3288 – dane katalogowe. Online: https://intercars.pl/produkty/321869-klocki-hamulcowe-trw-automotive-gdb3288 oraz https://www.autodoc.pl/trw/2192809 (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[7] Tarcza hamulcowa przednia Toyota Corolla E12 Ø 275 mm (ATE 24.0125-0168.1, TRW DF4366, A.B.S. 17544) – dane katalogowe. Online: https://motostacja.com/p/trw-2df4366gdb3288/ ; https://www.winparts.co.uk/brake-system/brake-discs/c632/brake-disc-sport-brake-disc-coat-z-590-2576-52-zimmermann/p846973.html (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[8] Schaeffler – katalog zastosowań: TOYOTA COROLLA (_E12_) 1.6 VVT-i (ZZE121_). Online: https://shop.rolling.hu/catalog/SCHAEFFLER/toyota-corolla-saloon-e12-16-vvt-i-zze121 (dostęp: 09.10.2026).', { size: 20, after: 80 }),
  p('[9] Toyota Motor Europe: portal informacji technicznej dla niezależnych operatorów. Online: https://www.toyota-tech.eu (dostęp: 09.10.2026).', { size: 20, after: 80 }),
];

// ---------- dokument ----------
const doc = new Document({
  creator: 'Projekt – karta technologiczna',
  title: 'Karta technologiczna – wymiana klocków hamulcowych – Toyota Corolla E12',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] }] },
  sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      footers: { default: pageFooter },
      children: [...titlePage, ...intro] },
    { properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
        margin: { top: 1900, bottom: 1500, left: 850, right: 850, header: 400, footer: 300 }, pageNumbers: { start: 1 } } },
      headers: { default: kartaHeader }, footers: { default: kartaFooter },
      children: [
        p('6. Karta instrukcyjna operacji wymiany klocków hamulcowych osi przedniej', { size: 24, bold: true, after: 60 }),
        p('Tabela 4. Karta instrukcyjna operacji wymiany klocków hamulcowych – Toyota Corolla IX (E12) 1.6 VVT-i, oś przednia [1]–[7]', { size: 18, italics: true, after: 80 }),
        kartaTable,
      ] },
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      footers: { default: pageFooter },
      children: endPart },
  ],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('ok', OUT, `figures: ${fig - 1}`); });

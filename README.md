# Projekt procesu technologicznego naprawy – wymiana przednich klocków hamulcowych

**Pojazd:** Toyota Corolla IX (E12) Hatchback 1.6 VVT-i (silnik 3ZZ-FE, 81 kW / 110 KM), kod ZZE121
**Plik projektu:** `Projekt_wymiana_klockow_Toyota_Corolla_E12.docx` (35 stron, układ jak we wzorze z zajęć)

**Temat dla prowadzącego:** Jakub Suszek – Proces technologiczny weryfikacji i wymiany przednich klocków hamulcowych wraz z kontrolą jakości naprawy – Toyota Corolla IX (E12) 1.6 VVT-i

## Zawartość dokumentu
1. Wstęp · 2. Opis pojazdu (Tabela 1) · 3. Układ hamulcowy (3.1 wymiary nominalne/dopuszczalne, 3.2 momenty + klucze)
4. Diagnoza (4.1 objawy, 4.2 metody, 4.3 przypadek) · 5. Narzędzia i oprogramowanie (Tabela 4)
6. Karta instrukcyjna – schemat procesu + 16 arkuszy, 32 czynności
7. **Kontrola jakości po naprawie** (wyróżnik: pedał wg Toyoty, stanowisko rolkowe ≥50% / ≤30%, protokół)
8. Koszty (≈383 zł brutto) · 9. Wnioski · Bibliografia (23 pozycje)

## Kluczowe dane (z instrukcji naprawy Toyota)
| Parametr | Wartość nominalna | Wartość dopuszczalna | Źródło |
|---|---|---|---|
| Grubość okładziny klocka przedniego | 11,0 mm | min. 1,0 mm | [overhaul-1145](https://www.tcorolla.net/overhaul-1145.html) |
| Grubość tarczy przedniej | 25,0 mm | min. 23,0 mm | [overhaul-1145](https://www.tcorolla.net/overhaul-1145.html) |
| Bicie tarczy (10 mm od krawędzi) | – | max 0,05 mm | [replacement-1121](https://www.tcorolla.net/replacement-1121.html) |
| Śruby zacisku (klucz 13 mm + kontra 17 mm) | 34,3 N·m | – | [overhaul-1145](https://www.tcorolla.net/overhaul-1145.html) |
| Nakrętki kół (nasadka 21 mm) | 103 N·m | – | [overhaul-1145](https://www.tcorolla.net/overhaul-1145.html) |
| Śruby jarzma (poza zakresem karty) | 106,8 N·m | – | [overhaul-1145](https://www.tcorolla.net/overhaul-1145.html) |
| Wysokość pedału / luz / rezerwa | M/T 134,9–144,9 mm; luz 1–6 mm; > 70 mm | – | [adjustment-1133](https://www.tcorolla.net/adjustment-1133.html) |

Rozmiary kluczy i zdjęcia krok po kroku pochodzą z poradnika AUTODOC CLUB dla Corolla IX Hatchback (E120):
https://club.autodoc.co.uk/manuals/how-to-change-front-brake-pads-on-toyota-corolla-ix-hatchback-e120-replacement-guide-25473
(wersja polska na club.autodoc.pl). Producent pojazdu nie podaje rozmiarów kluczy.

## Do zrobienia przed oddaniem
- [ ] Wpisać nazwę uczelni, wydziału i zakładu w nagłówku karty i na stronie tytułowej.
- [ ] Wkleić zdjęcia w ramki „MIEJSCE NA ZDJĘCIE” (w każdej ramce jest opis, co wkleić i skąd).
- [ ] Otworzyć każdy link z bibliografii i porównać z nim wartości z tabel 2 i 3.
- [ ] Po wklejeniu zdjęć sprawdzić numery stron w spisie treści (są wpisane ręcznie).

`narzedzia/build_projekt.js` generuje plik .docx (biblioteka `docx` dla Node.js), `narzedzia/schemat.html` → `schemat.png` to schemat procesu.

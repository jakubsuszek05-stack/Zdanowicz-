# Projekt procesu technologicznego naprawy – wymiana przednich klocków hamulcowych

**Pojazd:** Toyota Corolla IX (E12) Hatchback 1.6 VVT-i (silnik 3ZZ-FE, 81 kW / 110 KM), kod ZZE121
**Plik projektu:** `Karta_technologiczna_wymiana_klockow_Toyota_Corolla_E12.docx`

## Zawartość dokumentu
1. Cel i zakres
2. Identyfikacja obiektu naprawy
3. Wymiary nominalne i dopuszczalne (tabela)
4. Momenty dokręcania i rozmiar oraz typ kluczy (tabela)
5. Narzędzia, materiały, BHP
6. Karta instrukcyjna: 32 czynności, 30 ramek na zdjęcia (format jak we wzorze z zajęć)
7. Uwagi końcowe
8. Bibliografia

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
- [ ] Wkleić zdjęcia w 30 ramek „MIEJSCE NA ZDJĘCIE” (opis zdjęcia i źródło są pod każdą ramką).
- [ ] Otworzyć każdy link z bibliografii i porównać z nim wartości z tabel 2 i 3.
- [ ] W Wordzie zaznaczyć wszystko (Ctrl+A) i nacisnąć F9, żeby odświeżyć liczbę arkuszy w nagłówku.

`narzedzia/build_karta.js` to skrypt, który generuje plik .docx (biblioteka `docx` dla Node.js).

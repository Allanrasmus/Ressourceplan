# Testplan
- Bevar original baseline og generér ny fil med versionsnavn. Kontrollér ZIP-integritet, XML-parsing, openpyxl genindlæsning, gyldige formler, celletyper, tabelreferencer og synlige filterpile. Undersøg tidligere Excel-reparationsfejl; ingen ugyldige fejlværdier. Undgå at forveksle strukturkontrol med faktisk Excel-beregning.
- Kontrollér dropdowns, indtastningsfelter, skjult ID og beregningsrækker, kvartals-/datoformler, formateringsregler for alle typer, tilbudsprocentens visning og projektvælger.
- Test 37 normtimer, 20 Sag, 7 Ferie, 10 Tilbud (50 %): kapacitet 30, booket 20, ledig 10, belægning 66,7 %, brutto 10, vægtet 5. Andet tæller som booket; Fravær/Kursus reducerer kapacitet. Test også nul kapacitet med positiv belastning og ledig/overbelastning separat for to personer.
- Test projektoversigt med flere medarbejdere, ændring af type, bevarelse af eksisterende tilbudstimer og navneskift. Test 2026 uge 53, 2027 og 2028, årsskifte og filtrering. Årsvalg må ikke ændre lagrede timer. Ingen uge 53 i år uden den. Kontrollér at summer ikke brydes ved filtrering; hvis sortering er usikker, dokumentér det.
- Test udvidelse/aktivering af ledig medarbejderplads, eksisterende datavalidering og bevarelse af baseline. Dokumentér hvordan man udvider over 30 medarbejdere og 2031.
- Brug automatiseret test af logik og, hvor tilgængeligt, en faktisk beregningsmotor; openpyxl beregner ikke formler. Sæt genberegning ved åbning. Manuel accept: åbning uden reparationsdialog, direkte timeindtastning, filtre, skift år og type, farver, nøgletal, samtidig redigering og personlige arkvisninger i Excel Desktop/Online. Angiv særskilt hvad der ikke er testet.

## Testjournal for 3.3

Den 24. november 2026 bestod alle 9 automatiske tests. Beregningstestene
omfattede 7 genberegnede filer og fandt ingen celler med formelfejl. Kørslen
brugte Python 3.12.14, openpyxl 3.1.5 og LibreOffice Calc 26.8 (udviklingsbuild).
I et nyt virtuelt miljø med de fastlåste pakker bestod desuden alle 4
strukturtests, herunder sammenligning med den udgivne fil.

Kør `python -m unittest discover -s tests -v` fra projektmappen. Testene bruger
fiktive data og ændrer kun midlertidige kopier.

| Tjek | Dækning |
| --- | --- |
| Fil og migration | ZIP/OOXML, formulatokens, tabeloverskrifter, filtermetadata, datavalideringer, genindlæsning og genbygning. SHA-256 for originalen og sammenligning af alle inputfelter og registre. |
| Standardvisning og formatering | Gemt filter for 2026 og Medarbejder 1, 12 synlige planrækker, skjult ID, fastlåste kolonner, skjulte beregninger og regler for farver og procentvisning. |
| Referenceberegning | 20 booket, 30 disponibel, 10 ledig, 66,7 % belægning, 10 brutto og 5 vægtede tilbudstimer i uge 48. |
| Aktivitet og prognose | Andet, Fravær, Kursus, nul kapacitet, særskilt ledig/overbelastning for to personer, blank og eksplicit nul tilbudsprocent, samt område-, faggruppe- og afdelingsfiltrering. |
| Projekt og historik | Omdøbning, aktivering af p06, vundet tilbud flyttet til en anden person med uændret projekttotal, og særskilte historiske årsværdier. |
| ISO-uger | Samtlige mandage og kvartaler i 2026, 2027 og 2028. Uge 53 med i 2026, udeladt i 2027 og 2028 selv med indsatte værdier. Datoer udeladt ved visning af flere år. |

LibreOffice Calc genberegner testfilerne, og testen læser beregnede værdier tilbage
og kontrollerer regnearksfejl. Filen til udgivelse er openpyxl-builden.

### Manuel accept før drift

Åbn den udgivne fil i Excel Desktop og Excel Online, og kontrollér åbning uden
reparationsdialog, synlige filterpile, valg af ét år og navneskift i filtret,
timeindtastning, dropdowns, typefarver og skjult tilbudsprocent. Prøv derefter
samtidig redigering og personlige arkvisninger med to brugere i Teams/SharePoint.
Strukturtests verificerer filter- og formateringsreglerne i filen. Den visuelle
gengivelse og brugerinteraktionen i Microsoft Excel kræver denne manuelle accept.

## Testjournal for 3.3.1

Den 24. november 2026 bestod alle 7 nye tests og alle 9 tidligere tests.
De nye beregningstests omfattede 11 scenarier med LibreOffice Calc 26.8
(udviklingsbuild) uden celler med formelfejl. De tre nye strukturtests bestod
desuden med openpyxls alternative XML-skriver uden lxml.

Genbyg med `python migrate_3_3_1.py --output <ny-fil>`. Kør hele testpakken med
`python -m unittest discover -s tests -v`, eller brug
`python -m unittest discover -s tests -p test_ressourceplan_3_3_1.py -v` til
de nye regressionskontroller.

| Kontrol | Dækning |
| --- | --- |
| OOXML | ZIP, XML-dele, tabelhenvisninger, vinduesvisninger, valg i frosne ruder, gyldige DXF-indeks og regelprioriteter. Differentialfarver kontrolleres i den rå XML som `patternFill/bgColor` uden `patternType` eller `fgColor`. Den almindelige cellestil er særskilt. |
| Data og standardvisning | SHA-256 af 3.2 og 3.3. Cellevis sammenligning af plan og registre. Gemt medarbejder- og årsfilter, frossen rude, skjult ID, 12 synlige planrækker og skjulte beregninger. |
| Aktivitet og belægning | Alle seks aktivitetstyper, Sag → Tilbud → Sag, 0 %, 50 %, 85 %, 100 % og 118 %, samt positiv booking uden kapacitet. Regeludtryk genberegnes på testkopier sammen med kalender og periodemarkering. |
| Projekt | To sagsnumre, flere medarbejdere, udeladelse af medarbejdere uden periodetimer samt kontrol af medarbejder-, uge- og periodetotal. |
| Kalender | 2026–2028, alle fire kvartaler, 4/8/12/26 uger, årsskifte og hele året med 52 og 53 uger. Uge 48–51 i de første fire kolonner og blanke værdier efter periodens afslutning. |

### Kontrol af Excel-filens struktur og testkopier

Ved sammenligning med 3.3 blev en manglende workbook-visning fundet, selv om
arkene henviste til visning 0. 3.3.1 tilføjer visningen og gemmer markeringer i
de tilsvarende fastlåste ruder. De rapporterede farver var gemt med
`patternType="solid"` og `fgColor` i DXF. 3.3.1 anvender samme
baggrundsstruktur som flere af farverne i 3.2 og en uafhængig prøve genereret
af XlsxWriter 3.2.9. Planens procentregel ligger før
farvereglerne og tillader, at de kombineres.

En særskilt skemakontrol med Apache POI 5.4.1 og XMLBeans 5.3.0 accepterede
10 SpreadsheetML-dele i hver af 3.3 og 3.3.1 (workbook, styles, fem ark og tre
tabeller). Det forklarer, hvorfor XML-skemaet alene kan overse problemet fra
accepttesten. Den præcise årsag til Excels rapporterede reparationsproblem
kan først bekræftes med Excel og eventuelt dets reparationslog.

Calc genberegnede og gemte separate kopier. Ved sammenligning ændrede Calc
blandt andet DXF-mønsteret til `solid` med `fgColor` og ændrede de interne
XML-filnavne for tabellerne. Udgivelsesfilen kommer direkte fra openpyxl.
Calc-kopierne bruges til at læse beregnede resultater.

### Manuel accept i Microsoft Excel for 3.3.1

1. Åbn den frigivne fil i Excel Desktop. Kontrollér, at der ikke kommer en
   reparationsdialog. Hvis Excel tilbyder reparation, gem loggen med
   oplysninger om den berørte pakkedel. Åbn også filen i Excel Online.
2. I **Min ressourceplan**, kontrollér startfilter samt direkte indtastning,
   farver for alle typer og blå/hvide sagsstriber. Skift Sag → Tilbud → Sag
   i samme række og kontrollér farve og tilbudsprocent ved hvert skift.
3. I **Ressourceoverblik**, kontrollér belægningsfarverne ved grænserne og ved
   Over, kvartalsfarver, orange periodemarkering, læsbare overskrifter,
   procent i ugekolonnernes I ALT og skjulte beregninger.
4. I **Projektoversigt**, vælg 2026, uge 48 og 4 uger. Kontrollér uge 48–51
   fra venstre, at medarbejdere uden timer er udeladt, og at timerne summerer.
   Skift sagsnummer, år og periode og rul vandret for at kontrollere projektvalget.
5. Kontrollér navn- og sagsopslag fra begge registre, og prøv personlige
   arkvisninger og samtidig redigering i Teams/SharePoint.

De automatiske resultater dokumenterer pakkestruktur, beregnede værdier og
betingelsesudtryk. Visuel gengivelse og åbning uden reparationsdialog
kræver manuel accept i Microsoft Excel.

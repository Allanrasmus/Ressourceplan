# Ressourceplan
Excel-baseret ressourceplanlægning for medarbejdere, ledere og projektledere. HTML v2.8 er design- og funktionsreference. Version 3.4 bygger videre på den manuelt testede 3.3.1. Se AGENTS.md, Projektkrav.md, DESIGN.md, TESTPLAN.md og CHANGELOG.md. Brug kun fiktive data og godkendte tjenester.

## Brug version 3.4

Åbn `Ressourceplan_3_4.xlsx` i Microsoft Excel. I Teams eller SharePoint opretter
du en personlig arkvisning via **Vis → Arkvisning → Ny**, før du ændrer filtrene.
Filtre styrer visningen og giver alle med adgang til filen mulighed for at se data.
Ved første åbning er planen filtreret på Medarbejder 1 og 2026. Brug tabelpilene i
**Planår** og **Medarbejder**, og indtast timer direkte i ugekolonnerne. Vælg ét
planår for at se kvartaler og mandagsdatoer. Uge 53 er grå og udeladt af summen i
år med 52 ISO-uger. Indsætning via udklipsholderen kan omgå Excels datavalidering.

- **Leder** udfylder Type og Sagsnr. i medarbejderens ledige række i det relevante
  år. En blank tilbudsprocent tæller som 50 %. Angiv `0 %` eksplicit for nul
  sandsynlighed. Ved typeskift skjules den tidligere procent og ignoreres i
  beregningerne. Ryd den manuelt, hvis den ikke skal genbruges ved et senere skift
  tilbage til Tilbud.
- **Ved vundet tilbud** aftales flytningen med modtageren. Find en ledig række
  hos modtageren for samme planår. Gem en midlertidig kopi af kildens planår,
  sagsnummer, procent og ugefelter, så du kan gendanne kilden ved en fejl. Ryd så
  sagsnummer, procent og timer i kilden, og skift kildens type til **Sag**.
  Indsæt derefter **kun værdierne** for Sagsnr. (E) og ugefelterne (I–BI) hos
  modtageren, og vælg typen **Sag** dér. Rydning før indsætning undgår
  dobbeltregning. Kontrollér i projekt- og ressourceoversigten, at timerne står
  én gang, som bookede timer hos modtageren. Gentag pr. planår ved behov.
- **Ressourceoverblik** har år, startuge, periode, område, faggruppe og afdeling som
  valg. Nøgletallene viser antal medarbejdere, periodens bookede, disponible, ledige
  og overbelastede timer, tilbud brutto/vægtet samt belægning. Årsoversigten
  samler de valgte medarbejdere og markerer periodeugerne med orange. **I ALT** følger sidste person og viser
  periodetimer i timekolonnerne og samlet belægning i ugekolonnerne. Ledige timer
  og overbelastning summeres hver for sig pr. medarbejder og uge.
- **Projektoversigt** har en fastlåst dropdown med formatet `Projektnavn · Sagsnr.`.
  Listen følger Projektregister. Ved omdøbning viser **Aktuelt projekt** straks
  det nye navn, og beregningen følger fortsat sagsnummeret. Vælg projektet igen
  for at opdatere den gemte tekst i selve dropdown-feltet. Perioden starter i første ugekolonne, og medarbejdere med timer
  står samlet øverst. Timerne omfatter Sag og Tilbud. Kolonnen **Periodetimer (t)**
  viser totalen pr. medarbejder og samlet. Ugekolonnerne viser ugetotalerne i
  **I ALT** direkte under sidste medarbejder. Opret projekter i en ledig række
  i Projektregister, og behold altid sagsnummeret ved navneændringer.

Et kvartal følger mandagens måned. Derfor hører uge 1 i nogle ISO-år til Q4,
eksempelvis uge 1 i 2026, som starter 29. december 2025. Ressourceoverblikket
afkorter perioden ved udgangen af planåret. Projektoversigtens faste ugeperioder
fortsætter over årsskiftet, når næste planår findes i filen. Kvartaler og
**Hele året** følger de uger, der hører til det valgte ISO-planår.

### Rækker, historik og udvidelse

Brug filtrering i Min ressourceplan. Undgå sortering i både Min ressourceplan og
Medarbejdere. Beregningerne forudsætter faste medarbejderblokke, og den skjulte
ID-kolonne, planår, opslagsformler og placering må ikke redigeres. Hvis H6 viser
STOP, skal rækkeplaceringen kontrolleres mod en tidligere intakt fil. Det er
muligt, at en sortering inden for én medarbejders blok ikke udløser advarslen.

Aktivér en ledig medarbejderplads ved at udfylde navn, område, faggruppe,
afdeling og normtid på den tildelte ID-række i Medarbejdere. De 12 planrækker
pr. år er allerede oprettet for alle 30 ID'er. Behold også tidligere medarbejderes
ID og registerdata, hvis deres historik skal fremgå af rapporterne. Ændringer af
normtid og andre registerfelter slår igennem i alle år, da registeret kun har ét
sæt felter pr. medarbejder.

Udvidelse over 30 personer eller 2031 kræver en separat migration med en ny
versionsfil. Arbejd på en kopi af den aktuelle driftsfil, og match eksisterende
indtastninger på ID, planår og række inden for året. Opret nye ID'er eller år
med egne planrækker. Ved flere år udvides hver medarbejders blok, og alle faste
blokgrænser skal genberegnes. Ved flere medarbejdere udvides register, blokke og
begge rapporters synlige rækker, totaler og skjulte beregninger. Opdatér også
tabeller, filtre, årslister, formelområder, valideringer og formatering. Gentag
tests af databevarelse, sidste eksisterende og første nye ID/år samt uge 53, før
en ny version tages i brug.

## Byg og test

Kør fra projektmappen med Python 3.12. Brug et nyt outputfilnavn ved genbygning.

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
build_dir=$(mktemp -d)
python migrate_3_4.py --source Ressourceplan_3_3_1.xlsx --output "$build_dir/Ressourceplan_3_4.xlsx"
python -m unittest discover -s tests -v
```

Migrationen afviser at overskrive eksisterende filer. Testene sammenligner
XML-indholdet i en ny build med den udgivne fil. Sammenligningen ser bort fra
tidsstempler og forskelle i, hvordan ens XML skrives. På Windows kan du bruge
`.venv\Scripts\python.exe` og angive en ny filsti med `--output`.

LibreOffice Calc på PATH (`soffice` eller `libreoffice`) bruges til
beregningstestene. Disse springes tydeligt over, når motoren mangler. Calc kører
kun på midlertidige testkopier. Den udgivne fil skrives med openpyxl og
genberegnes ved åbning. Slutbrugere behøver kun Excel.

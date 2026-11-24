# Ressourceplan
Excel-baseret ressourceplanlægning for medarbejdere, ledere og projektledere. HTML v2.8 er design- og funktionsreference. Version 3.5 bygger videre på den manuelt testede 3.4. Se AGENTS.md, Projektkrav.md, DESIGN.md, TESTPLAN.md og CHANGELOG.md. Brug kun fiktive data og godkendte tjenester.

## Brug version 3.5

Åbn `Ressourceplan_3_5.xlsx` i Microsoft Excel. I Teams eller SharePoint opretter
du en personlig arkvisning via **Vis → Arkvisning → Ny**, før du ændrer filtrene.
Filtre styrer visningen og giver alle med adgang til filen mulighed for at se data.
Ved første åbning er planen filtreret på Medarbejder 1 og 2026. Brug tabelpilene i
**Planår** og **Medarbejder**, og indtast timer direkte i ugekolonnerne. Vælg ét
planår for at se kvartaler og mandagsdatoer. Uge 53 er grå og udeladt af summen i
år med 52 ISO-uger. Indsætning via udklipsholderen kan omgå Excels datavalidering.

- **Leder** vælger aktivitetstype og ét **Projektvalg** i formatet
  `Projektnavn · Sagsnr.`. Sag og Tilbud kan begge vælge et projekt. Ferie,
  Fravær, Kursus og Andet bruger aktivitetstypen uden projekt. En tom
  tilbudsprocent betyder 50 %. Ved typeskift gemmes en eventuel tidligere
  procent, men den indgår kun i beregningen for Tilbud.
- **Projektvalg** viser den tekst, der blev valgt. Feltet **Projekt / aktivitet**
  viser det aktuelle projektnavn. Omdøbninger ændrer dropdown-listen og
  navneopslaget, mens timer stadig følger sagsnummeret. Vælg samme projekt igen,
  hvis du vil opdatere teksten i selve dropdown-feltet.
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

### Tilføj en planrække

1. Højreklik i en synlig datarække **i** tabellen `PlanTimer` på Min ressourceplan.
   Vælg **Indsæt → Tabelrækker ovenfor**, hvis Excel tilbyder handlingen.
   Mens filtrene er aktive, skal du stadig vælge nøglefelterne udtrykkeligt.
2. Udfyld **Planår**, **Medarbejdervalg** og **Type** i den indsatte række.
   Kontrollér den viste **Medarbejder** og planåret. Vælg eventuelt et projekt
   under **Projektvalg**, og indtast derefter timerne. Medarbejdervalget viser
   navn sammen med den stabile kode, så ens navne kan skelnes. Den beregnede
   ID-kolonne er skjult.
3. Genanvend de relevante filtre, og kontrollér at rækken kommer frem under
   korrekt medarbejder og år. Hvis indsættelse er utilgængelig under filter,
   eller filteret viser nul rækker, så ryd filtrene midlertidigt i din egen
   arkvisning. Tilføj rækken inde i tabellen. Du kan også bruge
   **Tabeldesign → Tilpas størrelse på tabel** til at medtage næste datarække
   på tværs af tabellens fulde bredde til kolonne BM. Genanvend filtrene efter
   indtastningen.

Sortér hele tabellen med en pil i tabeloverskriften. Hver række bærer sine egne
nøgler og ugentlige timer. Feltet **Medarbejder** viser det aktuelle navn fra
registeret. Genanvend eller opdatér medarbejderfilteret efter en omdøbning.
Planår og medarbejdervalg skal være udfyldt, før en række tæller i oversigterne.

### Tilføj en medarbejder eller et projekt

- **Medarbejdere:** Udfyld en ledig række med et unikt, varigt ID, f.eks.
  `p31`, og registrér navn, område, faggruppe, afdeling og normtid. Ved brug af
  en plads med et forudfyldt ID beholdes det ID. Personen fremgår af
  Medarbejdervalg og af Ressourceoverblik, når filteret passer. Projektoversigt
  viser personen, når vedkommende har timer på det valgte projekt i perioden.
- **Projektregister:** Udfyld et nyt, entydigt sagsnummer og et projektnavn
  i en ledig registerrække. Begge projektvælgere opdaterer listen automatisk.
  To projekter kan have samme navn, fordi sagsnummeret adskiller dem.

Behold registrerede ID'er og sagsnumre, også når navnene ændres. Behold de
registre, der knytter sig til historiske plantimer. Ændring af normtid og
øvrige medarbejderoplysninger slår også igennem på historiske år, da de læses
fra det aktuelle register. Ved kopiering med udklipsholderen er det særligt
vigtigt at kontrollere de unikke nøgler, da indsætning kan omgå Excels
datavalidering.

Planens Excel-tabel kan få flere rækker uden ændring af beregningsformlerne.
Den leverede formatering og validering dækker 20.000 datarækker. Arbejdsbogen
har plads til 100 medarbejdere, 500 projekter og planårene 2026–2031. Udvid
register- og visningskapaciteten i en ny version ved behov. Migreringen bevarer
indtastningerne fra 3.4 og fjerner de tomme rækker med genkendelig fabriksopsætning.

Modellen kan flytte eksisterende timer ved at ændre rækkens medarbejdervalg.
Ændring fra Tilbud til Sag flytter timerne fra tilbudspipeline til booket tid.
En dedikeret overførselsfunktion kan komme i en senere version.

## Byg og test

Kør fra projektmappen med Python 3.12. Brug et nyt outputfilnavn ved genbygning.

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
build_dir=$(mktemp -d)
python migrate_3_5.py --source Ressourceplan_3_4.xlsx --output "$build_dir/Ressourceplan_3_5.xlsx"
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

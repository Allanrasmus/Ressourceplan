# Ressourceplan
Excel-baseret ressourceplanlægning for medarbejdere, ledere og projektledere. HTML v2.8 er design- og funktionsreference. Version 3.5.1 retter fejl fra den manuelle accepttest af 3.5. Se AGENTS.md, Projektkrav.md, DESIGN.md, TESTPLAN.md og CHANGELOG.md. Brug kun fiktive data og godkendte tjenester.

## Brug version 3.5.1

Åbn `Ressourceplan_3_5_1.xlsx` i Microsoft Excel. I Teams eller SharePoint opretter
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

### Opret medarbejder og første ressourceplan

1. Åbn **Medarbejdere**. Udfyld en ledig række med navn, område, faggruppe,
   afdeling og normtid. Behold pladsens forudfyldte ID, eller angiv et nyt,
   entydigt ID i en række uden ID, f.eks. `p31`.
2. Åbn **Planstart** og vælg **Planår** og **Medarbejder**. Registerets dropdown
   viser også personer, der endnu ikke har planrækker. Koden efter navnet
   gør medarbejdere med samme navn entydige.
3. Planstart viser kun manglende rækker blandt en ledig **Sag**, **Ferie**,
   **Fravær**, **Kursus** og **Andet**. **Tilbud** oprettes efter behov.
   Formlerne viser et forslag. Følg trinnene nedenfor for at oprette de
   redigerbare rækker. I en delt fil er Planstarts valg fælles, så kontrollér
   år og medarbejder i de viste rækker, før du kopierer dem.

### Indsæt standardrækker eller opret et nyt planår

1. Gør plads til det viste antal tomme datarækker **inde i** `PlanTimer`.
   Ryd om nødvendigt planens filtre midlertidigt i din egen arkvisning.
   Højreklik i en datarække og vælg **Indsæt → Tabelrækker ovenfor**.
2. Kopiér kun de udfyldte **A:B**-celler i Planstarts resultattabel.
   Indsæt **som værdier** i A:B i de nye planrækker.
3. Kopiér **D** for de samme hjælperækker. Indsæt **som værdier** i D i
   de samme planrækker. Lad den skjulte kolonne C være urørt. Planstart
   viser nu, at standardrækkerne findes.
4. Genanvend filteret på det valgte planår og medarbejder i planen.
   Ved et helt nyt år, f.eks. 2032, skal du først tilføje året i en
   ledig celle under **Tilgængelige planår** på Planstart. Gentag
   for de medarbejdere, der skal have planen for det nye år.

### Tilføj en ekstra projektrække under filter

Højreklik i en synlig datarække i `PlanTimer` og vælg **Indsæt → Tabelrækker
ovenfor**, hvis Excel tilbyder valget under de aktive filtre. Angiv **Planår**,
**Medarbejder** og **Type** udtrykkeligt, vælg **Projektvalg**, og indtast
timerne. Hvis der vises nul rækker eller indsættelse er utilgængelig, ryd
filtrene i din egen arkvisning. Indsæt i tabellen og filtrér igen bagefter.
Du kan også bruge **Tabeldesign → Tilpas størrelse på tabel** til at medtage
næste datarække på tværs af tabellens fulde bredde til kolonne BM.

Filterpilen viser værdier, der allerede findes i tabelrækkerne. Andre aktive
filtre kan begrænse listen. Ryd medarbejder- og eventuelle øvrige filtre, og
behold planåret, når du vil se alle personer med rækker for året. Brug Planstart
til personer, som endnu ikke har rækker. Sortér hele tabellen fra en tabelpil,
så medarbejder, projekt og timer følges ad.

### Omdøb medarbejdere og projekter

Behold altid medarbejder-ID og sagsnummer ved omdøbning. Nyoprettede projekter
og medarbejdere kommer automatisk i dropdown-listerne. Eksisterende
medarbejdervalg og projektvalg i planen er gemt som tekst, mens beregningerne
fortsat bruger de stabile nøgler. Oversigterne viser det aktuelle navn.
Genvælg det nye navn på eksisterende rækker for at opdatere den tekst, som
planens medarbejderfilter viser. ID-kolonnen og det tekniske navneopslag
er skjult.

Arbejdsbogen har plads til 100 medarbejdere, 500 projekter og 20 konfigurerbare
planår. Formatering og validering dækker 20.000 planrækker. De seks oprindelige
planår 2026–2031 er allerede konfigureret. Behold registre, der knytter sig til
historiske timer. Ændringer af normtid og øvrige medarbejderoplysninger slår
igennem på historiske visninger, da de læses fra det aktuelle register.
Indsætning via udklipsholderen kan omgå Excels datavalidering.

## Byg og test

Kør fra projektmappen med Python 3.12. Brug et nyt outputfilnavn ved genbygning.

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
build_dir=$(mktemp -d)
python migrate_3_5_1.py --source Ressourceplan_3_5.xlsx --output "$build_dir/Ressourceplan_3_5_1.xlsx"
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

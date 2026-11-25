# Ressourceplan
Excel-baseret ressourceplanlægning for medarbejdere, ledere og projektledere. HTML v2.8 er design- og funktionsreference. Version 3.6.3 leveres som en komplet Excel-fil med demonstrationsdata, tre driftsscripts og et opgraderingsscript til eksisterende arbejdsfiler. Grupperingen leveres som en særskilt prototype. Se AGENTS.md, Projektkrav.md, DESIGN.md, TESTPLAN.md og CHANGELOG.md. Brug kun fiktive data og godkendte tjenester.

## Office Scripts 3.6.4

Office Scripts 3.6.4 retter indsættelse i en filtreret PlanTimer-tabel.
Behold Excel 3.6.3, og udskift de tre driftsscripts efter
[opdateringsvejledningen](office-scripts/3.6.4/README.md).

## Brug Ressourceplan med Office Scripts 3.6.3

Hent [den komplette 3.6.3-pakke](Ressourceplan_3_6_3.zip) til en ny
installation. Åbn `Ressourceplan_3_6_3.xlsx` i Microsoft Excel, og tilknyt
de tre scripts efter [installationsvejledningen](office-scripts/3.6.3/README.md).
Opgradér en eksisterende arbejdsfil med samme vejledning, hvis den indeholder
jeres registrerede timer og stamdata. I Teams eller SharePoint opretter
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

### Opret og planlæg via Office Scripts

Knyt de tre scripts til projektmappen i Excel. Brug derefter knapperne, eller
kør scripts fra **Automatiser**. De markerede pladser i udgivelsesfilen bliver
først aktive, når knapperne er tilføjet i Microsoft 365.
[Installér eller opgradér til 3.6.3](office-scripts/3.6.3/README.md)

- **Opret medarbejder** på fanen Medarbejdere registrerer navn, område,
  faggruppe, afdeling og normtid, tildeler ID og opretter standardrækker
  for det angivne planår.
- **Tilføj projekt** på Min ressourceplan tilføjer én Sag eller ét Tilbud
  for den valgte medarbejder, det angivne år og det valgte sagsnummer.
  Skriv timerne direkte i den nye række.
- **Opret nyt planår** på Min ressourceplan tilføjer kun de standardrækker,
  der mangler for medarbejderen og året. Nye planår kommer med i årslisten.

Markér en celle under Medarbejder for at genbruge navn og planår, eller angiv et
entydigt navn. Ved ens navne kan du markere den rigtige række eller bruge hele
etiketten med kode. Rækker tilføjes også, når de ligger uden for din aktuelle filtrering. Vælg derefter
medarbejderen og året i din personlige arkvisning. Tabelfilteret viser
eksisterende tabelværdier og kan være begrænset af andre aktive filtre.

### Administrativ reserve

**Planstart** er bevaret som skjult ark. Ved behov kan en administrator vise
arket og bruge den manuelle vejledning fra [3.5.1](https://github.com/Allanrasmus/Ressourceplan/blob/ressourceplan-3.5.1/README.md)
til at indsætte standardrækker som værdier. Arket indeholder årslisten og
formelskabelonen til Office Scripts. Skjul det igen efter brug.

Manuel indsættelse af en enkelt projektrække fungerer gennem
**Indsæt → Tabelrækker ovenfor**. Hvis Excel ikke tilbyder indsættelse under
filter, ryd filtrene midlertidigt i din egen arkvisning, indsæt en række inde i
tabellen og genanvend filtrene bagefter. Angiv planår, medarbejder og type
før projekttimerne. Sortér hele tabellen med en tabelpil.

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

## Gruppeprototype

Åbn gruppeprototypen separat og prøv kvartaler og detaljer med Excels plus- og
minusknapper. Ugefelterne på Kvartalsprototype er redigerbare kopier til
afprøvning. Medarbejderprototype viser et udsnit, der vælges ved bygning.
[Forudsætninger og afprøvning](DESIGN.md#outline-prototype-i-361)

## Byg og test

Kør fra projektmappen med Python 3.12. Brug et nyt outputfilnavn ved genbygning.

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
build_dir=$(mktemp -d)
python migrate_3_6_1.py --source Ressourceplan_3_6.xlsx --output "$build_dir/Ressourceplan_3_6_1.xlsx"
python migrate_3_6_3.py --source "$build_dir/Ressourceplan_3_6_1.xlsx" --output "$build_dir/Ressourceplan_3_6_3.xlsx"
python office-scripts/3.6.3/build.py
python prototype_outline_3_6_1.py --source "$build_dir/Ressourceplan_3_6_1.xlsx" --output "$build_dir/Gruppeprototype.xlsx" --year 2026
python -m unittest discover -s tests -v
```

Migrationen afviser at overskrive eksisterende filer. Testene sammenligner
XML-indholdet i en ny build med den udgivne fil. Sammenligningen ser bort fra
tidsstempler og forskelle i, hvordan ens XML skrives. På Windows kan du bruge
`.venv\Scripts\python.exe` og angive en ny filsti med `--output`.

LibreOffice Calc på PATH (`soffice` eller `libreoffice`) bruges til
beregningstestene. Disse springes tydeligt over, når motoren mangler. Calc kører
kun på midlertidige testkopier. Den udgivne fil skrives med openpyxl og
genberegnes ved åbning. Slutbrugere behøver Excel med Office Scripts til
automatiseringen. Beregninger og direkte timeindtastning kører i regnearket.

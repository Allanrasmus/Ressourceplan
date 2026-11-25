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

## Testjournal for 3.4

Den 24. november 2026 bestod alle 22 automatiske tests: 6 for 3.4 og de
16 tidligere tests. De fem nye scenarier blev genberegnet på separate
LibreOffice Calc 26.8-kopier uden formelfejl. Apache POI 5.4.1 og XMLBeans
5.3.0 skemavaliderede 10 SpreadsheetML-dele i udgivelsesfilen. Layoutet på
alle fem faner blev gennemgået i PDF-forhåndsvisninger fra separate Calc-kopier.

Build med `python migrate_3_4.py --source Ressourceplan_3_3_1.xlsx --output <ny-fil>`.
Kør `python -m unittest discover -s tests -v` for at teste 3.4 og de tidligere
versioner. Version 3.3.1 er manuelt testet af brugeren i Excel Desktop med
39 af 40 testpunkter godkendt. Det sidste punkt om overskrifter blev ikke
testet. Den accept gælder referencefilen.

| Kontrol | Dækning |
| --- | --- |
| Filtrering og total | Medarbejder 1, 5 og 9 samles fra første datarække. Nul, ét og 30 match kontrolleres i begge oversigter. I ALT placeres lige efter de synlige medarbejdere. |
| Kapacitet og projekt | Nøgletal, timer pr. række og uge, periodetotal, ledig kapacitet og overbelastning kontrolleres. Perioden over et ISO-årsskifte har fortsat korrekt år, uge og timer. |
| Projektregister | Listeværdier med navn og sagsnummer, omdøbning med bevaret valg og projekttimer samt oprettelse og valg af projekt efter tomme registerrækker. |
| Databevarelse | SHA-256 af 3.3.1, sammenligning af alle plan- og registerdata samt kontrol af aktive planfiltre, tabeller og datavalidering. |
| Excel-pakke | ZIP, XML, differentialfarver, prioritet, kanter, skjulte hjælperækker, opdaterbart navngivet dropdown-område og reproducerbar byggeproces. Kompakt visning bruger almindelige cellereferencer og formler. Der er ingen dynamiske matrixudvidelser i Excel-pakken. |

### Manuel accept i Excel Desktop og Excel Online

1. Åbn 3.4, og kontrollér åbning uden reparation, og at alle fem faner er
   læsbare ved almindelig zoom. Gennemgå feltbredder, ombrydning,
   dropdown-pile, kanter, aktive farver og totalrækker.
2. Filtrér på område, faggruppe og afdeling med nul, ét og flere match.
   Kontrollér, at I ALT står direkte efter sidste medarbejder, og at nøgletallene
   stemmer. Kontrollér belægningsfarver og orange ugeoverskrifter.
3. Omdøb et projekt og tilføj et nyt i en ledig registerrække. Kontrollér, at
   begge vises med navn og sagsnummer i dropdown-listen. Kontrollér det aktuelle
   navn og timetallet ved et eksisterende valg, genvælg projektet, og rul vandret.
4. Kontrollér de oprindelige filtre, timeregistrering, typeskift og uge 53 i
   Min ressourceplan. Gentag relevant afprøvning med flere brugere i Teams.

LibreOffice-kopierne bruges til formelkontrol. Visuel gengivelse,
dropdown-interaktion og åbning uden reparation for 3.4 skal bekræftes i Excel.

## Testjournal for 3.5

Den 24. november 2026 bestod alle 29 automatiske tests, fordelt på 7 nye
og 22 tidligere tests. De fem 3.5-scenarier blev genberegnet på separate
LibreOffice Calc 26.8-kopier uden celler med formelfejl. Apache POI 5.4.1
og XMLBeans 5.3.0 validerede 10 centrale SpreadsheetML-dele. En
Calc-forhåndsvisning viste de nye valgfelter og den skjulte ID-kolonne.

Brug `python -m unittest discover -s tests -v`. For kun 3.5 bruges
`python -m unittest discover -s tests -p test_ressourceplan_3_5.py -v`.
Version 3.4 er ifølge brugerens manuelle accepttest godkendt på 46 af 47
punkter uden konstaterede fejl. Punktet om skjult ID blev ikke testet.
Resultatet gælder referencefilen.

| Kontrol | Dækning |
| --- | --- |
| Databevarelse | Sammenlign alle bevarede planinput og registerdata med 3.4, samt SHA-256 af referencefilen. Fjern kun genkendelige standardrækker uden indtastning. |
| Planrækker | Udvid en kopi af PlanTimer under det gemte medarbejder- og årsfilter. Udfyld eksplicit planår, medarbejder og type, giv én person 25 projektrækker, og genberegn. Test også en ufuldstændig række. |
| Sortering | Sortér hele testrækker og sammenlign nøgle-/timekombinationer, ressource- og projekttotaler med usorteret kopi. |
| Projekter og medarbejdere | Nye projekter, samme navn med forskellige sagsnumre, omdøbning af projekt og medarbejder samt tilføjelse af medarbejder 31. |
| Summer og pakke | Referenceberegninger, ledig kapacitet, overbelastning, tilbud og omfordeling. XML, beregnede tabelkolonner, tabelreferencer, validering, skjult ID og genbygning kontrolleres. |

### Manuel accept i Excel Desktop og Excel Online

1. Åbn 3.5 og kontrollér, at der ikke vises en reparationsdialog. Bekræft
   standardfilteret, skjult ID, kvartaler, aktivitetstyper og direkte
   indtastning. Brug to planår og bekræft, at datoer kun vises for ét
   filtreret år.
2. Indsæt en tabelrække under det aktive medarbejder-/årsfilter i begge miljøer.
   Følg README, og kontrollér også reservevejen via egen arkvisning, hvis
   menuvalget er utilgængeligt. Bekræft, at Excel udfylder de beregnede
   tabelkolonner og bevarer validering, formatering og valg på den nye række.
3. Sortér via en tabelpil, og kontrollér projekttid i begge oversigter.
   Opret og omdøb en medarbejder og et projekt. Genanvend filtre efter
   navneskift. Kontrollér den opdaterede dropdown og uændrede timers nøgler.
4. Test en række med nyt planår og en med tilbud. Kontrollér optælling,
   kapacitet og projekttotaler efter filter og genåbning, samt ved samtidig
   brug af personlige arkvisninger.

I de automatiske testkopier udvides tabellens referencer, og Excel-operationer
som indsættelse, sortering og filtervisning efterlignes i OOXML. Calc læser og
genberegner kun testkopier. Interaktion, farvegengivelse og åbning uden
reparation i Microsoft Excel kræver den manuelle accept.

## Testjournal for 3.5.1

Den 24. november 2026 bestod alle 37 automatiske tests, heraf 8 nye og
29 tidligere. Seks nye scenarier blev genberegnet i separate LibreOffice
Calc 26.8-kopier uden formelfejl. Apache POI 5.4.1 og XMLBeans 5.3.0
skemavaliderede 11 centrale OOXML-dele. Version 3.5's hash er uændret.

Version 3.5 fik 40 godkendte punkter, to fejl og ét ikke-testet punkt ved
brugerens manuelle accept i Excel Desktop. Brug
`python -m unittest discover -s tests -v` til alle tests eller begræns til
`-p test_ressourceplan_3_5_1.py`.

### Excel-reparation og AutoSave

| Del | Fund |
| --- | --- |
| `sheet3.xml` | Projektregister. Version 3.5 anvendte `SagRegister[Sagsnr.]` direkte i den brugerdefinerede datavalideringsformel. Version 3.4 anvendte en lokal reference i kolonne A. |
| `sheet5.xml` | Medarbejdere. Version 3.5 tilføjede validering med `PersonRegister[ID]` direkte i formlen. Validering af område, faggruppe og normtid var særskilt. |
| Reparationen | De direkte strukturerede referencer er den konkrete fælles ændring, der er identificeret som årsag til den Excel-specifikke valideringsregression. 3.5.1 bruger 3.4-formen med lokale, absolutte områder og relativt input i A7. Det bevarer også forskellen mellem `00123` og `123`. Områder og tællere er kontrolleret for overlap og struktur. |
| AutoSave | XML i 3.4, 3.5 og 3.5.1 indeholder hverken skrivebeskyttelse, læseanbefaling eller eksterne forbindelser. Reparationsloggen viser fjernet validering og giver ikke i sig selv en årsag til AutoSave-tilstanden. Bekræft AutoSave i en reparationsfri kopi, gemt som regneark i OneDrive eller SharePoint, med den relevante Excel-konto. En lokal download kan åbne med AutoSave slået fra. |

Læs om betingelserne for AutoSave hos Microsoft.
[Hvad er AutoSave?](https://support.microsoft.com/en-us/office/what-is-autosave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5)

### Automatisk dækning

- Sammenlign de fejlramte OOXML-regler med 3.4 og 3.5. Kontrollér de nye
  valideringsformler, fortsat adskilte områder, navngivne dropdown-kilder
  og beregnet unikhed ved dubletter og foranstillede nuller.
- Kontrollér alle 32 tidligere planrækker celle for celle og fem
  standardaktiviteter for hver af de fem registrerede medarbejdere i seks år.
  Gennemfør initialisering to gange, og kontrollér nul nye dubletter.
- Afprøv Planstart med delvist udfyldt plan, medarbejder 5 uden projekt, ny
  medarbejder 6 og et nyt år. Brug de beregnede værdier som input til
  efterfølgende testkopier og kontrollér manglende rækker igen.
- Kontrollér identiske navne, omdøbning, sortering, ekstra projektrækker under
  gemte filtre og omfordeling med summer i begge oversigter. Kun separate
  kopier åbnes til genberegning i Calc.

### Manuel accept i Microsoft Excel

1. Åbn den frigivne 3.5.1 i Excel Desktop og Excel Online. Kontrollér, at
   der ikke vises en reparationsdialog. Prøv dubletter i begge registre
   og kontrollér område-, faggruppe-, medarbejder- og projektdropdowns.
2. Gem en separat kopi i Teams/SharePoint, åbn den igen og bekræft AutoSave.
   Hvis AutoSave fortsat er slået fra, notér Excels forklaring ved kontakten
   og bekræft konto og placering.
3. Filtrér 2026 i planen, ryd øvrige filtre og kontrollér medarbejder 5.
   Opret medarbejder 6 og gennemfør Planstart med værdikopiering. Kontrollér,
   at medarbejderen nu findes i tabelfilteret, og at en ny kørsel viser nul
   manglende standardrækker.
4. Kontrollér én synlig medarbejderkolonne, skjult ID, indsættelse under
   aktive filtre, farver og fungerende genberegning efter sortering. Prøv
   et nyt planår og kontrollér gamle timer efter genåbning.

De automatiske kontroller kan påvise strukturelle fejl og regnefejl.
Excels dialoger, dropdown-interaktion og AutoSave kræver den manuelle accept.

## Testjournal for 3.6

Den 24. november 2026 bestod 42 automatiske tests, heraf 5 nye og
37 tidligere. De anvendte API-signaturer blev afstemt med Microsofts
reference. Apache POI og XMLBeans skemavaliderede 11 centrale OOXML-dele.
Version 3.5.1's SHA-256 er uændret.

Byg Excel-filen med `python migrate_3_6.py` og de selvstændige scripts med
`python office-scripts/build.py`. Brug nyt outputfilnavn, hvis filen allerede
findes. Kør `python -m unittest discover -s tests -v`. Office Scripts-testene
kræver Node og TypeScript 4.0.3 på PATH eller via `TSC`. Calc er påkrævet til
beregningstesten.

- De tre scripts kompileres separat med TypeScript i streng tilstand.
  Testens ExcelScript-typer er et udsnit af Microsofts offentliggjorte
  API-signaturer fra `office-scripts-docs-reference`, commit
  `97f8f415842e13e004d7c53c1d24fe45653e50c0`.
- En lokal simulator bruger en kopi af udgivelsesfilens tabeller og
  indstillinger. Den kontrollerer p31, ledige ID-reserver, identiske navne,
  standardaktiviteter, nyt år, gentagen kørsel, projekt- og tilbudsrækker,
  standard/0 %/25 %, inputfejl og genkørsel efter en indsættelsesfejl.
- Scriptændringerne anvendes på en særskilt Excel-testkopi, derefter
  genberegnes den i Calc. Testen kontrollerer begge oversigter, uændrede
  historiske rækker og 53-uges beregning. Den grå regel er sand for en
  52-ugers tabelrække og falsk for en 53-ugers tabelrække, en tom række
  og et år indtastet neden for tabellen.
- Den gamle uge 53-regel testede ISO-spændet over hele formateringsområdet.
  Den nye regel kræver et årstal og en række inden for tabellens aktuelle
  slutrække. Uge 53-valideringen og timerne i beregningerne er uændret.
- Outline er afprøvet separat, herunder kolonne- og rækkegrupper i OOXML
  og genberegning af prototypen i Calc. [Resultater og udestående kontroller](DESIGN.md#undersøgelse-af-excel-gruppering-i-36)

### Manuel accept i Microsoft 365

1. Åbn en acceptkopi i Excel Desktop og Excel Online, og kontrollér åbning
   uden reparation, især de bevarede registervalideringer.
2. Installér, tilknyt og placér de tre knapper efter
   [installationsvejledningen](office-scripts/README.md). Kontrollér
   dialoger og inputfelter ved kørsel fra både knap og Automatiser.
3. Opret en medarbejder og et nyt planår med aktivt medarbejder- og årsfilter.
   Genkør, og kontrollér ingen dubletter. Tilføj et projekt og et tilbud.
   Kontrollér valideringer, formler, direkte timeindtastning og farver
   på de faktisk indsatte rækker. Kontrollér uge 53 i 2026 og 2027.
4. Prøv to brugere med hver sin arkvisning i Teams/SharePoint. Kontrollér
   at scriptet ikke ændrer filtervalget, og at rækkerne vises efter
   genanvendelse af de relevante filtre. Afprøv fejlsituationen på en
   testkopi, herunder kontrolleret genkørsel.
5. Kontrollér AutoSave og genåbning i Teams/SharePoint. Ved en senere
   outline-prototype skal samtidig ind- og udfoldning også afprøves
   med to brugere, før gruppestrukturen kan frigives.

Simulatoren og Calc verificerer kode og beregninger. Excel-værten afgør
installation, knapper, filtreret rækkeindsættelse og delte arkvisninger.

## Testjournal for 3.6.1

Den 24. november 2026 bestod 48 automatiske tests, heraf 6 nye og 42
tidligere. Alle tre scripts blev kompileret separat med TypeScript 4.0.3
mod signaturer fra Microsofts ExcelScript-reference. Separate Calc-kopier
bekræftede scriptresultater og grupperede beregninger uden formelfejl.
Apache POI og XMLBeans skemavaliderede 11 dele i driftsfilen og 14 i
prototypen. Begge filer kunne genbygges med samme indhold.

COWI bekræftede i Excel Desktop for version 3.6, at alle tre scripts
kunne tilknyttes. Medarbejderoprettelse, standardaktiviteter og
planårsinitialisering fungerede. Tilføj projekt fungerede efter en
manuel rettelse af Sag-valideringen. Denne accept gælder 3.6.

### Automatiske kontroller

- Kør `python -m unittest discover -s tests -v` med Node, TypeScript 4.0.3
  og LibreOffice Calc. Sæt `TSC` til kompilerens sti ved behov.
- Simulér første ledige p06, p06 med historik, en allerede oprettet p31,
  medarbejdervalg med ens navne, en samtidig registerændring og genkørsel
  efter indsættelsesfejl. Kontrollér, at eksisterende ID'er og timer bevares.
- Simulér Sag med og uden tilbudsprocent. Tilbud testes med
  tom tekst, manglende værdi, null, 0, 25, 50 og 100 samt ugyldige værdier.
  Afprøv projektnavn, sagsnummer med foranstillet nul, nye registerposter,
  forskellige markeringer, år fra planrække og gentagen initialisering.
- Genberegn en separat kopi med scriptresultaterne i Calc og kontrollér
  de faktiske oversigtssummer. Kontrollér 3.6's uændrede indhold og
  den afgrænsede uge 53-regel.
- Kontrollér native outline-segmenter for 2026 og 2027 og et kompakt
  Vest-udsnit. Fold uger med faktiske timer, filtrér prøveplanen og
  kildetabellen på en anden medarbejder, og genberegn med skjulte
  uger og detaljer. Kontrollér, at første medarbejders detaljer
  fortsat viser korrekt belægning, kapacitet og bookede timer.

### Manuel accept efter opdatering

1. Opdatér de samme tre scripts på en prøvekopi af jeres eksisterende
   arbejdsfil. Prøv knapperne og **Automatiser** i Excel Desktop og Online.
   Bekræft valgfri tekst i planår/procent, tom procent, 0 % og den
   validerede markering. Afprøv også ny 3.6.1-fil med tilknytning til
   samme scriptidentiteter. [Opdateringsvejledning](office-scripts/3.6.1/README.md)
2. Opret en person på første sikre plads. Prøv et ubrugt reserveret ID,
   et med historik og en eksisterende p31 på testdata. Tilføj Sag,
   Tilbud og nyt planår under aktive filtre, og kontrollér summer og
   timeindtastning.
3. Brug to konti i hver sin arkvisning i Teams/SharePoint. Bekræft
   personlige filtre før og efter kontrollerede samtidige forsøg på
   medarbejder- og projektoprettelse. Sammenlign dubletter, ændrede
   celler og versionshistorik. Kontrollér AutoSave og genåbning.
4. Prøv plus/minus for ét kvartal og alle kvartaler i separate
   52- og 53-ugers prototypekopier. Prøv én medarbejder og alle
   medarbejdere på niveau 1 og 2. Kontrollér tid, sortering, filtre,
   farver og detaljer efter genåbning. Prøv fra to særskilte
   arkvisninger og registrér, om den anden konto ser gruppeændringen.

Automatiske tests verificerer kode, beregninger og OOXML. Microsoft 365's
dialoger, knapper, AutoSave og samtidige outline-visninger kræver
manuel accept før fælles drift med de nye funktioner.

Ved overdragelsen den 24. november 2026 afventer alle manuelle
acceptpunkter for 3.6.1 fortsat afprøvning. De 48 beståede automatiske
tests ovenfor gælder kode og Excel-filer fra implementeringscommit
`bf9bff2`. Overdragelsen ændrer dokumentationen. [Næste trin](STATUS.md)

## Testjournal for 3.6.2

Den 25. november 2026 bestod 54 automatiske tests. De 48 historiske
tests og seks nye blev kørt. `sha256` for arbejdsfilen og prototypen
matcher 3.6.1. I den efterfølgende Excel Desktop-accept godkendte COWI
26 af 27 punkter. Afdelingsfejlen følges op i 3.6.3.

### Rapporterede fejl og ændringer

Excel Desktop markerede `selectedRecord(plan.table, …)` i
`add_project.ts` på linje 32, kolonne 34, og linje 282, kolonne 22,
samt i `initialize_plan_year.ts` på linje 244, kolonne 22. Fællesårsagen
var at gemme et `ExcelScript.Table`-objekt på `Plan.table` og videregive
objektet via den egenskab. Det mønster bestod almindelig TypeScript, men
Excels scripteditor meldte en aliasing-fejl. `Plan` indeholder nu
kolonnenavne og formelskabeloner. Hjælperne henter tabeller og områder
direkte via ExcelScript og udveksler celleværdier, kolonne- og rækkeindekser
samt tabelnavne. Samme gennemgang er udført for Opret medarbejder.
[Microsofts TypeScript-begrænsninger](https://learn.microsoft.com/en-us/office/dev/scripts/develop/typescript-restrictions)

Excel advarede yderligere om læsning i løkker på linje 161 og 283 i
det gamle planårsscript. Ved 20.000 planrækker er en læsning af planens
nøglefelter reduceret fra 20 blokke med ni kolonner til fem med fem
kolonner. En markeret planrække læses én gang for medarbejder, år og
projekt. For at kende en ledig Sag kontrolleres de rå ugetimer for
relevante rækker i blokke på op til 500 rækker. Uge 53 og værdier,
der udligner hinanden, kan være skjult af den beregnede rækkesum.
Syv individuelle formelskrivninger er samlet i fire blokke. Formel- og
procentformatering udføres efter kontrol af de indsatte værdier.

Læsning i blokløkker kan stadig udløse editorens læseadvarsel. Blokkene
begrænser svartørrelsen, og kontroller umiddelbart før og efter
indsættelse beskytter mod dubletter og synlige konflikter. Optællingerne
kommer fra kode og simulering. Kørselstid og faktiske advarsler skal
måles i Excel. [Microsofts vejledning om ydeevne](https://learn.microsoft.com/en-us/office/dev/scripts/develop/web-client-performance)

### Automatiske resultater

| Kontrol | Resultat |
| --- | --- |
| Almindelig TypeScript | Alle tre selvstændige filer kompilerede i streng tilstand med TypeScript 4.0.3 mod et udsnit af Microsofts API-typer. |
| Statisk kompatibilitetskontrol | Kontrollen genfandt alle tre rapporterede kald i 3.6.1 og accepterede de tre nye scripts. Den kontrollerer, hvor ExcelScript-objekter lagres, tildeles og videregives. |
| Simuleret ExcelScript | Testene dækkede eksplicitte parametre, medarbejdervalg med ens navne, markering i planen og projektregisteret, ugyldig markering, aktivt filter, historik, Sag, tom og eksplicit 0 % ved Tilbud, dubletter og genkørsel. |
| Simuleret stor tabel | 20.000 rækker blev behandlet med fem læsninger af nøglefelter pr. opslag og højst 41 bloklæsninger af rå timer i det simulerede scenarie. En ændring af rækkeantallet før indsættelse afbrød skrivningen. |
| Beregning og build | En separat LibreOffice-kopi gav 28 bookede timer i det testede ressourceoverblik, 36 timer i den testede projektperiode og ingen formelfejl. Genbygning af scripts matcher de publicerede filer. |

TypeScript kontrollerer typer og syntaks. Den statiske kontrol dækker de
konkrete risikomønstre. Simulatoren afprøver scriptlogikken, og Calc
beregner den særskilte testkopi. Microsofts Excel-editor og faktiske
runtime skal kontrolleres i den manuelle accept nedenfor.

Kør `python -m unittest discover -s tests -v` med Node, TypeScript 4.0.3
og LibreOffice Calc. Brug `-p test_ressourceplan_3_6_2.py` til at køre
de seks nye tests alene. `TSC` kan angive stien til kompilatoren.

### Manuel accept i Excel Desktop og Excel Online

1. Gem alle tre scripts i Microsofts editor, og start dem fra både
   **Automatiser** og eksisterende knapper på en prøvekopi. Registrér
   Excel-klient, fejl og eventuelle læseadvarsler.
   [Opdateringsvejledning](office-scripts/3.6.2/README.md)
2. Opret en medarbejder på første sikre plads med aktivt filter og
   kontrollér ID, historiske timer og standardaktiviteter. Tilføj Sag
   og Tilbud med eksplicitte input og markering i både plan og
   projektregister. Prøv tom procent og 0 %, samt ugyldig markering og
   to ens medarbejdernavne. Kontrollér direkte timeindtastning og summer.
3. Opret et nyt planår fra en markeret medarbejderrække. Kør samme
   handling igen. Kontrollér filtre, dubletter, projekter og gamle timer.
   Afprøv en stor testkopi og registrér kørselstid for planår og projekt.
4. Prøv separate arkvisninger og kontrollerede samtidige oprettelser på
   en kopi i Teams/SharePoint. Kontrollér konfliktbeskeder, gemte filtre,
   AutoSave og versionshistorik. Afprøv grupperingsprototypen separat
   som beskrevet i testjournalen for 3.6.1.

## Testjournal for 3.6.3

COWIs manuelle Excel Desktop-test af version 3.6.2 godkendte 26 af 27
punkter. Alle punkter blev afprøvet. Afdelingsnummer med punktum kunne
blive tolket som decimaltal og vist med komma. Knapper, AutoSave,
standardrækker, markering, projektopslag og oversigter fungerede i testen.

Den 25. november 2026 bestod de ni 3.6.3-kontroller af den samlede
leverance. De omfatter syv tidligere kontroller fra 3.6.3 og to nye
kontroller af pakken og den sorterede Excel-fil. Apache POI 5.4.1 og
XMLBeans 5.3.0 skemavaliderede 11 SpreadsheetML-dele i Excel-filen. En særskilt
Calc-prøve med 20.000 syntetiske planrækker gav den forventede paritet
i alle rækker, også efter skjulte rækker. Kørsel og visning i Microsoft
Excel afventer manuel accept af 3.6.3.

### Automatiske kontroller

Kør `python -m unittest discover -s tests -v` med Node, TypeScript 4.0.3
og LibreOffice Calc. Brug `-p test_ressourceplan_3_6_3.py` for at køre
de nye tests alene. `TSC` kan angive kompilerens sti.

| Kontrol | Omfang |
| --- | --- |
| TypeScript | Fire selvstændige scripts kompileres i streng tilstand mod de dokumenterede ExcelScript-signaturer. |
| Statisk kompatibilitet | Kontrol for de tidligere aliasing-mønstre omfatter de nye typer til betinget formatering og navngivne områder. Den eksisterende regression genfinder alle tre fejlsteder i 3.6.1. |
| Simuleret ExcelScript | Lagring af `7134.01`, `1734.03`, `7134.10` og `01734.03` under simuleret lokal talfortolkning. Genbrug af ledig Sag, beskyttelse af formler og rå timer, fem sager, flere tilbud, filtre, flere personer og år, stabile nøgler, ukendt projekt, dubletter og ændret indsættelsessted. |
| Opgradering | Genkørsel og afvisning ved eksisterende brugerdata i Planstart. Kontrol af tekstformat, eksisterende værdier, regler, kanter og validering. |
| Beregning | Separate LibreOffice-kopier med skjulte rækker, sorteret rækkefølge, udskiftede rækker og afdelingsfiltre. Beregnet sagsparitet, tilbud, kapacitet og oversigter kontrolleres. |
| Excel-fil og ZIP-pakke | Reproducerbar XML og ZIP, alle seks versionsoverskrifter, de tre knappladser og identiske scriptfiler i pakken. Kontrol af komplette rækker før og efter sortering, tabeller, filtre, valideringer, farver og prioriteter. Tidligere fil og grupperingsprototype har samme SHA-256. |
| Ny installation | Simulering med den pakkede Excel-fil genbruger den ledige Sag og indsætter Tilbud før standardaktiviteterne. Calc sammenligner summer og filtre med 3.6.1 efter sortering og kontrollerer de synlige sagsstriber. |

De statiske kontroller dækker kendte mønstre. Simulatoren afprøver logik og
kald, mens Calc efterprøver formler i testkopier. Microsoft Excels
kørsel, visning og samredigering vurderes via den manuelle accept.

### Manuel accept i Microsoft Excel

Afprøv også en ny installation fra ZIP-pakken. Åbn Excel-filen,
kontrollér 3.6.3 på de synlige faner, og tilknyt alle tre scripts i de
klargjorte pladser. Kør hver knap og kontrollér sagsrækkefølge, striber,
direkte timeindtastning og oversigter.

1. Brug en kopi af arbejdsfilen og separate testscripts. Kør
   [opgraderingen](office-scripts/3.6.3/README.md) to gange. Kontrollér
   åbning, AutoSave, knapper, tidsdata, stamdata, valideringer og filtre.
   Gem og kør alle tre driftsscripts både fra knap og **Automatiser**.
2. Brug dansk regional indstilling. Opret testprofiler med de fire
   afdelingskoder ovenfor. Kontrollér tekstværdien i redigeringslinjen og
   med Excels teksttest. Brug afdelingsdropdownen i Ressourceoverblik.
   Afprøv også ID/sagsnummer med foranstillet nul og kontrollér numerisk
   normtid, timer, planår og tilbudsprocent. Gentag i Excel Online.
3. Brug aktive medarbejder- og årsfiltre. Genbrug en tom Sag, tilføj fem
   forskellige sager og to tilbud i blandet rækkefølge til en medarbejder.
   Prøv en medarbejder med tidligere projekter, en uden projekter og et
   nyt år. Afprøv dublet, ukendt projekt, markering, blank og 0 %
   tilbudsprocent. Indtast og genlæs timer, og kontrollér begge oversigter.
4. Kontrollér synlige blå/hvide sager med varierende skjulte rækker,
   aktivitetsfarver, kanter, uge 53, kvartaler, belægning og totaler.
   Indsæt og slet en ledig tabelrække. Sortér en særskilt testkopi med
   hele PlanTimer efter vejledningen, og bekræft timer, stabile nøgler,
   formler, valideringer og format efter genåbning.
5. Prøv to brugere med personlige arkvisninger i Teams/SharePoint.
   Afprøv samtidige tilføjelser og filtre, og kontrollér, at sagsstriberne
   følger hver visning, og at timer og nøgler forbliver intakte. Kontrollér
   AutoSave, versionshistorik, eventuelle Office Scripts-advarsler og
   kørselstid ved op til 20.000 rækker. Godkend før merge og fælles drift.

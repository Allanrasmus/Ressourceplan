# Design
Reference: Ressourceplan_v2_8_medarbejderadministration.html. Dansk brugerflade, Segoe UI, mørkeblå #1C2B54, orange #F04E23, baggrund #F6F7F8, tekst #192339, muted #64748B, linjer #E1E5EC. Mørk topbjælke med orange kant, tydelig navigation, lyse tabeller, ingen synlige gitterlinjer hvor praktisk. Almindelige sager blå/hvide striber (#D2EAF7); tilbud #FCE3E3; ferie og fravær #FFF4D9; kursus #E3F0DF; andet #EDE5F5. Betinget formatering overstyrer striber. Kvartaler Q1 #D2EAF7, Q2 #D5F0E9, Q3 #EFE1D6, Q4 #FCE3E3. Belægning: nul projekttimer #EAF2FC, under 85 % #DFF2E9, 85–100 % #FFF1CF, over 100 % #F9DADD. Kvartal, ugenummer og dato for mandag synlige samtidig; kvartal efter mandagens måned som HTML-reference. Fremhæv projektvalg med orange label og tydeligt inputfelt. Skjul tekniske beregninger og ID, ikke funktionalitet.

## Layout i Excel 3.4

COWI-båndet og den orange kant følger HTML-referencen på alle fem ark.
Sidens titel og vejledning står i et roligt, lyst felt. Filtre på de to
oversigter står i en selvstændig række med ▾ ved hver overskrift og lyseblå,
indrammede felter under. Projektvælger, år, startuge og periode ligger i den
fastlåste del af Projektoversigt. Registrene og Min ressourceplan bevarer
deres redigerbare Excel-tabeller og de tilhørende filterpile.

De to oversigter har et selvstændigt nøgletalsbånd. Ressourceoverblik viser
også en farveforklaring og beholder kvartaler, datoer og orange uger.
Min ressourceplan har en særskilt aktivitetsforklaring. Planens timefelter
og eksisterende rækker forbliver redigerbare.

Udfyldte medarbejderrækker vises fra første datarække, efterfulgt af **I ALT**.
Ved nul match vises besked og total på første datarække. Blå og hvide striber
og tynde grå kanter hjælper med at følge rækkerne. Mint markerer totalrækken
i Projektoversigt samt totalens medarbejder- og periodefelter i
Ressourceoverblik. I sidstnævntes ugekolonner viser belægningsfarverne fortsat
belastningen. Almindelige gitterlinjer er slået fra.

Beregningsmodellen fra 3.3.1 ligger i skjulte rækker i de to oversigter.
Visningen bruger almindelige Excel-formler, som udpeger de relevante rækker
og totalens placering inden for kapaciteten på 30 medarbejdere. Formler og
tabeller i Min ressourceplan bruger de eksisterende medarbejder-ID'er.

Projektregister leverer en navngiven dropdown-liste med `Projektnavn · Sagsnr.`.
Rapporten bruger sagsnummeret internt. **Aktuelt projekt** afspejler navneskift
med det samme, mens Excels gemte dropdown-tekst opdateres ved genvalg.

## Undersøgelse af Excel-gruppering i 3.6

En separat prototype gemte native kolonnegrupper i Min ressourceplan og
Ressourceoverblik og en rækkegruppe under en opsummering på et prøveark.
Ugekolonnerne kunne samles uden at medtage medarbejder- og projektfelter.
OOXML viste outline-niveauer, skjulte kolonner/rækker og bevarede
tabelreferencer og filtre. Ved genberegning i Calc bevarede den sammenfoldede
prototype 37 synlige plantimer, 20 bookede timer og projekttotalen på 20 timer.
Prøvearkets skjulte detaljer summerede fortsat til 27 timer. Prototypen
brugte tekstbeskrivelsen af plus- og minusknapperne. Skærmbilledet var ikke
tilgængeligt i denne opgave.

### Kvartalsgrupper

Excel gemmer gruppen på fysiske kolonner. Vores kvartaler følger mandagens
dato i det valgte planår. I 2026 svarer Q2 til uge 15–27, mens det svarer
til uge 14–26 i 2027. I 2026 tilhører uge 1 Q4 fra december 2025, og
Q4 optræder igen i uge 41–53. Derfor kræver korrekte grupper, at de
opdateres for et kendt planår. Opret grupper kun på ugekolonnerne,
og test både 52/53 uger, årsfilter, synligheden af de fastlåste felter og
direkte redigering efter udvidelse. Ved visning af flere år har grupperne
ingen entydig kvartalsbetydning.

### Medarbejdergrupper

Ressourceoverblikket viser allerede én ugentlig belægningsprocent pr.
medarbejder. En native rækkegruppe kræver et sammenhængende sæt eksplicitte
detaljerækker under hver summeringsrække. En senere prototype kan samle
projekt- og aktivitetsrækker via medarbejder-ID, år og uge, vise kapacitet
og bookede timer og bevare den eksisterende belægningsrække som
gruppehoved. Grupper skal bygges for det aktuelle filter og følge
medarbejder-ID, når den kompakte liste ændrer rækkefølge. Test gruppevis
foldning og foldning af alle uden at ændre nøgletal eller skjule en anden
medarbejders rækker.

### Office Scripts og arkvisninger

ExcelScript har `Range.group` og `Range.ungroup` for rækker og kolonner.
`Worksheet.showOutlineLevels` kan folde niveauer ud og ind.
`Worksheet.enterTemporaryNamedSheetView` giver en midlertidig arkvisning.
Gruppeopbygningen gemmes i selve arkets OOXML. API-dokumentationen for
`NamedSheetView` beskriver sorterings- og filterregler, men fastslår ikke,
at opbygning eller ændring af outline-grupper er isoleret pr. bruger.
Inden frigivelse skal to brugere i Excel Online og Excel Desktop kontrollere,
om den ene kan skifte år, gengruppere og folde ind/ud uden at ændre den andens
visning. Kontrollér samtidig filtre, sortering, direkte timeindtastning
og resultater efter genåbning. Grupperingen er derfor endnu en prototype.
[Range](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts)
[Worksheet](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet?view=office-scripts)
[NamedSheetView](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.namedsheetview?view=office-scripts)

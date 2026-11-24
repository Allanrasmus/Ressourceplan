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

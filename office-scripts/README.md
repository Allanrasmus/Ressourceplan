# Ressourceplan 3.6 i Excel

Installér scripts i jeres Microsoft 365-miljø med en konto, der har adgang
til **Automatiser** og skriverettigheder til projektmappen. Trinnene gælder
Excel Desktop og Excel Online, hvor Office Scripts og deling er tilgængelige.
[Microsofts vejledning til scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

## Installér og tilknyt tre scripts

1. Gem `Ressourceplan_3_6.xlsx` i jeres Teams-, SharePoint- eller OneDrive-bibliotek.
   Åbn projektmappen i Excel.
2. Vælg **Automatiser → Nyt script**. Kopiér hele den relevante fil fra
   tabellen nedenfor ind i editoren, erstat eksempelkoden, og gem scriptet
   med det anførte navn. Brug de selvstændige filer i denne mappe.
3. Åbn scriptets oplysninger. Under **Del dette script** slå
   **Tilknyt til projektmappe** til. Vælg det relevante ark, og klik
   **Føj knap til regneark**. Placér knappen over den markerede knapplads.
   Gentag for hvert script. Kør det fra **Automatiser**, hvis jeres politik
   tillader scriptkørsel, men begrænser tilknytning eller knapper.
4. Prøv hver handling i en særskilt acceptkopi. Afprøv både Excel Online
   og jeres relevante Excel Desktop-version før fælles drift.

| Kilde | Navn i Excel | Ark og knapplads |
| --- | --- | --- |
| [create_employee.ts](create_employee.ts) | Opret medarbejder | Medarbejdere, A5:B5 |
| [add_project.ts](add_project.ts) | Tilføj projekt | Min ressourceplan, A2:B2 |
| [initialize_plan_year.ts](initialize_plan_year.ts) | Opret nyt planår | Min ressourceplan, D2:F2 |

Scriptparametre indtastes i dialogen, når en knap eller et script køres.
Knappladserne i downloadfilen er almindelige formaterede celler.
[Microsofts vejledning til scriptparametre](https://learn.microsoft.com/en-us/office/dev/scripts/develop/user-input)

## Brug handlingerne

**Opret medarbejder** tager navn, område, faggruppe, afdeling, normtid (0–80
timer) og planår. Scriptet bruger en helt ledig plads først og tildeler
et nyt ID. Hvis der kun er ubrugte reserverede ID'er tilbage, kan et sådant
ID tildeles sin tilsvarende ledige plads. Identisk profil ved genkørsel
bruger det eksisterende ID og udfylder kun manglende standardrækker.
Ved identiske navne kan du bruge ID'et til at skelne. Hvis to personer har
samme fulde profil, opret en af dem i registeret med eget ID og kør
**Opret nyt planår** på det ID.

**Tilføj projekt** tager medarbejder, planår, projekt og typen Sag eller Tilbud.
Brug helst medarbejder-ID og sagsnummer, f.eks. `p31` og `24015`. Indtast
sagsnumre med foranstillede nuller som tekst i feltet. Et entydigt navn eller
hele etiketten med kode kan også bruges. For Tilbud kan feltet
`tilbudsprocent` stå tomt for 50 %, eller angives som 0–100, f.eks. 25 for
25 %. En kombination af medarbejder, år, projekt og aktivitetstype tilføjes
kun én gang af scriptet. Indtast timerne i planen efter oprettelsen.

**Opret nyt planår** tager medarbejder og planår. Scriptet registrerer året,
hvis det mangler, og tilføjer kun en ledig Sag, Ferie, Fravær, Kursus og
Andet, hvor de relevante standardrækker mangler. Et eksisterende projekt
optager ikke den ledige sagsrække. Historiske timer bevares ved genkørsel.

## Filtre og samtidige brugere

Brug **Vis → Arkvisning** i delte projektmapper og behold egne filtre.
Scripts ændrer ingen filtre og bruger ingen aktiv celle eller fælles
indtastningsfelter. Genanvend din visning efter en kørsel, og vælg det år og
den medarbejder, du angav. Append under aktive filtre afhænger af Excel;
hvis Excel afviser indsættelsen, prøv igen efter at have ryddet filtrene
midlertidigt i din egen arkvisning.

Rækkeindsættelse og ændringer i medarbejder- og årslisten er ændringer i
den fælles projektmappe. Kontrollér, at kolleger er færdige med at køre
scripts, før du kører samme handling, eller sorterer eller redigerer
registrene. Scripts genlæser nøgleoplysninger omkring oprettelse, men
Excel-tabeller har ingen transaktioner eller lås, der garanterer mod
samtidige dubletter. Ved en afbrudt medarbejderoprettelse med en registreret
profil, genkør samme profil eller brug **Opret nyt planår** for det
tildelte ID. Hvis Excel allerede har indsat rækker ved en fejl, kontrollér
medarbejder, år, type og formler i de berørte rækker inden genkørsel og
timeindtastning.

Standardplanen er klargjort til 100 medarbejdere, 500 projekter, 20 planår og
20.000 planrækker med validering og formatering.

## Genbyg og test scriptkilder

Redigér kilderne under `src`, og kør `python office-scripts/build.py` fra
repositoryets rod. Dette samler separate filer uden importafhængigheder,
så de kan indsættes direkte i Excels editor.

Til automatiske udviklingstests bruges Node, TypeScript 4.0.3 og eventuelt
LibreOffice Calc. Angiv stien til TypeScript-kompilatoren med `TSC`,
hvis `tsc` ikke er på PATH. Kør
`python -m unittest discover -s tests -p test_ressourceplan_3_6.py -v`.
Automatiske tests bruger dokumenterede API-signaturer og en lokal
simulator. Kørsel i Microsoft Excel og knapper kontrolleres manuelt.

# Installér Ressourceplan 3.6.7

ZIP-filen indeholder demonstrationsprojektmappen, tre driftsscripts og et
særskilt opgraderingsscript. Brug en acceptkopi i Microsoft 365 og særskilte
testscripts, da en kopieret projektmappe kan pege på eksisterende scripts.

## Ny installation

Åbn `Ressourceplan_3_6_7.xlsx`. Opret hvert af de tre driftsscripts under
**Automatiser → Nyt script**, kopiér hele den tilsvarende TypeScript-fil,
og knyt scriptet til projektmappen og den viste knapplads.

| Driftsscript | Fil | Knapplads |
| --- | --- | --- |
| Opret medarbejder | [create_employee.ts](create_employee.ts) | Medarbejdere, A5:B5 |
| Tilføj projekt | [add_project.ts](add_project.ts) | Min ressourceplan, A2:B2 |
| Opret nyt planår | [initialize_plan_year.ts](initialize_plan_year.ts) | Min ressourceplan, D2:F2 |

## Opgradér en eksisterende 3.6.3-arbejdsfil

1. Lav en versionssikret kopi i Microsoft 365. Aftal et tidspunkt uden andre
   redigeringer. Brug kopien til at afprøve opgradering og driftsscripts.
2. Opret et særskilt script fra [upgrade_workbook.ts](upgrade_workbook.ts).
   Kør det på kopien. Scriptet kontrollerer versionsoverskrifter, F8 og
   samtlige formler i Ressourceoverblik!H590:BH689 før skrivning. Det
   retter kun den kontrollerede formelblok og versionsoverskrifterne.
   Hvis layoutet afviger, standses opgraderingen. Behold arbejdsfilen
   og afklar afvigelsen på kopien. Scriptet kan genkøres efter en afbrydelse,
   hvis de kontrollerede celler stadig matcher en af versionerne.
3. Kontrollér periode, medarbejderfiltre, Bruttotilbud og Vægtede tilbud på
   kopien. Sammenlign eksisterende timer og registre med originalen.
   Afprøv alle tre driftsscripts i Desktop og Online og med to konti på
   SharePoint. [Manuel accept](https://github.com/Allanrasmus/Ressourceplan/blob/ressourceplan-3.6.7/TESTPLAN.md#testjournal-for-367)
4. Efter accept køres den samme opgradering på den sikkerhedskopierede
   arbejdsfil. Gem de gamle scripts til gendannelse. Erstat hele koden i
   hvert eksisterende driftsscript under **Automatiser → Alle scripts**,
   og gem med samme scriptidentitet, så knapperne kan genbruges.

## Filtre, tilbud og beskeder

Ved indsættelse midt i filtreret PlanTimer vælger du først
**Vis → Arkvisning** på Min ressourceplan og aktiverer din personlige
visning. Sæt først derefter `bekraeftPersonligArkvisning` til sand i
scriptdialogen. Scriptet kontrollerer kriterierne og din bekræftelse.
Bekræftelsen bygger på dit valg i Excel. Samtidig kørsel og AutoSave skal
afprøves med to konti før produktionsbrug. Aftal én administrator ad gangen.

3.6.7 understøtter ét Custom-lighedskriterium på Medarbejder og Planår
hver for sig eller sammen. Andre filtre giver en fejl før ændring.
Filtre ryddes kun ved indsættelse før en eksisterende række.
Et nyt tilbud uden procent får 50 %. Du kan angive 0–100 som heltal
eller decimaltal. Sag ignorerer procentfeltet.

Tilføj projekt afslutter med en fejlbesked ved dublet eller manglende
projekt i Projektregister. Ved succes returnerer og logger scriptet en
kvittering. Hvor meget Excel viser direkte fra knappen afhænger af
klienten. Åbn kørselsresultatet under **Automatiser**, hvis knappen kun viser
kørselsstatus. Office Scripts har ingen selvstændig dialog til denne
kvittering. [Microsoft om scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

Byg demonstrationsfilen med `python migrate_3_6_7.py`. Byg scripts og
ZIP med `python office-scripts/3.6.7/build.py`.

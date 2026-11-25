# Opgradér Ressourceplan til 3.6.3

Behold jeres eksisterende arbejdsfil med timer, stamdata og scriptknapper.
Versionen består af tre opdaterede driftsscripts og et særskilt opgraderingsscript
til tekstformater og sagsstriber. [Excel-referencefilen](../../Ressourceplan_3_6_3.xlsx)
er bygget med projektets demonstrationsdata til nye prøver.

## Prøv først i Excel

1. Lav en acceptkopi af arbejdsfilen i virksomhedens Microsoft 365-miljø.
   Kontrollér, at den kan gendannes fra versionshistorikken. Brug separate
   testscriptfiler under **Automatiser**. En kopieret arbejdsfil kan stadig
   referere til de delte produktionsscripts, så brug de særskilte testscripts
   ved afprøvningen.
2. Opret et særskilt script fra hele koden i [upgrade_workbook.ts](upgrade_workbook.ts).
   Kør det som administrator på acceptkopien. Det bevarer værdier, eksisterende
   formler og alle tabelrækker. Det opdaterer de to sagsregler, anvender
   tekstformat på registrenes tekstfelter og afdelingsvælgeren samt lægger
   de tekniske stribeformler i det skjulte Planstart-ark. Du kan køre det igen.
3. Installér de tre scripts fra tabellen og gennemfør
   [accepttesten](../../TESTPLAN.md#testjournal-for-363). Kontrollér både
   **Automatiser** og knapper, aktiv filtrering, farver og AutoSave.
4. Efter godkendelse køres opgraderingsscriptet én gang i produktionsfilen i
   et aftalt tidsrum. Erstat hele koden i hvert eksisterende driftsscript,
   og gem under samme identitet. Så kan knapperne genbruges. Bekræft
   tilknytningerne og gem de hidtidige scriptversioner til gendannelse.

| Eksisterende script | Færdigsamlet kode |
| --- | --- |
| Opret medarbejder | [create_employee.ts](create_employee.ts) |
| Tilføj projekt | [add_project.ts](add_project.ts) |
| Opret nyt planår | [initialize_plan_year.ts](initialize_plan_year.ts) |

Skriv afdelingskoden med punktum og alle nuller i **Opret medarbejder**.
Ressourceoverblikkets **Afdeling** bruger de værdier, der står i registret.
Allerede lagrede tal eller kommaer i gamle afdelingskoder bevares. Kontrollér
sådanne poster mod det autoritative register, før I eventuelt retter dem manuelt.

Hvis opgraderingen finder ændrede eller opsplittede sagsregler, standser den
før opgradering. Sammenlign reglerne via **Betinget formatering → Administrer
regler → Dette regneark** på en kopi. En administrator kan afstemme de to
Sag-regler mod en ren 3.6.1-reference, inklusive området fra A10 til mindst
BJ20009 og rækkekanterne. Kør igen, når reglerne matcher. Behold reglerne for
øvrige aktiviteter, kvartaler, tilbudsprocent og uge 53.

## Ret tidligere rækkefølge efter behov

**Tilføj projekt** bruger en sikker tom Sag på det relevante sted eller
indsætter en række ved medarbejderens og planårets aktiviteter. Et projekt
skal allerede være oprettet i **Projektregister**. En gammel række, der
allerede ligger efter en standardaktivitet, bevarer sin placering, indtil I
retter den særskilt.

Brug Excel Desktop på en kopi først. Aftal et eksklusivt vedligeholdelsesvindue
ved sortering i produktionsfilen, da rækkefølgen deles. Vis midlertidigt den
tekniske **ID**-kolonne C. Vælg en celle i **PlanTimer**, og brug **Data →
Sortér** på hele tabellen med overskrifter. Sortér på **Planår**, **ID** og
derefter **Type** med en brugerdefineret liste med disse værdier i rækkefølge.

```text
Sag
Tilbud
Ferie
Fravær
Kursus
Andet
```

Skjul ID igen og genanvend relevante filtre i den personlige arkvisning.
Kontrollér antal rækker, timer og oversigtssummer før og efter. Brug hele
tabellen ved sortering, så medarbejder-ID, projektnøgle og alle ugefelter
flytter samlet. Afprøv arbejdsgangen på en kopi, før den bruges i
produktion. [Microsoft om brugerdefineret sortering](https://support.microsoft.com/en-us/office/sort-data-using-a-custom-list-2f4615ee-6f17-4960-a05f-782d6449a812)

## Genbygning

Kør `python office-scripts/3.6.3/build.py` fra repositoryets rod. Kør
`python migrate_3_6_3.py --output Staging_3_6_3.xlsx` for en ny reference
med demonstrationsdata. Office Script-opgraderingen foregår direkte i
Microsoft Excel og kræver ingen Python.

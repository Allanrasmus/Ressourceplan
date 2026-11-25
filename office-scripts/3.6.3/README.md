# Installér eller opgradér Ressourceplan 3.6.3

3.6.3 samler en komplet Excel-fil, tre driftsscripts og et særskilt
opgraderingsscript. Den bygger på Excel 3.6.1 og Office Scripts 3.6.2.
I kan installere den medfølgende Excel-fil fra bunden eller opgradere
jeres eksisterende arbejdsfil med registrerede timer og stamdata.

## Installér den nye Excel-fil

1. Pak ZIP-filen ud, og åbn `Ressourceplan_3_6_3.xlsx` i Microsoft Excel.
   Filen indeholder projektets demonstrationsdata. Til fælles afprøvning
   placeres den i virksomhedens godkendte Microsoft 365-miljø.
2. Opret de tre scripts under **Automatiser → Nyt script**. Kopiér hele
   indholdet af hver TypeScript-fil fra tabellen. Gem hvert script med
   det viste navn.
3. Åbn scriptet under **Automatiser → Alle scripts**, og tilknyt det til
   projektmappen med **Føj til projektmappe** (*Add in workbook*).
   Tilføj en knap, og placér den på den viste plads. Gentag for alle tre,
   og prøv både knap og **Automatiser**. Knappladserne er klargjort i
   regnearket. Selve tilknytningen oprettes i Microsoft 365.
4. Kontrollér, at alle faner viser 3.6.3, og prøv **Tilføj projekt** ved
   at vælge en eksisterende sag i Projektregister. Test Sag og Tilbud
   i den aftalte rækkefølge og kontrollér sagsstriberne under filtrering.
   Den nye fil er klargjort med tekstformater, sagsstriber og rækkefølge.
   [Microsofts vejledning til scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

| Driftsscript | Færdigsamlet kode | Knapplads |
| --- | --- | --- |
| Opret medarbejder | [create_employee.ts](create_employee.ts) | Medarbejdere, A5:B5 |
| Tilføj projekt | [add_project.ts](add_project.ts) | Min ressourceplan, A2:B2 |
| Opret nyt planår | [initialize_plan_year.ts](initialize_plan_year.ts) | Min ressourceplan, D2:F2 |

## Opgradér den eksisterende arbejdsfil

Behold jeres Excel 3.6.1-fil med COWIs data og installerede knapper.
Denne arbejdsgang bevarer timer, stamdata, rækker og scriptidentiteter.

1. Lav en acceptkopi af arbejdsfilen i virksomhedens Microsoft 365-miljø.
   Kontrollér, at den kan gendannes fra versionshistorikken. Brug separate
   testscriptfiler under **Automatiser**. En kopieret arbejdsfil kan stadig
   referere til de delte produktionsscripts, så brug de særskilte testscripts
   ved afprøvningen.
2. Opret et særskilt script fra hele koden i [upgrade_workbook.ts](upgrade_workbook.ts).
   Kør det på acceptkopien med en bruger, der kan redigere projektmappen.
   Opgraderingen bevarer værdier, eksisterende formler og alle tabelrækker.
   Den opdaterer de to sagsregler og tekstformaterne i registrene og
   afdelingsvælgeren. De tekniske stribeformler ligger i det skjulte
   Planstart-ark. Du kan køre scriptet igen.
3. Installér de tre testscripts fra tabellen. Kontrollér både
   **Automatiser** og knapper, aktiv filtrering, farver og AutoSave.
   [Accepttest](https://github.com/Allanrasmus/Ressourceplan/blob/ressourceplan-3.6.3/TESTPLAN.md#testjournal-for-363)
4. Efter godkendelse køres opgraderingsscriptet én gang i produktionsfilen i
   et aftalt tidsrum. Erstat hele koden i hvert eksisterende driftsscript,
   og gem under samme identitet. Så kan knapperne genbruges. Bekræft
   tilknytningerne og gem de hidtidige scriptversioner til gendannelse.

Skriv afdelingskoden med punktum og alle nuller i **Opret medarbejder**.
Ressourceoverblikkets **Afdeling** bruger de værdier, der står i registret.
Allerede lagrede tal eller kommaer i gamle afdelingskoder bevares. Kontrollér
sådanne poster mod det autoritative register, før I eventuelt retter dem manuelt.

Hvis opgraderingen finder ændrede eller opsplittede sagsregler, standser den
før opgradering. Sammenlign reglerne via **Betinget formatering → Administrer
regler → Dette regneark** på en kopi. En administrator kan afstemme de to
Sag-regler mod en ren Excel 3.6.1-fil, inklusive området fra A10 til mindst
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

Kør `python migrate_3_6_3.py --output Staging_3_6_3.xlsx` fra
repositoryets rod for at bygge Excel-filen fra 3.6.1-demonstrationsdata.
Kør `python office-scripts/3.6.3/build.py` for at samle scripts. Tilføj
`--package` for at pakke den udgivne Excel-fil, vejledningen og scripts
som `Ressourceplan_3_6_3.zip`. Opgraderingen af en eksisterende
arbejdsfil foregår direkte i Excel.

# Opdatér Office Scripts til 3.6.4

Brug den eksisterende Excel 3.6.3-projektmappe. 3.6.4 udskifter de tre
driftsscripts og kræver ingen ændringer i projektmappen eller engangsmigrering
fra Excel 3.6.3.

| Driftsscript | Færdigsamlet kode | Knapplads |
| --- | --- | --- |
| Opret medarbejder | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.4/office-scripts/3.6.4/create_employee.ts) | Medarbejdere, A5:B5 |
| Tilføj projekt | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.4/office-scripts/3.6.4/add_project.ts) | Min ressourceplan, A2:B2 |
| Opret nyt planår | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.4/office-scripts/3.6.4/initialize_plan_year.ts) | Min ressourceplan, D2:F2 |

1. Lav en acceptkopi af Excel 3.6.3 i det godkendte Microsoft 365-miljø. Opret
   tre særskilte testscripts under **Automatiser → Nyt script**, og kopiér hele
   indholdet fra hver fil. En kopi kan stadig referere til produktionsscripts.
   Kør testscriptene under **Automatiser**, eller tilknyt særskilte testknapper.
2. Åbn **Min ressourceplan**, og vælg **Vis → Arkvisning → Ny** eller en af dine
   eksisterende arkvisninger. Sæt filtre, og afprøv alle tre scripts efter
   [testplanen](../../TESTPLAN.md#testjournal-for-office-scripts-364).
3. Efter accept gemmes de tidligere scriptversioner til gendannelse. Erstat hele
   koden i hvert eksisterende driftsscript under **Automatiser → Alle scripts**.
   Gem under samme scriptidentitet, og kontrollér, at knapperne kører 3.6.4.
   [Microsofts vejledning til scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

## Filtre og samtidig redigering

Ved indsættelse før en eksisterende række i filtreret **PlanTimer** skal Excel
kunne bekræfte en aktiv personlig arkvisning på **Min ressourceplan**. Brug en
navngiven eller midlertidig arkvisning. I en separat kopi til enebrug kan du
alternativt rydde filtrene manuelt før indsættelsen. Genbrug af en tom sagsrække,
tilføjelse nederst og indsættelse i en ufiltreret tabel kræver intet visningsskift.

Scriptet genskaber kriterierne for de kolonner, det forsøgte at rydde, også
ved en håndterbar Excel-fejl. Hvis Excel afviser et kriterium, viser scriptet
kolonnen i fejlbeskeden og forsøger at genskabe de øvrige. Kontrollér rækker,
relevante registerposter og planårslisten før genkørsel. En fejl kan opstå efter
at Excel har gemt ændringer. Genskab eventuelle manglende filtre manuelt. Ved
en tvungen afbrydelse kan oprydningen blive afbrudt.

Arkvisninger gemmer visningens filtre og sortering. Data- og tabelændringer
sker i den fælles projektmappe. Aftal, at administration af PlanTimer køres af
én bruger ad gangen, og undgå samtidig sortering og andre ændringer i tabellen
under kørslen. Flerbrugeradfærd og AutoSave afventer manuel SharePoint-test.
[Microsoft om arkvisninger](https://support.microsoft.com/en-us/excel/create-and-manage-sheet-views-in-excel)

Kør `python office-scripts/3.6.4/build.py` fra repositoryets rod for at
gensamle de tre scripts fra kildekoden.

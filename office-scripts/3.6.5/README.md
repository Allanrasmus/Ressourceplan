# Installér Office Scripts 3.6.5

Brug den eksisterende Excel 3.6.3-projektmappe. Opret en acceptkopi med
fiktive data i det godkendte Microsoft 365-miljø. Brug særskilte
testscriptfiler, da en kopieret projektmappe kan referere til de delte
produktionsscripts.

1. Hent [filterprøven](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.5/office-scripts/3.6.5/filter_probe.ts).
   Kopiér hele filen til et nyt script under **Automatiser**. Vælg en personlig
   arkvisning på **Min ressourceplan**, filtrér PlanTimer på planår og
   medarbejder, og kør prøven særskilt i Desktop og Online. Den rydder og
   genskaber filtre og sammenligner de aflæste kriterier. Kontrollér også
   visningen og filtre i brugerfladen.
2. Efter en bestået filterprøve oprettes hvert af de tre scripts i tabellen
   som et særskilt testscript. Kopiér hele filen. Afprøv alle tre i
   Excel Desktop og Online og følg SharePoint-punkterne i testjournalen.
3. Efter manuel accept gemmes de hidtidige scriptversioner til gendannelse.
   Erstat hele koden i hvert driftsscript under **Automatiser → Alle scripts**.
   Gem med samme scriptidentitet, og kontrollér, at knapperne kører 3.6.5.

| Driftsscript | Færdigsamlet kode | Knapplads |
| --- | --- | --- |
| Opret medarbejder | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.5/office-scripts/3.6.5/create_employee.ts) | Medarbejdere, A5:B5 |
| Tilføj projekt | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.5/office-scripts/3.6.5/add_project.ts) | Min ressourceplan, A2:B2 |
| Opret nyt planår | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.5/office-scripts/3.6.5/initialize_plan_year.ts) | Min ressourceplan, D2:F2 |

Brug en personlig arkvisning ved filtreret indsættelse før eksisterende
PlanTimer-rækker. Ved fejl kontrolleres planrækker, relevante registerposter,
planårslisten og kriterier før genkørsel. Genskab filtre manuelt, hvis Excel
afviser gendannelsen eller afbryder scriptet. Brug én administrator ad gangen
ved indsættelse. Arkvisninger gemmer filtre og sortering, mens tabelændringer
gemmes i den fælles projektmappe.

[Testjournal og manuel accept](https://github.com/Allanrasmus/Ressourceplan/blob/ressourceplan-3.6.5/TESTPLAN.md#testjournal-for-office-scripts-365)

Kør `python office-scripts/3.6.5/build.py` fra repositoryets rod for at
gensamle tre driftsscripts, filterprøven og ZIP-filen med driftsscripts og
denne vejledning.

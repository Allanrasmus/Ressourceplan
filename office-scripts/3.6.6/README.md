# Installér Office Scripts 3.6.6

Brug en acceptkopi af `Ressourceplan_3_6_3.xlsx` i det godkendte
Microsoft 365-miljø. Opret særskilte testscripts under **Automatiser**.
Kopiér hele indholdet af hver TypeScript-fil til sit eget script.

| Script | Færdigsamlet kode | Knapplads |
| --- | --- | --- |
| Opret medarbejder | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.6/office-scripts/3.6.6/create_employee.ts) | Medarbejdere, A5:B5 |
| Tilføj projekt | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.6/office-scripts/3.6.6/add_project.ts) | Min ressourceplan, A2:B2 |
| Opret nyt planår | [Download](https://raw.githubusercontent.com/Allanrasmus/Ressourceplan/ressourceplan-3.6.6/office-scripts/3.6.6/initialize_plan_year.ts) | Min ressourceplan, D2:F2 |

På **Min ressourceplan** vælger du din personlige arkvisning. Ved
indsættelse midt i PlanTimer understøttes **Medarbejder** og **Planår**
med ét Custom-lighedskriterium pr. kolonne, hver for sig eller sammen.
Excel Desktop kan registrere almindelige enkeltvalg som Custom, f.eks.
`=2026`. Hvis Excel registrerer et andet kriterium, stopper scriptet før ændring af PlanTimer. Fjern det ikke-understøttede filter manuelt i
din egen arkvisning, hvis du vil fortsætte. Genbrug af en ledig sagsrække
og tilføjelse nederst ændrer ingen filtre, når kriterierne er understøttet.

Afprøv de tre testscripts i Desktop og Online og med to konti på SharePoint.
Kør administrationen én ad gangen i drift. Arkvisninger giver ingen
transaktionslås. Ved fejl skal du kontrollere filtre, PlanTimer og eventuelle
ændringer i medarbejderregisteret eller planårslisten, før du prøver igen.
[Fuld testplan](https://github.com/Allanrasmus/Ressourceplan/blob/ressourceplan-3.6.6/TESTPLAN.md#testjournal-for-office-scripts-366)

Efter accept gemmes de tidligere scriptversioner til gendannelse. Erstat
hele koden under **Automatiser → Alle scripts**, og gem med samme
scriptidentitet, så de eksisterende knapper kan genbruges.

Kør `python office-scripts/3.6.6/build.py` for at bygge scripts og ZIP-fil.

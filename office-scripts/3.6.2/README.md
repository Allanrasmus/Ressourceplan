# Opdatér de eksisterende Office Scripts til 3.6.2

Brug jeres eksisterende Ressourceplan 3.6 eller 3.6.1 med indtastede
timer og stamdata. Version 3.6.2 består af tre scriptfiler.

| Kode til Excel | Eksisterende script |
| --- | --- |
| [create_employee.ts](create_employee.ts) | Opret medarbejder |
| [add_project.ts](add_project.ts) | Tilføj projekt |
| [initialize_plan_year.ts](initialize_plan_year.ts) | Opret nyt planår |

1. Aftal et tidspunkt med de andre brugere af de delte scripts. Brug
   en acceptkopi af arbejdsfilen til prøvekørsler. Hvis de eksisterende
   knapper refererer til delte scripts, får alle tilknyttede projektmapper
   adgang til den opdaterede kode, når scriptet gemmes under samme identitet.
2. Åbn **Automatiser → Alle scripts** som ejer eller med redigeringsret.
   Åbn det eksisterende script i kodeeditoren. Kopiér hele indholdet af
   den tilsvarende fil ovenfor, og erstat koden i editoren.
3. Gem det eksisterende script under samme identitet, så knapperne kan
   genbruges. Gentag for alle tre. Kontrollér, at Excel kan gemme og køre
   hvert script uden aliasing-fejl.
4. Prøv de eksisterende knapper og **Automatiser** på acceptkopien.
   Vælg én udfyldt celle i **Medarbejder** eller **Projektvalg** i en
   planrække, eller en projektcelle i **Projektregister**, når Tilføj
   projekt skal genbruge en markering. Kør fra **Automatiser**, hvis en
   knap ændrer markeringen. Dialogfelterne er de samme som i 3.6.1.
   Ved afvisning kan scriptets ejer gendanne koden fra versionsmappen,
   der var installeret før opdateringen.

[Arbejdsgange og parameterbeskrivelser](../3.6.1/README.md)
· [Accepttest for 3.6.2](../../TESTPLAN.md#testjournal-for-362)
· [Microsoft om scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

Genbyg de tre filer fra kilderne i `src` med følgende kommando fra repositoryets rod.

```bash
python office-scripts/3.6.2/build.py
```

# Ressourceplan – status og overdragelse

**25. november 2026 · Office Scripts 3.6.4 afventer manuel accept**

Udviklingsgren `ressourceplan-3.6.4` bygger på `27e9b56` fra 3.6.3.
De tre driftsscripts håndterer nu aktive filtre ved indsættelse før
eksisterende rækker i PlanTimer. Excel 3.6.3 bruges uændret.
[Download og installation](office-scripts/3.6.4/README.md)

Alle 72 automatiske tests bestod. De ni nye tests dækker filtreret
indsættelse, fejlhåndtering og scriptkompatibilitet. Microsoft Excel Desktop og Online samt flerbrugertest
på SharePoint afventer manuel accept. Arkvisninger adskiller visninger,
mens tabelændringer gemmes i den fælles fil.
[Testjournal](TESTPLAN.md#testjournal-for-office-scripts-364)

Brug en acceptkopi og separate testscripts. Kør
`python office-scripts/3.6.4/build.py` og `python -m unittest discover -s tests -v`
med Node, TypeScript 4.0.3 og LibreOffice. Sæt `TSC` efter behov.

[Status og åbne punkter fra 3.6.3](https://github.com/Allanrasmus/Ressourceplan/blob/27e9b56fe6616a47818ac13ea511d4c380a95894/STATUS.md)

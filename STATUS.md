# Ressourceplan – status og overdragelse

**25. november 2026 · Office Scripts 3.6.5 afventer manuel accept**

Excel Desktop afviste den fælles 3.6.4-filterkode i **Opret medarbejder** og
**Tilføj projekt** med en aliasingfejl. **Opret nyt planår** havde samme kald
i den færdigsamlede kode. 3.6.5 gemmer kriterier som lokale dataværdier og
opbygger kriterieargumentet direkte ved gendannelse. Excel 3.6.3-projektmappen
genbruges.

Alle 79 automatiske tests bestod. Kør den særskilte filterprøve i Excel før accept af driftsscriptene.
Automatiske testresultater og vejledning til manuel accept i Excel Desktop,
Excel Online og med to konti på SharePoint står i testjournalen.
[Download og installation](office-scripts/3.6.5/README.md)
[Testjournal](TESTPLAN.md#testjournal-for-office-scripts-365)

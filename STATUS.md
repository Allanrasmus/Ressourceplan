# Ressourceplan – status og overdragelse

**25. november 2026 · 3.6.3 afventer manuel accept**

Udviklingsgren `ressourceplan-3.6.3`. COWIs manuelle Excel Desktop-test af
3.6.2 godkendte 26 af 27 testpunkter. Alle punkter blev afprøvet. Den
resterende fejl var konverteringen af et afdelingsnummer fra punktum til komma.
`main` afventer manuel accept af 3.6.3.
[Udviklingsgren på GitHub](https://github.com/Allanrasmus/Ressourceplan/tree/ressourceplan-3.6.3)

## Leveret

- **Opret medarbejder** gemmer afdelingsnummeret som tekst og kontrollerer
  den gemte værdi. Medarbejder-ID, projekt- og afdelingskoder bevares.
- **Tilføj projekt** genbruger en sikkert tom sagsrække på relevant plads
  eller indsætter Sag og Tilbud lokalt. Registrerede timer flyttes med rækken
  af Excel. Kun projekter fra Projektregister accepteres.
- **Opret nyt planår** indsætter manglende standardaktiviteter i samme
  rækkefølge. De tre driftsscripts genbruger de eksisterende scriptknapper.
- Engangsscriptet opdaterer registerformater og synlige sagsstriber i
  den eksisterende fil. Sagsfarverne er #EAF2FC og #FFFFFF. Den separate
  [3.6.3-reference](Ressourceplan_3_6_3.xlsx) indeholder demonstrationsdata.
  [Opgraderingsvejledning](office-scripts/3.6.3/README.md)

## Gennemførte og afventende tests

61 automatiske tests bestod. Testjournalen dækker scriptkode, simuleret afdelingskonvertering
og lokale indsættelser, Excel-pakkens struktur samt filtre, farveregler og
beregninger i separate LibreOffice-kopier. Manuel accept af 3.6.3 i Excel
Desktop med dansk formatering og i Excel Online udestår.
[Testresultater og manuel accept](TESTPLAN.md#testjournal-for-363)

## Begrænsninger og videre udvikling

- Historiske fejlplacerede rækker bevarer deres placering, indtil en
  administrator sorterer hele PlanTimer i et særskilt vedligeholdelsesvindue.
  Eksisterende afdelingsværdier rettes kun efter opslag i den autoritative
  kilde.
- PlanTimer har ingen transaktionslås. Kontrollér indsættelse, genbrug,
  farver, personlige arkvisninger og AutoSave med to samtidige konti.
- [Grupperingsprototypen](Ressourceplan_3_6_1_Gruppeprototype.xlsx) er
  uændret. Kvartalsgruppering og medarbejderdetaljer afventer særskilt
  afprøvning. Øvrige åbne ønsker er dynamiske dialogvalg, overførsel af
  sager og tilbud, masseindtastning, ledelsesdashboard og særskilt
  **Opret projekt** i Projektregister.
- Fælles produktionsfil og produktionsscripts på SharePoint afventer
  test og godkendelse. Test og produktion bruger særskilte scripts, og
  virkelige data bliver i det godkendte Microsoft 365-miljø.
  [Projektkrav](Projektkrav.md#fremtidig-drift-i-microsoft-365)

## Genoptag arbejdet

Start på `ressourceplan-3.6.3` og læs [AGENTS.md](AGENTS.md). Kør
`python office-scripts/3.6.3/build.py` og
`python -m unittest discover -s tests -v` med Node, TypeScript 4.0.3 og
LibreOffice. Sæt `TSC` efter behov. Godkend først opgraderingen og
knapperne på en acceptkopi med separate testscripts. Registrér
resultaterne i testjournalen, før driftsopsætning eller merge.

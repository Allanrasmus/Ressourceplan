# Ressourceplan – status og overdragelse

**24. november 2026 · Version 3.6.1 afventer manuel accept**

Arbejdet fortsætter på `ressourceplan-3.6.1`. Regneark, kildekode og tests
til 3.6.1 blev leveret i implementeringscommit
`bf9bff212933628a98608393ed30c2fb0f64b15c`. De tidligere versioner er
bevaret. Merge til `main` afventer manuel godkendelse af version 3.6.1.
[Udviklingsgren på GitHub](https://github.com/Allanrasmus/Ressourceplan/tree/ressourceplan-3.6.1)

## Leverance og implementerede funktioner

- [Ressourceplan_3_6_1.xlsx](Ressourceplan_3_6_1.xlsx) bevarer direkte
  timeindtastning, PlanTimer, kompakte oversigter, stabile ID'er og den
  afgrænsede uge 53-regel. Planstart er skjult som administrativ reserve.
- **Opret medarbejder** vælger første sikre plads, springer reserverede
  ID'er med historik over og bevarer tildelte ID'er som p31.
- **Tilføj projekt** ignorerer tilbudsprocent ved Sag. Ved Tilbud giver
  tomt felt 50 %, og indtastet 0 bevares. Valideret markering i planen
  eller projektregisteret kan levere kontekst.
- **Opret nyt planår** kan genbruge medarbejderen fra en markeret planrække
  og tilføjer kun manglende standardaktiviteter. Genkørsel bevarer timerne.
- [Gruppeprototypen](Ressourceplan_3_6_1_Gruppeprototype.xlsx) har egne
  outline-ark for kvartaler og medarbejderdetaljer. Ugefelterne på
  Kvartalsprototype er redigerbare kopier. Medarbejderprototype beregner
  timer fra PlanTimer for et udsnit, der er valgt ved bygning.

## Office Scripts

De tre færdigsamlede scripts og deres kilder ligger i
[office-scripts/3.6.1](office-scripts/3.6.1). De er beregnet til både
3.6 og 3.6.1. COWI afprøvede de tidligere 3.6-scripts i Excel Desktop.
Tilføj projekt virkede dér efter en manuel rettelse af Sag-valideringen.
De opdaterede scripts og parameterdialoger i 3.6.1 afventer accept.
Behold arbejdsfilen med COWI's egne timer og opdatér koden i de samme
scriptidentiteter på en prøvekopi først. En ny downloadet arbejdsfil skal
tilknyttes de eksisterende scripts i Microsoft 365.
[Opdaterings- og brugervejledning](office-scripts/3.6.1/README.md)

## Gennemførte tests

48 automatiske tests bestod, heraf 6 nye og 42 tidligere. De tre
3.6.1-scripts blev kompileret med TypeScript 4.0.3 mod dokumenterede
ExcelScript-signaturer. Simulering og genberegning af separate
LibreOffice-kopier dækkede blandt andet historiske ID'er, procentvalg,
genkørsel, markering, filtrering og summer med skjulte rækker og kolonner.
11 OOXML-dele i driftsfilen og 14 i prototypen bestod skemavalidering.
Begge filer kunne genbygges med samme indhold. [Testjournal](TESTPLAN.md#testjournal-for-361)

## Afventende manuelle tests

1. Afprøv den nye kode i de eksisterende scripts og knapper i Excel
   Desktop og Online. Kontrollér den ændrede parameterdialog,
   tom procent, eksplicit 0, markering og tilføjelse under aktive filtre.
2. Afprøv 3.6.1-filen på en acceptkopi med samme scriptidentiteter.
   Kontrollér åbning, direkte timeindtastning og genberegning.
3. Test to brugere i Teams eller SharePoint med hver sin arkvisning.
   Kontrollér samtidige oprettelser, filtre, AutoSave og versionshistorik.
4. Prøv uafhængigt plus/minus i 52- og 53-ugers prototypekopier.
   Kontrollér, om en anden konto ser ændringer i grupperingen.

## Kendte fejl, begrænsninger og åbne ønsker

- Den rapporterede Sag-fejl fra 3.6 er rettet og regressionstestet
  automatisk. Manuel bekræftelse i COWI's Excel-miljø udestår.
- Office Scripts-dialogen har ingen dokumenteret dynamisk dropdown fra
  registertabeller. De validerede markeringer og tabellens egne dropdowns
  giver sikre arbejdsgange. Knapadfærd ved markering afventer accept.
- Excel-tabeller giver ingen transaktionslås til samtidige
  scriptkørsler. Risikoen for dubletter og delvist gennemførte
  handlinger består. Inden fælles drift skal de samtidige scenarier
  ovenfor afprøves.
- Kvartals- og medarbejdergruppering i produktionsarkene er et åbent
  ønske. Prototypens medlemskab og projektetiketter fastlægges ved
  bygning. Et nyt år eller udsnit kræver genbygning. Dynamisk
  vedligeholdelse og isolation pr. bruger er endnu uafklaret.
  [Design og prototype](DESIGN.md#outline-prototype-i-361)
- Udfasning af ubrugte, forudreserverede ID'er er et senere ønske.
  Eksisterende ID'er og historiske relationer skal fortsat bevares.

## Genoptag arbejdet

Start en ny samtale på `ressourceplan-3.6.1`, og læs denne fil samt
[AGENTS.md](AGENTS.md). [README](README.md) beskriver lokal genbygning af
regneark og scripts. Kør testen fra repositoryets rod med
`python -m unittest discover -s tests -v` og med Node, TypeScript 4.0.3
og LibreOffice Calc tilgængelige. Brug `TSC` til TypeScript-kompilatorens sti,
hvis den ikke er på PATH.

Prioritér først de manuelle acceptpunkter for 3.6.1 og dokumentér
resultaterne i TESTPLAN.md. Ret eventuelle fejl på udviklingsgrenen.
Fortsæt derefter den separate outline-afprøvning, især de to brugeres
visning. Tag først stilling til produktionsgruppering efter den test.

# Ressourceplan – status og overdragelse

**25. november 2026 · Office Scripts 3.6.2 afventer manuel accept**

Arbejdet fortsætter på `ressourceplan-3.6.2`. Den seneste arbejdsfil er
[3.6.1](Ressourceplan_3_6_1.xlsx). Scriptrettelsen bruges også med 3.6.
Merge til `main` afventer manuel godkendelse.
[Udviklingsgrenen på GitHub](https://github.com/Allanrasmus/Ressourceplan/tree/ressourceplan-3.6.2)

## Office Scripts og leverance

- COWI rapporterede tre aliasing-fejl i 3.6.1 ved accepttest i Excel Desktop.
  Tilføj projekt og Opret nyt planår sendte et indlejret Excel-tabelobjekt
  videre til en hjælpefunktion. 3.6.2 bruger tabelnavne, indekser og
  almindelige celleværdier i den fælles kode.
  [Fejlsteder, ydeevne og statisk regression](TESTPLAN.md#testjournal-for-362)
- **Opret medarbejder** bruger den første sikkert ledige plads og bevarer
  stabile ID'er og historik. Scriptet tilføjer manglende standardaktiviteter.
- **Tilføj projekt** understøtter eksplicitte input og valideret markering.
  Sag ignorerer procentfeltet. Ved Tilbud giver et tomt felt 50 %, og et
  indtastet 0 bevares.
- **Opret nyt planår** genbruger medarbejderen fra en valideret planrække og
  tilføjer manglende standardaktiviteter. Genkørsel bevarer projekter og timer.
- De tre scripts ligger i [office-scripts/3.6.2](office-scripts/3.6.2).
  Opdatér dem under deres eksisterende identiteter, så knapperne kan genbruges,
  og prøv dem med en kopi af arbejdsfilen.
  [Opdateringsvejledning](office-scripts/3.6.2/README.md)

## Gennemførte tests

54 automatiske tests bestod den 25. november 2026. De 48 tidligere tests
og seks nye blev kørt. De nye omfatter statisk genfinding af de tre
fejlsteder, særskilt TypeScript-kompilering af hvert script, simulering
med op til 20.000 planrækker og beregning i en separat LibreOffice-kopi.
SHA-256 for driftsfilen og grupperingsprototypen matcher 3.6.1.
[Testjournal og accepttrin](TESTPLAN.md#testjournal-for-362)

## Afventende manuelle tests

1. Prøv at gemme og køre alle tre scripts i Excel Desktop og Online.
   Kontrollér fejl og eventuelle læseadvarsler i editoren.
2. Prøv eksisterende knapper og **Automatiser** med eksplicitte input,
   relevante markeringer, Sag og Tilbud, aktivt filter og genkørsel.
3. Mål kørsel på en stor acceptkopi, og prøv to samtidige brugere med
   separate arkvisninger i Teams eller SharePoint. Kontrollér AutoSave,
   versionshistorik og bevarede filtre, formler og timer.
4. Afprøv grupperingsprototypen separat i 52- og 53-ugers år. Kontrollér
   plus/minus og om en anden konto ser ændringer i grupperingen.

## Begrænsninger og åbne udviklingsønsker

- Sikkerhedsopslag før skrivning reducerer risikoen for konflikter.
  Excel-tabeller har ingen transaktionslås til samtidige scriptkørsler.
- Dynamiske registerbaserede dropdowns i scriptdialogen er et åbent
  ønske. Markering og tabellernes egne dropdowns er tilgængelige.
- Overførsel af tilbud mellem medarbejdere, masseindtastning af timer
  og ledelsesdashboard afventer udvikling.
- Den separate [grupperingsprototype](Ressourceplan_3_6_1_Gruppeprototype.xlsx)
  er bevaret. Kvartalsarket har redigerbare kopier af ugefelter.
  Medarbejderarket beregner fra PlanTimer for det udsnit, der blev
  valgt ved bygning. Dynamisk vedligeholdelse, produktionsgruppering
  og isolation af brugernes visninger afventer afprøvning.
  [Design og prototype](DESIGN.md#outline-prototype-i-361)
- Udfasning af ubrugte, forudreserverede ID'er er et senere ønske.
  Eksisterende ID'er og deres historiske relationer skal bevares.

## Genoptag arbejdet

Start på `ressourceplan-3.6.2`, og læs [AGENTS.md](AGENTS.md).
Byg scripts med `python office-scripts/3.6.2/build.py`. Kør
`python -m unittest discover -s tests -v` fra repositoryets rod med Node,
TypeScript 4.0.3 og LibreOffice Calc tilgængelige. Sæt `TSC` til
TypeScript-kompilatorens sti efter behov.

Gennemfør Excel-accept først, og registrér resultaterne i
[TESTPLAN.md](TESTPLAN.md#testjournal-for-362). Afprøv derefter
grupperingens visning med to brugere, før produktionsgruppering vurderes.

# Versionshistorik (fra samtalen)
- HTML v2.8: lokal prototype med fem visninger, projektregister, medarbejderadministration, tilbudspipeline og prognose. Ikke fælles datalager.
- Excel 2.0–2.8: overgang til makrofri Excel; direkte indtastning, design fra HTML. 2.8 gav reparationsdialog og inputproblemer.
- 2.9: synlige filtre var et mål, men bruger rapporterede manglende pile; fireugers oversigter.
- 3.0: 30 pladser, 2026–2031 som rækker, 53 ugekolonner. Faste rækkeområder, sortering usikker. Bruger oplevede mange gentagne rækker og problemer med årsvalg.
- 3.1: blå striber, kvartalsfarver og synlige summer. Oprydning og automatisk standardfilter ikke færdige.
- 3.2: tilbudsfarve, skjult tilbudsprocent ved andre typer, dynamiske datoer, ID og tydeligere projektvalg forsøgt. Bruger rapporterede manglende kvartaler og synlige beregningstabeller.
## Excel 3.3

- 179 tomme standardtilbud ændret til ledige sagsrækker. Eksisterende sagsnumre,
  timer, tilbudsprocenter og registerdata er bevaret.
- Kvartalsoverskrifter, ISO-mandage og uge 53 beregnes ud fra planåret. Datoer
  udelades, når planen viser flere år.
- Dynamiske aktivitetsfarver, blå sagsstriber, validering af timer og tilbudsprocent,
  skjult ID og et gemt medarbejder-/årsfilter i planen.
- Ressourceoverblik med medarbejderfiltre, periodemarkering, kapacitet, særskilt
  ledig tid og overbelastning samt brutto- og vægtede tilbud. Beregningsrækker
  efter I ALT er skjult.
- Projektoversigt med markeret sagsvalg, aktuelt projektnavn, periodetimer og
  årstotal. Genbygning og validering kan køres fra repositoryet.

## Excel 3.3.1

- Betingede farver bruger Excel-kompatibel differentialformatering. Planens
  tabelstil giver plads til aktivitetsfarver og blå/hvide sagsstriber.
- Tilføjet workbook-visning og sammenhængende valg i fastlåste ruder.
- Ressourceoverblikket har bredere overskrifter, farvekoder for kvartal, periode
  og belægning samt ensartede time- og procentenheder i I ALT.
- Projektoversigten samler uger og medarbejdere fra venstre og oppefra.
  Fastlåst projektvalg og totaler pr. medarbejder, uge og periode.
- Tidligere versionsfiler og plan-/registerdata bevares. Se TESTPLAN.md for
  automatiske kontroller og manuel accept i Microsoft Excel.

## Excel 3.4

- Fælles COWI-layout på alle fem faner, med tydelige valgfelter og adskilte
  nøgletal. Tekniske beregningsfelter er skjult.
- Ressourceoverblik samler filtrerede medarbejdere, viser antal og har I ALT
  direkte efter sidste række. Projektoversigt bruger den samme totalplacering.
- Projektlisten viser projektnavn før sagsnummer og følger oprettelser og
  navneændringer i Projektregister. Sagsnummeret er fortsat beregningsnøglen.
- Ensartede tabelkanter, striber og mintfarvede totalfelter. Aktivitets-,
  kvartals- og belægningsfarver er bevaret.
- Version 3.3.1 bevares som stabil reference. Dens inputdata og
  beregningsregler er videreført.

## Excel 3.5

- Min ressourceplan har én voksende Excel-tabel, beregnede tabelkolonner og
  nøgler på hver række. Flytning og sortering af hele rækker bevarer timernes
  tilknytning. Tomme standardrækker fra 3.4 er fjernet.
- Medarbejdervalg og projektvalg læses automatisk fra registrene. Det skjulte
  ID og sagsnummer bruges i beregningerne. Navne kan ændres uden at ændre
  nøglen.
- Ressourceoverblik og Projektoversigt medtager de nye planrækker samt op
  til 100 medarbejdere. Projektregisteret har plads til 500 projekter.
- Tilføjet brugervejledning og regressionstest for ekstra rækker, sortering,
  omdøbning, identiske projektnavne og medarbejder nummer 31.

## Excel 3.5.1

- Erstattet direkte tabelreferencer i registerarkenes unikke nøglevalidering
  med lokale celleområder, som i 3.4. Dropdowns og normtidsvalidering er bevaret.
- Én synlig medarbejderkolonne til valg og filtrering. Det beregnede ID og
  navneopslag er skjult og følger den stabile nøgle ved omdøbning.
- Genskabt 149 manglende standardrækker uden at ændre eksisterende timer.
  Planstart gør det muligt at foreslå og indsætte manglende rækker for nye
  medarbejdere eller nye planår.
- Nye planår tilføjes i Planstart. Alle tidligere planår og den dynamiske
  plantabel bevares, sammen med projekt- og ressourceoversigterne.

## Excel 3.6

- Indført selvstændige Office Scripts til oprettelse af medarbejder,
  tilføjelse af projekt og initialisering af planår.
- Klargjort tre knappladser. Tilknytning og knapper installeres i Excel.
  Planstart er skjult som administrativ reserve.
- Grå markering i uge 53 er afgrænset til tabellens relevante rækker i
  år med 52 uger. Validering og sumlogik er bevaret.
- Undersøgt native række- og kolonnegrupper. Se DESIGN.md for
  prototype og resterende testpunkter før en frigivelse.

## Excel 3.6.1

- Medarbejderoprettelse vælger første sikkert ledige plads. Reserverede ID'er
  med historik springes over. Eksisterende ID'er bevares.
- Tilføj projekt ignorerer tilbudsprocent på Sag. Tilbud skelner mellem tom
  tekst, 0 % og øvrige gyldige procenter.
- Projekt og planår kan bruge en valideret markering. Et entydigt navn kan
  bruges i stedet for en teknisk kode. Fælles inputfelter er undgået.
- Tilføjet opdateringsvejledning for eksisterende Office Scripts og knapper.
- Leveret separat outline-prototype med kvartaler og medarbejderdetaljer.
  Driftsfilen bevarer de kompakte oversigter og direkte timeindtastning.
- Den 24. november 2026 bestod 48 automatiske tests. Den efterfølgende
  manuelle Excel Desktop-test fandt tre aliasing-fejl, som følges op i 3.6.2.

## Office Scripts 3.6.2

- Rettet de tre fejlsteder i Tilføj projekt og Opret nyt planår, hvor
  et Excel-tabelobjekt fra `Plan.table` blev givet videre til en
  hjælpefunktion. Alle tre scripts er bygget fra den rettede fælleskode.
- Samlet læsninger af planens nøglekolonner, valgt række og rå ugefelter.
  Formelkolonner skrives i sammenhængende blokke. Dublet- og
  konfliktkontroller før skrivning er bevaret.
- Excel-arbejdsfilen og grupperingsprototypen fra 3.6.1 er uændrede.
  Behold arbejdsfilen med jeres data, og gem koden i de eksisterende scripts.
  [Opdateringsvejledning](office-scripts/3.6.2/README.md)
- Den 25. november 2026 bestod 54 automatiske tests. COWI godkendte
  efterfølgende 26 af 27 punkter i Excel Desktop. Afdelingsfejlen
  følges op i 3.6.3. [Resultater](TESTPLAN.md#testjournal-for-362)

## Ressourceplan 3.6.3

- Opret medarbejder skriver afdelingskoder som tekst og verificerer den
  gemte værdi. Punktum og nuller bevares, og normtid forbliver numerisk.
- Tilføj projekt genbruger en sikker tom sagsrække på den relevante plads
  eller indsætter Sag og Tilbud lokalt i aktiviteternes rækkefølge.
  Opret nyt planår og Opret medarbejder bruger samme placeringsprincip
  til standardrækker. Ukendte projekter henvises til Projektregister.
- Sagsstriber bruger #EAF2FC og #FFFFFF efter synlige sager. Regler,
  farver og kanter for øvrige aktiviteter, uge 53 og oversigter bevares.
- Tilføjet engangsscript til eksisterende filer og en separat Excel-reference
  med demonstrationsdata. Dokumenteret særskilt sortering af historiske
  rækker og fremtidig drift på fælles SharePoint.
  [Opgraderingsvejledning](office-scripts/3.6.3/README.md)
- 61 automatiske tests bestod, og 11 OOXML-dele blev skemavalideret.
  Manuel accept i Microsoft Excel udestår. [Testjournal](TESTPLAN.md#testjournal-for-363)


Den samlede udgave supplerer opgraderingsvejen med en Excel-fil, hvor de
eksisterende demonstrationsrækker står i den aftalte rækkefølge fra start.
Version 3.6.3 vises på alle seks faner. ZIP-pakken indeholder Excel-filen,
de tre driftsscripts, engangsscriptet og vejledningen til installation og
opgradering. [Samlet pakke](Ressourceplan_3_6_3.zip)

## Office Scripts 3.6.4

- Indsættelse før eksisterende PlanTimer-rækker rydder midlertidigt aktive
  kolonnefiltre i den aktive personlige arkvisning. Kriterierne forsøges
  genskabt efter både succes og fejl. Genbrug og tilføjelse nederst
  ændrer ingen filtre.
- De tre driftsscripts bruger den fælles rettelse til `InsertDeleteConflict`.
  Ved afvist genskabelse angiver fejlbeskeden de berørte kolonner.
- Excel-filen og de tidligere scriptversioner er bevaret. Manuel
  flerbrugeraccept på SharePoint udestår.
  [Opdateringsvejledning](office-scripts/3.6.4/README.md)

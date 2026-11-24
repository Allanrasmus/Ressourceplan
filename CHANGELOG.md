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

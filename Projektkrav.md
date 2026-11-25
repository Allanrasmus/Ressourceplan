# Projektkrav
## Rammer
- Fem faner: Min ressourceplan, Ressourceoverblik, Projektregister, Projektoversigt, Medarbejdere.
- 30 medarbejderpladser i første version, mulighed for dokumenteret udvidelse. 2026–2031, senere udvidelse uden at historik overskrives. ISO-uger, inkl. uge 53 hvor relevant. Højst 53 ugekolonner i medarbejderplanen, år som fast nøgle i rækker.
- Medarbejdere indtaster kun timer i Min ressourceplan. Direkte redigerbare celler, filtrering på navn og planår, personlige arkvisninger i delt fil. Ingen formelbaseret FILTER-visning som foregiver at være redigerbar. Filtre er ikke adgangskontrol. Åbn helst med én medarbejder og ét år; vis tydelig vejledning. Færre tomme rækker og ingen unødige gentagelser. Bevar blå striber. Fastlås relevante kolonner/overskrifter.
- Projektregister med unikt stabilt sagsnummer og navn; navneændringer slår igennem. Medarbejderregister med stabilt ID, navn, område Øst/Vest, faggruppe EL/Belysning/BMS/CTS/Sikring, afdeling og normtid. ID skjules i planen.
- Aktivitetstyper: Sag, Tilbud, Ferie, Fravær, Kursus, Andet. Timer 0–80 pr. række/uge. Tilbudssandsynlighed som procent, standard 50 % hvis ikke angivet, kun synlig/relevant ved tilbud.
- Bookede timer = Sag + Andet. Disponibel kapacitet = max(0,normtid minus Ferie, Fravær og Kursus). Tilbud separat, brutto og sandsynlighedsvægtet. Ved nul kapacitet og positive bookede timer vis overbelastning særskilt. Ledige timer og overbelastning summeres separat pr. person/uge, ikke netto.
- Ressourceoverblik: filtre område, faggruppe, afdeling, år, startuge og periode 4/8/12/26 uger, Q1–Q4 eller hele året; årsoversigt med farvekoder og valgt periode markeret orange. Nøgletal og kapacitetsprognose; senere medarbejderkort og projektfordeling. Projektoversigt: tydeligt projektvalg, år/periode, timer pr. person/uge og total. Projektbudget/faktisk forbrug er senere muligheder, ikke krav til 3.3.
## Version 3.3 – første opgave
1. Undersøg baseline 3.2. Bevar eksisterende tilbud med timer; omdan tomme standardtilbudsrækker til ledige sagsrækker. Ingen fast tilbudsrække pr. person. Leder kan oprette tilbud i en ledig række. På sigt skal et vundet tilbud kunne overføres ved valg af medarbejder og manuel ændring til Sag. Nu er faste rækkeområder en kendt begrænsning: ændr ikke medarbejder-ID/navn direkte før beregningsarkitektur er gjort sikker; dokumentér sikker kopi af indtastningsfelter til modtagerens række og rydning af kilden. Undgå dobbeltregning.
2. Genskab kvartalsoverskrifter med korrekte ISO-ugestartdatoer for valgt år. Viser flere år, skal tvetydige datoer undgås; år med 52 uger må ikke tælle uge 53.
3. Dynamisk rækkefarve efter type (DESIGN.md), også når type ændres. Tilbud % skjules ved andre typer; uden makro kan manuel værdi ikke automatisk slettes, så ignorer den i beregninger og dokumentér dette.
4. Skjul beregningstabeller under I alt i Ressourceoverblik, bevar formler og links.
5. Bevar skjult ID, synlige filtre, tydeligt projektvalg, eksisterende timer, formler og 30 medarbejderpladser. Undersøg og ret eventuelle regressionsfejl.
## Senere
Udvid medarbejderpladser og år sikkert, forbedr direkte omfordeling af tilbud via stabile nøgler, håndtér helligdage, beskyttelse af formler uden at blokere redigering/filtre, bedre dashboards, eventuel tidsregistreringseksport. Test i Teams/Excel Online og Desktop før drift.

## Version 3.5 – dynamisk administration

- Den tidligere blok med 12 planrækker pr. medarbejder og år erstattes af en
  voksende, redigerbar Excel-tabel. Hver række bærer et eksplicit planår, en
  medarbejderidentitet, aktivitetstype og sine timer.
- Beregningerne bruger hele tabellen og stabile medarbejder- og projektnøgler.
  Sortering af tabellen bevarer sammenhængen mellem rækker og input.
- Projektnavn og sagsnummer vælges i samme dropdown. De beregnede felter bruger
  sagsnummeret og viser det aktuelle navn. Projekter med ens navn adskilles
  af nummeret. Medarbejdervalg og opslag i registeret bruger et varigt ID.
- Brug en ledig registerrække til nye medarbejdere og projekter. Den leverede
  kapacitet er 100 medarbejdere, 500 projekter og 20.000 validerede planrækker.
  Der forudoprettes ingen planrækker pr. ny medarbejder.
- Bevar indtastninger, årshistorik, aktivitetsfarver, direkte indtastning og
  begge kompakte oversigter. Data fra 3.4 er migrationsgrundlag. Rækkens nøgle
  muliggør fremtidig overførsel af samme række med bevarede timer.

## Version 3.5.1 – validering og første ressourceplan

- Nøglevalidering i Projektregister og Medarbejdere bruger lokale cellereferencer.
  Unikke medarbejder-ID'er og sagsnumre kontrolleres fortsat.
- Medarbejder vælges og filtreres i samme synlige kolonne. Den gemte tekst
  indeholder navn og kode. ID-formlen og opslag af aktuelt navn er skjult.
- Bevar alle eksisterende timer. For hver registreret medarbejder og hvert
  understøttet planår skal planen som udgangspunkt have en ledig Sag, Ferie,
  Fravær, Kursus og Andet. Tilbud oprettes manuelt efter behov.
- Et makrofrit Planstart-ark foreslår kun manglende standardrækker og kan
  oprette valgmuligheder for nye år. Brugeren indsætter forslag som værdier
  i den dynamiske plan, uden at ændre historikken.
- Bevar de godkendte funktioner og kapaciteten fra 3.5. Afprøv åbning,
  registervalidering og AutoSave i Microsoft Excel med filen i SharePoint.

## Version 3.6 – Office Scripts

- Arbejdsbogen er fortsat .xlsx og uden VBA. Office Scripts må anvendes
  via Excel Desktop og Excel Online, hvor de understøttes. Scriptkørsler
  bruger projektmappen i Microsoft 365, og beregningerne kræver ingen
  eksterne databaser eller tjenester.
- Lever selvstændige TypeScript-scripts til oprettelse af medarbejder,
  tilføjelse af projekt og initialisering af planår. Brug scriptparametre,
  bevar stabile nøgler, og kontrollér input og dubletter før indsættelse.
- Bevar direkte indtastning, dynamisk plantabel, oversigter og filtre.
  Scriptkørsler må ikke rydde eller ændre filtre globalt. Brugeren
  tilknytter selv scripts og knapper i Microsoft Excel.
- Afgræns uge 53-markeringen til de relevante tabelrækker. Validering og
  beregning følger fortsat årets faktiske antal ISO-uger.
- Behold Planstart skjult som administrativ reserve, indtil scripts er
  manuelt godkendt. Indfør først kvartals- og medarbejdergruppering efter
  stabilitetstest med Excel og delte projektmapper.

## Version 3.6.1 – sikre valg, medarbejderpladser og prototyper

- Vælg den første sikkert ledige plads. Brug en genkendelig reserveret kode,
  når pladsen er ubrugt og koden ikke findes i planhistorikken.
  Eksisterende ID'er og historiske rækker bevares.
- Sag gemmer ingen tilbudsprocent. Tilbud accepterer 0–100 som tekst,
  herunder 0, og bruger 50 % ved tom tekst.
- Undgå delt input til scripts. Brug registeropslag og en valideret enkelt
  celle i PlanTimer eller Projektregister til kontekst. Uændret manuel
  timeindtastning bruger tabellens dynamiske dropdowns.
- Bevar Planstart skjult og uge 53 afgrænset til den eksisterende tabel.
  Scriptkoderne kan opdateres på et kompatibelt 3.6-dokument uden at migrere timers data.
- Kvartals- og medarbejdergrupper leveres separat, indtil filtre, årsændringer
  og samtidige brugeres visning er afprøvet i Microsoft Excel.
- Den næste milepæl er manuel accept af version 3.6.1 på en kopi af
  arbejdsfilen samt en separat afprøvning af gruppeprototypen. Merge til
  `main` kræver manuel godkendelse af 3.6.1. [Aktuel status](STATUS.md)

## Version 3.6.2 – kompatibilitetsrettelse til Office Scripts

- Ret Excel-editorens aliasing-fejl i den fælles kildekode. Hjælperne
  henter Excel-objekter gennem ExcelScript og udveksler almindelige
  værdier, tabelnavne og indekser. Læs data i blokke, og behold
  dublet- og konfliktkontroller før skrivning.
- Fortsæt med eksisterende arbejdsfiler, scriptknapper, inputmuligheder og
  kapacitet på 100 medarbejdere, 500 projekter og 20.000 planrækker.
  Bevar den separate 3.6.1-grupperingsprototype. Afprøv begge Excel-klienter
  før merge til `main`. [Acceptpunkter](TESTPLAN.md#testjournal-for-362)
- Udviklingsplanen omfatter fortsat dynamiske dropdowns i scriptdialogen,
  overførsel af tilbud mellem medarbejdere, masseindtastning af timer,
  produktionsgruppering og ledelsesdashboard. [Aktuel status](STATUS.md)

## Version 3.6.3 – tekst, rækkefølge og sagsstriber

- Medarbejder-ID, medarbejdernavn, område, faggruppe, afdelingsnummer,
  sagsnummer og projektnavn behandles som tekst. Normtid, planår, planlagte
  timer og tilbudsprocent er numeriske. Bevar allerede registrerede værdier
  og datavalidering.
- Vis Sag, Tilbud, Ferie, Fravær, Kursus og Andet i den rækkefølge pr.
  medarbejder og år. Genbrug en sikkert tom sagsrække på relevant plads.
  Brug ellers lokal tabelindsættelse. Eksisterende fejlplacering korrigeres
  separat med hele planlægningstabellen og stabile ID'er.
- Sagsstriberne bruger #EAF2FC og #FFFFFF for skiftevis synlige sager.
  Særlige aktivitetsfarver, kanter, kvartaler, belægning og totaler bevares.
- Tilføj projekt henter altid et eksisterende projekt fra Projektregister.
  **Opret projekt** er et senere, særskilt script i Projektregister, der
  skal validere sagsnummer og navn, afvise dubletter og gøre nye projekter
  tilgængelige i de eksisterende vælgere.
- Lever en komplet `Ressourceplan_3_6_3.xlsx` bygget fra Excel 3.6.1
  sammen med de tre driftsscripts videreudviklet fra 3.6.2. Fanerne viser
  3.6.3, og Excel-filen har sagsstriber, ordnet PlanTimer og tre klargjorte
  knappladser. En særskilt opgradering af eksisterende arbejdsfiler bevarer
  timer, stamdata og scriptidentiteter. Godkend 3.6.3 manuelt før merge til
  `main`. [Installations- og opgraderingsvejledning](office-scripts/3.6.3/README.md)

## Fremtidig drift i Microsoft 365

Produktionsfilen placeres på sigt på fælles SharePoint. Produktionsscripts
placeres så vidt muligt samme sted med fælles ejerskab. Driften skal være
uafhængig af en enkelt medarbejders personlige OneDrive. Test og produktion
bruger separate scriptfiler. GitHub og Codex er til udvikling og
versionsstyring med fiktive data. Virkelige medarbejdere, projekter og timer
forbliver i virksomhedens godkendte Microsoft 365-miljø. Afprøvning og
godkendelse af driftsstrukturen udestår.

## Ressourceplan 3.6.7 og senere projektadministration

Bruttotilbud summerer periodens tilbudstimer uden vægtning. Vægtede tilbud
summerer timer gange tilbuddets procent. Et nyt tilbud får 50 % som
standard, og eksplicit 0 % bevares. Begge tilbudstal beregnes særskilt fra
bookede timer og almindelig belægning.

Projektregister er fortsat forudsætningen for både Sag og Tilbud. En
statuskolonne og et særskilt script til oprettelse af projekter og tilbud
i Projektregister hører til den fremtidige udviklingsplan.

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

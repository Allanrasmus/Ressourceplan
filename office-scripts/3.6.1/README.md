# Opdatering og brug af Office Scripts 3.6.1

Kilderne i denne mappe er færdigsamlede til Microsoft Excel. Scripts kan
bruges med Ressourceplan 3.6 og 3.6.1.

| Kilde | Eksisterende navn | Knapplads i en ny projektmappe |
| --- | --- | --- |
| [create_employee.ts](create_employee.ts) | Opret medarbejder | Medarbejdere, A5:B5 |
| [add_project.ts](add_project.ts) | Tilføj projekt | Min ressourceplan, A2:B2 |
| [initialize_plan_year.ts](initialize_plan_year.ts) | Opret nyt planår | Min ressourceplan, D2:F2 |

## Opdatér de allerede installerede scripts

1. Åbn jeres arbejdskopi i Microsoft Excel. Find det eksisterende script
   under **Automatiser → Alle scripts**, og redigér det som ejer eller med
   tilladelse til at redigere scriptfilen.
2. Kopiér hele koden fra den tilsvarende fil ovenfor, erstat den
   eksisterende kode, og gem scriptet under samme identitet. Gentag
   for alle tre scripts. Vælg **Gem** i det eksisterende script frem for
   **Lav en kopi**, hvis I vil bevare de eksisterende tilknytninger.
3. Prøv den eksisterende knap, og kontrollér inputdialogen. Knapper er
   knyttet til det delte script og bør fortsat referere til det, så længe
   scriptet gemmes samme sted med samme identitet. Tilføj projekt har
   ændrede, valgfrie tekstparametre for planår og tilbudsprocent. Kontrollér især,
   at dialogen viser tekstfelter, og at tom procent og eksplicit 0 giver
   hver sin virkning. Åbn eventuelt filen igen, eller kør via **Automatiser**,
   hvis en ældre dialog stadig vises. Tilknyt knappen igen, hvis den fortsat
   viser den gamle signatur.
4. Behold den arbejdsfil, der indeholder jeres egne timer og medarbejdere.
   De opdaterede scripts kan køres direkte på en kompatibel 3.6-fil. Ved
   skifte til en ny arbejdsfil skal eksisterende data medbringes. Brug
   migrationskoden på en separat kopi af jeres 3.6-fil, hvis I vil have
   3.6.1-layoutet med jeres data. Kontrollér knapperne bagefter. Den
   versionsbyggede .xlsx i repositoryet har klargjorte knappladser.

Microsoft gemmer Office Scripts som separate filer og knytter dem til
projektmappen. Ved en ny downloadet Excel-version åbner I den nye fil,
tilknytter det samme script med **Add in workbook** eller **Associate with
workbook**, og tilføjer knappen på det relevante ark. Dette er en tilknytning
til det eksisterende script. Til fælles ejerskab kan scriptfiler ligge i
SharePoint med redigeringsret for de relevante vedligeholdere. Hvis en
scriptfil flyttes eller erstattes af en kopi, kontrollér tilknytninger til
alle arbejdsfiler før ibrugtagning.
[Microsoft om lager og deling](https://learn.microsoft.com/en-us/office/dev/scripts/overview/script-storage)
[Microsoft om scriptknapper](https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons)

## Brug scripts uden tekniske ID'er

**Opret medarbejder** tager navn, område, faggruppe, afdeling, normtid og
planår. Scriptet vælger den første sikkert ledige plads. En reserveret
plads med planhistorik springes over. En identisk profil genbruger det
allerede tildelte ID og opretter kun manglende standardrækker. Et ID som
p31 flyttes aldrig over på en anden kode.

**Tilføj projekt** kan bruges på tre måder:

- Markér én synlig celle i **Medarbejder** i en eksisterende planrække.
  Kør scriptet fra **Automatiser** og efterlad medarbejder og planår tomme.
  Vælg typen og skriv et entydigt projektnavn eller hele projektetiketten,
  for eksempel `Hospital Nord · 24015`.
- Markér en celle i **Sagsnr.** eller **Projektnavn** på den ønskede række i
  **Projektregister**. Kør scriptet fra **Automatiser**. Lad projektfeltet
  stå tomt, og angiv medarbejderens entydige navn og planår. Nye projekter
  kan bruges, så snart de er registreret.
- Markér en udfyldt **Projektvalg**-celle i en eksisterende planrække, hvis
  den række skal levere medarbejder, år og projekt. Vælg aktivitetstypen
  i dialogen. Scriptet kontrollerer for dubletter. Du kan også skrive
  alle valg udtrykkeligt uden at markere en celle.

Til **Sag** ignoreres enhver tilbudsprocent. Til **Tilbud** kan teksten stå
tom for standarden på 50 %. Skriv `0` for 0 % eller f.eks. `25` for 25 %.
Dansk decimalkomma accepteres. Entydige navne virker som input. Ved ens
navne kan du markere den relevante række eller bruge hele etiketten med
kode. Brug enten en kompatibel markering eller angiv år og medarbejder
udtrykkeligt. Scriptet henter aldrig et år fra en række for en anden
medarbejder.

**Opret nyt planår** bruger navn og år fra en markeret planrække, når de
tilsvarende felter er tomme. Markér medarbejderen og skriv et nyt år,
f.eks. `2032`, for at oprette manglende standardaktiviteter i det år.
Genkørsel opretter ingen dubletter.

Scriptets aktive markering kan afhænge af Excel-klient og knapadfærd.
Brug **Automatiser** efter markeringen, hvis en knap mister den. Eksisterende
dynamiske dropdowns er tilgængelige ved direkte indsættelse af en ny
planrække. Du kan dér vælge medarbejder, år og projekt uden at kende en kode.
Ændr kun disse felter på din egen nye række, når formålet er at tilføje en post.

## Dynamiske valglister og samtidig redigering

Office Scripts' egen parameterdialog kan vise lister fra faste TypeScript-
værdier som Sag/Tilbud. Den har ingen dokumenteret funktion til at udfylde
parametervalg fra levende Excel-tabeller. Medarbejder og projekt findes i de
aktuelle registre, og årslisten ligger i den skjulte Planstart. En kontrolleret
markering genbruger native tabelvalg. Personlige arkvisninger kan adskille
filtre, mens værdier i fælles inputceller fortsat deles. Derfor bruger
løsningen ingen fælles inputformular. Valg i planrække eller register er
læsekontekst, og scripts justerer ingen aktive filtre.
[Microsoft om scriptparametre](https://learn.microsoft.com/en-us/office/dev/scripts/develop/user-input)

Microsoft dokumenterer valgfrie parametre, men angiver ingen entydig
værdi for tomme numeriske felter på tværs af Excel-klienter. I COWI's 3.6-test
udløste Sag den gamle numeriske validering. Den nye tekstparameter skelner
tom tekst fra `0` og genkender også `null` og manglende værdi. Den konkrete
dialog i jeres Microsoft 365-miljø skal stadig prøves efter opdateringen.

Brug personlige arkvisninger, genanvend filteret efter tilføjelse og kør
fælles skrivehandlinger én ad gangen i samme projektmappe. Scripts
validerer før ændringer, genlæser ledig plads og dubletter umiddelbart før
oprettelse og kontrollerer de indsatte rækker før formelskrivning.
Excel-tabeller har ingen transaktionslås mellem disse trin. Samtidige
kørsler, sortering eller ændring af registre kan derfor give konflikt
eller en delvist gennemført handling. Ved afbrudt medarbejderoprettelse
kan samme profil genkøres. Kontrollér nye planrækker og deres formler
før timeindtastning, hvis Excel rapporterede en delvis indsættelse.

AutoSave er et separat forhold ved samredigering. ExcelScript har
`Workbook.getAutoSave()` til at læse status. Scripts ændrer ikke den
indstilling. Kontrollér AutoSave og versionshistorik ved åbning af jeres
arbejdsfil fra Teams eller SharePoint, og udfør to-brugertesten i TESTPLAN.md.
[Microsoft om AutoSave](https://support.microsoft.com/en-us/office/what-is-autosave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5)

## Genbyg scriptfilerne

Kør `python office-scripts/3.6.1/build.py` fra repositoryets rod efter
ændringer under `src`. 3.6-kilderne i den overordnede mappe er den
tidligere version. Til udviklingstest bruges Node og TypeScript 4.0.3.
Angiv `TSC` til kompilatoren, hvis den ikke ligger på PATH.

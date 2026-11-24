# Testplan
- Bevar original baseline og generér ny fil med versionsnavn. Kontrollér ZIP-integritet, XML-parsing, openpyxl genindlæsning, gyldige formler, celletyper, tabelreferencer og synlige filterpile. Undersøg tidligere Excel-reparationsfejl; ingen ugyldige fejlværdier. Undgå at forveksle strukturkontrol med faktisk Excel-beregning.
- Kontrollér dropdowns, indtastningsfelter, skjult ID og beregningsrækker, kvartals-/datoformler, formateringsregler for alle typer, tilbudsprocentens visning og projektvælger.
- Test 37 normtimer, 20 Sag, 7 Ferie, 10 Tilbud (50 %): kapacitet 30, booket 20, ledig 10, belægning 66,7 %, brutto 10, vægtet 5. Andet tæller som booket; Fravær/Kursus reducerer kapacitet. Test også nul kapacitet med positiv belastning og ledig/overbelastning separat for to personer.
- Test projektoversigt med flere medarbejdere, ændring af type, bevarelse af eksisterende tilbudstimer og navneskift. Test 2026 uge 53, 2027 og 2028, årsskifte og filtrering. Årsvalg må ikke ændre lagrede timer. Ingen uge 53 i år uden den. Kontrollér at summer ikke brydes ved filtrering; hvis sortering er usikker, dokumentér det.
- Test udvidelse/aktivering af ledig medarbejderplads, eksisterende datavalidering og bevarelse af baseline. Dokumentér hvordan man udvider over 30 medarbejdere og 2031.
- Brug automatiseret test af logik og, hvor tilgængeligt, en faktisk beregningsmotor; openpyxl beregner ikke formler. Sæt genberegning ved åbning. Manuel accept: åbning uden reparationsdialog, direkte timeindtastning, filtre, skift år og type, farver, nøgletal, samtidig redigering og personlige arkvisninger i Excel Desktop/Online. Angiv særskilt hvad der ikke er testet.

## Testjournal for 3.3

Den 24. november 2026 bestod alle 9 automatiske tests. Beregningstestene
omfattede 7 genberegnede filer og fandt ingen celler med formelfejl. Kørslen
brugte Python 3.12.14, openpyxl 3.1.5 og LibreOffice Calc 26.8 (udviklingsbuild).
I et nyt virtuelt miljø med de fastlåste pakker bestod desuden alle 4
strukturtests, herunder sammenligning med den udgivne fil.

Kør `python -m unittest discover -s tests -v` fra projektmappen. Testene bruger
fiktive data og ændrer kun midlertidige kopier.

| Tjek | Dækning |
| --- | --- |
| Fil og migration | ZIP/OOXML, formulatokens, tabeloverskrifter, filtermetadata, datavalideringer, genindlæsning og genbygning. SHA-256 for originalen og sammenligning af alle inputfelter og registre. |
| Standardvisning og formatering | Gemt filter for 2026 og Medarbejder 1, 12 synlige planrækker, skjult ID, fastlåste kolonner, skjulte beregninger og regler for farver og procentvisning. |
| Referenceberegning | 20 booket, 30 disponibel, 10 ledig, 66,7 % belægning, 10 brutto og 5 vægtede tilbudstimer i uge 48. |
| Aktivitet og prognose | Andet, Fravær, Kursus, nul kapacitet, særskilt ledig/overbelastning for to personer, blank og eksplicit nul tilbudsprocent, samt område-, faggruppe- og afdelingsfiltrering. |
| Projekt og historik | Omdøbning, aktivering af p06, vundet tilbud flyttet til en anden person med uændret projekttotal, og særskilte historiske årsværdier. |
| ISO-uger | Samtlige mandage og kvartaler i 2026, 2027 og 2028. Uge 53 med i 2026, udeladt i 2027 og 2028 selv med indsatte værdier. Datoer udeladt ved visning af flere år. |

LibreOffice Calc genberegner testfilerne, og testen læser beregnede værdier tilbage
og kontrollerer regnearksfejl. Filen til udgivelse er openpyxl-builden.

### Manuel accept før drift

Åbn den udgivne fil i Excel Desktop og Excel Online, og kontrollér åbning uden
reparationsdialog, synlige filterpile, valg af ét år og navneskift i filtret,
timeindtastning, dropdowns, typefarver og skjult tilbudsprocent. Prøv derefter
samtidig redigering og personlige arkvisninger med to brugere i Teams/SharePoint.
Strukturtests verificerer filter- og formateringsreglerne i filen. Den visuelle
gengivelse og brugerinteraktionen i Microsoft Excel kræver denne manuelle accept.

# Instruktioner til Codex
Svar på dansk. Udvikl et Excel-værktøj (.xlsx) uden VBA til fælles ressourceplanlægning i Teams/SharePoint. Brug Python og openpyxl som udgangspunkt; behold en reproducerbar build- eller migrationskode og automatiske tests i repository. Undersøg eksisterende workbook og HTML før ændringer. Brug kun fiktive data. Office Scripts er tilladt fra version 3.6. Gem selvstændige scripts under office-scripts og brug den officielle ExcelScript API. Ingen VBA, eksterne databaser eller tjenester som funktionel forudsætning ud over den understøttede Microsoft 365-platform. Slutbrugeren skal kunne bruge projektmappen uden Python. Bevar originalfiler og indtastede data; skriv nye versionsfiler. Foretag små verificerbare ændringer, dokumenter begrænsninger og rapporter faktiske testresultater. Undgå at påstå at Excel Desktop/Online er testet, hvis det ikke er tilfældet. Kontrollér OOXML, formler, valideringer, filtre og beregningslogik. Se Projektkrav.md, DESIGN.md og TESTPLAN.md. Prioritér stabilitet og direkte indtastning frem for kosmetik. Foretag ikke ændringer i hovedbranch uden brugerens godkendelse.

For 3.6.1 og nyere skal Office Scripts genbruge en eksplicit, valideret
markering eller lokale scriptparametre. Brug ikke fælles indtastningsceller
som personlig scriptdialog. Bevar tidligere scriptversioner. Hold grupperings-
prototyper adskilt fra driftsfilen, indtil samtidige arkvisninger er afprøvet.

Hold test- og produktionsscripts i separate Microsoft 365-scriptfiler.
Produktionsfil og fælles produktionsscripts placeres på sigt på et
SharePoint-websted med fælles ejerskab, uafhængigt af en personlig OneDrive.
Brug kun fiktive data i GitHub og Codex. Virkelige medarbejder-, projekt-
og timedata forbliver i virksomhedens godkendte Microsoft 365-miljø.
Gem identificerende registerfelter som tekst og timer, planår og procenter
som tal. Brug lokal tabelindsættelse og kontrollér nøgler og rækkeantal før
skrivning. Historisk sortering af hele tabellen er en særskilt
vedligeholdelseshandling, som kræver afprøvning og en eksklusiv arbejdsperiode.

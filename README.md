# Handleiding voor de leraar — Les 12: Aan, CC, BCC en een bijlage

<!-- Beginvragen (stap 1 van CONTEXT.md), 21-10-2026. Jonas vroeg "Je mag beginnen met de lessen voor volgende week. Buiten
     3MWWE"; hij koos in de chat (21-10-2026) telkens het voorstel:
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 12 van het jaarplan (module 3), uur 1 van W09: 4ORLOb di 27 oktober (8ste uur),
     3ORLO di 27 oktober (9de uur), 4ORLOa wo 28 oktober (4de uur) · vorige les: les 11 (De zakelijke e-mail), gegeven op
     wo 21-10 (3ORLO, 4ORLOb) en do 22-10 (4ORLOa), dus vóór les 12 in elke klas → lesstart = ophalen uit les 11 ·
     toestel lokaal 18, Windows 11 · werkdocument CO2_Aan-CC-BCC-en-een-bijlage (reeks Communicatie, opdracht
     *Communicatie 2 — Aan, CC, BCC en een bijlage*, onderwerp *Digitale competenties*), 20 punten, deadline vr 30-10 20.00 ·
     de e-mail in een nagebootst venster Nieuw bericht in het werkdocument (Aan, Cc, Bcc, Onderwerp, Bijlage), want de
     leerlingen hebben geen Gmail (Jonas, 09-10-2026); er wordt niets verstuurd · bijlage: hun **eigen pdf uit les 04**
     (de onthaalbrochure), als chip (@) in het vak Bijlage; wie ze niet vindt, typt de naam (Jonas, 21-10-2026) ·
     schrijfhulp zoals les 11 sinds 21-10: per bouwsteen de feiten en een zinsbegin, en *Mag je overnemen* bij de
     voorbeeldzinnen (Jonas, 21-10-2026; voorstel uit de na de les van 3MWWE, 15-10, en 3ORLO, 21-10) · geen zip ·
     context NovaDepot · repository Groepsmail-NovaDepot. -->

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, *Lotte mailt de onthaalbrochure naar vijf jobstudenten* — module 3 van het jaarplan, les 12 (reeks *Communicatie 2*)
**Toestel:** **lokaal 18, Windows 11** — alles in Chrome, geen bestanden, geen Gmail
**Lesdag:** 4ORLOb di 27 oktober (8ste uur) · 3ORLO di 27 oktober (9de uur) · 4ORLOa wo 28 oktober (4de uur)
**Kernleerplandoel:** `BV2_04.01` (toepassen) · ondersteunend `BV2_04.04`, `BK2_02.01`
**Deadline:** vrijdag 30 oktober 2026, 20.00 uur · 20 punten
**Pages:** https://jonasdaltongent.github.io/Groepsmail-NovaDepot/ · dia's: …/presentatie.html

Lotte Maes (personeelsdienst) wil de onthaalbrochure naar vijf nieuwe jobstudenten mailen, en Tom (hun verantwoordelijke) in kopie.
De leerling kiest eerst bij vijf situaties *Aan*, *Cc* of *Bcc*, leest wie welk adres ziet, en schrijft dan de mail **voor Lotte** in het
venster *Nieuw bericht* van het werkdocument (**Aan, Cc, Bcc, Onderwerp en Bijlage**): *Aan* leeg, Tom in *Cc*, de jobstudenten in
*Bcc*, en de eigen pdf uit les 04 als chip in het vak *Bijlage*. Er wordt niets verstuurd. Opbouw en bronnen: `lesvoorbereiding.md`.

---

## 1. Inhoud van het pakket

```text
W09 - Les 12 - ORLO - Aan CC BCC en een bijlage/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 10 dia's: 1 Zo start je · 2 lesstart (les 11) · 3 lesdoel · 4 de situatie · 5 rollenspel Wie ziet welk adres? · 6 de bijlage (demo) · 7 controle (Bcc) · 8 stap 1 samen · 9 kloktijden · 10 exitvraag
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/                      # logo's, lettertypes, mascotte, het filmpje Inleveren, één schermafbeelding (stapjes/s5-codewoord.png)
├── werkdocument/
│   ├── CO2_Aan-CC-BCC-en-een-bijlage.docx  # het werkdocument dat de leerling INLEVERT (kopie per leerling)
│   ├── maak_werkdocumenten.py              # maakt het opnieuw (python-docx)
│   ├── maak_stapjes.py                     # maakt de uitsnede bij het stapje "Typ het codewoord" uit de pdf van Google Documenten
│   └── huisstijl.py, go-dalton-gent.png
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Geen zip en geen bronbestanden. De leerlingen hebben alleen hun **pdf uit les 04** nodig (die staat in hun Drive).

### 1b. Hoe de lespagina werkt

Zoals les 11: links de **route** (*Wie krijgt de mail?*: stap 1–2 · *De mail van Lotte*: stap 3–5), midden **één stap** met *Toon stapjes* per
handeling (de les begint in *Kort*, zonder knop), rechts de **checklist** met 17 taken. Geen *Hulp nodig?*: grijze regels onder de handeling en één tip per stap.

- **Startpagina** met drie beelden (de opdracht in Classroom → deze lespagina → je werkdocument CO2) en *Dit maak je vandaag*: het venster met vijf velden (de chip bij *Bijlage*) en de zes bouwstenen als balkjes.
- **Stap 2**: het infokader *Aan, Cc of Bcc?* met een tekening van drie vakjes met een oog (open = iedereen ziet dit adres, doorgestreept = niemand ziet de andere adressen in Bcc), en de mail van Lotte uit deel 2.
- **Stap 3 en 4: de feiten bij elke handeling** (`.feiten`, zichtbaar in *Kort*, zoals les 11 sinds 21-10-2026) met het groene label **Mag je overnemen** bij de zinnen die de leerling letterlijk mag nemen (het adres van Tom, het begin van het onderwerp, *Beste jobstudenten,*, de zinsbegin). Geen volledige zinnen voor de hele mail.
- **Stap 3, handeling Bijlage**: drie stapjes met tekeningen (`@` typen, *TV3* typen, de pdf kiezen = een chip) en de uitweg *Typ de naam: Onthaalbrochure NovaDepot.pdf*.
- **Stap 5: inleveren pas na de zelftest, met een codewoord** (zoals les 09 en 11): drie vragen (Bcc bij een groep, wie Emre ziet, de bijlage in de tekst), daarna het blok *Afwerken en inleveren*; de groene regel toont **PAPERCLIP**, dat de leerling bovenaan *Deel 4 — Nalezen* typt.
- **Extra** (telt niet mee): de naam van de pdf volgens de afspraak van les 06 (*Wat - Van wie*). De pagina zegt hoe de leerling opnieuw inlevert (**Ongedaan maken**, dan weer **Inleveren**).
- `localStorage` bewaart alleen vinkjes en de laatste stap (voorvoegsel `novadepot_groepsmail_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository `jonasdaltongent/Groepsmail-NovaDepot`, Pages op `main`, map `/ (root)`; adres hierboven. Dat adres staat op dia 9, in `classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling

`classroom.json`: **Communicatie 2 — Aan, CC, BCC en een bijlage**, onderwerp *Digitale competenties*, 3ORLO, 4ORLOa en 4ORLOb, 20 punten, deadline vrijdag 30 oktober 2026, 20.00 uur.
Bijlagen: de link naar de lespagina en `CO2_Aan-CC-BCC-en-een-bijlage` (**Een kopie maken voor elke leerling**). Geen zip. Status: zie de regel in `_index/logboek.md` (21-10-2026); de lesbouwer maakt alleen het **concept**, inplannen doet de hoofdsessie (een half uur vóór de les: di 27-10 om 14.05 voor 4ORLOb en 14.55 voor 3ORLO, wo 28-10 om 10.45 voor 4ORLOa).

Instructietekst:

```text
1. Open de lespagina (link) en je werkdocument CO2_Aan-CC-BCC-en-een-bijlage.
2. Je schrijft de mail in je werkdocument. Er wordt niets verstuurd.
3. Klik op Inleveren, ten laatste vrijdag 30 oktober 2026 om 20.00 uur.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/CO2_Aan-CC-BCC-en-een-bijlage.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads converteren* aan) en voeg het toe met **Bijvoegen** › **Drive** › **Een kopie maken voor elke leerling**.
2. Voeg de lespagina toe met **Link**.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 9 klopt.
- [ ] De opdracht staat in de drie klassen, met twee bijlagen (de link en het werkdocument).
- [ ] Een testleerling krijgt een eigen kopie van het werkdocument, met de eigen naam in de titel.
- [x] `@` stelt een pdf uit Drive voor en maakt er een chip van, ook in een tabelcel (getest met demoleerling3, 21-10-2026; § 2b).
- [ ] Les 11 is in die klas gegeven (dia 2 haalt er een vraag uit op); zo niet: sla dia 2 over. Kijk in de app *Na de les* of er over les 11 iets bij te sturen is.
- [ ] Vier kaartjes (A5) voor het rollenspel op dia 5: *Lotte*, *Tom*, *Lina*, *Emre*.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Werkdocument openen | in de opdracht rechts bij **Jouw werk** op de naam klikken | `_afspraken/klikpaden.md` |
| Een bijlage als chip toevoegen | in het vak **Bijlage** `@` typen, een stuk van de bestandsnaam typen (*TV3*), het bestand in de lijst kiezen: er komt een **chip** | [Documenten-help 10710316](https://support.google.com/docs/answer/10710316?hl=nl) (*Smartchips invoegen*), gelezen 21-10-2026, en **getest met demoleerling3 op 21-10-2026** (een test-pdf *Demo Leerling - TV3_Documentopmaak.pdf* in Drive, daarna naar de prullenbak): `@TV3` toont onder **Bestanden** de pdf met een **rood pdf-icoon** (ook als *TV3* midden in de naam staat), daaronder **Plaatsen** van Google Maps (*TV3* in Spanje, Ghana …). Klikken geeft een chip, ook in een tabelcel; in een smal vak wordt de naam afgekort met *…*. In de Drive van een echte leerling staat ook het Google-document van les 04 met dezelfde naam (blauw icoon): de pagina zegt *kies het rode pdf-icoon*. Uitweg: de naam typen |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** › status **Ingeleverd** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl); het venster gezien met demoleerling3 (15-10-2026) |
| Opnieuw inleveren (Extra) | **Ongedaan maken**, daarna weer **Inleveren** | zelfde bron en dezelfde proef |

Geen knop *Sturen* of *Verzenden*: de mail wordt niet verstuurd.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten (1'). Dia 2 lesstart (3'): welke bouwstenen van les 11 ontbreekt (*Slot*), met vingers; de brug *"Toms mail is vertrokken; vandaag mailt Lotte vijf mensen tegelijk, met een bijlage"*. Dia 3–6 (5'): lesdoel; de situatie; **rollenspel *Wie ziet welk adres?*** met vier kaartjes (lees de tekening met de ogen de eerste keer hardop voor); **demo van de bijlage** (`@`, *TV3*, de pdf kiezen). Dia 7 controle (1'): *Ziet Noor het adres van Kobe, die ook in Bcc staat?* |
| **Samen** | **3'** | Dia 8: iedereen doet stap 1; duim op. |
| **Keuzewerktijd** | **tot het einde** | Dia 9 blijft staan, met **kloktijden** (`_tools/zet_lesmomenten.py`); stap 2 tot 5, inleveren inbegrepen. Dia 10 in de laatste minuut, met *Vertrek je? Log uit, of Windows + L*. |

Stap 1–5: 3 + 9 + 8 + 12 + 8 = **40 minuten**.

**Eerste rondgang:** heeft iedereen het **eigen** werkdocument open (naam in de titel)? **Tweede rondgang, rond stap 3:** staan de jobstudenten in **Bcc** en Tom in **Cc**? Zit er een bijlage in het vak? Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert. Minimumroute: deel 1, deel 3 en inleveren; deel 2 mag later, vóór de deadline. Wie de pdf niet vindt, typt de naam: dat telt even goed. Zoek niet mee in Drive.

---

## 4. Verbetersleutel

- **Codewoord** bovenaan *Deel 4 — Nalezen*: **PAPERCLIP**. Ontbreekt het, dan is de zelftest van stap 5 niet (juist) gemaakt.

De leerling levert alleen het werkdocument in (20 punten: 4 · 3 · 13). De volledige verdeling staat in `lesvoorbereiding.md` § 22.

### Deel 1 — Aan, Cc of Bcc? (4)

Rij 1 staat er al (Kim Wouters: **Aan**). 1 punt per rij 2 tot 5, alleen voor het veld; een reden in plaats van het veld = 0.

| Rij | Situatie | Veld |
|---|---|---|
| 2 | Sara Claes (preventie) moet de koelcel-onderhoud alleen weten | **Cc** |
| 3 | nieuwjaarswens aan 40 klanten die elkaar niet kennen | **Bcc** |
| 4 | Lotte vraagt Tom het uurrooster na te kijken | **Aan** |
| 5 | planning naar de 8 collega's van Toms ploeg | **Aan** (collega's die samenwerken) |

### Deel 2 — Wie ziet welk adres? (3)

| Vraag | Antwoord | Punten |
|---|---|---|
| a. Ziet Lina het adres van Tom? | **ja** (Cc ziet iedereen) | 1 |
| b. Ziet Lina het adres van Emre? | **nee** (Bcc ziet niemand van de anderen) | 1 |
| c. Waarom Bcc? (één zin) | ze kennen elkaar niet, en hun privéadres is een persoonsgegeven | 1 |

*Omdat het moet* of *zo staat het in de les* telt niet.

### Deel 3 — De mail van Lotte (13)

| Criterium | Punten |
|---|---|
| *Aan* leeg, Tom (`tom.peeters@novadepot.example`) in *Cc*, de vijf adressen in *Bcc*, letter voor letter | 3 |
| een precies onderwerp | 1 |
| de brochure in het vak *Bijlage* (chip of de naam) en de tekst zegt dat ze in de bijlage zit | 2 |
| *Beste jobstudenten,*, een slotgroet zonder afkorting, en een inleiding die zegt wie schrijft en waarom | 2 |
| de boodschap: maandag 21 december, 7.30 uur, het onthaal, de identiteitskaart, de infosessie om 8.00 uur | 2 |
| slot (tegen vrijdag 11 december antwoorden) en overal *u*, zonder chattaal of emoji's | 2 |
| handtekening met Lotte Maes, personeelsdienst, NovaDepot en toestel 205 | 1 |

Een goede mail (een van vele):

```text
Aan:        (leeg)
Cc:         tom.peeters@novadepot.example
Bcc:        lina.haddad@mail.example, emre.yilmaz@mail.example, noor.verbeke@mail.example,
            kobe.desmet@mail.example, ayla.janssens@mail.example
Onderwerp:  Onthaalbrochure en eerste werkdag op maandag 21 december
Bijlage:    [chip: … TV3_Documentopmaak … .pdf]

Beste jobstudenten,

Mijn naam is Lotte Maes, van de personeelsdienst van NovaDepot. Welkom bij NovaDepot!

In de bijlage vindt u onze onthaalbrochure. Uw eerste werkdag is maandag 21 december. U komt om 7.30 uur naar het
onthaal en brengt uw identiteitskaart mee. Om 8.00 uur is er een infosessie in de kantine.

Kunt u mij tegen vrijdag 11 december laten weten dat u komt? Alvast bedankt.

Met vriendelijke groeten,

Lotte Maes
Personeelsdienst · NovaDepot · toestel 205
```

De controletabel onderaan het werkdocument heeft dezelfde criteria (samen 20); de leerling typt een x, jij zet de punten in dezelfde rij.

### Essentiële fouten — geef hier altijd feedback op

- Een jobstudent in *Aan* of *Cc*: iedereen ziet dan de privéadressen.
- Geen bijlage, en niets over de bijlage in de tekst.
- Geen datum of geen uur van de eerste werkdag.
- *U* en *je* door elkaar, chattaal of emoji's.
- Geen naam onder de mail.

Verbeteren en opnieuw inleveren mag. De ontvangers en de bijlage komen terug in E1 (les 15).

---

## 5. Schermafbeeldingen (optioneel)

Eén: `assets/stapjes/s5-codewoord.png`, een uitsnede van *Deel 4 — Nalezen* uit de pdf van het Google-document. Alle andere beelden zijn tekeningen in HTML en CSS (ook de lijst met suggesties en de chip: dat zijn **tekeningen**, geen echte schermafbeeldingen). Zie `assets/screenshots/LEESMIJ.md`.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt `CO2_Aan-CC-BCC-en-een-bijlage.docx` opnieuw. Vereist `python-docx` en `Pillow`. De uitsnede bij het codewoord: `python3 werkdocument/maak_stapjes.py <pdf van het Google-document>` (vereist `PyMuPDF`).

`werkdocument/huisstijl.py` is de versie van het sjabloon, met de parameter `velden` bij `mailvenster` (hier: Aan, Cc, Bcc, Onderwerp, Bijlage).

---

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json`: `BV2_04.01` (kern), `BV2_04.04` en `BK2_02.01` (ondersteunend). De repository krijgt de `pre-push` hook uit `_sjablonen/pre-push`.

---

## 8. Wat nog moet blijken in de klas

- **De chip (`@`) met een pdf: getest met demoleerling3 op 21-10-2026** (§ 2b): de pdf staat in de lijst en wordt een chip, ook in een tabelcel. Wat nog moet blijken: kiezen de leerlingen de **pdf** (rood icoon) en niet het Google-document van les 04 met dezelfde naam (blauw icoon) of een plaats van Google Maps? En wat leest de verbetertool uit een chip (de naam van het bestand, verwacht)?
- **Staat de pdf uit les 04 in de Drive van de leerling?** Ze voegden ze toe via *Jouw werk* › *Toevoegen of maken* › *Bestand*; wie ze niet maakte of niet inleverde, heeft ze niet. Hoeveel leerlingen typen de naam?
- **Haalbaarheid in 40 minuten.** Stap 4 (12 minuten) vraagt zes stukken tekst, plus de feiten lezen; les 11 was al krap. Zakt het tempo, dan is de minimumroute de uitweg.
- **Copy/paste van de feiten.** Lezen ze de feiten per bouwsteen, en schrijven ze nog eigen zinnen? (*Mag je overnemen* is nieuw in ORLO.)
- **Bcc: vijf adressen kopiëren** uit het kader van het werkdocument in Google Documenten, met een komma ertussen. Tikfouten? Lukt kopiëren uit een tabelkader?
- **Schrijven in het venster** van het werkdocument met vijf velden: groeit het vak *Bijlage* mee met een chip met een lange naam? Alleen nagekeken in de opbouw, niet met een leerling aan het toetsenbord.
- **Rollenspel (dia 5) en dia 6** zijn nieuwe soorten tekeningen (ogen, lijst met suggesties): hoeven ze meer uitleg dan de notities geven?
- **De Extra en *Ongedaan maken*:** zoals les 11: klikken leerlingen op *Ongedaan maken* en vergeten ze opnieuw in te leveren?
- De lespagina en dia's zijn gemeten met scripts en in headless Chrome bekeken, niet door een mens op een Windows-toestel.

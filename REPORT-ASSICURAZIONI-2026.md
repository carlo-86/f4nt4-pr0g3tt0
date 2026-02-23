# Report Definitivo Assicurazioni - Mercato Invernale 2026

## Fonti dati utilizzate
- **DB completo**: 06/02/2026 (foglio DB → FVM Prop., foglio SQUADRE → status assicurativo pre-mercato)
- **Rose**: 17/02/2026 (TutteLeRose → Spesa aggiornata post-mercato per TUTTI i giocatori)
- **Conteggi svincoli**: 07/02/2026
- **Comunicazioni assicurazioni**: 13-14/02/2026
- **Scambi sessione mercato invernale 2026**: PDF con tutti i trasferimenti

## Formula costo assicurativo
Estratta dalla colonna E del foglio ROSA del DB Excel:
```
Se Spesa ≤ FVM Prop.:  Costo = MAX(FVM_Prop / 10, 1)
Se Spesa > FVM Prop.:  Costo = MAX(MEDIA(Spesa, FVM_Prop) / 10, 1)
```
- **Spesa** = prezzo d'asta (dal file Rose 17/02, che include acquisizioni dell'asta riparazione)
- **FVM Prop.** = Fantacalcio Valuation Model Proprietà (dal foglio DB, colonne BK:CB, indipendente dalla squadra)

---

## SCAMBI SESSIONE MERCATO INVERNALE 2026

### Fanta Tosti 2026 (1 scambio)
| Data | Da | A | Giocatore | Crediti |
|------|-----|-----|-----------|---------|
| 05/02 | PARTIZAN | FC CKC 26 | Immobile | +23cr a CKC |
| 05/02 | FC CKC 26 | PARTIZAN | Adams C. | |

### FantaMantra Manageriale (7 operazioni + svincoli)
| Data | Da | A | Giocatore | Crediti |
|------|-----|-----|-----------|---------|
| 03/02 | Hellas Madonna | Minnesota al Max | Lazaro | |
| 03/02 | Minnesota al Max | Hellas Madonna | Coco | |
| 04/02 | Minnesota al Max | Mastri Birrai | Hojlund | |
| 04/02 | Mastri Birrai | Minnesota al Max | David | |
| 04/02 | Legenda Aurea | H-Q-A Barcelona | Odgaard | +15cr a HQA |
| 04/02 | H-Q-A Barcelona | Legenda Aurea | Barella | |
| 06/02 | Legenda Aurea | Minnesota al Max | Bernabè | |
| 06/02 | Minnesota al Max | Legenda Aurea | Cancellieri | |
| 11/02 | Minnesota al Max | Papaie Top Team | Hien | +22cr a Minnesota |
| 13/02 | Minnesota al Max | Papaie Top Team | Diego Carlos | +2cr a Minnesota |
| 13/02 | Minnesota al Max | Lino Banfield FC | Mazzitelli, Tavares N., Koopmeiners, Bernabè | |
| 13/02 | Lino Banfield FC | Minnesota al Max | Fagioli, Bellanova, Gimenez, Miller L. | |

**Svincoli FM (06/02):**
- Legenda Aurea svincola: Bianchetti, Angelino, Zerbin, Pobega, Sohm
- FC CKC 26 svincola: Zanoli, Carboni V., Lang

---

## CALCOLO COSTI ASSICURATIVI

### FANTA TOSTI 2026

#### Hellas Madonna — Subtotale: 11.0 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Sportiello | 1 | 1 | MAX(1/10, 1) | **1.0** |
| Circati | 1 | 13 | 13/10 | **1.3** |
| Berisha M. | 1 | 20 | 20/10 | **2.0** |
| Moreo | 1 | 18 | 18/10 | **1.8** |
| Durosinmi ¹ | 3 | 49 | 49/10 | **4.9** |

¹ Scritto "Duronisimi" nella comunicazione

#### PARTIZAN — Subtotale: 5.6 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Belghali | 7 | 18 | 18/10 | **1.8** |
| Strefezza | 29 | 27 | (29+27)/2/10 | **2.8** |
| Przyborek | 1 | 5 | MAX(5/10, 1) | **1.0** |

#### Kung Fu Pandev — Subtotale: 15.2 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Malen | 173 | 60 | (173+60)/2/10 | **11.7** | |
| Vergara | 31 | 20 | (31+20)/2/10 | **2.5** | |
| Beukema | 1 | 10 | 10/10 | **1.0** | Rinnovo preventivo triennale |
| ~~Kouamè~~ | — | 1 | — | **NON ASSICURABILE** | Non più listato su Leghe Fantacalcio |

#### FC CKC 26 — Subtotale: 15.2 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Tiago Gabriel | 1 | 18 | 18/10 | **1.8** | |
| Vaz | 3 | 18 | 18/10 | **1.8** | |
| Muharemovic | 3 | 18 | 18/10 | **1.8** | |
| Baldanzi | 3 | 21 | 21/10 | **2.1** | Già assicurato |
| Santos A. | 1 | 13 | 13/10 | **1.3** | Comunicato come "Allison S." |
| Bijlow | 2 | 10 | 10/10 | **1.0** | |
| Bernasconi | 1 | 18 | 18/10 | **1.8** | |
| Konè I. | 1 | 36 | 36/10 | **3.6** | Già assicurato |

#### muttley superstar — Subtotale: 5.6 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Ostigard | 2 | 22 | 22/10 | **2.2** |
| Luis Henrique ³ | 17 | 13 | (17+13)/2/10 | **1.5** |
| Solomon | 21 | 16 | (21+16)/2/10 | **1.9** |

³ Scritto "Luis Enrique" nella comunicazione, nel DB è "Luis Henrique"

#### FCK Deportivo — 0.0 cr
Nessuna assicurazione richiesta (tutti già assicurati).

#### Millwall — Subtotale: 15.4 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Muric | 21 | 12 | (21+12)/2/10 | **1.6** |
| Celik | 2 | 12 | 12/10 | **1.2** |
| Ratkov | 2 | 22 | 22/10 | **2.2** |
| Zaragoza | 46 | 13 | (46+13)/2/10 | **3.0** |
| Perrone | 9 | 36 | 36/10 | **3.6** |
| Paleari | 1 | 10 | 10/10 | **1.0** |
| Boga | 1 | 12 | 12/10 | **1.2** |
| Holm | 5 | 16 | 16/10 | **1.6** |

#### Papaie Top Team — Subtotale: 1.9 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Hien | 1 | 19 | 19/10 | **1.9** |

#### Legenda Aurea — Subtotale: 44.5 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Di Gregorio | 4 | 60 | 60/10 | **6.0** |
| Sommer | 57 | 57 | 57/10 | **5.7** |
| Martinez Jo. | 22 | 24 | 24/10 | **2.4** |
| Kalulu | 57 | 49 | (57+49)/2/10 | **5.3** |
| Bartesaghi | 3 | 23 | 23/10 | **2.3** |
| Lovric | 1 | 11 | 11/10 | **1.1** |
| Taylor K. | 42 | 34 | (42+34)/2/10 | **3.8** |
| Fagioli | 1 | 24 | 24/10 | **2.4** |
| Ekkelenkamp | 4 | 24 | 24/10 | **2.4** |
| Miretti | 1 | 23 | 23/10 | **2.3** |
| Bonazzoli | 2 | 20 | 20/10 | **2.0** |
| Raspadori | 50 | 61 | 61/10 | **6.1** |
| Vitinha O. | 5 | 27 | 27/10 | **2.7** |

#### A.S. Tronzano — 0.0 cr
Nessuna comunicazione ricevuta.

### **TOTALE FT: 114.4 crediti** (44 giocatori assicurati + 1 NON assicurabile)

---

### FANTAMANTRA MANAGERIALE

#### Kung Fu Pandev — Subtotale: 14.7 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Konè I. | 41 | 41 | 41/10 | **4.1** | Già assicurato |
| Raspadori | 119 | 61 | (119+61)/2/10 | **9.0** | |
| ~~Posch~~ | — | — | — | **RESPINTO** | Svincolato da KFP, non assicurabile |
| Ferguson | 17 | 14 | (17+14)/2/10 | **1.6** | Già assicurato |
| ~~Kouamè~~ | — | 1 | — | **NON ASSICURABILE** | Non più listato su Leghe Fantacalcio |

#### FC CKC 26 — Subtotale: 16.3 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Durosinmi | 19 | 46 | 46/10 | **4.6** |
| Vergara | 71 | 20 | (71+20)/2/10 | **4.5** |
| Zaniolo | 38 | 71 | 71/10 | **7.1** |

#### H-Q-A Barcelona — Subtotale: 54.5 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Holm | 1 | 16 | 16/10 | **1.6** | Già assicurato |
| Ndicka | 38 | 25 | (38+25)/2/10 | **3.1** | Già assicurato |
| Gallo | 2 | 12 | 12/10 | **1.2** | Già assicurato |
| Vasquez | 1 | 23 | 23/10 | **2.3** | Già assicurato |
| Gudmundsson A. ⁴ | 183 | 65 | (183+65)/2/10 | **12.4** | Già assicurato |
| Frendrup ⁵ | 1 | 25 | 25/10 | **2.5** | Già assicurato |
| Britschgi | 9 | 13 | 13/10 | **1.3** | |
| Sulemana I. | 1 | 10 | 10/10 | **1.0** | Già assicurato |
| Taylor K. | 52 | 34 | (52+34)/2/10 | **4.3** | |
| Malen | 292 | 59 | (292+59)/2/10 | **17.6** | |
| Sommer | 88 | 57 | (88+57)/2/10 | **7.3** | Già assicurato |

⁴ Scritto "Gudmusson" / ⁵ Scritto "Frendup"

#### Hellas Madonna — Subtotale: 40.9 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| David | 307 | 51 | (307+51)/2/10 | **17.9** | Già assicurato |
| Cheddira | 1 | 9 | MAX(9/10, 1) | **1.0** | |
| Zaragoza | 10 | 13 | 13/10 | **1.3** | |
| Ekkelenkamp ⁷ | 6 | 24 | 24/10 | **2.4** | |
| Brescianini | 4 | 25 | 25/10 | **2.5** | |
| Belghali | 1 | 18 | 18/10 | **1.8** | |
| Scamacca | 190 | 90 | (190+90)/2/10 | **14.0** | Già assicurato |

Scritto "Davids" nella comunicazione, corretto in "David". Scritto "Ekkelekamp", corretto in "Ekkelenkamp".

#### FICA — Subtotale: 4.1 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Luis Henrique | 1 | 13 | 13/10 | **1.3** |
| Fullkrug | 1 | 28 | 28/10 | **2.8** |

#### Lino Banfield FC — Subtotale: 20.9 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Celik | 18 | 15 | (18+15)/2/10 | **1.6** | |
| Obert | 1 | 15 | 15/10 | **1.5** | |
| Marcandalli | 1 | 11 | 11/10 | **1.1** | |
| Bernasconi | 3 | 18 | 18/10 | **1.8** | |
| Bowie | 1 | 12 | 12/10 | **1.2** | |
| Caprile | 28 | 25 | (28+25)/2/10 | **2.6** | Già assicurato |
| Cambiaghi | 1 | 25 | 25/10 | **2.5** | Già assicurato |
| Vaz | 5 | 18 | 18/10 | **1.8** | |
| Baldanzi | 52 | 21 | (52+21)/2/10 | **3.6** | Già assicurato |
| Koopmeiners | 12 | 11 | (12+11)/2/10 | **1.1** | Da Minnesota (scambio 13/02) |
| ~~Tavares N.~~ | 1 | 9 | — | **RESPINTO** | Triennale non decorso (scade 14/08/2027), NON assicurabile |
| Mazzitelli | 1 | 19 | 19/10 | **1.9** | Da Minnesota (scambio 13/02) |

#### Minnesota al Max — Subtotale: 14.3 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Montipò | 1 | 2 | MAX(2/10, 1) | **1.0** | Già assicurato |
| Marianucci ⁸ | 4 | 11 | 11/10 | **1.1** | Già assicurato |
| Cataldi | 1 | 25 | 25/10 | **2.5** | Già assicurato |
| Fagioli | 29 | 24 | (29+24)/2/10 | **2.6** | Da Lino (scambio 13/02), già assicurato |
| Miller L. | 1 | 15 | 15/10 | **1.5** | Da Lino (scambio 13/02) |
| Bakola | 1 | 7 | MAX(7/10, 1) | **1.0** | |
| Adzic | 2 | 1 | MAX((2+1)/2/10, 1) | **1.0** | |
| Ratkov | 1 | 19 | 19/10 | **1.9** | |
| Bellanova | 19 | 14 | (19+14)/2/10 | **1.6** | Da Lino (scambio 13/02), già assicurato |

⁸ Potrebbe essere "Marinucci"

#### Papaie Top Team — Subtotale: 17.1 cr
| Giocatore | Spesa | FVM | Formula | Costo | Note |
|-----------|-------|-----|---------|-------|------|
| Kolasinac | 60 | 15 | (60+15)/2/10 | **3.8** | Già assicurato |
| Hien | 1 | 19 | 19/10 | **1.9** | Da Minnesota (acquisto 11/02 per 22cr), già assicurato |
| Pasalic | 50 | 19 | (50+19)/2/10 | **3.5** | Già assicurato |
| Nicolussi Caviglia | 4 | 15 | 15/10 | **1.5** | Già assicurato |
| Solomon | 28 | 18 | (28+18)/2/10 | **2.3** | |
| Vlahovic | 1 | 42 | 42/10 | **4.2** | Già assicurato |

#### Legenda Aurea — Subtotale: 25.4 cr
| Giocatore | Spesa | FVM | Formula | Costo |
|-----------|-------|-----|---------|-------|
| Nelsson | 1 | 12 | 12/10 | **1.2** |
| Dossena | 1 | 4 | MAX(4/10, 1) | **1.0** |
| Bartesaghi | 40 | 23 | (40+23)/2/10 | **3.1** |
| Gandelman | 2 | 12 | 12/10 | **1.2** |
| Barbieri | 1 | 14 | 14/10 | **1.4** |
| Leao | 1 | 161 | 161/10 | **16.1** |
| Zappa | 1 | 13 | 13/10 | **1.3** |

#### Mastri Birrai — 0.0 cr
Nessuna comunicazione ricevuta.

### **TOTALE FM: 208.1 crediti** (59 giocatori assicurati + Posch respinto + Kouamè non assicurabile + Tavares respinto)

---

## RIEPILOGO PER SQUADRA

### Fanta Tosti 2026
| Squadra | Giocatori | Costo Totale |
|---------|-----------|-------------|
| Hellas Madonna | 5 | 11.0 cr |
| PARTIZAN | 3 | 5.6 cr |
| Kung Fu Pandev | 3 (+1 non assicurabile) | 15.2 cr |
| FC CKC 26 | 8 | 15.2 cr |
| muttley superstar | 3 | 5.6 cr |
| FCK Deportivo | 0 | 0.0 cr |
| Millwall | 8 | 15.4 cr |
| Papaie Top Team | 1 | 1.9 cr |
| Legenda Aurea | 13 | 44.5 cr |
| A.S. Tronzano | 0 | 0.0 cr |
| **TOTALE** | **44** | **114.4 cr** |

### FantaMantra Manageriale
| Squadra | Giocatori | Costo Totale |
|---------|-----------|-------------|
| Kung Fu Pandev | 3 (+Posch respinto, +Kouamè non assicurabile) | 14.7 cr |
| FC CKC 26 | 3 | 16.3 cr |
| H-Q-A Barcelona | 11 | 54.5 cr |
| Hellas Madonna | 7 | 40.9 cr |
| FICA | 2 | 4.1 cr |
| Lino Banfield FC | 11 (+Tavares respinto) | 20.9 cr |
| Minnesota al Max | 9 | 14.3 cr |
| Papaie Top Team | 6 | 17.1 cr |
| Legenda Aurea | 7 | 25.4 cr |
| Mastri Birrai | 0 | 0.0 cr |
| **TOTALE** | **59** | **208.1 cr** |

---

## PUNTI RISOLTI

1. **FT CKC - Allison S. = Santos A.**: Confermato che "Allison S." nella comunicazione corrisponde a "Santos A." nel DB. FVM=13, Sp=1, costo=1.3 cr. Aggiornato nel calcolo CKC.
2. **FT CKC - Konè I.**: Confermato che si tratta di Konè I. (non Konè M.). FVM=36 nel DB FT, costo=3.6 cr (invariato).
3. **FM Hellas - David**: Chiarito che Hellas Madonna è un nome team presente in entrambe le leghe (stesso fantallenatore). David è correttamente di proprietà di Hellas Madonna in FM (Sp=307, FVM=51, costo=17.9 cr).
4. **Kouamè (FT KFP + FM KFP)**: Confermato NON assicurabile. Il giocatore è legittimamente in rosa (non svincolato), ma non è più listato su Leghe Fantacalcio. Il regolamento prevede che i calciatori non listati non siano assicurabili/rinnovabili. Rimosso dal calcolo costi.

## PUNTI RISOLTI (aggiornamento sessione 22/02/2026)

5. **Sportiello (FT Hellas)**: FVM=1 nel DB. Verificato: presente nel foglio "Tutti" del listone (P, Atalanta, FVM=1) → assicurabile, costo minimo 1.0 cr.
6. **Adzic (FM Minnesota)**: FVM=1 nel DB. Verificato: presente nel foglio "Tutti" del listone (C, Juventus, FVM=1) → assicurabile, costo minimo 1.0 cr.
7. **Tavares N. (FM Lino)**: Triennale NON decorso (acquisto 14/08/2024, scadenza 14/08/2027). RESPINTO — non assicurabile. Rimosso dai calcoli.
8. **Minnesota al Max deficit**: RISOLTO — crediti effettivi dal file Rose 17/02 sono 24 (non 9). 24 disponibili vs ~15 di costo assicurazioni → nessun deficit.
9. **Beukema (FT KFP)**: Assicurabile preventivamente (clausola regolamentare: triennio scade 12/09/2026, prima della finestra estiva 2026).

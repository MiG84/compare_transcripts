# compare_transcripts

Primerja dva transkripta (**NOV** in **TRENUTNI**) z referenčnimi podnapisi VTT
ter oba transkripta še med seboj (**TRENUTNI vs NOV**, brez reference).
Izpiše odstotke podobnosti in razlike, razvrščene po kategorijah.

## Uporaba

```bash
python compare_transcripts.py --nov nov.docx --trenutni trenutni.docx --vtt podnapisi.vtt
```

Podprti formati vhodnih datotek: `.docx`, `.txt`, `.vtt`

### Brez reference

`--vtt` lahko izpustiš — takrat se naredi samo medsebojna primerjava obeh transkriptov:

```bash
python compare_transcripts.py --nov nov.txt --trenutni trenutni.txt
```

Izpis in poročilo v tem primeru vsebujeta le razdelek **TRENUTNI vs NOV** (podobnost,
razčlenitev razlik, primeri), brez odstotkov glede na referenco in **brez zmagovalca** —
brez podnapisov ni podlage za trditev, kateri transkript je pravilnejši. Odstotki so
glede na dolžino TRENUTNI.

### Argumenti

| Argument | Opis |
|---|---|
| `--nov` | pot do NOV transkripta (obvezno) |
| `--trenutni` | pot do TRENUTNI transkripta (obvezno) |
| `--vtt` | pot do referenčnih VTT podnapisov (**neobvezno**) |
| `--primeri N` | število primerov razlik na kategorijo (privzeto 8) |
| `--izhod MAPA` | mapa za poročili `.md`/`.json` (privzeto mapa datoteke NOV) |
| `--shrani-txt` | shrani očiščeno besedilo vseh treh vhodov kot `.txt` |
| `--brez-barv` | izpis brez ANSI barv (za preusmeritev v datoteko) |

### Priporočeno: delaj na `.txt`

`.docx` je za primerjavo slaba osnova — ne vidiš, kaj je skript dejansko prebral
(odstavki, tabele, sledi sprememb). Zato enkrat poženi z `--shrani-txt`, ki zapiše
natanko tisto besedilo, ki se primerja (brez oznak govorcev, ena poved na vrstico),
naslednje zagone pa delaj na teh `.txt`. Rezultat je identičen, je pa berljiv,
diffable in ne potrebuje `python-docx`.

```bash
python compare_transcripts.py --nov Files/nov.docx --trenutni Files/trenutni.docx \
    --vtt Files/podnapisi.vtt --shrani-txt
```

## Kategorije razlik

- **samo ločila** — vejica, pika, velika začetnica
- **zapis številk** — `38 %` proti `osemintrideset odstotkov`
- **okrajšave** — `oz.`, `itn.`, `npr.` proti polnim besedam
- **vsebinske razlike** — napačno prepoznana beseda

> Stolpec _blokov_ ni primerljiv med transkripti (en blok obsega 1 ali 30 besed) —
> primerljiv je stolpec _besed_ oz. odstotek glede na referenco.

## Branje rezultatov

- **Podobnost** se izpisuje na dve decimalki, razlika med transkriptoma (»za X točk
  boljši«) pa se vedno izračuna iz teh **izpisanih** vrednosti — sicer se lahko zgodi,
  da sta oba prikazana kot `81.5 %`, razlika pa je enkrat `6.2` in drugič `6.3`.
- **Vsebinske napake v besedah** so edina številka, ki je primerljiva med različnimi
  NOV datotekami. Število _blokov_ je odvisno od tega, kako `difflib` združi sosednje
  razlike, zato lahko zelo niha, tudi če je podobnost enaka.
- Ime primerjane datoteke je v glavi poročila — pred primerjavo dveh poročil preveri,
  da gre res za isto NOV datoteko.

## Izhod

Poleg izpisa v konzoli se poleg NOV datoteke zapišeta še:

- `<nov>.json` — strojno berljivi rezultati (`nov_vs_vtt`, `trenutni_vs_vtt`, `trenutni_vs_nov`)
- `<nov>.md` — Markdown poročilo s primeri razlik in medsebojno primerjavo

Če bi izhodna datoteka povozila katero od vhodnih (npr. `--nov porocilo.md`), se ime
dopolni v `<nov>_porocilo.md`.

## Zahteve

- Python 3
- [`python-docx`](https://pypi.org/project/python-docx/) — samo za branje `.docx`:

```bash
pip install python-docx
```

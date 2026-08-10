# compare_transcripts

Primerja dva transkripta (**NOV** in **TRENUTNI**) z referenčnimi podnapisi VTT.
Izpiše odstotke podobnosti in razlike, razvrščene po kategorijah.

## Uporaba

```bash
python compare_transcripts.py --nov nov.docx --trenutni trenutni.docx --vtt podnapisi.vtt
```

Podprti formati vhodnih datotek: `.docx`, `.txt`, `.vtt`

### Argumenti

| Argument | Opis |
|---|---|
| `--nov` | pot do NOV transkripta (obvezno) |
| `--trenutni` | pot do TRENUTNI transkripta (obvezno) |
| `--vtt` | pot do referenčnih VTT podnapisov (obvezno) |
| `--primeri N` | število primerov razlik na kategorijo (privzeto 8) |
| `--brez-barv` | izpis brez ANSI barv (za preusmeritev v datoteko) |

## Kategorije razlik

- **samo ločila** — vejica, pika, velika začetnica
- **zapis številk** — `38 %` proti `osemintrideset odstotkov`
- **okrajšave** — `oz.`, `itn.`, `npr.` proti polnim besedam
- **vsebinske razlike** — napačno prepoznana beseda

> Stolpec _blokov_ ni primerljiv med transkripti (en blok obsega 1 ali 30 besed) —
> primerljiv je stolpec _besed_ oz. odstotek glede na referenco.

## Izhod

Poleg izpisa v konzoli se poleg NOV datoteke zapišeta še:

- `<nov>.json` — strojno berljivi rezultati
- `<nov>.md` — Markdown poročilo s primeri razlik

## Zahteve

- Python 3
- [`python-docx`](https://pypi.org/project/python-docx/) — samo za branje `.docx`:

```bash
pip install python-docx
```

import os
import sys
import tempfile
import zipfile
import lxml.etree as ET
from StringIO import StringIO

# Adapted from https://stackoverflow.com/questions/25738523
def update_zip(zipname, filename, cb):
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    with zipfile.ZipFile(zipname, 'r') as zin:
        with zipfile.ZipFile(tmpname, 'w') as zout:
            zout.comment = zin.comment
            for item in zin.infolist():
                d = zin.read(item.filename)
                if item.filename == filename:
                    d = cb(d)
                zout.writestr(item, d)

    os.rename(zipname, "." + zipname + "-orig")
    os.rename(tmpname, zipname)

def NS_office(tag):
    return "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}" + tag

def NS_table(tag):
    return "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}" + tag

def NS_text(tag):
    return "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}" + tag

def cell_text(elem):
    ret = ""
    for ch in elem.getchildren():
        ret = ret + ch.text
    return ret

def process_invoices(orig):
    f = StringIO(orig)
    tree = ET.parse(f)
    root = tree.getroot()
    for table in root.findall(NS_office("body") + "/" +
                              NS_office("spreadsheet") + "/" +
                              NS_table("table")):
        print table
        for row in table.findall(NS_table("table-row")):
            cells = row.findall(NS_table("table-cell"))

            # Skip trivial cases. There tends to be a million or so empty rows
            # after the main document body, don't bother enumerating them.
            if len(cells) == 1 and cells[0].getchildren() == []:
                continue

            row_repeat = int(row.get(NS_table("number-rows-repeated"), "1"))
            for _ in range(row_repeat):
                for cell in cells:
                    cell_repeat = int(cell.get(NS_table("number-columns-repeated"), "1"))
                    text = cell_text(cell)
                    if text == "":
                        continue
                    for _ in range(cell_repeat):
                        sys.stdout.write("\t" + text.encode("utf-8"))
                sys.stdout.write("\n")
    return orig

update_zip(sys.argv[1], 'content.xml', process_invoices)

# Program zpracovava hodnoty nastavene jednotlivym zakaznikum, resp. fakturam.
# Jmena nastavovanych hodnot jsou uvedena ve sloupci pod sebou. Hodnoty tykajici
# se jednotlivych faktur nebo zakazniku jsou po prave strane toho sloupce, v
# kazdem sloupci hodnota pro jednoho zakaznika. Hodnoty tykajici se vsech faktur
# jsou vlevo od toho sloupce.
#
# Hodnoty tykajici se zakazniku se uvadeji stejne. Aby program rozlisil ktere
# jsou ktere, ma dva pojmenovane vyrazy: "CustHead" pro prvni hlavicku
# zakaznickych udaju, a "InvHead" pro prvni hlavicku fakturacnich udaju.
#
# Program postupne prochazi hlavicky nastavovanych hodnot. Aby poznal, kde konci
# zaznam jedne faktury, hlida, jestli se predpis nepokusi nastavit jiz
# nastavenou hodnotu. Pokud ano, bylo zadavani udaju dokonceno, a momentalne
# nastavene hodnoty je mozno zpracovat. Podobne kdyz hlavicky skonci, je mozno
# jiz zadane hodnoty zpracovat. Po zpracovani jsou fakturni udaje smazany, a
# pokracuje se znovu.
#
# Zpracovani udaju spociva v rozhodnuti, zda je treba k dane mnozine udaju
# priradit ID faktury a vygenerovat PDF. Pokud je uvedena aspon jedne
# nezakaznicka hodnota specificky pro danou fakturu (tedy ne pro vsechny
# faktury), pak je treba zalozit fakturu. Pokud je k fakture vedena hodnota
# "ID", byla faktura jiz vygerenovana, a lze ji preskocit. Jinak se priradi ID,
# a faktura je zapamatovana jako nova.
#
# ID faktury se prirazuje pomoci pojmenovaneho vyrazu InvNextID, ktere obsahuje
# cislo pristi faktury. Po prirazeni se hodnota v policku zvysi na nasledujici
# neobsazene ID (prirazovani ID faktur je tedy treba odlozit az pote, co jsou
# vsechny faktury naparsovany).
#
# Po ukonceni zpracovani program vypise nove prirazena cisla faktur, a
# aktualizuje dokument.
#
# Ve druhem rezimu program neprovadi prirazeni cisel fakturam, a misto toho
# vypise hodnoty promennych (jak zakaznickych, tak fakturnich) pro zadane cislo
# faktury. Predpoklada se, ze tyto hodnoty prevezme jiny skript, ktery z nich
# vygeneruje fakturu.

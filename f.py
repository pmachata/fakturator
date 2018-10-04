import os
import string
import sys
import tempfile
import zipfile
import lxml.etree as ET
from StringIO import StringIO
import ab

# Adapted from https://stackoverflow.com/questions/25738523
def update_zip(zipname, filename, cb):
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    try:
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

    except:
        os.unlink(tmpname)
        raise

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

class Row:
    def __init__(self, root):
        cells = []
        for cell_el in root.findall(NS_table("table-cell")):
            cell_repeat = int(cell_el.get(NS_table("number-columns-repeated"), "1"))
            cells += [cell_el] * cell_repeat

        self.el = root
        self.cells = cells

    def __iter__(self):
        return iter(self.cells)

class Ref:
    def parse_tablename_1(self, s):
        quoted = False
        escape = False
        name = ""
        for i, c in enumerate(s):
            if c == "'":
                if quoted:
                    if escape:
                        name = name + "'"
                        escape = False
                    else:
                        escape = True
                else:
                    quoted = True

            elif escape:
                # Non-quote after a quote: we are done and this ought to be a
                # table name separator.
                assert c == '.'
                return name, s[i+1:]

            elif not quoted and c == '.':
                return name, s[i+1:]

            else:
                name = name + c

        if quoted:
            # We never found the ending quote!
            raise ValueError(u"Invalid reference: " + unicode(s))

        # xxx is this a column coordinate?
        assert False

    def parse_tablename(self, s):
        assert s != ""
        fixed = s[0] == '$'
        if fixed:
            s = s[1:]
        return fixed, self.parse_tablename_1(s)

    def parse_dollar(self, s):
        assert s != ""
        if s[0] == '$':
            return True, s[1:]
        else:
            return False, s

    def parse_col(self, s):
        fixed, s = self.parse_dollar(s)
        ss = s.lstrip(string.ascii_letters)
        return fixed, (s[:len(s) - len(ss)], ss)

    def parse_row(self, s):
        fixed, s = self.parse_dollar(s)
        ss = s.lstrip(string.digits)
        return fixed, (s[:len(s) - len(ss)], ss)

    def escape_tn(self, tn):
        if '.' in tn or "'" in tn:
            return "'%s'" % tn.replace("'", "''")
        else:
            return tn

    def __init__(self, s):
        assert s != ""
        self.tn_fixed, (self.tn, s) = self.parse_tablename(s)
        self.col_fixed, (col, s) = self.parse_col(s)
        self.row_fixed, (row, s) = self.parse_row(s)

        if col == "":
            self.col = None
        else:
            self.col = ab.abn(col)

        if row == "":
            self.row = None
        else:
            self.row = int(row)

    def __str__(self):
        return "".join(["$" if self.tn_fixed else "",
                        (self.escape_tn(self.tn) + ".") \
                            if self.tn is not None else "",
                        "$" if self.col_fixed else "",
                        ab.nab(self.col) if self.col is not None else "",
                        "$" if self.tn_fixed else "",
                        str(self.row) if self.row is not None else ""])

print Ref("Aa4.Aa4")
raise SystemExit()

class Table:
    def __init__(self, root):
        rows = []
        for row_el in root.findall(NS_table("table-row")):
            cell_els = row_el.findall(NS_table("table-cell"))

            # Skip trivial cases. There tends to be a million or so empty rows
            # after the main document body, don't bother enumerating them.
            if len(cell_els) == 1 and cell_els[0].getchildren() == []:
                continue

            row_repeat = int(row_el.get(NS_table("number-rows-repeated"), "1"))
            rows += [Row(row_el)] * row_repeat

        named = {}
        for named_el in root.findall(NS_table("named-expressions") + "/" +
                                     NS_table("named-range")):
            name = named_el.get(NS_table("name"))
            addr = named_el.get(NS_table("cell-range-address"))
            print name, addr, Ref(addr)

        self.el = root
        self.rows = rows
        self.name = root.get(NS_table("name"))

    def __iter__(self):
        return iter(self.rows)

def process_invoices(orig):
    f = StringIO(orig)
    tree = ET.parse(f)
    root = tree.getroot()
    for el in root.findall(NS_office("body") + "/" +
                           NS_office("spreadsheet") + "/" +
                           NS_table("table")):
        table = Table(el)
        for row in table:
            for cell in row:
                text = cell_text(cell)
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

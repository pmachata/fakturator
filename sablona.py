# -*- coding: utf-8 -*-
import string

class InvoiceTemplate(string.Template):
    pattern = r"""
        \#\#\#(?:
            (?P<named>[_a-z][_a-z0-9]*)  | # Unbraced identifiers
            {(?P<braced>[^}]*)}          | # Arbitrary braced strings
            (?P<invalid>)                  # Other ill-formed delimiter exprs
        )\#\#\#"""

s = file("sablona.tex").read()
t = InvoiceTemplate(s)
print t.substitute({"ID": 123,
                    "Období": "Říjen 2018",
                    "Odběratel": "Roman Vonášek",
                    "Adresa": "Bublaninová 123/45, Praha 5 - Chlebodary"})

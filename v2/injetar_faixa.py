# -*- coding: utf-8 -*-
"""
Injeta a faixa de opcoes (ribbon) do IMTSolver num arquivo .xlsm ou .xlam.

Um arquivo do Office e um zip. A faixa personalizada e uma parte dentro
desse zip (customUI/customUI14.xml) mais uma entrada de relacionamento
apontando para ela. Nenhuma biblioteca de planilha faz isso, entao
mexemos no zip diretamente.

Uso:
    python injetar_faixa.py MinhaPasta.xlsm [customUI14.xml]
"""

import os
import re
import shutil
import sys
import zipfile

PARTE = "customUI/customUI14.xml"
TIPO_REL = ("http://schemas.microsoft.com/office/2007/relationships/ui/"
            "extensibility")


def injetar(arquivo, xml_faixa):
    if not os.path.exists(arquivo):
        print(f"ERRO: nao achei {arquivo}")
        return 1
    if not zipfile.is_zipfile(arquivo):
        print(f"ERRO: {arquivo} nao parece um arquivo do Office")
        return 1

    xml = open(xml_faixa, encoding="utf-8").read()

    backup = arquivo + ".bak"
    shutil.copy2(arquivo, backup)
    print(f"backup: {backup}")

    origem = zipfile.ZipFile(arquivo, "r")
    itens = {n: origem.read(n) for n in origem.namelist()}
    origem.close()

    # 1) a parte com a definicao da faixa
    itens[PARTE] = xml.encode("utf-8")

    # 2) o relacionamento, em _rels/.rels
    rels = itens.get("_rels/.rels", b"").decode("utf-8")
    if TIPO_REL not in rels:
        ids = re.findall(r'Id="rId(\d+)"', rels)
        novo = max((int(i) for i in ids), default=0) + 1
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rId{novo}" Type="{TIPO_REL}" '
            f'Target="{PARTE}"/></Relationships>')
        itens["_rels/.rels"] = rels.encode("utf-8")
        print(f"relacionamento rId{novo} acrescentado")
    else:
        print("relacionamento ja existia")

    # 3) o tipo de conteudo, para o Office saber ler a parte
    ct = itens.get("[Content_Types].xml", b"").decode("utf-8")
    if 'PartName="/customUI/customUI14.xml"' not in ct:
        if 'Extension="xml"' not in ct:
            ct = ct.replace(
                "<Types ", '<Types ', 1)
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/customUI/customUI14.xml" '
            'ContentType="application/xml"/></Types>')
        itens["[Content_Types].xml"] = ct.encode("utf-8")
        print("tipo de conteudo declarado")

    destino = zipfile.ZipFile(arquivo, "w", zipfile.ZIP_DEFLATED)
    for nome, dados in itens.items():
        destino.writestr(nome, dados)
    destino.close()

    print(f"\nOK - faixa IMTSolver injetada em {arquivo}")
    print("Feche e abra o arquivo no Excel: a aba IMTSolver deve aparecer.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    alvo = sys.argv[1]
    faixa = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "customUI14.xml")
    sys.exit(injetar(alvo, faixa))

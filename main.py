#!/usr/bin/python
# -*- coding: UTF-8 -*-

import sys
import argparse
import docx
import json
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, ns
from docx.shared import Cm, Pt

# Инициализация входных параметров.
def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument('inputfile', type=str)
    parser.add_argument('data', type=str)
    parser.add_argument('otputfile', type=str)
    return parser

# Основное действие.
if __name__ == '__main__':
    
    # Парсер параметров.
    parser = createParser()
    arguments = parser.parse_args(sys.argv[1:])

    # Чтение документа-исходника.
    doc = docx.Document(arguments.inputfile)

    # Чтение входящих параметров.
    data = json.load(open(arguments.data, encoding='utf-8'))

    # Сохраним документ.
    doc.save(arguments.otputfile)

    doc = docx.Document(arguments.inputfile)

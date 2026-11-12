# -*- coding: utf-8 -*-
"""Реестр приложений к отчету о выполнении работ.

Каждое приложение соответствует работе календарного плана и называется так же, как эта работа (дословно).
Объекты нумеруются внутри приложения, литера ставится после номера: «Таблица 1А», «Рисунок 2Б».
Реестр общий для отчета (ссылки вида {A:ключ}) и для файла приложений, поэтому номера не расходятся.
"""
import re
import facts as F

LETTERS = "АБВГД"

# (литера, название = работа календарного плана, [(вид, ключ), …] в порядке размещения)
APPS = [
    ("А", F.PLAN[0], [("tbl", "requirements"), ("tbl", "materials"), ("tbl", "equipment"), ("tbl", "recipes"),
                      ("tbl", "cook"), ("tbl", "variants"), ("fig", "selection"), ("tbl", "final")]),
    ("Б", F.PLAN[1], [("tbl", "batches"), ("fig", "techscheme"), ("tbl", "batch2"), ("fig", "ph_process"),
                      ("fig", "ph_batch2"), ("fig", "ph_batch3"), ("fig", "ph_pack")]),
    ("В", F.PLAN[2], [("tbl", "criteria"), ("tbl", "final_check")]),
    ("Г", F.PLAN[3], [("tbl", "site_approaches"), ("fig", "hosting"), ("tbl", "site_timeline"),
                      ("fig", "ph_site_main"), ("fig", "ph_site_palette"), ("fig", "ph_site_form"),
                      ("fig", "ph_site_metrika")]),
    ("Д", F.PLAN[4], [("fig", "campaign"), ("tbl", "syn_timeline"), ("fig", "pipeline"), ("tbl", "radar"),
                      ("fig", "ph_syn_card"), ("fig", "ph_direct"), ("fig", "ph_syn_metrika"),
                      ("fig", "ph_billboard1"), ("fig", "ph_billboard2"), ("fig", "ph_release")]),
]

LABELS = {}   # ключ -> «1А»
KINDS = {}    # ключ -> "tbl" | "fig"
for letter, _title, objs in APPS:
    n = {"tbl": 0, "fig": 0}
    for kind, key in objs:
        n[kind] += 1
        assert key not in LABELS, key
        LABELS[key] = f"{n[kind]}{letter}"
        KINDS[key] = kind


def label(key):
    if key not in LABELS:
        raise KeyError(f"нет объекта приложения «{key}»")
    return LABELS[key]


def resolve(text):
    """Подставляет номера объектов приложений вместо {A:ключ}."""
    return re.sub(r"\{A:([a-z0-9_]+)\}", lambda m: label(m.group(1)), text)


def all_labels():
    """[(«Таблица 1А» | «Рисунок 1А»), …] — все объекты приложений."""
    return [("Таблица " if KINDS[k] == "tbl" else "Рисунок ") + v for k, v in LABELS.items()]

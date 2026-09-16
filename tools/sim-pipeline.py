# -*- coding: utf-8 -*-
"""Точная модель пайплайна Query-ImportJSON (v5) на Python.
Воспроизводит семантику M-кода, а не переписывает его по-своему."""
import json, datetime as dt
from collections import defaultdict

def load(path):
    return json.load(open(path, encoding='utf-8'))

# --- Query-ImportJSON: SampleFields = объединение полей первых 1000 записей
def expand(records):
    fields = []
    for r in records[:1000]:
        for k in r.keys():
            if k not in fields:
                fields.append(k)
    return [{f: r.get(f, None) for f in fields} for r in records], fields

# --- fnNormalizeFields
def to_datetime(v):
    """ToDateTime: en-US, затем ru-RU, затем null"""
    if v is None or v == "":
        return None
    if isinstance(v, str):
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%m/%d/%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S",
                    "%Y-%m-%d %H:%M:%S", "%d.%m.%Y"):
            try:
                return dt.datetime.strptime(v, fmt)
            except ValueError:
                continue
        return None
    return None

def post_normalize(p):
    s = p if p is not None else ""
    low = s.lower()
    if "стк" in low:
        return "СТК"
    if "прк" in low:
        return "ПРК"
    return s

def normalize_fields(rows, fields):
    has_post = "post" in fields
    out = []
    for r in rows:
        n = dict(r)
        for c in ("date", "status_date"):
            if c in fields:
                n[c] = to_datetime(r.get(c))
        n["postN"] = post_normalize(r.get("post") if has_post else None)
        try:
            y = int(r.get("year_status") or 0)
        except (TypeError, ValueError):
            y = 0
        try:
            w = int(r.get("week_status") or 0)
        except (TypeError, ValueError):
            w = 0
        n["yearWeek"] = y * 100 + w
        out.append(n)
    return out

# --- fnComputeKey
def compute_key(row):
    def F(name):
        v = row.get(name, None)
        return "" if v is None else str(v)
    d = row.get("date", None)
    dtx = "" if d is None else d.strftime("%Y-%m-%dT%H:%M:%S")
    return f'{F("number")}|{dtx}|{F("ready_for")}|{F("direction")}'

# --- fnComputeGroupMetrics
def norm_status(v):
    return "" if v is None else str(v).replace("ё", "е").lower()

def is_start(v):
    return norm_status(v).startswith("готов к приемке")

def is_end(v):
    return norm_status(v).startswith("готов к выбытию")

def group_metrics(rows):
    agg = defaultdict(lambda: {"start": [], "end": []})
    for r in rows:
        k = (r.get("number"), r.get("direction"))
        sd = r.get("status_date")
        if sd is not None:
            if is_start(r.get("ready_for")):
                agg[k]["start"].append(sd)
            if is_end(r.get("ready_for")):
                agg[k]["end"].append(sd)
    delta = {}
    for k, v in agg.items():
        ts = min(v["start"]) if v["start"] else None
        te = max(v["end"]) if v["end"] else None
        delta[k] = None if (ts is None or te is None) else (te - ts).total_seconds() / 3600.0
    out = []
    for r in rows:
        n = dict(r)
        k = (r.get("number"), r.get("direction"))
        n["deltaHours"] = delta.get(k) if is_end(r.get("ready_for")) else None
        out.append(n)
    return out

# --- fnUpsert
def upsert(new_rows, existing_rows):
    if not existing_rows:
        return list(new_rows)
    new_keys = {r["Key"] for r in new_rows}
    kept = [r for r in existing_rows if r["Key"] not in new_keys]
    return list(new_rows) + kept

def pipeline(path, existing=None):
    recs = load(path)
    rows, fields = expand(recs)
    rows = normalize_fields(rows, fields)
    for r in rows:
        r["Key"] = compute_key(r)
    rows = group_metrics(rows)
    return upsert(rows, existing or [])

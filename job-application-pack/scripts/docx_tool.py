#!/usr/bin/env python3
"""Правка резюме в .docx без потери оформления.

Команды:
  dump  <in.docx>                       — список абзацев с id, стилем и текстом
  apply <in.docx> <plan.json> <out.docx> — применить план правок
  check <out.docx>                      — найти плейсхолдеры, посчитать страницы

id абзацев:
  p12        — абзац №12 в теле документа
  t0.0.1.2   — таблица 0, строка 0, ячейка 1, абзац 2 (шапка с фото и контактами)

В тексте ** выделяет жирный фрагмент: "**Инструменты:** Notion • Figma".
Жирные фрагменты получают формат первого жирного прогона абзаца,
обычные — формат первого обычного. Так сохраняются шрифт, кегль и цвет.

Формат plan.json — список операций, id берутся из dump исходного файла:
  {"op": "set",          "id": "p15", "text": "— Новый текст буллета"}
  {"op": "delete",       "id": "p21"}
  {"op": "insert_after", "id": "p16", "like": "p15", "text": "— ..."}
      (копирует оформление абзаца like и вставляет после id;
       несколько вставок после одного id идут в порядке плана)
"""
import copy
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import docx
from docx.oxml.ns import qn

W_P, W_TBL, W_R, W_HL = qn("w:p"), qn("w:tbl"), qn("w:r"), qn("w:hyperlink")


def index_paragraphs(doc):
    """Вернуть упорядоченный dict id -> элемент <w:p>."""
    out, t_i, p_i = {}, 0, 0
    for el in doc.element.body.iterchildren():
        if el.tag == W_P:
            out[f"p{p_i}"] = el
            p_i += 1
        elif el.tag == W_TBL:
            for r_i, tr in enumerate(el.iterchildren(qn("w:tr"))):
                for c_i, tc in enumerate(tr.iterchildren(qn("w:tc"))):
                    for cp_i, p in enumerate(tc.iterchildren(W_P)):
                        out[f"t{t_i}.{r_i}.{c_i}.{cp_i}"] = p
            t_i += 1
    return out


def p_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def run_is_bold(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return False
    b = rpr.find(qn("w:b"))
    return b is not None and b.get(qn("w:val")) not in ("0", "false")


def all_runs(p):
    """Все прогоны абзаца, включая вложенные в sdt/smartTag (экспорт из Google Docs)."""
    return [r for r in p.iter(W_R)
            if not any(a.tag == W_HL for a in r.iterancestors())]


def richest(runs):
    """Прогон с самым полным форматированием (явный шрифт и кегль) — образец для новых."""
    def score(r):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            return 0
        return (rpr.find(qn("w:rFonts")) is not None) * 2 + (rpr.find(qn("w:sz")) is not None)
    return max(runs, key=score) if runs else None


def marked_text(p):
    """Текст абзаца с ** вокруг жирных фрагментов."""
    parts = []
    for r in all_runs(p):
        t = "".join(x.text or "" for x in r.iter(qn("w:t")))
        if not t:
            continue
        b = run_is_bold(r)
        if parts and parts[-1][0] == b:
            parts[-1][1] += t
        else:
            parts.append([b, t])
    return "".join(f"**{t}**" if b else t for b, t in parts)


def style_name(doc, p):
    ppr = p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
        sid = ppr.find(qn("w:pStyle")).get(qn("w:val"))
        try:
            return doc.styles.get_by_id(sid, 1).name if hasattr(doc.styles, "get_by_id") else sid
        except Exception:
            return sid
    return "normal"


def flags(p):
    f = []
    if p.find(".//" + W_HL) is not None:
        f.append("LINK")
    if p.find(".//" + qn("w:drawing")) is not None or p.find(".//" + qn("w:pict")) is not None:
        f.append("IMAGE")
    return f


def cmd_dump(path):
    doc = docx.Document(path)
    for pid, p in index_paragraphs(doc).items():
        fl = flags(p)
        txt = marked_text(p) if "LINK" not in fl else p_text(p)
        print(f"{pid}\t[{style_name(doc, p)}]{' ' + ','.join(fl) if fl else ''}\t{txt}")


def make_run(template_r, text):
    r = copy.deepcopy(template_r)
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    t = r.makeelement(qn("w:t"), {})
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def set_text(p, text):
    if p.find(".//" + W_HL) is not None:
        raise SystemExit("Абзац со ссылкой (LINK) менять нельзя: так потеряется ссылка. Оставьте его как есть.")
    if p.find(".//" + qn("w:drawing")) is not None:
        raise SystemExit("Абзац с изображением (IMAGE) менять нельзя.")
    runs = all_runs(p)
    text_runs = [r for r in runs if r.find(qn("w:t")) is not None] or runs
    bold_tpl = richest([r for r in text_runs if run_is_bold(r)])
    norm_tpl = richest([r for r in text_runs if not run_is_bold(r)])
    base = norm_tpl if norm_tpl is not None else bold_tpl
    if base is None:  # пустой абзац без прогонов
        base = p.makeelement(W_R, {})
    base = copy.deepcopy(base)
    bold_tpl = copy.deepcopy(bold_tpl) if bold_tpl is not None else None
    norm_tpl = copy.deepcopy(norm_tpl) if norm_tpl is not None else None
    keep = (qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"))
    for child in list(p):
        if child.tag not in keep:
            p.remove(child)
    segments = re.split(r"(\*\*.+?\*\*)", text)
    for seg in segments:
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**"):
            tpl = bold_tpl
            seg = seg[2:-2]
            if tpl is None:
                tpl = copy.deepcopy(base)
                rpr = tpl.find(qn("w:rPr"))
                if rpr is None:
                    rpr = tpl.makeelement(qn("w:rPr"), {})
                    tpl.insert(0, rpr)
                rpr.append(rpr.makeelement(qn("w:b"), {}))
        else:
            tpl = norm_tpl
            if tpl is None:
                tpl = copy.deepcopy(base)
                rpr = tpl.find(qn("w:rPr"))
                if rpr is not None:
                    for b in rpr.findall(qn("w:b")) + rpr.findall(qn("w:bCs")):
                        rpr.remove(b)
        p.append(make_run(tpl, seg))


def cmd_apply(src, plan_path, dst):
    doc = docx.Document(src)
    idx = index_paragraphs(doc)
    plan = json.loads(Path(plan_path).read_text(encoding="utf-8"))
    for op in plan:
        if op["id"] not in idx:
            raise SystemExit(f"Нет абзаца {op['id']}: {op}")
    last_inserted = {}
    deletes = []
    for op in plan:
        p = idx[op["id"]]
        kind = op["op"]
        if kind == "set":
            set_text(p, op["text"])
        elif kind == "delete":
            deletes.append(p)
        elif kind == "insert_after":
            like = idx[op.get("like", op["id"])]
            new_p = copy.deepcopy(like)
            for hl in new_p.findall(".//" + W_HL):
                new_p.remove(hl)
            anchor = last_inserted.get(op["id"], p)
            anchor.addnext(new_p)
            last_inserted[op["id"]] = new_p
            set_text(new_p, op["text"])
        else:
            raise SystemExit(f"Неизвестная операция: {kind}")
    for p in deletes:
        parent = p.getparent()
        # в ячейке таблицы должен остаться хотя бы один абзац
        if parent.tag == qn("w:tc") and len(parent.findall(W_P)) == 1:
            for r in list(p.iterchildren(W_R)):
                p.remove(r)
        else:
            parent.remove(p)
    doc.save(dst)
    print(f"Сохранено: {dst}")


PLACEHOLDERS = [
    "Имя Фамилия", "Результат работы", "Компания 1", "Компания 2", "Компания 3",
    "Описание компании или проекта", "месяц год", "Название вуза", "Релевантный для должности курс",
    "Язык (уровень", "почта • телеграм", "liza-career.ru", "Полный гайд по составлению резюме",
    "[", "TODO", "XXX",
]


def cmd_check(path):
    doc = docx.Document(path)
    problems = []
    empty_run = 0
    for pid, p in index_paragraphs(doc).items():
        t = p_text(p)
        for ph in PLACEHOLDERS:
            if ph in t:
                problems.append(f"{pid}: плейсхолдер «{ph}» — {t[:80]}")
        if t.strip() in ("Фото", "Должность"):
            problems.append(f"{pid}: плейсхолдер «{t.strip()}»")
        empty_run = empty_run + 1 if not t.strip() and not flags(p) else 0
        if empty_run == 3:
            problems.append(f"{pid}: три пустых абзаца подряд — вероятно, остаток удалённого блока")
    print("\n".join(problems) if problems else "Плейсхолдеров не найдено.")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice не найден — число страниц не проверено.")
        return
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, path],
                       capture_output=True, timeout=180)
        pdfs = list(Path(td).glob("*.pdf"))
        if not pdfs:
            print("Не удалось сконвертировать в PDF.")
            return
        data = pdfs[0].read_bytes()
        pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
        print(f"Страниц: {pages}")
        keep = Path(path).with_suffix(".preview.pdf")
        shutil.copy(pdfs[0], keep)
        print(f"PDF для визуальной проверки: {keep}")


if __name__ == "__main__":
    a = sys.argv[1:]
    if not a or a[0] not in ("dump", "apply", "check"):
        print(__doc__)
        sys.exit(1)
    {"dump": lambda: cmd_dump(a[1]),
     "apply": lambda: cmd_apply(a[1], a[2], a[3]),
     "check": lambda: cmd_check(a[1])}[a[0]]()

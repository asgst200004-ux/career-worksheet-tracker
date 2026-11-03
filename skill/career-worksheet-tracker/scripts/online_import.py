"""구글시트(온라인 제출) 내보내기 xlsx → 반별 누적 엑셀에 이어 붙이기.

사용(모두 python3 -I 로 실행):
  1) 추출: python3 -I online_import.py extract <구글시트.xlsx> <작업폴더> <반=반파일.xlsx> [<반=반파일.xlsx> ...]
       → <작업폴더>/new.json (반별로 '학생 시트에 아직 없는' (학번, 차시)의 최신 제출만)
  2) 읽기: python3 -I online_import.py view <작업폴더> <반>
       → <작업폴더>/v<반>.txt (판단을 쓰기 위해 읽는 요약본)
  3) 판단 파일 작성: <작업폴더>/j<반>.py  (아래 형식, Claude가 v<반>.txt를 읽고 직접 씀)
       J = {("11201", 4): ("상", "특이사항", "비고", "간단기록 구절"), ...}   # 기록정도는 None 가능
       TOPIC_FIX = {"11201": "바뀐 주제"}   # 선택: 확인표 주제 칸 갱신
       TRACK_FIX = {"11201": "주제탐구"}    # 선택: 확인표 트랙 칸 갱신
       def post(wb, entries, J, data, label): ...  # 선택: 요약 시트(차시 목록 등) 후처리
  4) 반영: python3 -I online_import.py build <작업폴더> <반> <반파일.xlsx> <저장할.xlsx>

'접속코드' 등 학생 비밀번호가 든 시트는 읽지 않는다(이 스크립트는 '제출기록' 시트만 연다).
"""
import collections, importlib.util, json, os, re, shutil, sys

H = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, H)
import openpyxl
from excel_helpers import append_row, find_student_sheet, read_progress, update_progress, write_progress_sheet
from online_fmt import summarize, teacher_line

SHEET = "제출기록"
PAPER_SCAN = ("[종이 활동지 제출] 이번 차시는 종이 활동지로 내서 구글시트에는 스캔본 1장만 연결되어 있음.\n"
              "· 글 내용이 시트에 없어 이번 정리에는 응답을 옮기지 못함 — 스캔본을 주시면 이어서 정리")


def _int(v):
    return int(str(v).split(".")[0])


def extract(export, work, pairs):
    ws = openpyxl.load_workbook(export, read_only=True)[SHEET]
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h or "") for h in next(rows)]
    col = lambda key: next(i for i, h in enumerate(hdr) if h.startswith(key))
    c = {k: col(k) for k in ["최종 제출 시각", "학번", "이름", "반", "트랙", "차시", "제출 내용", "교사 확인",
                             "교사 메모", "원본 데이터", "교사용 참고"]}
    recs = collections.defaultdict(list)
    for r in rows:
        if not r[c["학번"]]:
            continue
        sid, ban, ch = str(r[c["학번"]]).split(".")[0], _int(r[c["반"]]), _int(r[c["차시"]])
        text = r[c["제출 내용"]] or ""
        raw = r[c["원본 데이터"]]
        recs[(ban, sid, ch)].append(dict(
            time=str(r[c["최종 제출 시각"]]), sid=sid, name=r[c["이름"]], ban=ban, track=r[c["트랙"]], ch=ch,
            text=text, status=r[c["교사 확인"]], memo=r[c["교사 메모"]], paper="(종이 활동지" in text,
            note=r[c["교사용 참고"]], raw=(json.loads(raw) if raw else {})))
    out = {}
    for ban, fn in pairs:
        wb = openpyxl.load_workbook(fn, read_only=True)  # 한셀 저장 파일도 읽힘
        have = {}
        for n in wb.sheetnames:
            if n[:3].isdigit() and " " in n:
                have[n.split(" ")[0]] = {re.sub(r"\D", "", str(r[0]).split("(")[0])
                                         for r in wb[n].iter_rows(min_row=5, values_only=True)
                                         if r and r[0] and str(r[0]).upper().startswith("P")}
        new = []
        for (b, sid, ch), lst in recs.items():
            if b != ban:
                continue
            lst.sort(key=lambda x: x["time"])
            x = lst[-1]
            x["n_versions"] = len(lst)
            if str(ch) in have.get(sid, set()):
                continue  # 이미 정리된 차시
            x["in_file"] = sid in have
            new.append(x)
        new.sort(key=lambda x: (x["ch"], x["sid"]))
        out[str(ban)] = new
        print(f"{ban}반: 새 제출 {len(new)}건", dict(collections.Counter(x["ch"] for x in new)),
              "| 시트 없는 학번:", [x["sid"] for x in new if not x["in_file"]])
    os.makedirs(work, exist_ok=True)
    json.dump(out, open(os.path.join(work, "new.json"), "w"), ensure_ascii=False, indent=1)


def view(work, ban):
    d = json.load(open(os.path.join(work, "new.json")))
    lines = []
    for x in d[str(ban)]:
        _t, _d, s, _m = summarize(x)
        lines.append(f"### {x['sid']} {x['name']} | {x['track']} | {x['ch']}차시 | {teacher_line(x)} | 버전 {x['n_versions']}")
        lines.append(s)
        if x.get("note"):
            lines.append("NOTE: " + str(x["note"]))
        lines.append("")
    p = os.path.join(work, f"v{ban}.txt")
    open(p, "w").write("\n".join(lines))
    print(p, len(d[str(ban)]), "건")


def build(work, ban, src, out):
    spec = importlib.util.spec_from_file_location("j", os.path.join(work, f"j{ban}.py"))
    jm = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(jm)
    J = jm.J
    entries = json.load(open(os.path.join(work, "new.json")))[str(ban)]
    miss = [(x["sid"], x["ch"]) for x in entries if (x["sid"], x["ch"]) not in J]
    assert not miss, f"판단 누락: {miss}"
    if os.path.abspath(src) != os.path.abspath(out):
        shutil.copy(src, out)
    wb = openpyxl.load_workbook(out)
    prog = wb["프로젝트 진행 확인표"]
    label = "p" if str(prog.cell(4, 5).value).startswith("p") else "P"
    summary_name = wb.sheetnames[-1]
    n_students = sum(1 for n in wb.sheetnames if n[:3].isdigit() and " " in n)

    # 1) 학생 시트에 한 행씩
    for x in entries:
        ws = find_student_sheet(wb, x["sid"])
        _lv, note, remark, _ = J[(x["sid"], x["ch"])]
        if x["raw"].get("_paper") and not x["raw"].get("summary"):
            title, desc, resp = f"진로 프로젝트 {x['ch']}차시(종이 활동지)", "-", PAPER_SCAN
        else:
            title, desc, resp, _m = summarize(x)
            if not x["paper"]:
                resp = f"(온라인 제출 {x['time'][:10]})\n" + resp
        rm = " / ".join(t for t in [remark, teacher_line(x)] if t and t != "없음")
        n = resp.count("\n") + 1
        append_row(ws, {"차시": f"{label}-{x['ch']}", "활동명": title, "개요": desc, "응답요약": resp,
                        "특이사항": note, "비고": rm},
                   row_height=max(110, min(409, 24 + n * 30 + len(resp) // 3)))

    # 2) 진행 확인표
    data = read_progress(wb)
    for x in entries:
        lv, _n, _r, phrase = J[(x["sid"], x["ch"])]
        update_progress(data, x["sid"], sess_no=x["ch"], submitted="○", level=lv, note_add=phrase,
                        track=getattr(jm, "TRACK_FIX", {}).get(x["sid"]),
                        topic=getattr(jm, "TOPIC_FIX", {}).get(x["sid"]))
    done = []
    for ch in range(1, 9):  # 반 절반 이상 낸 차시 = 진행된 차시 → 안 낸 학생 ×
        subs = [s for s, rec in data.items() if rec["sess"].get(ch, (None,))[0] == "○"]
        if len(subs) >= n_students / 2:
            done.append(ch)
            for s, rec in data.items():
                if ch not in rec["sess"]:
                    update_progress(data, s, sess_no=ch, submitted="×", note_add=f"{label}-{ch} 미제출")
    if label == "p":
        for rec in data.values():
            rec["note"] = rec["note"].replace("P-", "p-")
    write_progress_sheet(wb, data, prog["A1"].value, label=label, summary_sheet_name=summary_name)

    # 3) 요약 시트 후처리
    if hasattr(jm, "post"):
        jm.post(wb, entries, J, data, label)
    wb.save(out)
    print("saved", out, "| 진행된 차시:", done)


if __name__ == "__main__":
    cmd, a = sys.argv[1], sys.argv[2:]
    if cmd == "extract":
        extract(a[0], a[1], [(int(p.split("=", 1)[0]), p.split("=", 1)[1]) for p in a[2:]])
    elif cmd == "view":
        view(a[0], a[1])
    elif cmd == "build":
        build(a[0], a[1], a[2], a[3])
    else:
        sys.exit(__doc__)

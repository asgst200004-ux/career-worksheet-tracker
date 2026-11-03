"""온라인 제출 원문(raw JSON) → 학생 시트용 응답 요약 / 활동명 / 개요."""
import json, re

DAYS = "월화수목금토일"


def cut(s, n=220):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def q(s, n=220):
    s = cut(s, n)
    return f"'{s}'" if s else "(미기재)"


TITLES = {
    ("주제탐구", 3): ("진로 프로젝트 3차시(주제탐구): 자료 수집",
                   "탐구 질문에 필요한 자료를 5~10개 조사, 출처 기록, 핵심 내용 요약, 나의 생각/메모"),
    ("습관형성", 3): ("진로 프로젝트 3차시(습관형성): 습관 실천 — 1주차 습관 기록표",
                   "이번 주 목표·인증 방법, 요일별 실천·인증·오늘의 한 줄, 실천 일수와 소감"),
    ("주제탐구", 4): ("진로 프로젝트 4차시(주제탐구): 심화 탐구(인터뷰/설문) 또는 추가 자료 수집",
                   "인터뷰·설문 대상과 질문 목록을 정해 답변을 요약하거나, 추가 자료를 수집하고 새롭게 알게 된 점 정리"),
    ("습관형성", 4): ("진로 프로젝트 4차시(습관형성): 습관 실천 — 2주차 습관 기록표",
                   "이번 주 목표(수정 시 수정 목표)·인증 방법, 요일별 실천·인증·오늘의 한 줄, 실천 일수와 소감"),
    ("진로독서", 4): ("진로 프로젝트 4차시(진로독서): 독서기록장 ① 인상 깊은 구절",
                   "인상 깊은 구절 3개(쪽수 포함)를 옮겨 적고 그 이유 쓰기"),
    ("주제탐구", 5): ("진로 프로젝트 5차시(주제탐구): 자료 분석 및 정리(마인드맵)",
                   "중심 주제와 가지(하위 질문)별 핵심 내용을 마인드맵으로 정리하고, 탐구의 핵심 메시지 도출"),
    ("습관형성", 5): ("진로 프로젝트 5차시(습관형성): 습관 실천 — 3주차 습관 기록표",
                   "이번 주 목표·인증 방법, 요일별 실천·인증·오늘의 한 줄, 실천 일수와 소감"),
    ("진로독서", 5): ("진로 프로젝트 5차시(진로독서): 독서기록장 ② 관점 확장 & 진로 연결",
                   "읽기 전과 달라진 생각(관점 확장)과 책 내용을 나의 진로와 연결하기"),
    ("주제탐구", 6): ("진로 프로젝트 6차시(주제탐구): 주제탐구보고서 작성",
                   "서론·본론(1~3)·결론·참고문헌을 갖춘 주제탐구보고서 작성"),
    ("습관형성", 6): ("진로 프로젝트 6차시(습관형성): 습관 실천 — 4주차 습관 기록표",
                   "이번 주 목표·인증 방법, 요일별 실천·인증·오늘의 한 줄, 실천 일수와 소감"),
}


def habit(k, week):
    days = []
    ok = 0
    for i in range(7):
        o = (k.get(f"d_{i}_ok") or "").strip()
        do, line = cut(k.get(f"d_{i}_do"), 70), cut(k.get(f"d_{i}_line"), 60)
        if o == "○":
            ok += 1
        if o or do or line:
            s = f"{DAYS[i]} {o or '-'} {do}".rstrip()
            if line:
                s += f" ('{line}')"
            days.append(s)
    out = [f"[습관형성] {week}주차 습관 기록표",
           f"· 이번 주 목표: {cut(k.get('goal'), 120) or '미기재'} / 확인(인증) 방법: {cut(k.get('cert'), 80) or '미기재'}",
           f"· 실천 ○ {ok}일 / 7일"]
    if days:
        out.append("· 요일별: " + " / ".join(days))
    out.append(f"· 느낀 점: {q(k.get('feel'), 260)}")
    return "\n".join(out), ok


def summarize(x):
    """x: new.json 항목 → (차시, 활동명, 개요, 응답요약, 메타)"""
    tr, ch = x["track"], x["ch"]
    raw = x.get("raw") or {}
    meta = {}
    if raw.get("_paper"):
        title = raw.get("title") or f"진로 프로젝트 {ch}차시"
        return title, "(종이 활동지 요약본)", raw.get("summary", ""), meta
    title, desc = TITLES.get((tr, ch), (f"진로 프로젝트 {ch}차시({tr})", ""))
    k = raw
    if tr == "습관형성" and ch >= 3:
        s, ok = habit(k, ch - 2)
        meta["ok"] = ok
    elif tr == "주제탐구" and ch == 3:
        lines = ["[주제탐구] 자료 수집"]
        n = 0
        for i in range(1, 6):
            if k.get(f"s{i}") or k.get(f"m{i}"):
                n += 1
                lines.append(f"· {'①②③④⑤'[i-1]} {cut(k.get(f's{i}'), 90) or '(출처 미기재)'}: {cut(k.get(f'm{i}'), 200) or '(요약 미기재)'}")
        lines.append(f"· 메모: {q(k.get('memo'))}")
        s = "\n".join(lines); meta["n"] = n
    elif tr == "주제탐구" and ch == 4:
        s = "\n".join(["[주제탐구] 심화 탐구 / 추가 자료",
                       f"· 대상(인터뷰·설문): {cut(k.get('who'), 100) or '미기재'}",
                       f"· 질문 목록·수집 자료: {cut(k.get('qs'), 600) or '미기재'}",
                       f"· 답변·자료 요약: {cut(k.get('ans'), 500) or '미기재'}",
                       f"· 새롭게 알게 된 점: {q(k.get('new'), 300)}"])
    elif tr == "주제탐구" and ch == 5:
        s = "\n".join(["[주제탐구] 마인드맵",
                       f"· 중심 주제: {cut(k.get('center'), 120) or '미기재'}",
                       f"· 가지 ①: {cut(k.get('b1'), 260) or '미기재'}",
                       f"· 가지 ②: {cut(k.get('b2'), 260) or '미기재'}",
                       f"· 가지 ③: {cut(k.get('b3'), 260) or '미기재'}",
                       f"· 핵심 메시지: {q(k.get('msg'), 260)}"])
    elif tr == "주제탐구" and ch == 6:
        s = "\n".join(["[주제탐구] 주제탐구보고서",
                       f"· 서론: {cut(k.get('intro'), 260) or '미기재'}",
                       f"· 본론 1: {cut(k.get('body1'), 260) or '미기재'}",
                       f"· 본론 2: {cut(k.get('body2'), 260) or '미기재'}",
                       f"· 본론 3: {cut(k.get('body3'), 260) or '미기재'}",
                       f"· 결론: {cut(k.get('concl'), 260) or '미기재'}",
                       f"· 참고문헌: {cut(k.get('refs'), 200) or '미기재'}"])
    elif tr == "진로독서" and ch == 4:
        lines = ["[진로독서] 인상 깊은 구절"]
        for i in range(1, 4):
            if k.get(f"p{i}") or k.get(f"r{i}"):
                lines.append(f"· 구절 {'①②③'[i-1]}: {q(k.get(f'p{i}'), 160)} → 이유: {cut(k.get(f'r{i}'), 200) or '미기재'}")
        s = "\n".join(lines)
    elif tr == "진로독서" and ch == 5:
        s = "\n".join(["[진로독서] 관점 확장 & 진로 연결",
                       f"· 달라진 생각: {q(k.get('diff'), 400)}",
                       f"· 진로 연결: {q(k.get('link'), 400)}"])
    elif tr == "진로독서" and ch == 3:
        s = "\n".join(["[진로독서] 독서 전 활동", f"· 예상한 내용: {q(k.get('guess'))}",
                       f"· 궁금한 목차: {q(k.get('toc'))}", f"· 알고 싶은 것: {q(k.get('want'))}"])
    else:
        s = x["text"]
    return title, desc, s, meta


def teacher_line(x):
    st, memo = x.get("status"), x.get("memo")
    if not st and not memo:
        return "교사 확인 전"
    t = f"교사 확인: {st or '-'}"
    if memo:
        t += f" / 교사 메모: '{cut(memo, 150)}'"
    return t

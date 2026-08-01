# -*- coding: utf-8 -*-
"""
K-Color 퍼스널컬러 앱 — 데이터 빌드 (v8 규칙)

입력  kcolor_data.json (776색) + K-color DB v5.xlsx (Pantone / Color Name)
출력  kcolor_typed_v8.json

v8 규칙 요약
  1) Pantone 열이 'x' 인 색은 제외              → 606색
  2) 색상별 웜/쿨 을 표로 직접 지정 (중성 없음)   → 웜 295 / 쿨 311
     · 무채색(C=0)은 쿨
  3) 유형은 톤(tone)으로 배정. 같은 온도 안에서만 중복 허용
  4) 가을 딥·겨울 딥의 어두운회(dkgy)·검은(bk) 은
     색상 표기가 2.5x / 7.5x 인 색 제외 (5x · 10x · N 만 허용)
"""
import json, collections, os, re

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = '/sessions/modest-vigilant-newton/mnt/K-Color/kcolor_data.json'
XLS = '/sessions/modest-vigilant-newton/mnt/K-Color/K-color DB v5.xlsx'
DST = os.path.join(BASE, 'kcolor_typed_v8.json')

EXCLUDE_MARK = 'x'          # Pantone 열이 이 값이면 팔레트에서 제외

# ── 1) 웜 / 쿨 색상 ─────────────────────────────────────────
WARM_HUES = ['7.5R', '10R',
             '2.5YR', '5YR', '7.5YR', '10YR',
             '2.5Y', '5Y', '7.5Y',
             '2.5GY', '5GY', '7.5GY', '10GY',
             '2.5G', '5G', '7.5G',
             '2.5B', '5B',
             '10P']
COOL_HUES = ['2.5R', '5R',
             '10Y',
             '10G',
             '2.5BG', '5BG', '7.5BG', '10BG',
             '7.5B', '10B',
             '2.5PB', '5PB', '7.5PB', '10PB',
             '2.5P', '5P', '7.5P',
             '2.5RP', '5RP', '7.5RP', '10RP']

# ── 2) 톤코드 (톤경계표_양식_v5.xlsx '1.규칙맵' 과 동일) ────────
TONE_NAMES = {
    'vv': '선명한', 'lt': '밝은', 'c': '기본', 'pl': '연한', 'sf': '흐린',
    'dl': '탁한', 'dp': '진한', 'dk': '어두운', 'wh': '흰', 'vltgy': '아주밝은회',
    'ltgy': '밝은회', 'gy': '회', 'dkgy': '어두운회', 'bk': '검은',
}

def tone_code(V, C):
    if C >= 11: return 'vv'
    if C >= 7:
        if V >= 7:   return 'lt'
        if V >= 4.5: return 'c'
        return 'dp'
    if C >= 3:
        if V >= 7.5: return 'pl'
        if V >= 6:   return 'sf'
        if V >= 4.5: return 'dl'
        if V >= 3.5: return 'dp'
        return 'dk'
    if V >= 9:   return 'wh'
    if V >= 8:   return 'vltgy'
    if V >= 7:   return 'ltgy'
    if V >= 4.5: return 'gy'
    if V >= 3:   return 'dkgy'
    return 'bk'

# ── 3) 유형 정의 + 유형별 주요 톤 ──────────────────────────────
TYPES = [
    ('SP_LT', '봄 라이트',    'Light Spring',  'spring', ['pl', 'lt', 'wh']),
    ('SP_BR', '봄 브라이트',  'Bright Spring', 'spring', ['c', 'vv']),
    ('SP_WM', '봄 웜',        'Spring Warm',   'spring', ['c', 'lt']),
    ('SU_LT', '여름 라이트',  'Light Summer',  'summer', ['wh', 'pl', 'vltgy', 'ltgy']),
    ('SU_MU', '여름 뮤트',    'Soft Summer',   'summer', ['sf', 'dl', 'gy']),
    ('SU_CL', '여름 쿨',      'Summer Cool',   'summer', ['sf', 'lt']),
    ('AU_MU', '가을 뮤트',    'Soft Autumn',   'autumn', ['sf', 'dl', 'gy', 'ltgy']),
    ('AU_DP', '가을 딥',      'Deep Autumn',   'autumn', ['dp', 'dk', 'dkgy', 'bk']),
    ('AU_WM', '가을 웜',      'Autumn Warm',   'autumn', ['c', 'dp', 'dl']),
    ('WI_BR', '겨울 브라이트','Bright Winter', 'winter', ['c', 'vv']),
    ('WI_DP', '겨울 딥',      'Deep Winter',   'winter', ['dp', 'dk', 'dkgy', 'bk']),
    ('WI_CL', '겨울 쿨',      'Winter Cool',   'winter', ['c', 'dp', 'dl']),
]
WARM_SEASONS = ('spring', 'autumn')

# ── 4) 딥 유형의 어두운회·검은 색상 제한 ────────────────────────
DEEP_TYPES = ('AU_DP', 'WI_DP')
DEEP_LIMIT_TONES = ('dkgy', 'bk')

HALF_STEP = re.compile(r'^\d+\.5')

def is_half_step(hue):
    """2.5R · 7.5YR 처럼 숫자가 .5 인 색상인지 (N 은 제외)"""
    return hue != 'N' and bool(HALF_STEP.match(hue))

# ---------------------------------------------------------------------------
def pantone_names():
    """K-color DB v5.xlsx 의 Pantone / Color Name 열을 K-Color ID 기준으로 읽는다."""
    import openpyxl
    ws = openpyxl.load_workbook(XLS, read_only=True, data_only=True)['Sheet1']
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        kid, pantone, cname = r[2], r[14], r[15]
        if not kid:
            continue
        excluded = str(pantone).strip().lower() == EXCLUDE_MARK if pantone is not None else True
        out[kid] = {'pantone': pantone, 'pname': cname or '', 'excluded': excluded}
    return out


def main():
    d = json.load(open(SRC, encoding='utf-8'))
    PN = pantone_names()

    out, cnt = [], collections.Counter()
    excluded, unassigned = [], []

    for c in d['colors']:
        info = PN.get(c['id']) or {}
        if info.get('excluded', True):
            excluded.append(c['id'])
            continue

        V, C, hue = c['value'], c['chroma'], c['hue']
        tone = tone_code(V, C)
        temp = 'W' if (C > 0 and hue in WARM_HUES) else 'C'   # 무채색(C=0)은 쿨

        types = []
        for code, ko, en, season, tones in TYPES:
            if (season in WARM_SEASONS) != (temp == 'W'):
                continue
            if tone not in tones:
                continue
            if code in DEEP_TYPES and tone in DEEP_LIMIT_TONES and is_half_step(hue):
                continue
            types.append(code)

        if not types:
            unassigned.append(c['id'])
        for t in types:
            cnt[t] += 1

        rec = {k: c[k] for k in ('id', 'hue', 'group', 'value', 'chroma', 'ks',
                                 'lab', 'rgb', 'cmyk', 'hex', 'ncs', 'name')}
        rec['pantone'] = info.get('pantone')
        rec['pname'] = info.get('pname', '')
        rec['tone'] = tone
        rec['temp'] = temp
        rec['types'] = types
        out.append(rec)

    payload = {
        'meta': dict(
            d['meta'],
            tone_names=TONE_NAMES,
            types=[{'code': a, 'ko': b, 'en': e, 'season': s} for a, b, e, s, _ in TYPES],
            tone_map={a: t for a, _, _, _, t in TYPES},
            warm_hues=WARM_HUES,
            cool_hues=COOL_HUES,
            achromatic_side='무채색(C=0)은 쿨',
            deep_five_only='가을 딥·겨울 딥의 어두운회(dkgy)·검은(bk) 은 2.5x/7.5x 색상 제외',
            excluded_count=len(excluded),
            excluded_reason='K-color DB v5.xlsx 의 Pantone 열이 x 로 표시된 색 제외',
            excluded_ids=excluded,
            unassigned_ids=unassigned,
        ),
        'colors': out,
    }
    json.dump(payload, open(DST, 'w', encoding='utf-8'), ensure_ascii=False)

    used = sum(1 for c in out if c['types'])
    warm = sum(1 for c in out if c['temp'] == 'W')
    print(f'원본 {len(d["colors"])} / Pantone=x 제외 {len(excluded)} / 사용 가능 {len(out)}')
    print(f'웜 {warm} · 쿨 {len(out) - warm} · 유형 배정된 색 {used} · 미배정 {len(unassigned)}')
    print('-' * 34)
    for code, ko, en, s, _ in TYPES:
        print(f'{ko:8s} {code:6s} {cnt[code]:4d}')
    print('-' * 34)
    print('합계(중복 포함)', sum(cnt.values()))


if __name__ == '__main__':
    main()

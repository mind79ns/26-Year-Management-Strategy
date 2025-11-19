#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
과거 전략 PPT 종합 분석
패턴, 트렌드, 핵심 키워드 추출
"""

import json
import re
from collections import Counter

def analyze_comprehensive():
    with open('전략PPT_분석결과.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    print("="*100)
    print("📊 과거 전략 PPT 종합 분석 보고서")
    print("="*100)

    # 1. 파일별 슬라이드 수
    print("\n1️⃣ 전략 PPT 기본 정보")
    print("-"*100)
    for ppt in data:
        if "error" not in ppt:
            print(f"  📄 {ppt['file_name']}: {ppt['total_slides']}개 슬라이드")

    # 2. 전체 텍스트 추출 및 키워드 분석
    all_texts = []
    for ppt in data:
        if "error" not in ppt:
            for slide in ppt["slides"]:
                all_texts.extend(slide["texts"])

    # 키워드 추출
    keywords = []
    for text in all_texts:
        # 괄호, 특수문자 제거
        text = re.sub(r'[\[\]\(\)]', ' ', text)
        # 단어 분리
        words = text.split()
        for word in words:
            if len(word) >= 2 and not word.startswith('[TABLE]'):
                keywords.append(word)

    keyword_counter = Counter(keywords)

    print("\n2️⃣ 핵심 키워드 TOP 30")
    print("-"*100)
    for i, (keyword, count) in enumerate(keyword_counter.most_common(30), 1):
        print(f"  {i:2d}. {keyword:20s} : {count:3d}회")

    # 3. 핵심 주제 분석
    print("\n3️⃣ 핵심 주제 분류")
    print("-"*100)

    themes = {
        "효율/가동율": ["효율", "가동율", "가동", "평가가동율"],
        "유실/손실": ["유실", "손실", "LOSS", "loss"],
        "점당가공비/비용": ["점당가공비", "가공비", "비용", "COST", "원가"],
        "품질/불량": ["품질", "불량", "PPM", "ppm", "양품"],
        "설비/CAPA": ["설비", "CAPA", "capa", "능력", "라인", "LINE"],
        "개선/과제": ["개선", "과제", "추진", "목표", "활동"],
        "SMD/공정": ["SMD", "AXIAL", "RADIAL", "IMT", "공정"],
        "MES/자동화": ["MES", "자동화", "시스템", "SYSTEM", "DATA"]
    }

    theme_counts = {}
    for theme_name, theme_keywords in themes.items():
        count = sum(keyword_counter.get(kw, 0) for kw in theme_keywords)
        theme_counts[theme_name] = count

    for theme, count in sorted(theme_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"  • {theme:20s} : {count:4d}회 언급")

    # 4. 연도별 주요 전략 요약
    print("\n4️⃣ 연도별 주요 전략 요약")
    print("-"*100)

    strategy_summaries = {
        "21년smd전략.pptx": "SMD 공정 점당 가공비 상승 원인 분석 및 유실 개선",
        "22년 제조1 경영전략 R2.pptx": "21년 성과 반성 및 22년 핵심 추진 과제",
        "하노이 법인 21년 경영 전략 20201217.pptx": "무삽 불량 개선 및 SMD 설비 유실 개선",
        "하노이 법인 21년 하반기 경영 전략_ 제조1_REV3.pptx": "MES System 정착 및 KPI 목표 달성",
        "하노이 법인 22년 하반기 경영 전략 제조1팀 R3.pptx": "자동화 공정 지표 개선 및 Main Line 혁신"
    }

    for filename, summary in strategy_summaries.items():
        print(f"\n  📅 {filename}")
        print(f"     → {summary}")

    # 5. 공통 패턴 및 트렌드
    print("\n5️⃣ 공통 패턴 및 트렌드 분석")
    print("-"*100)

    patterns = [
        "✓ 점당 가공비 절감이 핵심 목표로 지속 반복",
        "✓ 유실(손실) 개선이 주요 전략 과제",
        "✓ SMD, AXIAL, RADIAL 공정별 개선 활동",
        "✓ 평가가동율/효율 향상 KPI 설정",
        "✓ MES/자동화 시스템 구축 및 활용",
        "✓ 설비 CAPA 증가 및 최적화",
        "✓ 품질 불량 감소 목표",
        "✓ WORST LINE/MODEL 집중 개선",
        "✓ SPARE PART 비용 관리",
        "✓ 정량적 목표 설정 (%, ppm, 건수)"
    ]

    for pattern in patterns:
        print(f"  {pattern}")

    # 6. 핵심 KPI 항목
    print("\n6️⃣ 주요 KPI 항목")
    print("-"*100)

    kpis = [
        "평가가동율/효율 (%)",
        "유실률 (%)",
        "점당 가공비",
        "품질 불량률 (ppm)",
        "설비 CAPA",
        "C/T (Cycle Time)",
        "SPARE PART 비용",
        "노무비"
    ]

    for kpi in kpis:
        print(f"  • {kpi}")

    print("\n" + "="*100)
    print("✅ 분석 완료")
    print("="*100)

if __name__ == "__main__":
    analyze_comprehensive()

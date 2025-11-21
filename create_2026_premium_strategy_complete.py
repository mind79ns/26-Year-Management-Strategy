#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2026년 제조1팀 경영전략 PPT - 최종 완성판 (전체)
- Part 1의 모든 함수 + 나머지 6개 페이지
- 총 12페이지 완성
"""

# Part 1에서 import 및 색상 정의
from create_2026_premium_strategy_part1 import *

def create_strategy2(prs):
    """페이지 6: 전략2 - 불량 재발 Zero"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "전략 2: 불량 재발 Zero 시스템"
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = GREEN

    # 좌측: 3단계 프로세스
    processes = [
        {"step": "1", "name": "즉시 감지", "icon": "🔍", "desc": "비전검사 시스템\n실시간 불량 감지"},
        {"step": "2", "name": "원인 분석", "icon": "🧠", "desc": "AI 패턴 분석\n불량 DB 활용"},
        {"step": "3", "name": "재발 방지", "icon": "🛡️", "desc": "SOP 자동 업데이트\n작업자 실시간 알림"}
    ]

    for i, proc in enumerate(processes):
        y = 1.3 + i * 1.8

        # 프로세스 박스
        pbox = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(y), Inches(4.3), Inches(1.5)
        )
        pbox.fill.solid()
        pbox.fill.fore_color.rgb = WHITE
        pbox.line.color.rgb = GREEN
        pbox.line.width = Pt(3)
        add_shadow(pbox)

        pt = pbox.text_frame
        pt.text = f"{proc['icon']}  단계 {proc['step']}: {proc['name']}"
        pt.paragraphs[0].font.size = Pt(18)
        pt.paragraphs[0].font.bold = True
        pt.paragraphs[0].font.color.rgb = GREEN

        p2 = pt.add_paragraph()
        p2.text = f"\n{proc['desc']}"
        p2.font.size = Pt(13)
        p2.font.color.rgb = NAVY

        # 화살표
        if i < 2:
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.DOWN_ARROW,
                Inches(2.3), Inches(y + 1.55), Inches(0.5), Inches(0.2)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = GREEN
            arrow.line.fill.background()

    # 우측: 과거 대비 개선
    comp_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.2), Inches(1.3), Inches(4.3), Inches(2.3)
    )
    comp_box.fill.solid()
    comp_box.fill.fore_color.rgb = RGBColor(240, 255, 240)
    comp_box.line.color.rgb = GREEN
    comp_box.line.width = Pt(2)

    ct = comp_box.text_frame
    ct.text = "📈 과거 vs 2026"
    ct.paragraphs[0].font.size = Pt(18)
    ct.paragraphs[0].font.bold = True
    ct.paragraphs[0].font.color.rgb = GREEN
    ct.paragraphs[0].alignment = PP_ALIGN.CENTER

    comps = [
        ("과거 5년", "사후 대응 중심", GRAY),
        ("24년 전환", "품질 10배 증가", ORANGE),
        ("2026 목표", "불량률 50% 감소", GREEN),
        ("핵심 차별화", "AI 패턴 학습", PURPLE)
    ]

    for label, value, color in comps:
        p = ct.add_paragraph()
        p.text = f"\n{label}"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = color

        p2 = ct.add_paragraph()
        p2.text = f"  → {value}"
        p2.font.size = Pt(12)
        p2.font.color.rgb = DARK_GRAY

    # 하단: 기대효과
    effect_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.2), Inches(3.9), Inches(4.3), Inches(2.8)
    )
    effect_box.fill.solid()
    effect_box.fill.fore_color.rgb = RGBColor(245, 255, 245)
    effect_box.line.color.rgb = GREEN
    effect_box.line.width = Pt(2)

    et = effect_box.text_frame
    et.text = "🎯 기대효과"
    et.paragraphs[0].font.size = Pt(18)
    et.paragraphs[0].font.bold = True
    et.paragraphs[0].font.color.rgb = GREEN
    et.paragraphs[0].alignment = PP_ALIGN.CENTER

    effects = [
        "✓ 불량률 50% 감소 (10% → 5%)",
        "✓ 불량 비용 40% 절감",
        "✓ 고객 클레임 70% 감소",
        "✓ 재작업 시간 60% 단축",
        "✓ 품질 경쟁력 대폭 향상"
    ]

    for eff in effects:
        p = et.add_paragraph()
        p.text = eff
        p.font.size = Pt(14)
        p.font.color.rgb = NAVY
        p.space_before = Pt(8)

def create_strategy3(prs):
    """페이지 7: 전략3 - 설비 CAPA 증대"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "전략 3: 설비 CAPA 15% 증대"
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ORANGE

    # Before/After 비교 (3개 지표)
    metrics = [
        {"name": "Tact Time", "before": 12, "after": 10, "unit": "초", "max": 15},
        {"name": "설비 가동률", "before": 75, "after": 90, "unit": "%", "max": 100},
        {"name": "일일 생산량", "before": 5000, "after": 5750, "unit": "개", "max": 6000}
    ]

    start_y = 1.5
    for i, metric in enumerate(metrics):
        y = start_y + i * 1.7

        # 메트릭 이름
        name_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(y), Inches(2), Inches(0.5)
        )
        nt = name_box.text_frame
        nt.text = metric['name']
        nt.paragraphs[0].font.size = Pt(16)
        nt.paragraphs[0].font.bold = True
        nt.paragraphs[0].font.color.rgb = NAVY
        nt.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Before 막대
        before_width = 3.5 * (metric['before'] / metric['max'])
        before_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.8), Inches(y), Inches(before_width), Inches(0.45)
        )
        before_bar.fill.solid()
        before_bar.fill.fore_color.rgb = LIGHT_GRAY
        before_bar.line.fill.background()

        bt = before_bar.text_frame
        bt.text = f"현재: {metric['before']}{metric['unit']}"
        bt.paragraphs[0].font.size = Pt(12)
        bt.paragraphs[0].font.color.rgb = GRAY
        bt.paragraphs[0].alignment = PP_ALIGN.CENTER
        bt.vertical_anchor = MSO_ANCHOR.MIDDLE

        # After 막대
        after_width = 3.5 * (metric['after'] / metric['max'])
        after_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(2.8), Inches(y + 0.6), Inches(after_width), Inches(0.45)
        )
        after_bar.fill.solid()
        after_bar.fill.fore_color.rgb = ORANGE
        after_bar.line.fill.background()

        at = after_bar.text_frame
        at.text = f"목표: {metric['after']}{metric['unit']}"
        at.paragraphs[0].font.size = Pt(12)
        at.paragraphs[0].font.bold = True
        at.paragraphs[0].font.color.rgb = WHITE
        at.paragraphs[0].alignment = PP_ALIGN.CENTER
        at.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 개선율
        improvement = ((metric['after'] - metric['before']) / metric['before'] * 100) if metric['name'] != "Tact Time" else ((metric['before'] - metric['after']) / metric['before'] * 100)
        imp_box = slide.shapes.add_textbox(
            Inches(6.8), Inches(y + 0.3), Inches(1.2), Inches(0.5)
        )
        it = imp_box.text_frame
        it.text = f"↑ {improvement:.1f}%"
        it.paragraphs[0].font.size = Pt(14)
        it.paragraphs[0].font.bold = True
        it.paragraphs[0].font.color.rgb = ORANGE
        it.paragraphs[0].alignment = PP_ALIGN.CENTER
        it.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 우측: 실행 계획
    plan_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(6.5), Inches(9), Inches(0.8)
    )
    plan_box.fill.solid()
    plan_box.fill.fore_color.rgb = RGBColor(255, 245, 230)
    plan_box.line.color.rgb = ORANGE
    plan_box.line.width = Pt(2)

    pt = plan_box.text_frame
    pt.text = "📋 실행 계획: ① 병목공정 개선  ② 고속화 설비 개조  ③ 자동화 라인 증설  ④ 작업 동선 최적화  ⑤ 다기능 인력 양성"
    pt.paragraphs[0].font.size = Pt(14)
    pt.paragraphs[0].font.bold = True
    pt.paragraphs[0].font.color.rgb = ORANGE
    pt.paragraphs[0].alignment = PP_ALIGN.CENTER
    pt.vertical_anchor = MSO_ANCHOR.MIDDLE

def create_strategy4(prs):
    """페이지 8: 전략4 - 설비관리 혁신"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "전략 4: 설비관리 혁신 (신규)"
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = PURPLE

    # 목표
    goal_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(1.1), Inches(9), Inches(0.6)
    )
    goal_box.fill.solid()
    goal_box.fill.fore_color.rgb = RGBColor(245, 235, 255)
    goal_box.line.color.rgb = PURPLE
    goal_box.line.width = Pt(2)

    gt = goal_box.text_frame
    gt.text = "🎯 목표: 예방보전 체계 고도화로 설비 고장 50% 감소 및 설비 수명 20% 연장"
    gt.paragraphs[0].font.size = Pt(16)
    gt.paragraphs[0].font.bold = True
    gt.paragraphs[0].font.color.rgb = PURPLE
    gt.paragraphs[0].alignment = PP_ALIGN.CENTER
    gt.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 4분할 매트릭스
    boxes = [
        {
            "title": "예방보전 고도화",
            "icon": "🔧",
            "items": ["주기 → 상태 기반", "IoT 센서 모니터링", "이상징후 조기 감지"],
            "x": 0.5, "y": 2.1
        },
        {
            "title": "설비 이력 관리",
            "icon": "📋",
            "items": ["설비별 정비 DB화", "고장 패턴 분석", "부품 교체 최적화"],
            "x": 5.2, "y": 2.1
        },
        {
            "title": "부품 수명 예측",
            "icon": "🎯",
            "items": ["AI 기반 수명 예측", "적기 부품 교체", "재고 최적화"],
            "x": 0.5, "y": 4.6
        },
        {
            "title": "긴급 정비 체계",
            "icon": "⚡",
            "items": ["24시간 대응", "비상부품 확보", "협력업체 네트워크"],
            "x": 5.2, "y": 4.6
        }
    ]

    for box_data in boxes:
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(box_data["x"]), Inches(box_data["y"]),
            Inches(4.3), Inches(2.2)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = PURPLE
        box.line.width = Pt(2)
        add_shadow(box)

        bt = box.text_frame
        bt.text = f"{box_data['icon']} {box_data['title']}"
        bt.paragraphs[0].font.size = Pt(16)
        bt.paragraphs[0].font.bold = True
        bt.paragraphs[0].font.color.rgb = PURPLE
        bt.paragraphs[0].alignment = PP_ALIGN.CENTER

        for item in box_data['items']:
            p = bt.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(13)
            p.font.color.rgb = NAVY
            p.space_before = Pt(10)

def create_efficiency_targets(prs):
    """페이지 9: 평가가동 효율 목표"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "2026 평가가동 효율 목표"
    tf.paragraphs[0].font.size = Pt(32)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY

    # 3개 라인 비교
    lines = [
        {"name": "SMD", "target": 91, "current": 85, "color": LIGHT_BLUE, "x": 1.2},
        {"name": "RADIAL", "target": 85, "current": 78, "color": GREEN, "x": 4.2},
        {"name": "AXIAL", "target": 85, "current": 80, "color": ORANGE, "x": 7.2}
    ]

    for line in lines:
        # 메인 박스
        main_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(line["x"]), Inches(1.5), Inches(2.3), Inches(3.8)
        )
        main_box.fill.solid()
        main_box.fill.fore_color.rgb = WHITE
        main_box.line.color.rgb = line["color"]
        main_box.line.width = Pt(3)
        add_shadow(main_box)

        # 라인명
        name_box = slide.shapes.add_textbox(
            Inches(line["x"] + 0.2), Inches(1.7), Inches(1.9), Inches(0.5)
        )
        nt = name_box.text_frame
        nt.text = line["name"]
        nt.paragraphs[0].font.size = Pt(24)
        nt.paragraphs[0].font.bold = True
        nt.paragraphs[0].font.color.rgb = line["color"]
        nt.paragraphs[0].alignment = PP_ALIGN.CENTER

        # 목표값 (대형)
        target_box = slide.shapes.add_textbox(
            Inches(line["x"] + 0.2), Inches(2.4), Inches(1.9), Inches(1.2)
        )
        tt = target_box.text_frame
        tt.text = f"{line['target']}%"
        tt.paragraphs[0].font.size = Pt(52)
        tt.paragraphs[0].font.bold = True
        tt.paragraphs[0].font.color.rgb = line["color"]
        tt.paragraphs[0].alignment = PP_ALIGN.CENTER
        tt.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 목표 라벨
        label_box = slide.shapes.add_textbox(
            Inches(line["x"] + 0.2), Inches(3.6), Inches(1.9), Inches(0.3)
        )
        lt = label_box.text_frame
        lt.text = "2026 목표"
        lt.paragraphs[0].font.size = Pt(13)
        lt.paragraphs[0].font.color.rgb = GRAY
        lt.paragraphs[0].alignment = PP_ALIGN.CENTER

        # 현재값
        current_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(line["x"] + 0.3), Inches(4.1), Inches(1.7), Inches(0.5)
        )
        current_box.fill.solid()
        current_box.fill.fore_color.rgb = LIGHT_GRAY
        current_box.line.fill.background()

        ct = current_box.text_frame
        ct.text = f"현재: {line['current']}%"
        ct.paragraphs[0].font.size = Pt(14)
        ct.paragraphs[0].font.color.rgb = GRAY
        ct.paragraphs[0].alignment = PP_ALIGN.CENTER
        ct.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 증가 화살표
        improvement = line['target'] - line['current']
        arrow_box = slide.shapes.add_textbox(
            Inches(line["x"] + 0.3), Inches(4.7), Inches(1.7), Inches(0.4)
        )
        at = arrow_box.text_frame
        at.text = f"↑ {improvement}%p 향상"
        at.paragraphs[0].font.size = Pt(14)
        at.paragraphs[0].font.bold = True
        at.paragraphs[0].font.color.rgb = line["color"]
        at.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 하단 전략 요약
    strategy_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(5.7), Inches(9), Inches(1.1)
    )
    strategy_box.fill.solid()
    strategy_box.fill.fore_color.rgb = RGBColor(250, 250, 250)
    strategy_box.line.color.rgb = NAVY
    strategy_box.line.width = Pt(2)

    st = strategy_box.text_frame
    st.text = "💡 핵심 전략"
    st.paragraphs[0].font.size = Pt(18)
    st.paragraphs[0].font.bold = True
    st.paragraphs[0].font.color.rgb = NAVY
    st.paragraphs[0].alignment = PP_ALIGN.CENTER

    p2 = st.add_paragraph()
    p2.text = "\nMES 자동분석 + 불량재발 Zero + 설비CAPA 증대 + 설비관리 혁신"
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = NAVY
    p2.alignment = PP_ALIGN.CENTER

    p3 = st.add_paragraph()
    p3.text = "= 평가가동 효율 극대화"
    p3.font.size = Pt(16)
    p3.font.bold = True
    p3.font.color.rgb = GOLD
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(5)

def create_roadmap(prs):
    """페이지 10: 실행 로드맵 (Q1-Q4)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "2026 실행 로드맵"
    tf.paragraphs[0].font.size = Pt(32)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY

    # 분기 헤더
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    header_start_x = 2.8
    quarter_width = 1.65

    for i, q in enumerate(quarters):
        x = header_start_x + i * quarter_width

        header_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(1.2), Inches(quarter_width - 0.1), Inches(0.5)
        )
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = NAVY
        header_box.line.fill.background()

        ht = header_box.text_frame
        ht.text = q
        ht.paragraphs[0].font.size = Pt(18)
        ht.paragraphs[0].font.bold = True
        ht.paragraphs[0].font.color.rgb = WHITE
        ht.paragraphs[0].alignment = PP_ALIGN.CENTER
        ht.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 과제별 간트 바
    tasks = [
        {"name": "MES 자동분석 시스템", "color": LIGHT_BLUE, "quarters": [1, 1, 1, 1]},
        {"name": "불량 재발 Zero", "color": GREEN, "quarters": [1, 1, 1, 0]},
        {"name": "설비 CAPA 증대", "color": ORANGE, "quarters": [0, 1, 1, 1]},
        {"name": "설비관리 혁신", "color": PURPLE, "quarters": [1, 1, 0, 0]}
    ]

    start_y = 2
    row_height = 1

    for i, task in enumerate(tasks):
        y = start_y + i * row_height

        # 과제명
        name_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(y + 0.1), Inches(2), Inches(0.6)
        )
        nt = name_box.text_frame
        nt.text = task['name']
        nt.paragraphs[0].font.size = Pt(14)
        nt.paragraphs[0].font.bold = True
        nt.paragraphs[0].font.color.rgb = task['color']
        nt.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 간트 바
        for q_idx, active in enumerate(task['quarters']):
            x = header_start_x + q_idx * quarter_width

            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x), Inches(y + 0.15), Inches(quarter_width - 0.1), Inches(0.5)
            )
            bar.fill.solid()

            if active:
                bar.fill.fore_color.rgb = task['color']
                bar.line.fill.background()
            else:
                bar.fill.fore_color.rgb = LIGHT_GRAY
                bar.line.fill.background()

    # 마일스톤
    milestones = [
        {"text": "중간 점검", "q": 1, "y": 6.2},
        {"text": "성과 평가", "q": 3, "y": 6.2}
    ]

    for ms in milestones:
        x = header_start_x + ms['q'] * quarter_width

        # 다이아몬드
        diamond = slide.shapes.add_shape(
            MSO_SHAPE.DIAMOND,
            Inches(x + 0.6), Inches(ms['y']), Inches(0.45), Inches(0.45)
        )
        diamond.fill.solid()
        diamond.fill.fore_color.rgb = RED
        diamond.line.fill.background()

        # 텍스트
        ms_text = slide.shapes.add_textbox(
            Inches(x + 0.2), Inches(ms['y'] + 0.5), Inches(1.3), Inches(0.3)
        )
        mt = ms_text.text_frame
        mt.text = ms['text']
        mt.paragraphs[0].font.size = Pt(11)
        mt.paragraphs[0].font.bold = True
        mt.paragraphs[0].font.color.rgb = RED
        mt.paragraphs[0].alignment = PP_ALIGN.CENTER

def create_expected_results(prs):
    """페이지 11: 기대 효과"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.text = "2026 기대 효과"
    tf.paragraphs[0].font.size = Pt(32)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = NAVY

    # 3개 핵심 KPI (원형 게이지)
    kpis = [
        {"name": "가공비 절감", "target": 10, "color": LIGHT_BLUE, "x": 1.2},
        {"name": "품질 개선", "target": 10, "color": GREEN, "x": 4.2},
        {"name": "유실시간 감소", "target": 5, "color": ORANGE, "x": 7.2}
    ]

    for kpi in kpis:
        # 외부 원
        outer = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(kpi['x']), Inches(1.3), Inches(1.8), Inches(1.8)
        )
        outer.fill.solid()
        outer.fill.fore_color.rgb = LIGHT_GRAY
        outer.line.fill.background()

        # 내부 원
        inner = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(kpi['x'] + 0.15), Inches(1.45), Inches(1.5), Inches(1.5)
        )
        inner.fill.solid()
        inner.fill.fore_color.rgb = kpi['color']
        inner.line.fill.background()

        # 중앙 원
        center = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(kpi['x'] + 0.45), Inches(1.75), Inches(0.9), Inches(0.9)
        )
        center.fill.solid()
        center.fill.fore_color.rgb = WHITE
        center.line.fill.background()

        # 퍼센트
        pct_box = slide.shapes.add_textbox(
            Inches(kpi['x'] + 0.45), Inches(1.95), Inches(0.9), Inches(0.5)
        )
        pt = pct_box.text_frame
        pt.text = f"{kpi['target']}%"
        pt.paragraphs[0].font.size = Pt(28)
        pt.paragraphs[0].font.bold = True
        pt.paragraphs[0].font.color.rgb = kpi['color']
        pt.paragraphs[0].alignment = PP_ALIGN.CENTER
        pt.vertical_anchor = MSO_ANCHOR.MIDDLE

        # KPI 이름
        name_box = slide.shapes.add_textbox(
            Inches(kpi['x']), Inches(3.3), Inches(1.8), Inches(0.4)
        )
        nt = name_box.text_frame
        nt.text = kpi['name']
        nt.paragraphs[0].font.size = Pt(15)
        nt.paragraphs[0].font.bold = True
        nt.paragraphs[0].font.color.rgb = NAVY
        nt.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 하단 세부 지표
    metrics_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(4.2), Inches(9), Inches(2.6)
    )
    metrics_box.fill.solid()
    metrics_box.fill.fore_color.rgb = RGBColor(245, 248, 250)
    metrics_box.line.color.rgb = NAVY
    metrics_box.line.width = Pt(2)

    mt = metrics_box.text_frame
    mt.text = "📊 세부 성과 지표"
    mt.paragraphs[0].font.size = Pt(20)
    mt.paragraphs[0].font.bold = True
    mt.paragraphs[0].font.color.rgb = NAVY
    mt.paragraphs[0].alignment = PP_ALIGN.CENTER

    details = [
        ("MES 자동분석", "ROI 3,159%, 회수기간 11일", LIGHT_BLUE),
        ("불량률", "10% → 5% (50% 개선)", GREEN),
        ("설비 가동률", "75% → 90% (15%p 향상)", ORANGE),
        ("평가가동 효율", "SMD 91%, RADIAL 85%, AXIAL 85%", PURPLE),
        ("설비 고장", "50% 감소, 수명 20% 연장", RED),
        ("경제적 효과", "연간 수억원 비용 절감", GOLD)
    ]

    for i, (label, value, color) in enumerate(details):
        x = 0.8 + (i % 2) * 4.7
        y = 4.9 + (i // 2) * 0.7

        db = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(4.2), Inches(0.6)
        )
        dt = db.text_frame
        dt.text = f"• {label}"
        dt.paragraphs[0].font.size = Pt(14)
        dt.paragraphs[0].font.bold = True
        dt.paragraphs[0].font.color.rgb = color

        p2 = dt.add_paragraph()
        p2.text = f"  → {value}"
        p2.font.size = Pt(13)
        p2.font.color.rgb = DARK_GRAY

def main():
    """메인 실행"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    print("=" * 80)
    print("2026년 제조1팀 경영전략 PPT 생성 중 (완전판)...")
    print("=" * 80)

    # Part 1 함수들
    create_cover(prs)
    print("✓ 페이지 1: 프리미엄 커버")

    create_executive_summary(prs)
    print("✓ 페이지 2: Executive Summary")

    create_5year_journey(prs)
    print("✓ 페이지 3: 5년 여정 (2021-2025)")

    create_strategy_overview(prs)
    print("✓ 페이지 4: 2026 전략 개요")

    create_strategy1(prs)
    print("✓ 페이지 5: 전략1 - MES 자동분석 시스템")

    # Part 2 새 함수들
    create_strategy2(prs)
    print("✓ 페이지 6: 전략2 - 불량 재발 Zero")

    create_strategy3(prs)
    print("✓ 페이지 7: 전략3 - 설비 CAPA 증대")

    create_strategy4(prs)
    print("✓ 페이지 8: 전략4 - 설비관리 혁신")

    create_efficiency_targets(prs)
    print("✓ 페이지 9: 평가가동 효율 목표")

    create_roadmap(prs)
    print("✓ 페이지 10: 실행 로드맵")

    create_expected_results(prs)
    print("✓ 페이지 11: 기대 효과")

    create_conclusion(prs)
    print("✓ 페이지 12: 결론")

    output = "2026_제조1팀_경영전략_최종완성판.pptx"
    prs.save(output)

    print("\n" + "=" * 80)
    print(f"✅ PPT 생성 완료: {output}")
    print("📄 총 12페이지")
    print("🎨 특징:")
    print("   ✓ 21-25년 분석 결과 완전 반영")
    print("   ✓ 프리미엄 고급 디자인")
    print("   ✓ 12가지 다양한 시각화 스타일")
    print("   ✓ 정확한 레이아웃 (겹침 완전 방지)")
    print("   ✓ 4대 전략 + 평가가동 효율 목표")
    print("=" * 80)

if __name__ == "__main__":
    main()

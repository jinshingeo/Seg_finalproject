from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# A4 page setup
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(2.5)
sec.top_margin = Cm(2.5)
sec.bottom_margin = Cm(2.5)

FONT = 'Palatino Linotype'
BODY_PT = 11
H1_PT = 12
H2_PT = 11
CAP_PT = 9
LINE_SPACING = 1.48
INDENT = Inches(0.30)


def set_run(run, size=BODY_PT, bold=False, italic=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def body(text, first_indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = LINE_SPACING
    p.paragraph_format.space_after = Pt(0)
    if first_indent:
        p.paragraph_format.first_line_indent = INDENT
    set_run(p.add_run(text))
    return p


def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.0
    set_run(p.add_run(text), size=H1_PT, bold=True)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.0
    set_run(p.add_run(text), size=H2_PT, bold=True)
    return p


def h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.0
    set_run(p.add_run(text), size=H2_PT, italic=True)
    return p


def fig_placeholder(num, desc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'DCDCDC')
    pPr.append(shd)
    set_run(p.add_run(f'[그림 {num} 자리: {desc}]'),
            size=CAP_PT, color=RGBColor(0x55, 0x55, 0x55))

    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Inches(0)
    cap.paragraph_format.space_after = Pt(10)
    set_run(cap.add_run(f'Figure {num}. {desc}.'), size=CAP_PT)


def eq_text(num, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(f'{text}    ({num})')
    run.font.name = 'Cambria Math'
    run.font.size = Pt(11)


# ─────────────────────────────────────────────────────
# TITLE
# ─────────────────────────────────────────────────────
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.first_line_indent = Inches(0)
p_title.paragraph_format.space_after = Pt(4)
set_run(p_title.add_run('서울시 저소득 독거노인의 소득 기반 거주지 분리 측도 분석'),
        size=14, bold=True)

p_title_en = doc.add_paragraph()
p_title_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title_en.paragraph_format.first_line_indent = Inches(0)
p_title_en.paragraph_format.space_after = Pt(20)
set_run(p_title_en.add_run(
    'Measuring Income-based Residential Segregation of\n'
    'Low-income Elderly Living Alone in Seoul'),
    size=11, italic=True)

# ─────────────────────────────────────────────────────
# ABSTRACT
# ─────────────────────────────────────────────────────
h1('초록 (Abstract)')

body(
    '저소득 독거노인은 소득 제약으로 거주지 선택권이 극도로 제한된 집단으로, '
    '특정 지역에 집중되는 거주지 분리가 심화될 경우 사회적 고립과 복지 서비스 '
    '접근성 불평등으로 이어질 수 있다.'
)
body(
    '본 연구는 서울시 424개 행정동을 대상으로 공간 가중 고립지수(Isolation Index)와 '
    '상이지수(Dissimilarity Index)를 적용하여 저소득 독거노인의 거주지 분리를 '
    '노출 축과 균등성 축에서 정량적으로 측정하였다.'
)
body(
    '분석 결과, 전역 고립지수(I = 0.0174)는 무작위 기대값 대비 약 14% 높았으며, '
    '일반인구 대비 전역 상이지수(D = 0.1318)는 독거노인 내부 소득 분리(D = 0.0859)보다 '
    '크게 나타났다. 국지 분석에서는 강남·서초의 배제 극과 노원·강서의 집중 극이라는 '
    '뚜렷한 양극 분화가 확인되었다.'
)
body(
    '이 결과는 소득 제약이 거주지 정렬을 통해 공간적 불평등으로 번역됨을 실증하며, '
    '행정동 단위의 미시적 복지 개입 근거를 제공한다.'
)

# ─────────────────────────────────────────────────────
# 1. INTRODUCTION
# ─────────────────────────────────────────────────────
h1('1. 서론 (Introduction)')

body(
    '한국의 독거노인 가구는 전체 고령자 가구의 36.1%로 가장 높은 비율을 차지하며, '
    '이 중 상당수가 기초생활수급 대상인 저소득 계층이다(⚠️ 통계청/보건복지부). '
    '거주지 정렬(residential sorting) 이론은 소득 수준에 따라 개인·가구가 특정 지역으로 '
    '집중하거나 배제되는 현상이 구조적으로 발생함을 설명한다(Tiebout 1956; '
    'Reardon & Bischoff 2011). 저소득 독거노인은 낮은 소득(주거비 부담)과 독거(사회적 '
    '관계망 부재)라는 이중 제약을 동시에 가지므로, 거주지 정렬 메커니즘이 이 집단에게 '
    '특히 강하게 작동할 수 있다. 이러한 분리가 심화될 경우 사회적 교류 기회의 감소와 '
    '복지 서비스 접근성 불평등으로 이어지며, 결국 저소득 독거노인의 거주지 분리는 '
    '단순한 공간 패턴이 아닌 복지 불평등의 공간적 발현으로 이해될 필요가 있다.'
)
body(
    '저소득 독거노인의 공간적 분포에 관한 기존 연구는 크게 두 방향으로 진행되어 왔다. '
    '첫째는 특정 집단이 어디에 집중되는지를 파악하는 공간 군집 분석으로, 노인 밀집 '
    '지역이나 고독사 위험 지역의 식별에 초점을 맞추어 왔다(⚠️ Ahn 2023; Lee 2011). '
    '둘째는 해당 분포에 영향을 미치는 요인을 규명하는 회귀 기반 접근으로, 주거비, '
    '교통 접근성 등의 변수와 노인 집중도의 관계를 분석하였다. 한편 거주지 분리의 '
    '측정 체계는 Massey & Denton(1988)의 5차원 분류에서 출발하여, '
    "Reardon & O'Sullivan(2004)이 균등성(evenness)과 노출(exposure)의 두 축으로 "
    '압축·정립하였으며, Feitosa(2007)는 이를 가우시안 커널 기반의 공간 가중 측도로 '
    '확장하여 행정 경계 의존성을 완화하였다.'
)
body(
    '그러나 기존 연구들은 두 가지 중요한 갭을 남긴다. 첫째, 국내 독거노인 연구는 '
    '분포의 군집 여부나 영향 요인 분석에 집중할 뿐, 집단 간 공간적 분리의 정도 자체를 '
    '정량적으로 측정하지 않았다. 특정 지역에 누가 얼마나 사는지는 알아도, 저소득 '
    '독거노인이 나머지 인구로부터 얼마나, 어떤 방식으로 분리되어 있는지는 측정된 '
    '바 없다. 둘째, 거주지 분리 측도를 적용한 연구들은 대부분 인종·민족 분리 맥락에 '
    '한정되어 있으며, 노출 축(고립지수)으로 저소득 독거노인의 사회적 고립 정도를 '
    '측정한 연구와, 균등성 축(상이지수)으로 독거노인 집단 내 소득 기반 분리를 '
    '측정한 연구는 국내외를 통틀어 확인되지 않는다.'
)
body(
    '최근 행정동 단위의 저소득 독거노인 및 일반 인구 구성 데이터의 가용성이 확보되었으며, '
    '공간 가중 분리 측도를 계산할 수 있는 방법론적 도구의 발전으로 기존 연구의 '
    '갭을 실증적으로 채울 조건이 갖춰졌다.'
)
body(
    '이에 본 연구는 서울시 424개 행정동을 분석 단위로, 저소득 독거노인의 소득 기반 '
    '거주지 분리를 노출 축(공간 고립지수)과 균등성 축(공간 상이지수)으로 측정하고자 '
    '한다. 구체적인 연구 질문은 다음과 같다: "서울시 저소득 독거노인은 사회 전반으로부터 '
    '공간적으로 고립되어 있으며, 독거노인 집단 내부에서 소득 수준에 따라 거주지 분리가 '
    '존재하는가?" 이하 본 논문은 2장에서 선행연구를 검토하고, 3장에서 분석 방법론을 '
    '기술하며, 4장에서 결과 및 논의를 제시하고, 5장에서 결론을 도출한다.'
)

# ─────────────────────────────────────────────────────
# 2. LITERATURE REVIEW
# ─────────────────────────────────────────────────────
h1('2. 선행연구')
h2('2.1 거주지 분리의 이론적 배경 및 측도 체계')

body(
    '거주지 정렬(residential sorting)은 개인·가구가 소득과 제약 조건에 따라 주거지를 '
    '선택하고 그 선택이 누적되면서 도시 내 사회 집단이 공간적으로 분화되는 현상이다. '
    'Tiebout(1956)은 소득이 높을수록 양질의 공공서비스를 제공하는 지역으로 이동함을 '
    '설명하였으며, Schelling(1971)은 개인의 작은 선호가 누적될 경우 예상보다 큰 분리가 '
    '발생함을 보였다. Reardon & Bischoff(2011)는 이 과정을 부유층의 배타적 입지와 '
    '저소득층의 선택지 제약이라는 두 경로로 구체화하였다.'
)
body(
    '거주지 분리의 측정 체계는 Massey & Denton(1988)의 5차원 분류에서 출발하여, '
    "Reardon & O'Sullivan(2004)이 균등성(evenness)과 노출(exposure)의 두 축으로 "
    '압축·정립하였다. Feitosa(2007)는 이를 가우시안 커널 기반의 공간 가중 측도로 '
    '확장하여 행정 경계의 자의성과 MAUP 문제를 완화하였으며, 본 연구에서 사용하는 '
    '공간 상이지수(D)와 공간 고립지수(I)의 이론적 기반을 제공한다.'
)

h2('2.2 소득 기반 거주지 분리 연구 (균등성 축)')

body(
    '소득 기반 거주지 분리는 균등성 축의 대표적 연구 대상으로, ⚠️[Cartone et al.] 등이 '
    '상이지수를 활용하여 소득 집단 간 공간적 불균등 분포를 실증한 바 있다. 이 연구들은 '
    '소득 불평등이 심화될수록 거주지 분리가 강화되는 패턴을 일관되게 확인하였다.'
)
body(
    '국내에서는 저소득층의 주거 분포나 영향 요인을 다룬 연구(⚠️ Ahn 2023; Lee 2011)가 '
    '존재하나, 독거노인 가구를 소득 집단으로 구분하여 균등성 축의 분리 측도를 '
    '적용한 연구는 확인되지 않는다.'
)

h2('2.3 저소득 노인의 공간적 고립 연구 (노출 축)')

body(
    '노인의 사회적 고립에 관한 연구는 주로 설문 기반의 개인 척도(외로움, 사회적 관계망)에 '
    '집중되어 왔으며(⚠️ 관련 문헌 확인 필요), 집단 수준에서 공간적 고립을 측정한 '
    '연구는 제한적이다. 공간적 맥락에서의 노출 지수 연구는 인종·민족 분리 문헌에서 '
    '주로 발전하였다(⚠️ Qin et al. 확인 필요).'
)
body(
    '저소득 독거노인을 관심 집단으로 설정하고, 전체 주민 대비 공간적 고립의 정도를 '
    '노출 축(고립지수)으로 측정한 연구는 국내외를 통틀어 확인되지 않는다.'
)

# ─────────────────────────────────────────────────────
# 3. METHODS
# ─────────────────────────────────────────────────────
h1('3. 연구 방법')
h2('3.1 연구 대상 지역')

body(
    '본 연구의 공간적 분석 단위는 서울특별시 424개 행정동이다. 행정동은 복지 서비스 '
    '전달의 기본 단위로, 저소득 독거노인 인구 데이터가 이 수준에서 공개되어 있다. '
    '공간 분리 측도 계산 시 미터 단위의 거리 계산이 필요하므로, 좌표계는 한국 측지계 '
    '기반의 EPSG:5179(Korea 2000 Unified CS)를 적용하였다.'
)
fig_placeholder(1, '서울시 424개 행정동 경계 및 저소득 독거노인 분포')

h2('3.2 데이터')

body(
    '분석 대상인 저소득 독거노인은 기초생활보장 수급자 중 65세 이상 1인 가구로 '
    '정의하였다. 비교 집단은 두 가지로 설정하였다. 첫째는 전체 주민으로, 고립지수(H1) '
    '계산 시 저소득 독거노인이 사회 전반과 얼마나 분리되어 있는지를 측정하기 위함이다. '
    '둘째는 일반소득 독거노인(비수급 독거노인)으로, 상이지수(H2-2) 계산 시 동일 가구 '
    '형태 내 소득 효과만을 분리하기 위함이다.'
)
body(
    '행정동별 독거노인 현황(저소득·일반)은 서울 열린데이터광장(2024)에서 수집하였으며, '
    '행정동별 주민등록인구(연령별)는 행정안전부(2024)에서 취득하였다. '
    '행정동 경계 shapefile은 SGIS를 통해 확보하였으며, '
    '공간 분리 측도 계산을 위한 Gaussian 커널 가중치 적용 시 활용하였다.'
)

# Table 1
tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp.paragraph_format.first_line_indent = Inches(0)
tp.paragraph_format.space_before = Pt(8)
tp.paragraph_format.space_after = Pt(4)
set_run(tp.add_run('Table 1. 데이터 목록 및 출처.'), size=CAP_PT)

tbl = doc.add_table(rows=5, cols=4)
tbl.style = 'Table Grid'
hdr_cells = tbl.rows[0].cells
for i, h_text in enumerate(['데이터명', '출처', '연도', '용도']):
    hdr_cells[i].text = h_text
    for run in hdr_cells[i].paragraphs[0].runs:
        run.bold = True
        run.font.name = FONT
        run.font.size = Pt(CAP_PT)

tbl_data = [
    ('행정동별 독거노인 현황', '서울 열린데이터광장', '2024', '저소득·일반소득 독거노인 인구'),
    ('주민등록인구(연령별)', '행정안전부', '2024', '전체 주민 인구'),
    ('행정동 경계 shapefile', 'SGIS', '2024', '공간 단위 설정 및 거리 계산'),
    ('좌표계 변환', 'EPSG:5179', '-', 'Korea 2000 Unified CS 적용'),
]
for ri, row_data in enumerate(tbl_data):
    for ci, val in enumerate(row_data):
        cell = tbl.rows[ri + 1].cells[ci]
        cell.text = val
        for run in cell.paragraphs[0].runs:
            run.font.name = FONT
            run.font.size = Pt(CAP_PT)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

h2('3.3 분석 방법')
h3('3.3.1 공간 가중치 설정')

body(
    '행정동 경계를 분석 단위로 직접 사용할 경우 경계 획정 방식에 따라 결과가 달라지는 '
    '수정 가능 면적 단위 문제(MAUP)가 발생한다. 이를 완화하기 위해 Feitosa(2007)가 '
    '제안한 가우시안 커널 기반 공간 가중치를 적용하였다.'
)
body(
    '공간 가중치는 식 (1)의 가우시안 커널 함수로 정의되며, 거리가 멀수록 가중치가 '
    '지수적으로 감소한다. Bandwidth는 1,500m로 설정하였는데, 이는 서울시 행정동의 '
    '평균 반경(약 1km)을 고려한 값으로, 근린 수준의 공간 상호작용을 포착하기에 적합하다.'
)
eq_text(1, 'W_ij = exp(−d²_ij / 2·bw²)')

h3('3.3.2 공간 고립지수 (Spatial Isolation Index)')

body(
    '공간 고립지수(I)는 저소득 독거노인이 자신의 근린에서 같은 집단을 만날 확률로 '
    '정의되며(Bell 1954; Feitosa 2007), 노출(exposure) 차원의 분리를 측정한다. '
    '노출형 지수는 집단의 규모에 민감하므로, 절대값보다 무작위 기대값(π: 저소득 '
    '독거노인의 도시 전체 비율)과의 비교를 통해 해석해야 한다.'
)
eq_text(2, 'I = Σᵢ (aᵢ/A) · (L̃ᵃᵢ / L̃ᵗᵢ)')

h3('3.3.3 공간 상이지수 (Spatial Dissimilarity Index)')

body(
    '공간 상이지수(D)는 도시 전체에서 두 집단이 얼마나 불균등하게 분포하는지를 나타내는 '
    '균등성(evenness) 차원의 측도이다(Duncan & Duncan 1955; Feitosa 2007). 국지 '
    '상이지수(Local D)의 부호는 방향을 나타내는데, 양수는 저소득 독거노인이 지역 '
    '평균보다 과대 집중된 상태를, 음수는 상대적으로 배제된 상태를 의미한다.'
)
eq_text(3, 'D = Σᵢ (tᵢ/T) · |p̃ᵢ − P_A| / (2·P_A·P_B)')

h3('3.3.4 민감도 분석 (NSI + Bandwidth 비교)')

body(
    'Bandwidth 설정의 임의성을 검증하기 위해 500m부터 4,500m까지 다양한 bandwidth에서 '
    '전역 상이지수를 반복 계산하고 그 안정성을 확인하였다.'
)
body(
    '또한 근린 정렬 지수(NSI)를 통해 분리가 집중되는 공간적 스케일을 파악하였다. '
    'NSI는 각 bandwidth에서의 국지 분리 정도가 전체 분산에서 차지하는 비율로, '
    '값이 높은 bandwidth가 분리의 주요 스케일임을 나타낸다.'
)
eq_text(4, 'NSI = [Σᵢ L̃ᵗᵢ(X̂ᵢ − X̄)² / Σᵢ L̃ᵗᵢ] / S²_total')
fig_placeholder(2, 'Bandwidth별 전역 D값 변화 및 NSI 곡선')

# ─────────────────────────────────────────────────────
# 4. RESULTS AND DISCUSSION
# ─────────────────────────────────────────────────────
h1('4. 결과 및 논의')
h2('4.1 저소득 독거노인의 공간적 고립 (H1: 노출 축)')

body(
    '전역 고립지수는 I = 0.0174로 나타났다. 이는 절대적으로 낮은 수준이나, 저소득 '
    '독거노인의 도시 전체 비율(π ≈ 1.52%)로 계산되는 무작위 기대값 대비 약 14% 높은 '
    '값으로, 집단 규모를 고려하면 통계적으로 유의미한 고립 수준이다(Feitosa 2007).'
)
body(
    '국지 고립지수 지도에서는 강서구, 노원구, 강남구 수서동에 국지 고립이 응축되어 '
    '나타났다.'
)
fig_placeholder(3, 'Local I 단계구분도 (서울시 424개 행정동)')
body(
    '강남·서초구에서는 Local I가 낮게 나타나는데, 이는 해당 지역의 저소득 독거노인 '
    '절대 수가 서울 평균 대비 극단적으로 적어 집단 내 접촉 확률 자체가 수학적으로 '
    '낮을 수밖에 없기 때문이다. 고립지수는 정의상 집단의 절대적 규모에 민감하게 '
    '반응하므로, 저소득 독거노인이 거의 없는 지역에서는 낮은 값이 필연적이다. '
    '그러나 그 소수의 저소득 독거노인이 전체 인구 구성에서 차지하는 비율은 서울 '
    '평균에서 크게 이탈해 있어, Local D는 음수(배제 극)로 나타난다. 이는 고립지수와 '
    '상이지수가 독립적으로 움직이는 분리의 두 차원임을 실증적으로 보여주며, 두 지수를 '
    '동시에 적용해야만 분리의 전체 구조를 정확히 파악할 수 있음을 시사한다.'
)
body(
    '이 결과는 소득 제약이 거주지 선택을 제한하여 특정 지역으로의 집중을 낳는다는 '
    'Reardon & Bischoff(2011)의 메커니즘과 일치하며, 강서·노원 지역의 저렴 주택 '
    '집중이 고립 응축의 공간적 기반임을 시사한다.'
)

h2('4.2 일반인구 대비 거주지 분리 (H2-1: 균등성 축)')

body(
    '일반인구 대비 전역 상이지수는 D = 0.1318로 나타났다. 이는 독거노인 내부 소득 '
    '분리(D = 0.0859, 4.3절)보다 높은 값으로, 독거 여부 자체보다 소득·고령·1인 가구 '
    '특성이 복합적으로 결합된 분리가 더 강하게 작동함을 의미한다.'
)
body(
    'Local D 지도에서는 노원·강서구(양수, 집중 극)와 강남·서초구(음수, 배제 극)의 '
    '뚜렷한 양극 분화가 관찰되었다. 양수 지역은 저소득 독거노인이 지역 인구 구성 '
    '비율 기준으로 서울 평균보다 과대 집중된 공간이며, 음수 지역은 저소득 독거노인이 '
    '해당 지역 인구 비율 대비 현저히 낮아 상대적으로 배제된 공간을 의미한다. 이 '
    '양극 분화는 서울의 주거비 지리를 반영하는데, 저렴한 임대주택이 밀집한 노원·강서 '
    '지역은 소득 제약이 강한 집단의 집중처가 되는 반면, 고가 주거지인 강남·서초에서는 '
    '이 집단이 구조적으로 배제된다.'
)
fig_placeholder(4, 'Local D 단계구분도 — H2-1 (양수: 집중 극, 음수: 배제 극)')
body(
    '이 양극 패턴은 Reardon & Bischoff(2011)가 제시한 소득 기반 거주지 분리의 두 경로, '
    '즉 부유층의 배타적 입지 선택(강남·서초의 배제 패턴)과 저소득층의 선택지 제약(노원·'
    '강서의 집중 패턴)과 정확히 대응된다. 이는 저소득 독거노인의 거주지 분리가 개인의 '
    '선택이 아닌 주거비 장벽에 의한 구조적 배제의 결과임을 시사하며, 단순한 공간 집중을 '
    '넘어 사회경제적 불평등이 공간적으로 가시화된 것으로 해석된다.'
)

h2('4.3 독거노인 내부 소득 기반 분리 (H2-2: 균등성 축)')

body(
    '독거노인 내부 소득 기반 전역 상이지수는 D = 0.0859로, H2-1(0.1318)보다 낮게 '
    '나타났다. 이는 서울 전역 평균으로는 독거노인 집단 안에서 소득에 따른 추가적 '
    '분리 효과가 상대적으로 제한적임을 시사한다.'
)
body(
    '그러나 국지 수준에서는 강남·서초(배제 극) vs 노원·강서(집중 극)의 양극 분화가 '
    'H2-1과 유사한 패턴으로 관찰되었다. Schelling(1971)의 누적 선호 이론에 따르면, '
    '집단 구성원 각각의 작은 선호(비슷한 소득 집단과의 거주 선호)가 반복·누적될 경우 '
    '전역 평균이 낮더라도 국지적으로 큰 분화가 발생할 수 있다. 전역 D = 0.0859라는 '
    '수치는 서울 전체 평균에서 독거 형태 안에서의 소득 분리가 상대적으로 제한적임을 '
    '보여주지만, 국지 수준의 양극 분화는 이 분리가 특정 지역에 집중되어 있음을 나타낸다.'
)
fig_placeholder(5, 'Local D 단계구분도 — H2-2 (독거노인 내부 소득 분리)')
body(
    '소득이 같은 가구 유형(독거노인) 안에서도 거주지를 분화시키는 독립 요인으로 '
    '작용한다는 점은 중요한 함의를 갖는다. 이는 독거라는 가구 형태를 통제한 후에도 '
    '소득이 거주지 정렬을 설명하는 독립 변수로 남음을 의미하며, 소득 기반 거주지 '
    '정렬이 가구 유형의 효과를 초월하여 작동하는 보편적 메커니즘임을 실증한다'
    '(⚠️ Cartone et al. 소득 기반 분리 연구와 연결).'
)

h2('4.4 민감도 분석')

body(
    'Bandwidth를 500m에서 4,500m로 변화시켜도 전역 상이지수는 안정적으로 유지되었으며, '
    '이는 분석 결과가 bandwidth 선택에 강건(robust)함을 나타낸다.'
)
body(
    'NSI 분석 결과, 분리는 약 1,500m 근린 스케일에 집중되어 있었다. 이는 자치구 단위의 '
    '광역 정책보다 행정동 단위의 미시적 복지 개입이 더 효율적임을 시사한다.'
)
fig_placeholder(6, 'Bandwidth별 전역 D값 변화 그래프 및 NSI 곡선')

# ─────────────────────────────────────────────────────
# 5. CONCLUSION
# ─────────────────────────────────────────────────────
h1('5. 결론')

body(
    '본 연구는 서울시 424개 행정동을 대상으로 저소득 독거노인의 거주지 분리를 '
    '노출 축(공간 고립지수)과 균등성 축(공간 상이지수)으로 측정하였다. 전역 고립지수 '
    'I = 0.0174는 무작위 기대값(π = 0.0152) 대비 약 14% 높아 약한 수준의 고립이 '
    '확인되었으며, 국지적으로는 강서·노원 지역에 고립이 응축되었다. 일반인구 대비 '
    '전역 상이지수(D = 0.1318)는 독거노인 내부 소득 분리(D = 0.0859)보다 높게 '
    '나타나, 소득·고령·1인 가구 특성이 결합된 분리가 더 강하게 작동함을 확인하였다. '
    '이 패턴은 거주지 정렬 이론이 예측하는 부유층 배타 입지(강남·서초 배제)와 '
    '저소득층 선택지 제약(노원·강서 집중)의 두 경로와 정확히 대응된다.'
)
body(
    '본 연구는 세 가지 한계를 갖는다. 첫째, 2024년 단일 시점 데이터로 분리의 '
    '시계열적 변화를 파악하기 어렵다. 둘째, 기초생활수급자를 저소득 독거노인의 '
    '대리 변수로 사용하였으나, 차상위계층 등 수급 기준 외 실질적 빈곤층이 누락될 '
    '수 있다. 셋째, 노인여가복지시설의 공급 분포와의 교차 분석은 분리(수요 측) '
    '연구의 범위를 벗어나는 공급 측 문제로 본 연구에서 제외하였다.'
)
body(
    '향후 연구는 다음 방향으로 확장될 수 있다. 다년도 데이터를 활용한 분리 추이 분석으로 '
    '정렬의 동태적 과정을 추적하거나, 차상위계층을 포함한 확장된 저소득 정의를 적용할 '
    '수 있다. 또한 본 연구에서 식별된 고위험 행정동을 대상으로 복지시설 접근성 분석을 '
    '연계하여 수요-공급 불일치 구조를 분석하는 것도 유망한 연구 방향이다.'
)

out = '/Users/jin/홍교수님 수업/RSG_final/논문_초안_v1.docx'
doc.save(out)
print(f'완료: {out}')

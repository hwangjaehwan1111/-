import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
import textwrap
import matplotlib.font_manager as fm
import traceback
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
import time

# --- 1. 환경 및 폰트 설정 ---
font_path = "NanumSquareRoundB.ttf"
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
else:
    plt.rcParams['font.family'] = 'NanumGothic' if os.name != 'nt' else 'Malgun Gothic'

plt.rcParams['axes.unicode_minus'] = False
  
COLOR_NAVY = '#1F4E3D'; COLOR_RED = '#D97706'; COLOR_STUDENT = '#2F855A'
COLOR_AVG = '#9CA3AF'; COLOR_GRID = '#E5E7EB'; COLOR_BG = '#F9FAFB'
COLOR_CONCEPT = '#3B82F6' 
COLOR_APP = '#EF4444'     

# --- 2. 구글 스프레드시트 연동 ---
@st.cache_resource
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = None
    
    if "GOOGLE_JSON" in os.environ:
        try:
            creds_dict = json.loads(os.environ["GOOGLE_JSON"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except: pass
        
    if creds is None and os.path.exists("secrets.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
        
    if creds is None:
        try:
            if "GOOGLE_JSON" in st.secrets:
                creds_dict = json.loads(st.secrets["GOOGLE_JSON"])
                creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except: pass

    if creds is None:
        st.error("구글 인증 정보가 없습니다.")
        st.stop()
            
    client = gspread.authorize(creds)
    
    # 🌟 원장님의 찐 구글 시트 주소 🌟
    doc = client.open_by_url("https://https://docs.google.com/spreadsheets/d/1pFj7C3uv1Q7PffHsN8Spg2PmHbc1hSKYDgdtqT7rRDk/edit?gid=0#gid=0")
    return doc

@st.cache_data(ttl=60)
def fetch_all_dataframes():
    doc = get_google_sheet()
    try:
        df_info = pd.DataFrame(doc.worksheet('Hakryeok_Info').get_all_records())
        df_results = pd.DataFrame(doc.worksheet('Hakryeok_Results').get_all_records())
    except Exception as e:
        st.error("구글 시트에 'Hakryeok_Info'와 'Hakryeok_Results' 시트가 존재하는지 확인해주세요!")
        st.stop()
        
    df_results.rename(columns=lambda x: str(x).strip(), inplace=True)
    df_info.rename(columns=lambda x: str(x).strip(), inplace=True)
    
    for col in ['시험명', '이름', '학교', '학년']:
        if col in df_results.columns:
            df_results[col] = df_results[col].astype(str).str.strip()
            
    if '시험명' in df_info.columns:
        df_info['시험명'] = df_info['시험명'].astype(str).str.strip()

    df_results = df_results.replace('', 0).replace('nan', 0).fillna(0)
    return df_info, df_results

def load_data():
    doc = get_google_sheet()
    ws_info = doc.worksheet('Hakryeok_Info')
    ws_results = doc.worksheet('Hakryeok_Results')
    df_info, df_results = fetch_all_dataframes()
    return doc, ws_info, ws_results, df_info, df_results

# --- 3. 학력평가 (개념 vs 응용) PDF 생성 함수 ---
def generate_hakryeok_report(target_name, selected_test):
    try:
        _, _, _, df_info_all, df_results_all = load_data()
        df_info = df_info_all[df_info_all['시험명'] == selected_test.strip()].copy()
        df_results = df_results_all[df_results_all['시험명'] == selected_test.strip()].copy()
        
        df_results.columns = df_results.columns.astype(str)
        df_info['문항번호'] = df_info['문항번호'].astype(str).str.strip()
        df_info = df_info[df_info['문항번호'] != '']
        if '배점' not in df_info.columns: df_info['배점'] = 1
        
        # 🌟 [스마트 파트 자동 할당 로직: 중등 45문항 / 고등 35문항] 🌟
        if '파트' not in df_info.columns:
            df_info['파트'] = ''
            
        total_qs = len(df_info)
        def assign_part(row):
            current_part = str(row['파트']).strip()
            # 시트에 이미 '개념'이나 '응용'이라고 잘 적어두셨다면 무조건 그걸 따릅니다.
            if current_part != '' and current_part != 'nan': 
                return current_part 
                
            try:
                q_num = int(row['문항번호'])
                if total_qs == 45: # 중등부
                    return '개념' if q_num <= 25 else '응용'
                elif total_qs == 35: # 고등부
                    return '개념' if q_num <= 20 else '응용'
                else: # 그 외의 경우 기본값
                    return '개념' if q_num <= 20 else '응용'
            except:
                return '개념'
                
        df_info['파트'] = df_info.apply(assign_part, axis=1)
        
        q_cols = [str(q) for q in df_info['문항번호']]
        valid_cols = [col for col in df_results.columns if col in q_cols]
        
        def safe_to_int(val):
            try: return int(float(val))
            except: return 0
            
        df_scores = df_results[valid_cols].applymap(safe_to_int)
        avg_per_q = df_scores.mean()
        
        total_analysis = df_info.copy()
        total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0)) * total_analysis['배점']
        
        total_concept = total_analysis[total_analysis['파트'] == '개념']
        total_app = total_analysis[total_analysis['파트'] == '응용']
        
        avg_concept_score = int((total_concept['평균득점'].sum() / total_concept['배점'].sum() * 100)) if total_concept['배점'].sum() > 0 else 0
        avg_app_score = int((total_app['평균득점'].sum() / total_app['배점'].sum() * 100)) if total_app['배점'].sum() > 0 else 0
        
        student_found = False
        pdf_buffer = io.BytesIO()

        with PdfPages(pdf_buffer) as pdf:
            target_df = df_results if target_name == "전체" else df_results[df_results['이름'] == str(target_name).strip()]

            for _, s_row in target_df.iterrows():
                student_name = str(s_row.get('이름', '')).strip()
                if not student_name or student_name == '0': continue
                
                student_found = True
                student_grade = s_row.get('학년', '')
                
                analysis = df_info.copy()
                analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
                analysis['득점'] = analysis['정답여부'] * analysis['배점']
                
                if analysis['득점'].sum() == 0 and target_name != "전체": continue
                
                s_concept = analysis[analysis['파트'] == '개념']
                s_app = analysis[analysis['파트'] == '응용']
                
                score_c = int((s_concept['득점'].sum() / s_concept['배점'].sum() * 100)) if s_concept['배점'].sum() > 0 else 0
                score_a = int((s_app['득점'].sum() / s_app['배점'].sum() * 100)) if s_app['배점'].sum() > 0 else 0
                total_score = int((analysis['득점'].sum() / analysis['배점'].sum() * 100)) if analysis['배점'].sum() > 0 else 0
                
                fig = plt.figure(figsize=(8.27, 11.69))
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_NAVY, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)
                
                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    logo_ax = fig.add_axes([0.80, 0.915, 0.15, 0.045], zorder=15)
                    logo_ax.imshow(logo_img); logo_ax.axis('off')
                
                fig.text(0.15, 0.88, 'JEET', fontsize=32, fontweight='bold', color='red', ha='left')
                fig.text(0.28, 0.88, '학력평가 심층 분석 리포트', fontsize=30, fontweight='bold', color=COLOR_NAVY, ha='left')
                
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}  |  과정: {selected_test}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=14, fontweight='bold', color='#222')
                
                fig.text(0.25, 0.77, f"총점 성취도: {total_score}%", ha='center', fontsize=15, fontweight='bold', color=COLOR_NAVY)
                fig.text(0.5, 0.77, f"개념 성취도: {score_c}%", ha='center', fontsize=15, fontweight='bold', color=COLOR_CONCEPT)
                fig.text(0.75, 0.77, f"응용 성취도: {score_a}%", ha='center', fontsize=15, fontweight='bold', color=COLOR_APP)
                
                # --- 4분면 매트릭스 ---
                ax1 = fig.add_axes([0.15, 0.47, 0.35, 0.25])
                ax1.set_xlim(0, 100); ax1.set_ylim(0, 100)
                
                ax1.axhline(60, color=COLOR_GRID, linestyle='--', linewidth=1.5)
                ax1.axvline(60, color=COLOR_GRID, linestyle='--', linewidth=1.5)
                
                ax1.text(97, 97, "최상위권\n(개념+응용 완벽)", ha='right', va='top', fontsize=10, fontweight='bold', color=COLOR_NAVY, alpha=0.5)
                ax1.text(3, 97, "실전 감각 우수\n(개념 보완 필요)", ha='left', va='top', fontsize=10, fontweight='bold', color=COLOR_NAVY, alpha=0.5)
                ax1.text(3, 3, "기초 다지기\n(절대적 보완 필요)", ha='left', va='bottom', fontsize=10, fontweight='bold', color=COLOR_NAVY, alpha=0.5)
                ax1.text(97, 3, "개념 탄탄형\n(응용 훈련 필요)", ha='right', va='bottom', fontsize=10, fontweight='bold', color=COLOR_NAVY, alpha=0.5)
                
                ax1.text(57, 57, "도약 준비\n(기본기 안착)", ha='right', va='top', fontsize=10, fontweight='bold', color=COLOR_NAVY, alpha=0.5)
                
                ax1.scatter(avg_concept_score, avg_app_score, s=100, color=COLOR_AVG, marker='X', label='전체 평균', zorder=4)
                ax1.scatter(score_c, score_a, s=250, color=COLOR_RED, marker='*', edgecolor='white', linewidth=1.5, label='학생 위치', zorder=5, clip_on=False)
                
                ax1.set_xlabel("개념 이해 성취도 (%)", fontsize=11, fontweight='bold', color=COLOR_CONCEPT)
                ax1.set_ylabel("응용/심화 성취도 (%)", fontsize=11, fontweight='bold', color=COLOR_APP)
                
                ax1.legend(loc='lower center', bbox_to_anchor=(0.5, 1.02), ncol=2, frameon=True, fontsize=9)
                
                # --- 단원별 그래프 ---
                ax2 = fig.add_axes([0.62, 0.47, 0.28, 0.25])
                
                ordered_units = analysis['단원'].drop_duplicates().tolist()
                ordered_units.reverse() 
                
                unit_groups = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'}).reindex(ordered_units)
                units = unit_groups.index.tolist()
                y_pos = np.arange(len(units))
                
                unit_percents = (unit_groups['득점'] / unit_groups['배점'] * 100).fillna(0).values
                
                ax2.barh(y_pos, unit_percents, height=0.5, color=COLOR_STUDENT, alpha=0.8)
                ax2.set_yticks(y_pos)
                ax2.set_yticklabels([textwrap.fill(u, 6) for u in units], fontsize=9, fontweight='bold', color=COLOR_NAVY)
                ax2.set_xlim(0, 110)
                ax2.set_xlabel("단원별 성취도 (%)", fontsize=10, fontweight='bold')
                
                for i, v in enumerate(unit_percents):
                    ax2.text(v + 2, i, f"{int(v)}%", va='center', fontsize=9, fontweight='bold', color=COLOR_STUDENT)
                    
                ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
                
                # --- 🌟 종합 진단 🌟 ---
                rect_diag = plt.Rectangle((0.08, 0.11), 0.84, 0.28, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                
                fig.text(0.11, 0.36, f"▶ {student_name} 학생 [개념 vs 응용] 심층 분석", fontsize=14, fontweight='bold', color=COLOR_NAVY)
                
                if score_c >= 60 and score_a >= 60:
                    tier_title = "[1사분면] 최상위권 (개념 완벽 & 응용 탁월)"
                    diag_text = "개념과 응용 두 마리 토끼를 모두 잡은 최상위권의 면모를 보여주고 있습니다. 흔들림 없는 기본기를 바탕으로 고난도 심화 문제까지 막힘없이 해결하는 훌륭한 학업 밸런스를 갖추었습니다."
                    sol_text = "지금처럼 훌륭한 학습 패턴을 칭찬해 주세요. 이제는 익숙한 유형을 넘어, 여러 단원의 개념이 복합적으로 융합된 킬러 문항과 사고력 문제에 기분 좋게 도전하며 수학적 시야를 더욱 넓혀가면 좋겠습니다."
                
                elif score_c < 60 and score_a >= 60:
                    if score_c >= 50:
                        tier_title = "[2사분면] 실전 우수형 (응용 탁월 & 개념 완성 단계)"
                        # 🌟 '준수하며' -> '기본 틀을 갖춰가고 있으며'로 수정
                        diag_text = "응용문제를 훌륭하게 풀어내는 뛰어난 수학적 감각을 지녔습니다. 개념 이해도 역시 기본 틀을 갖춰가고 있으며, 교과 기본기의 빈틈만 살짝 채워주면 완벽한 1사분면(최상위권)이 될 수 있는 유망한 케이스입니다."
                        sol_text = "심화 문제를 푸는 능력은 이미 훌륭하게 갖춰져 있습니다! 가끔 발생하는 실수나 헷갈리는 필수 개념들만 오답 노트를 통해 꼼꼼하게 다듬어주면 점수가 더욱 수직 상승할 것입니다."
                    else:
                        tier_title = "[2사분면] 실전 우수형 (응용 탁월 & 개념 보완 요망)"
                        diag_text = "응용문제는 번뜩이는 센스로 잘 풀어내지만, 오히려 쉬운 개념 문항이나 기본 연산 과정에서 아까운 실수가 발생하는 실전형 우수 케이스입니다."
                        sol_text = "문제 해결력은 아주 뛰어나지만, 간혹 눈으로만 풀거나 직관에 의존하려는 경향이 있을 수 있습니다. 주요 공식을 빈 종이에 스스로 꼼꼼하게 적어보는 연습을 더해주면 훨씬 더 좋은 결과가 있을 것입니다."
                
                elif score_c >= 60 and score_a < 60:
                    if score_a >= 50:
                        tier_title = "[4사분면 상위] 개념 탄탄형 (응용 도약 단계)"
                        # 🌟 '준수한 성취도를 보이고 있어' -> '의미 있는 성장을 보이고 있어'로 수정
                        diag_text = "수학의 뼈대가 되는 개념이 매우 튼튼하게 잡혀 있습니다. 응용 문항에서도 점차 의미 있는 성장을 보이고 있어, 실전 응용력 또한 꾸준히 올라오고 있습니다."
                        sol_text = "기본기가 탄탄하므로 심화 문제에 대한 적응력도 금방 좋아질 것입니다. 낯설고 어려운 문제라고 지레 겁먹지 않고, 배운 개념들을 퍼즐 맞추듯 차근차근 적용해 보는 연습을 칭찬과 함께 지도하겠습니다."
                    else:
                        tier_title = "[4사분면 하위] 개념 탄탄형 (개념 완벽 & 응용 훈련 필요)"
                        diag_text = "기본적인 개념과 공식은 충실히 숙지하고 있으나, 조건이 복잡해지거나 낯선 형태로 변형된 응용문항에서 다소 아쉬운 성취를 보이고 있습니다. 배운 지식을 실전에 적용하는 연결 고리 훈련이 필요합니다."
                        sol_text = "문제 속 조건들을 끊어 읽으며 출제자의 의도를 파악하는 연습이 중요합니다. 해설지를 바로 보기보다, 문제의 설계도를 먼저 그려보는 심화 훈련을 집중적으로 진행하겠습니다."
                
                else:
                    if score_c >= 50 and score_a >= 50:
                        tier_title = "[3사분면 상위] 도약 준비 (기본기 안착 & 심화 훈련 시작)"
                        diag_text = "수학의 뼈대가 되는 기본기가 훌륭하게 안착된 상태입니다. 개념과 응용 모두 50점 이상을 기록하며 이제 한 단계 더 도약할 준비를 마쳤습니다."
                        sol_text = "자신감을 듬뿍 심어주세요! 튼튼한 기본기를 바탕으로 이제 조금 더 난이도 있는 심화 문제에 도전하며 실전 감각을 끌어올릴 황금 타이밍입니다."
                    elif score_c >= 50 and score_a < 50:
                        tier_title = "[3사분면 우측] 개념 안착형 (개념 훌륭, 응용 훈련 필요)"
                        # 🌟 '좋은 점수를 내며' -> '기본기를 다지며'로 수정
                        diag_text = "개념 부분에서 차근차근 기본기를 다지며 긍정적인 흐름을 보이고 있습니다. 다만, 아직 낯선 응용문제를 만났을 때 해결하는 실전 연습이 조금 더 필요합니다."
                        sol_text = "지금의 훌륭한 개념 이해도를 칭찬해 주세요! 배운 개념을 다양한 유형의 문제에 적용해 보는 연습을 집중적으로 진행하겠습니다. 틀려도 끝까지 풀어보는 끈기가 중요합니다."
                    elif score_c < 50 and score_a >= 50:
                        tier_title = "[3사분면 좌측] 실전 우수형 (응용력 훌륭, 개념 꼼꼼함 보완)"
                        diag_text = "전체 점수에 비해 응용문제를 풀어내는 감각과 직관력이 돋보입니다. 하지만 오히려 맞혀야 할 교과 기본 개념이나 연산에서 아쉬운 실수가 나오고 있습니다."
                        sol_text = "어려운 문제도 풀어낼 수 있는 잠재력이 훌륭합니다! 헷갈리는 필수 개념들을 다시 꼼꼼히 정리하고, 아는 문제에서 실수하지 않는 연습을 통해 점수를 크게 끌어올리겠습니다."
                    else:
                        tier_title = "[3사분면 하위] 기초 다지기 (기본 개념 & 연산 집중 보완)"
                        diag_text = "수학의 가장 기본이 되는 연산과 필수 개념을 한 번 더 단단하게 다져야 하는 시기입니다. 수학의 뼈대를 튼튼하게 세워가는 중요한 과정인 만큼, 아이가 위축되지 않도록 따뜻한 격려와 칭찬이 꼭 필요합니다."
                        sol_text = "지금 당장의 점수보다는 '매일 꾸준히 하는 습관'을 만들어 주는 것이 중요합니다. 예제 위주의 반복 학습을 통해 '나도 할 수 있다'는 성취감을 회복할 수 있도록 따뜻하게 이끌어주겠습니다."

                final_content = (
                    f"1. 현재 위치: {tier_title}\n\n"
                    f"2. 성취도 분석: {diag_text}\n\n"
                    f"3. JEET 전문가 솔루션: {sol_text}"
                )
                
                wrapped_lines = [textwrap.fill(p, width=52) for p in final_content.split('\n\n')]
                fig.text(0.11, 0.33, "\n\n".join(wrapped_lines), fontsize=10.5, linespacing=1.8, va='top', ha='left', color='#333')
                
                line_footer = plt.Line2D([0.05, 0.95], [0.09, 0.09], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure)
                fig.lines.append(line_footer)
                
                campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜령로 10")]
                for i, (name, addr) in enumerate(campuses):
                    fig.text([0.22, 0.50, 0.78][i], 0.065, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
                    fig.text([0.22, 0.50, 0.78][i], 0.045, addr, ha='center', fontsize=7.5, color='#555')
                
                pdf.savefig(fig); plt.close(fig)
            
        if not student_found: return False, None, "학생을 찾을 수 없습니다."
        pdf_buffer.seek(0)
        return True, pdf_buffer, "학력평가 리포트 생성 완료!"
    except Exception as e: return False, None, f"오류 발생: {traceback.format_exc()}"

# --- 4. Streamlit 웹 UI 구성 ---
st.set_page_config(page_title="JEET 학력평가 분석 시스템", layout="wide", page_icon="📝")
col1, col2 = st.columns([8, 2])
with col1: st.title("📝 JEET 학력평가 [개념/응용] 분석 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)

try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"구글 시트 로드 실패: {e}"); st.stop()

st.sidebar.header("📚 학력평가 선택")
test_list = df_info_all['시험명'].astype(str).str.strip().dropna().unique().tolist()
selected_test = st.sidebar.selectbox("평가 과정을 선택하세요:", test_list)

df_info_filtered = df_info_all[df_info_all['시험명'] == selected_test]

tab1, tab2 = st.tabs(["📝 신규 학력평가 성적 입력", "📑 개별 매트릭스 리포트 출력"])

with tab1:
    st.subheader(f"[{selected_test}] 성적 입력 (개념+응용)")
    
    if st.session_state.get('save_success', False):
        st.success("✅ 성적이 완벽하게 저장되었습니다! 이제 [리포트 출력] 탭의 명단에 즉시 나타납니다.")
        st.session_state['save_success'] = False
        
    question_numbers = [str(x) for x in df_info_filtered['문항번호'].tolist() if str(x).strip() != '']
    
    if question_numbers:
        with st.form("data_input_form", clear_on_submit=True):
            ci1, ci2, ci3 = st.columns(3)
            with ci1: input_name = st.text_input("이름")
            with ci2: input_school = st.text_input("학교")
            with ci3: input_grade = st.selectbox("학년", ["중1", "중2", "중3", "고1", "고2", "고3"])
            st.markdown("---")
            
            total_q = len(question_numbers)
            
            # 🌟 [스마트 안내 멘트: 중등/고등 감지] (물결표 취소선 에러 방지!) 🌟
            if total_q == 45:
                st.markdown("💡 **[중등부 감지 완료] 1번부터 25번은 개념문항, 26번부터 45번은 응용문항으로 자동 계산됩니다.**")
            elif total_q == 35:
                st.markdown("💡 **[고등부 감지 완료] 1번부터 20번은 개념문항, 21번부터 35번은 응용문항으로 자동 계산됩니다.**")
            else:
                st.markdown(f"💡 **총 {total_q}문항이 있습니다. (구글 시트 '파트' 열 기준 반영)**")
            
            answers = {}
            for i in range(0, len(question_numbers), 5):
                cols = st.columns(5)
                for j, q_num in enumerate(question_numbers[i:i+5]):
                    with cols[j]:
                        choice = st.radio(f"**{q_num}번**", options=["O", "X"], horizontal=True, key=f"q_{q_num}")
                        answers[str(q_num)] = 1 if choice == "O" else 0

            if st.form_submit_button("구글 시트에 성적 저장하기", type="primary"):
                clean_name = input_name.strip()
                if not clean_name: 
                    st.error("⚠ 이름을 입력해주세요.")
                else:
                    with st.spinner("구글 시트와 시스템을 완전히 동기화 중입니다. 잠시만 기다려주세요..."):
                        try:
                            header_row = ws_results.row_values(1)
                            new_row = []
                            for col_name in header_row:
                                col_str = str(col_name).strip()
                                if col_str == '시험명': new_row.append(selected_test) 
                                elif col_str == '이름': new_row.append(clean_name)
                                elif col_str == '학교': new_row.append(input_school)
                                elif col_str == '학년': new_row.append(input_grade)
                                elif col_str in answers: new_row.append(answers[col_str])
                                else: new_row.append("")
                            
                            ws_results.append_row(new_row)
                            last_row_index = len(ws_results.get_all_values())
                            
                            ws_results.format(f"A{last_row_index}:AZ{last_row_index}", {
                                "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.6}
                            })
                            
                            st.cache_data.clear()
                            time.sleep(1.5) 
                            st.session_state['save_success'] = True
                            
                            try:
                                st.rerun() 
                            except AttributeError:
                                st.experimental_rerun()
                                
                        except Exception as e: st.error(f"저장 중 오류: {e}")

with tab2:
    st.subheader(f"[{selected_test}] 매트릭스 분석 리포트 생성")
    
    raw_student_list = df_results_all[df_results_all['시험명'] == selected_test]['이름'].dropna().unique().tolist()
    clean_student_list = [str(name).strip() for name in raw_student_list if str(name).strip() not in ['', '0', 'nan', 'None']]
    clean_student_list = sorted(list(set(clean_student_list)))
    
    target_student = st.selectbox("출력할 학생을 선택하세요:", ["선택하세요..."] + clean_student_list)
    st.markdown("<br><br><br><br><br>", unsafe_allow_html=True)

    col1, spacer, col2 = st.columns([2.5, 5, 2.5])
    with col1: btn_single = st.button("🧑 개별 리포트 생성", type="primary", use_container_width=True)
    with spacer: st.empty() 
    with col2: btn_all = st.button("🌟 전체 학생 일괄 출력", use_container_width=True)

    if btn_single:
        if target_student == "선택하세요...": 
            st.warning("⚠️ 학생 이름을 먼저 선택해주세요.")
        else:
            with st.spinner(f"{target_student} 학생 리포트 생성 중..."):
                success, buf, msg = generate_hakryeok_report(target_student, selected_test)
                if success:
                    st.success(msg)
                    st.download_button("📥 PDF 다운로드", buf.getvalue(), f"{target_student}_학력평가_리포트.pdf", "application/pdf")
                else: st.error(msg)

    if btn_all:
        with st.spinner("전체 리포트 일괄 생성 중... (시간이 소요됩니다)"):
            success, buf, msg = generate_hakryeok_report("전체", selected_test)
            if success:
                st.success(msg)
                st.download_button("📥 전체 PDF 다운로드", buf.getvalue(), f"{selected_test}_학력평가_전체.pdf", "application/pdf")
            else: st.error(msg)

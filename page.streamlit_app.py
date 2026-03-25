import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
import textwrap
import matplotlib.font_manager as fm
import matplotlib.patheffects as path_effects
import traceback
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# --- 1. 환경 및 폰트 설정 ---
font_path = "NanumSquareRoundB.ttf"
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
else:
    plt.rcParams['font.family'] = 'Malgun Gothic'

plt.rcParams['axes.unicode_minus'] = False
  
COLOR_NAVY = '#1F4E3D'; COLOR_RED = '#D97706'; COLOR_STUDENT = '#2F855A'
COLOR_AVG = '#9CA3AF'; COLOR_GRID = '#E5E7EB'; COLOR_BG = '#F9FAFB'

# --- 2. 구글 스프레드시트 연동 및 캐시 설정 ---
@st.cache_resource
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if os.path.exists("secrets.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
    else:
        try:
            if "GOOGLE_JSON" in st.secrets:
                creds_dict = json.loads(st.secrets["GOOGLE_JSON"])
            elif "gcp_secret_string" in st.secrets:
                creds_dict = json.loads(st.secrets["gcp_secret_string"])
            elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                creds_dict = json.loads(st.secrets["connections"]["gsheets"].get("credentials", "{}"))
            else:
                st.error("구글 시트 인증 정보를 찾을 수 없습니다.")
                st.stop()
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except Exception:
            st.error("secrets.json 파일이나 스트림릿 클라우드 보안 키가 없습니다!")
            st.stop()
    client = gspread.authorize(creds)
    doc = client.open_by_url("https://docs.google.com/spreadsheets/d/1pFj7C3uv1Q7PffHsN8Spg2PmHbc1hSKYDgdtqT7rRDk/edit?gid=0#gid=0")
    return doc

@st.cache_data(ttl=120)
def fetch_all_dataframes():
    doc = get_google_sheet()
    df_info = pd.DataFrame(doc.worksheet('Test_Info').get_all_records())
    df_results = pd.DataFrame(doc.worksheet('Student_Results').get_all_records())
    df_results = df_results.replace('', 0).fillna(0)
    return df_info, df_results

def load_data():
    doc = get_google_sheet()
    ws_info = doc.worksheet('Test_Info')
    ws_results = doc.worksheet('Student_Results')
    df_info, df_results = fetch_all_dataframes()
    return doc, ws_info, ws_results, df_info, df_results

# --- 3. PDF 생성 함수 (일괄 출력 기능 탑재) ---
def generate_jeet_expert_report(target_name, selected_test):
    try:
        _, _, _, df_info, df_results = load_data()
        df_info = df_info[df_info['시험명'] == selected_test]
        df_results = df_results[df_results['시험명'] == selected_test]
        df_results.columns = df_results.columns.astype(str)
        df_info = df_info[df_info['문항번호'].astype(str).str.strip() != '']
        df_info['배점'] = 1
        unit_order = df_info['단원'].drop_duplicates().tolist()
        q_cols = [str(q) for q in df_info['문항번호']]
        valid_cols = [col for col in df_results.columns if col in q_cols]
        def safe_to_int(val):
            try: return int(float(val))
            except: return 0
        df_scores = df_results[valid_cols].applymap(safe_to_int)
        df_scores = df_scores[df_scores.sum(axis=1) > 0]
        avg_per_q = df_scores.mean()
        total_analysis = df_info.copy()
        total_analysis['평균득점'] = total_analysis['문항번호'].apply(lambda x: avg_per_q.get(str(x), 0)) * total_analysis['배점']
        avg_cat_ratio = (total_analysis.groupby('영역')['평균득점'].sum() / total_analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
        unit_avg_data = total_analysis.groupby('단원').agg({'평균득점': 'sum'})
        unit_avg_data = unit_avg_data.reindex([u for u in unit_order if u in unit_avg_data.index])
        student_found = False
        pdf_buffer = io.BytesIO()

        with PdfPages(pdf_buffer) as pdf:
            for _, s_row in df_results.iterrows():
                student_name = str(s_row.get('이름', '')).strip()
                if not student_name or student_name == '0': continue
                if target_name != "전체" and student_name != str(target_name).strip(): continue
                student_found = True
                student_grade = s_row.get('학년', '')
                analysis = df_info.copy()
                analysis['정답여부'] = [safe_to_int(s_row.get(str(q), 0)) for q in analysis['문항번호']]
                analysis['득점'] = analysis['정답여부'] * analysis['배점']
                if analysis['득점'].sum() == 0: continue
                cat_ratio = (analysis.groupby('영역')['득점'].sum() / analysis.groupby('영역')['배점'].sum() * 100).fillna(0)
                unit_data = analysis.groupby('단원').agg({'득점': 'sum', '배점': 'sum'})
                unit_data = unit_data.reindex([u for u in unit_order if u in unit_data.index])
                fig = plt.figure(figsize=(8.27, 11.69))
                border = plt.Rectangle((0.015, 0.015), 0.97, 0.97, fill=False, edgecolor=COLOR_RED, linewidth=5.0, transform=fig.transFigure, zorder=10)
                fig.patches.append(border)
                if os.path.exists("logo.png"):
                    logo_img = plt.imread("logo.png")
                    logo_ax = fig.add_axes([0.80, 0.915, 0.15, 0.045], zorder=15)
                    logo_ax.imshow(logo_img)
                    logo_ax.axis('off')
                fig.text(0.31, 0.88, 'JEET', fontsize=32, fontweight='bold', color='red', ha='right')
                fig.text(0.33, 0.88, '수학 능력 분석 리포트', fontsize=32, fontweight='bold', color=COLOR_NAVY, ha='left')
                info_text = f"학교: {s_row.get('학교', '')}  |  학년: {student_grade}  |  이름: {student_name}  |  과정: {selected_test}"
                fig.text(0.5, 0.84, info_text, ha='center', fontsize=15, fontweight='bold', color='#222')
                ax1 = fig.add_axes([0.15, 0.52, 0.32, 0.22], polar=True)
                all_cats = cat_ratio.index.tolist()
                ordered_labels = ['수리 연산'] + [c for c in all_cats if c != '수리 연산'] if '수리 연산력' in all_cats else all_cats
                s_ordered = cat_ratio.reindex(ordered_labels)
                a_ordered = avg_cat_ratio.reindex(ordered_labels)
                labels = s_ordered.index.tolist()
                s_vals = s_ordered.values.tolist() + [s_ordered.values[0]]
                a_vals = a_ordered.values.tolist() + [a_ordered.values[0]]
                angles = np.linspace(0, 2*np.pi, len(labels), endpoint=False).tolist() + [0]
                ax1.set_theta_direction(-1); ax1.set_theta_offset(np.pi/2.0)
                ax1.plot(angles, a_vals, color=COLOR_AVG, linewidth=1, linestyle='--', label='전체 평균')
                ax1.fill(angles, a_vals, color=COLOR_AVG, alpha=0.1)
                ax1.plot(angles, s_vals, color=COLOR_STUDENT, linewidth=2.5, label='학생 성취도', path_effects=[path_effects.SimpleLineShadow(shadow_color='#888888', alpha=0.6, offset=(2, -2)), path_effects.Normal()])
                ax1.fill(angles, s_vals, color=COLOR_STUDENT, alpha=0.15) 
                ax1.set_ylim(0, 110); ax1.set_xticks(angles[:-1]); ax1.set_xticklabels([]); ax1.set_yticklabels([]) 
                ax1.spines['polar'].set_visible(False); ax1.grid(color=COLOR_GRID, linestyle=':', linewidth=1) 
                for i in range(len(labels)):
                    angle = angles[i]; label_text = labels[i]
                    if angle == 0: ha, va, dist = 'center', 'bottom', 105
                    elif 0 < angle < np.pi: ha, va, dist = 'left', 'center', 120
                    elif angle == np.pi: ha, va, dist = 'center', 'top', 125
                    else: ha, va, dist = 'right', 'center', 120
                    ax1.text(angle, dist, label_text, fontsize=10, fontweight='bold', va=va, ha=ha, color=COLOR_NAVY)
                    s_v, a_v = int(s_vals[i]), int(a_vals[i])
                    td = s_v + 10 if s_v < 85 else s_v - 18
                    txt_s = ax1.text(angle, td, f"{s_v}%", fontsize=9, fontweight='bold', color=COLOR_STUDENT, va='center', ha='right')
                    txt_a = ax1.text(angle, td, f" ({a_v}%)", fontsize=9, fontweight='bold', color=COLOR_RED, va='center', ha='left')
                    for t in [txt_s, txt_a]: t.set_path_effects([path_effects.withStroke(linewidth=3, foreground='white')])
                ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.15), ncol=2, fontsize=8, frameon=False)
                ax2 = fig.add_axes([0.55, 0.52, 0.35, 0.20])
                x_pos = np.arange(len(unit_data))
                ax2.bar(x_pos, unit_avg_data['평균득점'], color=COLOR_AVG, alpha=0.25, width=0.55, label='전체 평균', zorder=2)
                ax2.bar(x_pos, unit_data['득점'], color=COLOR_STUDENT, alpha=0.9, width=0.25, label='학생 성취도', zorder=3)
                ax2.set_xticks(x_pos); ax2.set_xticklabels([textwrap.fill(str(l), 5) for l in unit_data.index], fontsize=8, fontweight='bold', color=COLOR_NAVY)
                max_v = unit_data['배점'].max(); max_v = 10 if pd.isna(max_v) or max_v == 0 else max_v
                ax2.set_ylim(0, max_v * 1.4); ax2.legend(loc='upper center', bbox_to_anchor=(0.5, 1.25), ncol=2, fontsize=8, frameon=False)
                ax2.grid(axis='y', color=COLOR_GRID, linestyle='--', linewidth=0.5, zorder=0)
                for i in range(len(unit_data)):
                    sv, av = int(unit_data['득점'].iloc[i]), int(unit_avg_data['평균득점'].iloc[i])
                    ax2.text(x_pos[i], sv + 0.3, f"{sv}", ha='center', va='bottom', fontsize=9, fontweight='bold', color=COLOR_STUDENT)
                    if sv != av: ax2.text(x_pos[i] + 0.28, av, f"({av})", ha='left', va='center', fontsize=8, fontweight='bold', color='#757575')
                ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False); ax2.spines['left'].set_visible(False); ax2.spines['bottom'].set_color(COLOR_GRID); ax2.set_yticks([]) 
                fig.text(0.31, 0.78, "▶ 영역별 핵심 역량 지표 (%)", fontsize=14, fontweight='bold', color=COLOR_NAVY, ha='center')
                fig.text(0.725, 0.78, "▶ 단원별 성취도", fontsize=14, fontweight='bold', color=COLOR_NAVY, ha='center')
                rect_diag = plt.Rectangle((0.08, 0.15), 0.84, 0.32, fill=True, facecolor=COLOR_BG, edgecolor=COLOR_GRID, transform=fig.transFigure)
                fig.patches.append(rect_diag)
                fig.text(0.11, 0.44, "▶ ", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                fig.text(0.13, 0.44, " JEET", fontsize=15, fontweight='bold', color='red')
                fig.text(0.20, 0.44, f" {student_name} 학생 심층 분석", fontsize=15, fontweight='bold', color=COLOR_NAVY)
                total_stu_score, total_max_score = analysis['득점'].sum(), analysis['배점'].sum()
                avg_val = int((total_stu_score / total_max_score) * 100) if total_max_score > 0 else 0
                total_avg_score, total_avg_max_score = total_analysis['평균득점'].sum(), total_analysis['배점'].sum()
                total_avg_val = int((total_avg_score / total_avg_max_score) * 100) if total_avg_max_score > 0 else 0
                diff_cats = s_ordered - a_ordered
                best_cat, worst_cat = diff_cats.idxmax(), diff_cats.idxmin()
                unit_diff = unit_data['득점'] - unit_avg_data['평균득점']
                worst_unit = unit_diff.idxmin() if not unit_diff.empty else "전반적인"
                if avg_val >= 90: eval_tier = "고난도 심화 개념까지 흔들림 없이 소화해 내는 최상위권 성취도"
                elif avg_val >= 75: eval_tier = "탄탄한 기본기를 바탕으로 안정적인 성장을 보여주는 우수한 성취도"
                elif avg_val >= 60: eval_tier = "핵심 개념을 내재화하며 다음 단계로 성실히 나아가고 있는 성장형 성취도"
                else: eval_tier = "수학적 잠재력을 깨우기 위해 개념의 뼈대를 견고하게 다져가는 발돋움 단계"
                if avg_val >= 80:
                    sol_dict = {
                        '수리 연산': "이제 단순한 계산을 넘어, 빠르고 정교한 연산 설계가 필요한 시점입니다. 고난도 문항을 풀 때 본인의 풀이 과정을 논리정연하게 식별하고 스스로 검산하는 루틴을 체화한다면 실전에서의 잔실수를 완벽히 차단할 수 있습니다.",
                        '개념 이해': "개념에 대한 뼈대가 이미 견고하게 잡혀 있습니다. 여기서 한 걸음 더 나아가, 배운 개념을 융합하여 문제를 변형해 보거나 남에게 직접 설명해 보는 '출제자 모드' 학습을 권장합니다. 이는 최상위권 굳히기의 확실한 열쇠가 될 것입니다.",
                        '논리 추론': "숨겨진 출제자의 의도를 파악하는 감각이 탁월합니다. 이제는 복잡한 조건이 다중으로 얽힌 킬러 문항에 적극적으로 도전할 때입니다. 조건을 잘게 쪼개어 분석하고 여러 개념을 유기적으로 연결하는 심화 훈련을 진행하길 바랍니다.",
                        '실전 응용': "배운 것을 실전에 적용하는 능력이 매우 훌륭합니다. 지금부터는 낯설고 생소한 신유형과 변형 문제를 다양하게 접하며 실전 적응력을 극대화해야 합니다. 특히 시간을 엄격하게 제한하는 타임어택 훈련을 통한 멘탈 관리가 동반되면 완벽합니다."
                    }
                else:
                    sol_dict = {
                        '수리 연산': "수학의 가장 든든한 무기인 '정확하고 신속한 연산력'을 최우선으로 끌어올려야 할 시기입니다. 눈으로만 풀지 않고 반드시 손으로 끝까지 계산해 내는 끈기가 필요하며, 매일 일정한 분량의 훈련을 통해 근본적인 자신감을 회복해야 합니다.",
                        '개념 이해': "기계적인 문제 풀이를 잠시 멈추고, 내가 무엇을 알고 모르는지 정확히 진단해야 합니다. 공식의 결과만 암기하기보다는, 원리와 증명 과정을 백지에 스스로 적어 내려갈 수 있을 때까지 '진짜 개념'을 다지는 인내의 시간이 필요합니다.",
                        '논리 추론': "문제가 조금만 길어지거나 낯설게 느껴져도 지레 포기하는 습관을 경계해야 합니다. 해설지에 의존하기보다는, 문제 속에 숨겨진 작은 힌트들을 찾아내어 내가 아는 개념과 어떻게든 연결해 보려는 치열한 고민의 과정이 사고력을 비약적으로 성장시킬 것입니다.",
                        '실전 응용': "아직은 무리하게 꼬인 변형 문제에 욕심내기보다는, 교과서와 기본서 수준의 필수 예제를 완벽하게 내 것으로 만드는 데 집중해야 합니다. 하나의 유형이라도 왜 이런 풀이 방식이 적용되는지 꼼꼼히 씹어 넘긴다면 응용력을 꽃피울 훌륭한 토양분이 될 것입니다."
                    }
                gap = s_ordered.max() - s_ordered.min()
                if avg_val >= 95 or gap <= 5:
                    bw_text = "영역별 분석 결과, 모든 영역에서 편차 없이 고르게 탁월한 학업 밸런스를 유지하고 있습니다. 지금의 훌륭한 학습 습관과 몰입도를 꾸준히 이어가면서, 더 넓고 깊은 심화 학습으로 지적 호기심을 확장해 나가길 적극 권장합니다."
                    worst_sol = "지금처럼 꾸준히 올바른 학습 태도를 유지한다면, 앞으로의 중등 과정에서도 최상위권의 자리를 굳건히 지킬 수 있을 것입니다."
                else:
                    bw_text = f"분석 결과, '{best_cat}' 영역에서 남다른 강점이 돋보이는 반면, 상대적으로 '{worst_cat}' 역량에서는 집중적인 보완이 이루어진다면 한 차원 더 높은 도약이 기대됩니다. 특히 이번 평가에서 '{worst_unit}' 단원의 오답률이 다소 눈에 띄므로, 해당 단원의 핵심 개념을 반드시 짚고 넘어가야 합니다."
                    worst_sol = sol_dict.get(worst_cat, "악어수학의 꼼꼼한 맞춤 클리닉을 통해 부족한 부분을 채워 나간다면 충분히 더 큰 성장을 이뤄낼 수 있습니다.")
                diag_c = f"1. 종합 진단: {student_name} 학생은 전체 평균({total_avg_val}%) 대비 성취도 {avg_val}%를 기록하며, 현재 [{eval_tier}]를 보여주고 있습니다.\n\n2. 강약점 분석: {bw_text}\n\n3. JEET 전문가 맞춤 솔루션: {worst_sol}"
                wrapped = [textwrap.fill(p, width=54) for p in diag_c.split('\n\n')]
                fig.text(0.11, 0.41, "\n\n".join(wrapped), fontsize=10.5, linespacing=1.8, va='top', ha='left', color='#333')
                line = plt.Line2D([0.05, 0.95], [0.12, 0.12], color=COLOR_NAVY, linewidth=1, transform=fig.transFigure)
                fig.lines.append(line)
                campuses = [("수지 캠퍼스: 276-8003", "풍덕천로 129번길 16-1"), ("죽전 캠퍼스: 263-8003", "기흥구 죽현로 29"), ("광교 캠퍼스: 257-8003", "영통구 혜령로 10")]
                for i, (name, addr) in enumerate(campuses):
                    fig.text([0.22, 0.50, 0.78][i], 0.08, name, ha='center', fontsize=10, fontweight='bold', color=COLOR_NAVY)
                    fig.text([0.22, 0.50, 0.78][i], 0.05, addr, ha='center', fontsize=7.5, color='#555')
                pdf.savefig(fig); plt.close(fig)
        if not student_found: return False, None, "학생을 찾을 수 없습니다."
        return True, pdf_buffer, "리포트 생성 완료!"
    except Exception as e: return False, None, f"오류 발생: {traceback.format_exc()}"

# --- 4. Streamlit 웹 UI 구성 ---
st.set_page_config(page_title="JEET수학 통합 관리 시스템", layout="wide", page_icon="📊")
col1, col2 = st.columns([8, 2])
with col1: st.title("📊 JEET수학 성적 통합 관리 시스템")
with col2: 
    if os.path.exists("logo.png"): st.image("logo.png", width=150)
try:
    doc, ws_info, ws_results, df_info_all, df_results_all = load_data()
except Exception as e:
    st.error(f"구글 시트 로드 실패: {e}"); st.code(traceback.format_exc()); st.stop()
st.sidebar.header("📚 시험 과정 선택")
test_list = df_info_all['시험명'].dropna().unique().tolist()
selected_test = st.sidebar.selectbox("분석할 시험 과정을 선택하세요:", test_list)
df_info_f = df_info_all[df_info_all['시험명'] == selected_test]
tab1, tab2 = st.tabs(["📝 신규 성적 입력", "📑 개별 리포트 출력"])
with tab1:
    st.subheader(f"[{selected_test}] 신규 학생 성적 입력")
    q_nums = df_info_f['문항번호'].tolist()
    if q_nums:
        with st.form("data_input_form", clear_on_submit=True):
            ci1, ci2, ci3 = st.columns(3)
            with ci1: in_name = st.text_input("이름")
            with ci2: in_school = st.text_input("학교")
            with ci3: in_grade = st.selectbox("학년", ["중1", "중2", "중3"])
            st.markdown("---")
            ans = {}
            for i in range(0, len(q_nums), 5):
                cols = st.columns(5)
                for j, q_num in enumerate(q_nums[i:i+5]):
                    with cols[j]:
                        choice = st.radio(f"**{q_num}번**", options=["O", "X"], horizontal=True, key=f"q_{q_num}")
                        ans[str(q_num)] = 1 if choice == "O" else 0
            if st.form_submit_button("구글 시트에 성적 저장하기", type="primary"):
                c_name = in_name.strip()
                if not c_name: st.error("⚠ 이름을 입력해주세요.")
                else:
                    try:
                        h_row = ws_results.row_values(1); new_row = []
                        for c_name_str in h_row:
                            col_str = str(c_name_str)
                            if col_str == '시험명': new_row.append(selected_test) 
                            elif col_str == '이름': new_row.append(c_name)
                            elif col_str == '학교': new_row.append(in_school)
                            elif col_str == '학년': new_row.append(in_grade)
                            elif col_str in ans: new_row.append(ans[col_str])
                            else: new_row.append("")
                        ws_results.append_row(new_row); st.success("성적이 저장되었습니다!"); st.cache_data.clear()
                    except Exception as ex: st.error(f"저장 중 오류: {ex}")
with tab2:
    st.subheader(f"[{selected_test}] 개별 심층 분석 리포트 생성")
    r_s_list = df_results_all[df_results_all['시험명'] == selected_test]['이름'].dropna().unique().tolist()
    c_s_list = [str(n).strip() for n in r_s_list if str(n).strip() not in ['', '0']]
    c_s_list.sort() 
    target_s = st.selectbox("리포트를 출력할 학생을 선택하세요:", ["선택하세요", "🌟 전체 학생 일괄 출력"] + c_s_list)
    if st.button("PDF 리포트 생성", type="primary"):
        if target_s == "선택하세요": st.warning("⚠️ 학생을 선택해주세요!")
        else:
            actual_t = "전체" if target_s == "🌟 전체 학생 일괄 출력" else target_s
            f_name = f"{selected_test}_전체.pdf" if actual_t == "전체" else f"{target_s}_리포트.pdf"
            l_msg = "전체 리포트 생성 중..." if actual_t == "전체" else f"{target_s} 리포트 생성 중..."
            with st.spinner(l_msg):
                succ, buf, m = generate_jeet_expert_report(actual_t, selected_test)
                if succ: st.success(m); st.download_button("📥 PDF 다운로드", buf.getvalue(), f_name, "application/pdf")
                else: st.error(m)

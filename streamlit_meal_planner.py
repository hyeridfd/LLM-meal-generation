# ============================================================================
# streamlit_meal_planner.py
#
# Streamlit 기반 고령자 맞춤형 식단 생성 웹앱
# 
# 실행: streamlit run streamlit_meal_planner.py
# ============================================================================

import streamlit as st
import pandas as pd
import json
from dataclasses import dataclass
from typing import List, Literal, Tuple, Dict, Optional
from openai import OpenAI
import plotly.graph_objects as go
import plotly.express as px


# ============================================================================
# 1. 페이지 설정
# ============================================================================

st.set_page_config(
    page_title="🍽️ 고령자 맞춤형 식단 생성 시스템",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 커스터마이징
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 16px;
        padding: 10px 20px;
    }
    .metric-box {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================================
# 2. 데이터 클래스 (기존 코드 활용)
# ============================================================================

ChewingStage = Literal[1, 2, 3]

@dataclass
class ElderlyProfile:
    name: str
    age: int
    sex: Literal["M", "F"]
    height_cm: float
    weight_kg: float
    pal: float
    diseases: List[str]
    chewing_stage: ChewingStage
    allergies: List[str]
    likes: List[str]
    dislikes: List[str]


@dataclass
class NutrientRequirement:
    energy_kcal: float
    protein_g: Tuple[float, float]
    fat_g: Tuple[float, float]
    carbs_g: Tuple[float, float]
    sodium_mg: Tuple[float, float]
    fiber_g: Tuple[float, float]
    calcium_mg: Tuple[float, float]
    vitamin_d_mcg: Tuple[float, float]


@dataclass
class MenuItem:
    food_id: str
    name: str
    category: Literal["밥", "국", "주찬", "부찬", "김치"]
    energy_kcal: float
    protein_g: float
    fat_g: float
    carbs_g: float
    sodium_mg: float
    calcium_mg: float
    vitamin_d_mcg: float
    fiber_g: float


# ============================================================================
# 3. 데이터 관리 클래스
# ============================================================================

class RecipeDataManager:
    """엑셀 데이터 관리"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.df = None
        self.all_recipes = {}
        self.load_and_process()
    
    def load_and_process(self):
        """데이터 로드"""
        self.df = pd.read_excel(self.excel_path)
        
        self.df.columns = [
            'food_id', 'name', 'energy', 'protein', 'fat', 'carbs',
            'sodium', 'calcium', 'vitamin_d', 'fiber', 'category'
        ]
        
        for category in ["밥", "국", "주찬", "부찬", "김치"]:
            cat_items = self.df[self.df['category'] == category]
            
            items = []
            for _, row in cat_items.iterrows():
                try:
                    item = MenuItem(
                        food_id=str(row['food_id']),
                        name=str(row['name']),
                        category=category,
                        energy_kcal=float(row['energy']),
                        protein_g=float(row['protein']),
                        fat_g=float(row['fat']),
                        carbs_g=float(row['carbs']),
                        sodium_mg=float(row['sodium']),
                        calcium_mg=float(row['calcium']),
                        vitamin_d_mcg=float(row['vitamin_d']),
                        fiber_g=float(row['fiber'])
                    )
                    items.append(item)
                except Exception as e:
                    pass
            
            self.all_recipes[category] = items
    
    def get_all_categories(self) -> Dict[str, List[MenuItem]]:
        return self.all_recipes


# ============================================================================
# 4. 영양 계산
# ============================================================================

def compute_daily_requirement(profile: ElderlyProfile) -> NutrientRequirement:
    """영양 기준 계산"""
    
    if profile.sex == "M":
        bmr = 10 * profile.weight_kg + 6.25 * profile.height_cm - 5 * profile.age + 5
    else:
        bmr = 10 * profile.weight_kg + 6.25 * profile.height_cm - 5 * profile.age - 161
    
    energy = bmr * profile.pal
    energy_min = energy * 0.95
    energy_max = energy * 1.05
    
    protein_kcal = energy * 0.15
    fat_kcal = energy * 0.30
    carbs_kcal = energy * 0.55
    
    protein_g = protein_kcal / 4
    fat_g = fat_kcal / 9
    carbs_g = carbs_kcal / 4
    
    p_min, p_max = protein_g * 0.8, protein_g * 1.2
    f_min, f_max = fat_g * 0.8, fat_g * 1.2
    c_min, c_max = carbs_g * 0.8, carbs_g * 1.2
    
    sodium_max = 2000.0
    if "고혈압" in profile.diseases:
        sodium_max = 1500.0
    
    fiber_min = 20.0
    
    calcium_target = 800.0 if profile.sex == "F" else 1000.0
    calcium_min = calcium_target * 0.8
    calcium_max = calcium_target * 1.2
    
    vitamin_d_min = 12.0
    vitamin_d_max = 20.0
    
    return NutrientRequirement(
        energy_kcal=energy,
        protein_g=(p_min, p_max),
        fat_g=(f_min, f_max),
        carbs_g=(c_min, c_max),
        sodium_mg=(0.0, sodium_max),
        fiber_g=(fiber_min, 999.0),
        calcium_mg=(calcium_min, calcium_max),
        vitamin_d_mcg=(vitamin_d_min, vitamin_d_max)
    )


# ============================================================================
# 5. 프롬프트 생성
# ============================================================================

def build_menu_list_for_prompt(categories: Dict[str, List[MenuItem]]) -> str:
    """메뉴 리스트 생성"""
    lines = []
    
    for cat in ["밥", "국", "주찬", "부찬", "김치"]:
        items = categories.get(cat, [])
        lines.append(f"\n[{cat}] - 총 {len(items)}개 항목")
        
        for item in items[:20]:
            line = (
                f"  {item.food_id[:10]}｜{item.name:20} "
                f"｜ {item.energy_kcal:6.1f}kcal "
                f"｜ 나트륨{item.sodium_mg:7.1f}mg"
            )
            lines.append(line)
        
        if len(items) > 20:
            lines.append(f"  ... 외 {len(items) - 20}개 더 있음")
    
    return "\n".join(lines)


def build_meal_planning_prompt(
    profile: ElderlyProfile,
    req: NutrientRequirement,
    categories: Dict[str, List[MenuItem]]
) -> str:
    """식단 생성 프롬프트"""
    
    menu_list = build_menu_list_for_prompt(categories)
    diseases = ", ".join(profile.diseases) if profile.diseases else "없음"
    allergies = ", ".join(profile.allergies) if profile.allergies else "없음"
    likes = ", ".join(profile.likes) if profile.likes else "없음"
    dislikes = ", ".join(profile.dislikes) if profile.dislikes else "없음"
    
    prompt = f"""당신은 고령자 영양관리 전문 영양사 팀입니다.
다음 고령자를 위한 최적화된 7일 식단을 설계해주세요.

【고령자 프로필】
- 이름: {profile.name}
- 나이: {profile.age}세
- 성별: {"남성" if profile.sex == "M" else "여성"}
- 신체: {profile.height_cm}cm, {profile.weight_kg}kg
- 신체활동계수(PAL): {profile.pal}
- 질환: {diseases}
- 저작 단계: {profile.chewing_stage}단계
- 알레르기: {allergies}
- 선호: {likes}
- 비선호: {dislikes}

【1일 영양 기준】
- 에너지: {req.energy_kcal*0.95:.0f}~{req.energy_kcal*1.05:.0f} kcal
- 단백질: {req.protein_g[0]:.1f}~{req.protein_g[1]:.1f}g
- 지방: {req.fat_g[0]:.1f}~{req.fat_g[1]:.1f}g
- 탄수화물: {req.carbs_g[0]:.1f}~{req.carbs_g[1]:.1f}g
- 나트륨: 0~{req.sodium_mg[1]:.0f}mg
- 식이섬유: {req.fiber_g[0]:.1f}g 이상

【끼니 구성】
- breakfast: 일일 에너지의 25%
- lunch: 일일 에너지의 35%
- dinner: 일일 에너지의 30%
- snack: 일일 에너지의 10%

【사용 가능한 레시피】
{menu_list}

【과제】
1. 7일 완벽한 식단 설계
2. 각 끼니는 반드시: 밥(1) + 국(1) + 주찬(1) + 부찬(2) + 김치(1)
3. 전체 1일 영양소를 범위에 맞춰서
4. 같은 음식 3회 이상 반복 금지
5. {likes}는 자주 포함, {dislikes}는 피하기

【출력 형식】
반드시 JSON만:
{{
  "weekly_plan": [
    {{
      "day": 1,
      "day_name": "월요일",
      "meals": [
        {{
          "meal_type": "breakfast",
          "items": {{
            "밥": ["food_id"],
            "국": ["food_id"],
            "주찬": ["food_id"],
            "부찬": ["food_id", "food_id"],
            "김치": ["food_id"]
          }}
        }},
        ... 점심, 저녁, 간식 ...
      ]
    }},
    ... day 2-7 ...
  ]
}}
"""
    
    return prompt


# ============================================================================
# 6. LLM 호출
# ============================================================================

@st.cache_resource
def get_openai_client():
    """OpenAI 클라이언트 (캐시)"""
    api_key = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    if not api_key:
        st.error("❌ OPENAI_API_KEY 환경변수를 설정해주세요")
        st.stop()
    return OpenAI(api_key=api_key)


def extract_json(text: str) -> str:
    """JSON 추출"""
    start = text.find("{")
    end = text.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError("JSON을 찾을 수 없습니다")
    return text[start:end]


def generate_meal_plan_with_llm(
    profile: ElderlyProfile,
    req: NutrientRequirement,
    categories: Dict[str, List[MenuItem]]
) -> dict:
    """LLM 호출"""
    
    client = get_openai_client()
    prompt = build_meal_planning_prompt(profile, req, categories)
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "당신은 고령자 영양관리 전문 영양사입니다. 반드시 유효한 JSON 형식으로만 응답하세요."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
            max_tokens=4000
        )
        
        content = response.choices[0].message.content
        json_str = extract_json(content)
        return json.loads(json_str)
        
    except Exception as e:
        st.error(f"❌ LLM 호출 오류: {str(e)}")
        return None


# ============================================================================
# 7. 분석 클래스
# ============================================================================

class MealPlanAnalyzer:
    """식단 분석"""
    
    def __init__(self, profile, req, categories, weekly_plan):
        self.profile = profile
        self.req = req
        self.categories = categories
        self.weekly_plan = weekly_plan
        self.food_map = self._create_food_map()
    
    def _create_food_map(self) -> Dict[str, MenuItem]:
        food_map = {}
        for items in self.categories.values():
            for item in items:
                food_map[item.food_id] = item
        return food_map
    
    def analyze_daily_nutrition(self, day_num: int) -> dict:
        """일별 영양소 분석"""
        day_data = self.weekly_plan["weekly_plan"][day_num - 1]
        
        total = {
            "energy_kcal": 0.0,
            "protein_g": 0.0,
            "fat_g": 0.0,
            "carbs_g": 0.0,
            "sodium_mg": 0.0,
            "calcium_mg": 0.0,
            "fiber_g": 0.0
        }
        
        for meal in day_data["meals"]:
            for category, food_ids in meal["items"].items():
                for food_id in food_ids:
                    if food_id in self.food_map:
                        item = self.food_map[food_id]
                        total["energy_kcal"] += item.energy_kcal
                        total["protein_g"] += item.protein_g
                        total["fat_g"] += item.fat_g
                        total["carbs_g"] += item.carbs_g
                        total["sodium_mg"] += item.sodium_mg
                        total["calcium_mg"] += item.calcium_mg
                        total["fiber_g"] += item.fiber_g
        
        return {
            "day": day_num,
            "nutrition": total,
            "day_name": ["월", "화", "수", "목", "금", "토", "일"][day_num - 1]
        }
    
    def get_all_daily_analysis(self) -> List[dict]:
        """전체 주간 분석"""
        return [self.analyze_daily_nutrition(day) for day in range(1, 8)]


# ============================================================================
# 8. Streamlit UI
# ============================================================================

def main():
    """메인 앱"""
    
    # 헤더
    st.title("🍽️ 고령자 맞춤형 7일 식단 생성 시스템")
    st.markdown("**LLM 기반 Hybrid Intelligence 접근법**")
    st.divider()
    
    # 세션 상태 초기화
    if 'data_manager' not in st.session_state:
        st.session_state.data_manager = None
    if 'weekly_plan' not in st.session_state:
        st.session_state.weekly_plan = None
    if 'analyzer' not in st.session_state:
        st.session_state.analyzer = None
    
    # 탭
    tab1, tab2, tab3, tab4 = st.tabs(["👤 사용자 정보", "📊 식단 생성", "📈 분석", "💾 결과"])
    
    # ========== TAB 1: 사용자 정보 ==========
    with tab1:
        st.header("고령자 프로필 입력")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📋 기본 정보")
            name = st.text_input("이름", value="홍길동")
            age = st.slider("나이", 60, 100, 78)
            sex = st.radio("성별", options=["M", "F"], format_func=lambda x: "남성" if x == "M" else "여성")
            
            st.subheader("📏 신체 정보")
            height = st.slider("키 (cm)", 140, 200, 155)
            weight = st.slider("체중 (kg)", 30, 120, 55)
        
        with col2:
            st.subheader("🏃 활동 수준")
            pal_options = {
                "거의 움직이지 않음 (침상)": 1.2,
                "앉아서 일함": 1.3,
                "약간의 활동": 1.4,
                "중간 활동": 1.55,
                "활동적": 1.7,
                "매우 활동적": 2.0
            }
            pal = st.selectbox("신체활동계수(PAL)", options=pal_options.keys())
            pal_value = pal_options[pal]
            
            st.subheader("💊 건강 상태")
            diseases = st.multiselect(
                "질환 (중복 선택 가능)",
                options=["당뇨", "고혈압", "신장질환", "심장질환", "골다공증"],
                default=["당뇨", "고혈압"]
            )
            
            chewing = st.selectbox(
                "저작 단계",
                options=[1, 2, 3],
                format_func=lambda x: f"Level {x}: " + {1: "연하곤란", 2: "저작곤란", 3: "정상"}[x]
            )
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("🚫 알레르기 및 선호도")
            allergies = st.multiselect(
                "알레르기",
                options=["우유", "계란", "견과류", "생선", "새우", "밀"],
                default=[]
            )
            
            likes = st.multiselect(
                "선호 음식",
                options=["생선", "달걀", "나물", "김", "두부", "채소"],
                default=["생선", "달걀"]
            )
        
        with col4:
            st.subheader("")
            dislikes = st.multiselect(
                "비선호 음식",
                options=["매운 음식", "기름진 음식", "냄새 강한 음식", "딱딱한 음식"],
                default=["매운 음식"]
            )
        
        # 프로필 저장
        if st.button("✅ 프로필 저장", use_container_width=True, type="primary"):
            st.session_state.profile = ElderlyProfile(
                name=name,
                age=age,
                sex=sex,
                height_cm=height,
                weight_kg=weight,
                pal=pal_value,
                diseases=diseases,
                chewing_stage=chewing,
                allergies=allergies,
                likes=likes,
                dislikes=dislikes
            )
            
            st.session_state.requirement = compute_daily_requirement(st.session_state.profile)
            st.success("✅ 프로필이 저장되었습니다!")
            
            # 영양 기준 표시
            req = st.session_state.requirement
            col_a, col_b, col_c, col_d = st.columns(4)
            with col_a:
                st.metric("일일 에너지", f"{req.energy_kcal:.0f} kcal")
            with col_b:
                st.metric("단백질", f"{req.protein_g[0]:.0f}~{req.protein_g[1]:.0f}g")
            with col_c:
                st.metric("나트륨 (max)", f"{req.sodium_mg[1]:.0f}mg")
            with col_d:
                st.metric("식이섬유 (min)", f"{req.fiber_g[0]:.0f}g")
    
    # ========== TAB 2: 식단 생성 ==========
    with tab2:
        st.header("🤖 LLM을 이용한 식단 생성")
        
        if 'profile' not in st.session_state:
            st.warning("⚠️ 먼저 [👤 사용자 정보] 탭에서 프로필을 입력하세요")
        else:
            col_left, col_right = st.columns([3, 1])
            
            with col_left:
                st.info(f"""
                **프로필 요약**
                - 👤 {st.session_state.profile.name} ({st.session_state.profile.age}세, {"남" if st.session_state.profile.sex == "M" else "여"})
                - 🏥 질환: {", ".join(st.session_state.profile.diseases)}
                - 🍖 선호: {", ".join(st.session_state.profile.likes)}
                """)
            
            with col_right:
                if st.button("🚀 식단 생성!", use_container_width=True, type="primary", key="generate"):
                    with st.spinner("🔄 AI가 식단을 생성 중입니다... (약 30초)"):
                        # 데이터 로드
                        if st.session_state.data_manager is None:
                            try:
                                st.session_state.data_manager = RecipeDataManager("merged_nutrient_with_category.xlsx")
                            except FileNotFoundError:
                                st.error("❌ 엑셀 파일을 찾을 수 없습니다")
                                st.stop()
                        
                        categories = st.session_state.data_manager.get_all_categories()
                        
                        # 식단 생성
                        weekly_plan = generate_meal_plan_with_llm(
                            st.session_state.profile,
                            st.session_state.requirement,
                            categories
                        )
                        
                        if weekly_plan:
                            st.session_state.weekly_plan = weekly_plan
                            st.session_state.analyzer = MealPlanAnalyzer(
                                st.session_state.profile,
                                st.session_state.requirement,
                                categories,
                                weekly_plan
                            )
                            st.success("✅ 식단 생성 완료!")
                        else:
                            st.error("❌ 식단 생성 실패")
            
            # 생성된 식단 표시
            if st.session_state.weekly_plan:
                st.divider()
                st.success("✅ 7일 식단이 생성되었습니다!")
                
                # 요일 선택
                day_select = st.selectbox(
                    "요일 선택",
                    options=[f"{i}일차 ({['월', '화', '수', '목', '금', '토', '일'][i-1]})" for i in range(1, 8)]
                )
                
                day_num = int(day_select.split("일")[0])
                day_data = st.session_state.weekly_plan["weekly_plan"][day_num - 1]
                
                # 끼니별 표시
                st.subheader(f"📅 {day_select} 식단")
                
                meal_names_kr = {
                    "breakfast": "🌅 아침",
                    "lunch": "🍽️ 점심",
                    "dinner": "🌆 저녁",
                    "snack": "☕ 간식"
                }
                
                for meal in day_data["meals"]:
                    with st.expander(f"{meal_names_kr[meal['meal_type']]}", expanded=True):
                        meal_cols = st.columns(5)
                        categories_list = ["밥", "국", "주찬", "부찬", "김치"]
                        
                        for idx, category in enumerate(categories_list):
                            with meal_cols[idx]:
                                st.markdown(f"**{category}**")
                                food_ids = meal["items"].get(category, [])
                                
                                if food_ids:
                                    for food_id in food_ids:
                                        # 음식 찾기
                                        if st.session_state.data_manager:
                                            for items in st.session_state.data_manager.get_all_categories().values():
                                                for item in items:
                                                    if item.food_id == food_id:
                                                        st.write(f"• {item.name}")
                                                        st.caption(f"{item.energy_kcal:.0f}kcal")
                                                        break
    
    # ========== TAB 3: 분석 ==========
    with tab3:
        st.header("📊 식단 분석")
        
        if not st.session_state.weekly_plan or not st.session_state.analyzer:
            st.warning("⚠️ 먼저 식단을 생성하세요")
        else:
            analyzer = st.session_state.analyzer
            req = st.session_state.requirement
            
            # 주간 분석 데이터
            daily_analysis = analyzer.get_all_daily_analysis()
            
            # 영양소별 그래프
            col_chart1, col_chart2 = st.columns(2)
            
            # 에너지 그래프
            with col_chart1:
                energy_data = [d["nutrition"]["energy_kcal"] for d in daily_analysis]
                day_labels = [d["day_name"] for d in daily_analysis]
                
                fig_energy = go.Figure()
                fig_energy.add_trace(go.Scatter(
                    x=day_labels, y=energy_data,
                    mode='lines+markers',
                    name='실제 에너지',
                    line=dict(color='#1f77b4', width=3)
                ))
                fig_energy.add_hline(y=req.energy_kcal, line_dash="dash", line_color="red", annotation_text="목표")
                fig_energy.update_layout(title="⚡ 일일 에너지", yaxis_title="kcal", height=400)
                st.plotly_chart(fig_energy, use_container_width=True)
            
            # 나트륨 그래프
            with col_chart2:
                sodium_data = [d["nutrition"]["sodium_mg"] for d in daily_analysis]
                
                fig_sodium = go.Figure()
                fig_sodium.add_trace(go.Scatter(
                    x=day_labels, y=sodium_data,
                    mode='lines+markers',
                    name='나트륨',
                    line=dict(color='#ff7f0e', width=3)
                ))
                fig_sodium.add_hline(y=req.sodium_mg[1], line_dash="dash", line_color="red", annotation_text="최대")
                fig_sodium.update_layout(title="🧂 일일 나트륨", yaxis_title="mg", height=400)
                st.plotly_chart(fig_sodium, use_container_width=True)
            
            # 상세 표
            st.divider()
            st.subheader("📋 주간 영양소 상세")
            
            table_data = []
            for d in daily_analysis:
                n = d["nutrition"]
                table_data.append({
                    "요일": f"{d['day']}일({d['day_name']})",
                    "에너지 (kcal)": f"{n['energy_kcal']:.0f}",
                    "단백질 (g)": f"{n['protein_g']:.1f}",
                    "나트륨 (mg)": f"{n['sodium_mg']:.0f}",
                    "식이섬유 (g)": f"{n['fiber_g']:.1f}"
                })
            
            df_table = pd.DataFrame(table_data)
            st.dataframe(df_table, use_container_width=True)
    
    # ========== TAB 4: 결과 ==========
    with tab4:
        st.header("💾 결과 저장")
        
        if not st.session_state.weekly_plan:
            st.warning("⚠️ 먼저 식단을 생성하세요")
        else:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # JSON 다운로드
                json_str = json.dumps(st.session_state.weekly_plan, ensure_ascii=False, indent=2)
                st.download_button(
                    label="📥 JSON 다운로드",
                    data=json_str,
                    file_name=f"{st.session_state.profile.name}_meal_plan.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            with col2:
                # CSV 다운로드
                if st.session_state.analyzer:
                    daily_analysis = st.session_state.analyzer.get_all_daily_analysis()
                    table_data = []
                    for d in daily_analysis:
                        n = d["nutrition"]
                        table_data.append({
                            "Day": d['day'],
                            "Day_Name": d['day_name'],
                            "Energy_kcal": round(n['energy_kcal'], 1),
                            "Protein_g": round(n['protein_g'], 1),
                            "Fat_g": round(n['fat_g'], 1),
                            "Carbs_g": round(n['carbs_g'], 1),
                            "Sodium_mg": round(n['sodium_mg'], 1),
                            "Fiber_g": round(n['fiber_g'], 1),
                        })
                    
                    df = pd.DataFrame(table_data)
                    csv_str = df.to_csv(index=False)
                    
                    st.download_button(
                        label="📊 CSV 다운로드",
                        data=csv_str,
                        file_name=f"{st.session_state.profile.name}_nutrition.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
            
            with col3:
                # 텍스트 리포트
                if st.button("📄 텍스트 리포트", use_container_width=True):
                    report = f"""
╔════════════════════════════════════════════════════════════════╗
║           고령자 맞춤형 7일 식단 계획서                          ║
╚════════════════════════════════════════════════════════════════╝

【대상자 정보】
- 이름: {st.session_state.profile.name}
- 나이: {st.session_state.profile.age}세
- 신체: {st.session_state.profile.height_cm}cm, {st.session_state.profile.weight_kg}kg
- 질환: {", ".join(st.session_state.profile.diseases)}
- 저작 단계: {st.session_state.profile.chewing_stage}단계

【영양 기준】
- 일일 에너지: {st.session_state.requirement.energy_kcal*0.95:.0f}~{st.session_state.requirement.energy_kcal*1.05:.0f} kcal
- 단백질: {st.session_state.requirement.protein_g[0]:.1f}~{st.session_state.requirement.protein_g[1]:.1f}g
- 나트륨: 0~{st.session_state.requirement.sodium_mg[1]:.0f}mg
- 식이섬유: {st.session_state.requirement.fiber_g[0]:.1f}g 이상

생성일시: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
                    st.text_area("리포트", value=report, height=300, disabled=True)
            
            st.divider()
            st.success("✅ 식단 생성 및 분석 완료!")
            st.info("""
            ### 🎉 다음 단계
            1. JSON 파일을 영양관리 시스템에 입력
            2. 배식 담당자가 식단 확인
            3. 환자에게 제공 및 반응 모니터링
            4. 필요시 프로필 수정 후 재생성
            """)


if __name__ == "__main__":
    main()

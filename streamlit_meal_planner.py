# ============================================================================
# streamlit_meal_planner_advanced.py
#
# 업그레이드 버전:
# 1. 엑셀 배치 처리 (여러 명의 식단 한번에 생성)
# 2. 현재식사현황 자동 인식 (일반밥/죽/미음, 일반찬/다진찬/갈찬)
# 3. 조리법 명시 (다진찬, 갈찬 표시)
# 4. LLM이 선택한 이유 설명
# ============================================================================

import streamlit as st
import pandas as pd
import json
from dataclasses import dataclass
from typing import List, Literal, Tuple, Dict, Optional
from openai import OpenAI
from dotenv import load_dotenv
import os
import plotly.graph_objects as go
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================================
# 1. 페이지 설정
# ============================================================================

st.set_page_config(
    page_title="🍽️ 요양원 식단 자동 생성 시스템",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 16px;
        padding: 10px 20px;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================================
# 2. 데이터 클래스
# ============================================================================

ChewingStage = Literal[1, 2, 3]

@dataclass
class ElderlyProfile:
    resident_id: str  # GC01, GC02 등
    name: str
    age: int
    sex: Literal["M", "F"]
    height_cm: float
    weight_kg: float
    pal: float
    diseases: List[str]
    chewing_stage: ChewingStage
    current_meal: str  # "일반밥/일반찬", "죽식/다진찬" 등
    meal_type: str    # "일반밥", "죽식", "미음"
    side_type: str    # "일반찬", "다진찬", "갈찬"


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
# 3. 엑셀 파일 파싱
# ============================================================================

def parse_elderly_excel(uploaded_file) -> List[ElderlyProfile]:
    """
    요양원 엑셀 파일 파싱
    """
    try:
        df = pd.read_excel(uploaded_file, sheet_name='고령자 데이터', header=None)
        
        # 데이터 시작 행 찾기 (컬럼명이 있는 행 = row 4)
        profiles = []
        
        for idx in range(5, len(df)):  # 5부터 시작 (0-indexed)
            row = df.iloc[idx]
            
            # 비어있는 행 스킵
            if pd.isna(row[1]):
                continue
            
            try:
                resident_id = str(row[1]).strip()
                name = resident_id  # 이름 정보 없으면 ID 사용
                sex_val = str(row[2]).strip()
                sex = "M" if sex_val == "남" else "F"
                age = int(row[3])
                height = float(row[4])
                weight = float(row[5]) if not pd.isna(row[5]) else 55.0
                
                # 질환 파싱
                diseases = []
                if not pd.isna(row[8]) and str(row[8]).strip() == "O":
                    diseases.append("당뇨")
                if not pd.isna(row[9]) and str(row[9]).strip() == "O":
                    diseases.append("고혈압")
                if not pd.isna(row[10]) and str(row[10]).strip() == "O":
                    diseases.append("신장질환")
                if not pd.isna(row[11]) and str(row[11]).strip() == "O":
                    diseases.append("연하장애")
                
                # 현재식사현황 파싱
                current_meal = str(row[13]).strip() if not pd.isna(row[13]) else "일반밥/일반찬"
                
                # 밥 타입 추출
                if "죽식" in current_meal:
                    meal_type = "죽식"
                elif "미음" in current_meal:
                    meal_type = "미음"
                else:
                    meal_type = "일반밥"
                
                # 반찬 타입 추출
                if "갈찬" in current_meal:
                    side_type = "갈찬"
                elif "다진찬" in current_meal:
                    side_type = "다진찬"
                else:
                    side_type = "일반찬"
                
                # 저작 단계 (활동정도 기반)
                activity = str(row[6]).strip() if not pd.isna(row[6]) else "3단계"
                if "1단계" in activity or "누워" in str(row[17]):
                    chewing = 1
                elif "2단계" in activity or "휠체어" in str(row[17]):
                    chewing = 2
                else:
                    chewing = 3
                
                profile = ElderlyProfile(
                    resident_id=resident_id,
                    name=name,
                    age=age,
                    sex=sex,
                    height_cm=height,
                    weight_kg=weight,
                    pal=1.4,  # 기본값
                    diseases=diseases if diseases else [],
                    chewing_stage=chewing,
                    current_meal=current_meal,
                    meal_type=meal_type,
                    side_type=side_type
                )
                
                profiles.append(profile)
                
            except Exception as e:
                continue
        
        return profiles
    
    except Exception as e:
        st.error(f"❌ 엑셀 파일 파싱 오류: {str(e)}")
        return []


# ============================================================================
# 4. 영양 계산
# ============================================================================

def compute_daily_requirement(profile: ElderlyProfile):
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
    
    return {
        "energy": energy,
        "energy_range": (energy_min, energy_max),
        "protein": (p_min, p_max),
        "fat": (f_min, f_max),
        "carbs": (c_min, c_max),
        "sodium": (0.0, sodium_max),
        "fiber": (20.0, 999.0),
    }


# ============================================================================
# 5. 데이터 로드 (캐시)
# ============================================================================

@st.cache_resource
def load_recipe_data():
    """932개 레시피 데이터 로드"""
    try:
        df = pd.read_excel("merged_nutrient_with_category.xlsx")
        df.columns = [
            'food_id', 'name', 'energy', 'protein', 'fat', 'carbs',
            'sodium', 'calcium', 'vitamin_d', 'fiber', 'category'
        ]
        return df
    except:
        return None


# ============================================================================
# 6. 프롬프트 생성
# ============================================================================

def build_advanced_prompt(
    profile: ElderlyProfile,
    requirement: dict,
    recipe_df: pd.DataFrame
) -> str:
    """고급 프롬프트 - 조리법 및 선택 이유 포함"""
    
    # 조리법 설명
    cooking_instruction = ""
    if profile.side_type == "다진찬":
        cooking_instruction = "모든 반찬은 잘게 다져서 제공하세요. 삼키기 쉽도록 준비하세요."
    elif profile.side_type == "갈찬":
        cooking_instruction = "모든 반찬은 고운 체로 걸러 부드럽게 갈아서 제공하세요. 연하곤란을 위해 준비하세요."
    
    # 밥 종류
    rice_type = ""
    if profile.meal_type == "미음":
        rice_type = "미음 (끓인 밥을 곱게 갈아 물을 섞은 것)"
    elif profile.meal_type == "죽식":
        rice_type = "죽 (밥 1:물 5 비율의 쌀죽)"
    else:
        rice_type = "일반 쌀밥"
    
    diseases = ", ".join(profile.diseases) if profile.diseases else "없음"
    
    prompt = f"""당신은 요양원의 영양사입니다. 
특정 어르신을 위한 7일 맞춤형 식단을 설계해야 합니다.

【대상 어르신】
- ID: {profile.resident_id}
- 나이: {profile.age}세, 성별: {"남" if profile.sex == "M" else "여"}
- 신체: {profile.height_cm}cm, {profile.weight_kg}kg
- 질환: {diseases}
- 현재 식사 현황: {profile.current_meal}

【식사 요구사항】
- 밥 종류: {rice_type}
- 반찬 조리법: {cooking_instruction}
- 영양 목표 에너지: {requirement['energy']:.0f}kcal
- 나트륨 제한: {requirement['sodium'][1]:.0f}mg 이하
- 식이섬유: 20g 이상

【중요 지시사항】
1. 각 끼니는 반드시 다음 구성으로:
   - {rice_type}
   - 국 또는 죽 (밥이 미음이면 생략 가능)
   - 주찬 1개
   - 부찬 2개
   - 김치

2. 조리법을 명시하세요:
   - 다진찬인 경우: "○○○ (다져서 제공)"
   - 갈찬인 경우: "○○○ (곱게 갈아서 제공)"

3. 각 메뉴 선택 이유를 설명하세요:
   - 왜 이 음식을 선택했는가
   - 어떻게 이 음식이 {profile.resident_id} 어르신의 건강에 도움이 되는가
   - 예: "고혈압이 있으신 분이라 나트륨이 낮은 음식 선택"

【출력 형식】
다음 JSON 형식만 출력:
{{
  "resident_id": "{profile.resident_id}",
  "weekly_plan": [
    {{
      "day": 1,
      "day_name": "월요일",
      "meals": [
        {{
          "meal_type": "breakfast",
          "items": {{
            "밥": [{{
              "name": "쌀죽",
              "cooking_method": "죽식 (밥 1:물 5)",
              "reason": "연하곤란이 있으신 분이라 부드러운 죽으로 준비"
            }}],
            "국": [{{"name": "계란국", "cooking_method": "다져서 제공", "reason": "..."}},
            ...
          }}
        }},
        ...
      ]
    }},
    ...
  ],
  "summary": {{
    "nutritional_notes": "이 식단의 영양학적 특징",
    "special_preparations": "특별히 준비할 사항"
  }}
}}
"""
    
    return prompt


# ============================================================================
# 7. LLM 호출
# ============================================================================

@st.cache_resource
def get_openai_client():
    load_dotenv()
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        st.error("❌ OPENAI_API_KEY 환경변수 필요")
        st.stop()
    return OpenAI(api_key=api_key)


def extract_json(text: str) -> dict:
    """JSON 추출"""
    start = text.find("{")
    end = text.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError("JSON을 찾을 수 없습니다")
    
    json_str = text[start:end]
    return json.loads(json_str)


def generate_meal_plan(
    profile: ElderlyProfile,
    requirement: dict
) -> Optional[dict]:
    """식단 생성"""
    
    recipe_df = load_recipe_data()
    if recipe_df is None:
        st.error("❌ 레시피 데이터를 로드할 수 없습니다")
        return None
    
    prompt = build_advanced_prompt(profile, requirement, recipe_df)
    
    try:
        client = get_openai_client()
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "당신은 요양원 영양사입니다. JSON만 출력하세요."
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=4000
        )
        
        content = response.choices[0].message.content
        return extract_json(content)
        
    except Exception as e:
        st.error(f"❌ LLM 오류: {str(e)}")
        return None


# ============================================================================
# 8. Streamlit UI
# ============================================================================

def main():
    st.title("🍽️ 요양원 어르신 맞춤형 식단 자동 생성 시스템")
    st.markdown("**대량의 입소자 데이터를 한번에 처리하여 7일 식단 자동 생성**")
    st.divider()
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["📤 데이터 업로드", "🤖 식단 생성", "📊 결과"])
    
    # ========== TAB 1: 데이터 업로드 ==========
    with tab1:
        st.header("요양원 고령자 데이터 업로드")
        
        st.info("""
        📋 요구 형식:
        - 엑셀 파일 (.xlsx)
        - 시트명: "고령자 데이터"
        - 컬럼: 수급자명, 성별, 나이, 신장, 체중, 활동정도, 
                당뇨병, 고혈압, 신장질환, 연하장애, 현재식사현황
        """)
        
        uploaded_file = st.file_uploader("엑셀 파일 선택", type=['xlsx'])
        
        if uploaded_file:
            with st.spinner("📂 파일 분석 중..."):
                profiles = parse_elderly_excel(uploaded_file)
            
            if profiles:
                st.success(f"✅ {len(profiles)}명의 어르신 정보 로드 완료!")
                
                # 프로필 미리보기
                st.subheader("📋 로드된 어르신 정보")
                preview_data = []
                for p in profiles[:10]:  # 처음 10명만 표시
                    preview_data.append({
                        "ID": p.resident_id,
                        "나이": f"{p.age}세",
                        "성별": "남" if p.sex == "M" else "여",
                        "체중": f"{p.weight_kg}kg",
                        "질환": ", ".join(p.diseases) if p.diseases else "없음",
                        "현재식사": p.current_meal
                    })
                
                df_preview = pd.DataFrame(preview_data)
                st.dataframe(df_preview, use_container_width=True)
                
                if len(profiles) > 10:
                    st.caption(f"... 외 {len(profiles) - 10}명 더 있음")
                
                # 세션에 저장
                st.session_state.profiles = profiles
                st.session_state.file_loaded = True
            else:
                st.error("❌ 파일 파싱 실패")
    
    # ========== TAB 2: 식단 생성 ==========
    with tab2:
        st.header("🤖 식단 자동 생성")
        
        if not st.session_state.get('file_loaded', False):
            st.warning("⚠️ 먼저 [📤 데이터 업로드] 탭에서 파일을 업로드하세요")
        else:
            profiles = st.session_state.profiles
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.info(f"📋 {len(profiles)}명의 어르신을 위한 식단을 생성합니다 (약 5-10분)")
            
            with col2:
                if st.button("🚀 식단 생성 시작!", use_container_width=True, type="primary"):
                    st.session_state.generating = True
            
            if st.session_state.get('generating', False):
                # 진행 상황 표시
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                all_results = {}
                
                for idx, profile in enumerate(profiles):
                    # 진행도 업데이트
                    progress = (idx + 1) / len(profiles)
                    progress_bar.progress(progress)
                    status_text.text(f"🔄 {idx + 1}/{len(profiles)} - {profile.resident_id} 어르신 식단 생성 중...")
                    
                    # 영양 기준 계산
                    requirement = compute_daily_requirement(profile)
                    
                    # 식단 생성
                    meal_plan = generate_meal_plan(profile, requirement)
                    
                    if meal_plan:
                        all_results[profile.resident_id] = {
                            "profile": profile,
                            "requirement": requirement,
                            "meal_plan": meal_plan
                        }
                
                status_text.text(f"✅ {len(all_results)}/{len(profiles)} 명 식단 생성 완료!")
                
                # 결과 저장
                st.session_state.all_results = all_results
                st.session_state.generating = False
                
                st.success(f"🎉 총 {len(all_results)}명의 식단 생성 완료!")
    
    # ========== TAB 3: 결과 ==========
    with tab3:
        st.header("📊 생성된 식단 확인")
        
        if not st.session_state.get('all_results'):
            st.warning("⚠️ 먼저 식단을 생성하세요")
        else:
            all_results = st.session_state.all_results
            
            # 어르신 선택
            resident_ids = list(all_results.keys())
            selected_id = st.selectbox("어르신 선택", resident_ids)
            
            if selected_id:
                result = all_results[selected_id]
                profile = result['profile']
                meal_plan = result['meal_plan']
                
                # 프로필 정보
                st.subheader(f"👤 {profile.resident_id} 어르신")
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("나이", f"{profile.age}세")
                with col2:
                    st.metric("체중", f"{profile.weight_kg}kg")
                with col3:
                    st.metric("식사 형태", profile.meal_type)
                with col4:
                    st.metric("반찬 형태", profile.side_type)
                with col5:
                    st.metric("질환", len(profile.diseases))
                
                # 7일 식단 표시
                st.divider()
                st.subheader("📅 7일 식단")
                
                if "weekly_plan" in meal_plan:
                    for day_data in meal_plan["weekly_plan"]:
                        day_num = day_data["day"]
                        day_name = day_data.get("day_name", ["월", "화", "수", "목", "금", "토", "일"][day_num-1])
                        
                        with st.expander(f"🗓️ {day_num}일 ({day_name}요일)", expanded=(day_num==1)):
                            
                            for meal in day_data.get("meals", []):
                                meal_type = meal["meal_type"]
                                meal_names = {
                                    "breakfast": "🌅 아침",
                                    "lunch": "🍽️ 점심",
                                    "dinner": "🌆 저녁",
                                    "snack": "☕ 간식"
                                }
                                
                                st.markdown(f"**{meal_names.get(meal_type, meal_type)}**")
                                
                                # 카테고리별 음식
                                items = meal.get("items", {})
                                for category, foods in items.items():
                                    if foods:
                                        st.markdown(f"*{category}*:")
                                        for food in foods:
                                            if isinstance(food, dict):
                                                name = food.get("name", "")
                                                cooking = food.get("cooking_method", "")
                                                reason = food.get("reason", "")
                                                
                                                if cooking and reason:
                                                    st.write(f"  • **{name}** ({cooking})")
                                                    st.caption(f"    💡 {reason}")
                                                else:
                                                    st.write(f"  • {name}")
                                            else:
                                                st.write(f"  • {food}")
                                
                                st.divider()
            
            # 일괄 다운로드
            st.divider()
            st.subheader("💾 결과 다운로드")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # JSON 다운로드
                json_data = json.dumps(all_results, ensure_ascii=False, indent=2, default=str)
                st.download_button(
                    label="📥 전체 JSON 다운로드",
                    data=json_data,
                    file_name="meal_plans_all.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            with col2:
                # Excel 다운로드
                if st.button("📊 Excel로 내보내기", use_container_width=True):
                    st.info("Excel 생성 중...")
                    # 구현은 별도로 진행

if __name__ == "__main__":
    # 세션 상태 초기화
    if 'profiles' not in st.session_state:
        st.session_state.profiles = None
    if 'all_results' not in st.session_state:
        st.session_state.all_results = None
    if 'file_loaded' not in st.session_state:
        st.session_state.file_loaded = False
    
    main()

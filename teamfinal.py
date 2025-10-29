import streamlit as st
from openpyxl import *
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) #현재 파일 경로 설정->엑셀/이미지 절대경로 생성

def load_recipe_data():
    try:
        excel_path = os.path.join(BASE_DIR, "recipes.xlsx")
        wb = load_workbook(excel_path)
        ws = wb.active
        recipes = []    #row단위로 레시피 dic생성
        for row in ws.iter_rows(min_row=2, values_only=True):   
            row = list(row) + [None] * (6 - len(row))
            time_tag, main, subs, name, steps, image = row
            if not name or not main:
                continue
            subs_list = []  
            if subs:
                for part in str(subs).split(";"):   #서브재료 ; 기준으로 구분
                    s = part.strip()
                    if s and s.lower() != "none":
                        subs_list.append(s)

            image_val = None    #이미지 파일 경로 넣을 변수
            image_original = None   #엑셀 원본경로 보존 변수
            if image:
                image_original = str(image).strip()
                # 엑셀에는 assets/xxx.jpg(상대경로)로 적혀있음
                candidate = os.path.join(BASE_DIR, image_original)
                # 정규화
                candidate = os.path.normpath(candidate)
                if os.path.exists(candidate) and os.path.isfile(candidate):
                    image_val = candidate  # 존재하면 절대경로로 저장
                else:
                    # 파일이 존재하지 않으면 None으로 
                    image_val = None

            recipe = {
                "time_tag": str(time_tag).strip() if time_tag else "",
                "main": str(main).strip(),
                "ingredients": [str(main).strip()] + [s.strip() for s in str(subs).split(";") if s],
                "name": str(name).strip(),
                "steps": [s.strip() for s in str(steps).split("|") if s],
                "image": image_val,"image_original": image_original
            }
            recipes.append(recipe)
        wb.close()
        return recipes
    except Exception as e:
        st.error(f"엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}")
        return []

# 레시피 DB 전역 로드
RECIPE_DB = load_recipe_data()

# 초기화
if "confirmed" not in st.session_state: #confirmed가 없으면 새 키 만들고 False로 설정, '진행합니다'버튼 누르기 전인지 여부 구분 위함
    st.session_state.confirmed = False
if "time_selected" not in st.session_state:
    st.session_state.time_selected = False   # 시간 입력 상태 초기화
if "main_stage" not in st.session_state:    #메인재료 선택 단계 플래그
    st.session_state.main_stage = False
if "time" not in st.session_state:
    st.session_state.time = ""
if "selected_time" not in st.session_state:  # datetime.time 저장용 새 변수
    st.session_state.selected_time = None
if "meal_time" not in st.session_state:     # 아침/점심/저녁 문자열 저장용 새 변수
    st.session_state.meal_time = None


if not st.session_state.main_stage:
    st.title('[레이어드 : 냉장고를 부탁해🥪]')
    st.header('🧊냉장고에 있는 재료를 선택하고 요리하기')


if not st.session_state.time_selected:
    selected_time = st.time_input("식사 시간을 선택해주세요", value=st.session_state.selected_time)
    if selected_time:
        st.session_state.selected_time = selected_time
        hour = selected_time.hour
        if 6 <= hour < 11:
            meal_time = "아침"
        elif 11 <= hour < 15:
            meal_time = "점심"
        elif 17 <= hour < 22:
            meal_time = "저녁"
        else:
            st.warning("선택하신 시간이 지정된 식사 시간(아침 6-11, 점심 11-15, 저녁 17-22) 범위 밖입니다. 범위 내 시간을 선택해주세요.")
            st.stop()  # 시간 범위 밖이면 진행 중단

        st.session_state.meal_time = meal_time  #아침/점심/저녁 문자열 저장
        st.write(f"자동으로 **{meal_time}** 시간대로 설정되었습니다. 진행하시겠습니까?")


        col1, col2=st.columns(2)
        with col1 :
            if st.button('변경을 원합니다'):
                st.session_state.selected_time = None  # 초기화
                st.session_state.meal_time = None
                st.session_state.time_selected = False
                st.session_state.confirmed = False
                st.session_state.main_stage = False
                st.rerun()
        with col2 :
            if st.button('진행합니다'):
                st.session_state.time_selected = True
                st.session_state.confirmed = True
                st.session_state.main_stage=True
                st.session_state.step = 1
                st.rerun()

elif st.session_state.main_stage:

    def score_recipe(user_has, recipe):
        need = set(recipe["ingredients"])
        have = sorted(need & user_has)
        missing = sorted(need - user_has)
        return {
            "name": recipe["name"],
            "have": have,
            "missing": missing,
            "have_count": len(have),
            "need_count": len(need),
            "steps": recipe["steps"]
        }
    # ---------------------------
    # 재료 선택 페이지
    # ---------------------------
    def page_select():
        st.subheader("1) 메인 재료 선택🍖")

        main_options = sorted(set([r["main"] for r in RECIPE_DB]))
        main_options = ["- 선택 -"] + main_options

        main_selected = st.selectbox("냉장고 속 재료 중 **메인 재료**로 사용할 아이템을 1개 골라주세요", options=main_options, index=0)
        
        st.subheader("2) 서브 재료 선택🥗")
        st.caption("**서브 재료**를 2~3개 골라보세요.")

        all_subs = set()
        if main_selected != "- 선택 -":
            for r in RECIPE_DB:
                if r["main"] == main_selected:
                    all_subs.update(r["ingredients"])
            all_subs.discard(main_selected)

        all_subs = sorted(all_subs)
        cols_chips = st.columns(4)
        chosen_subs = set()
        MAX_SUB = 3
        for i, item in enumerate(all_subs):
            col = cols_chips[i % 4]
            checked = col.checkbox(item, value=st.session_state.get(f"sub_{item}", False), key=f"sub_{item}")
            if checked:
                chosen_subs.add(item)

        if len(chosen_subs) > MAX_SUB:
            st.warning(f"서브재료는 최대 {MAX_SUB}개까지만 선택할 수 있습니다.")

        # 다음 버튼
        if st.button("다음 ➜"):
            if main_selected == "- 선택 -":
                st.warning("메인 재료를 먼저 선택해 주세요.")
                return
            if not (2 <= len(chosen_subs) <= MAX_SUB):
                st.warning("서브 재료는 2~3개를 선택해 주세요.")
                return
            # 선택 저장
            st.session_state.main_selected = main_selected
            st.session_state.chosen_subs = list(chosen_subs)  # set은 저장/직렬화 이슈 있으니 list로
            st.session_state.step = 2
            st.rerun()

                
    def page_result():
        st.subheader("👩‍🍳레시피 추천 결과👩‍🍳")
        main_selected = st.session_state.get("main_selected", "- 선택 -")
        chosen_subs = set(st.session_state.get("chosen_subs", []))
        user_has = {main_selected} | chosen_subs
        meal_time=st.session_state.meal_time

        candidates = []
        for r in RECIPE_DB:
            if (
                r["main"] == main_selected
                and (len(set(r["ingredients"]) & chosen_subs) >= 1)
                and (r["time_tag"] == meal_time)
            ):
                candidates.append(r)

        if not candidates:
            st.info("조건에 맞는 레시피가 아직 없어요. 이전 화면에서 재료를 바꿔보세요.")
        else:
            scored = [score_recipe(user_has, r) for r in candidates]    #재료 계산 
            scored.sort(key=lambda x: (-x["have_count"], len(x["missing"]), x["name"]))

            for rec in scored:
                recipe_info = next((r for r in RECIPE_DB if r["name"] == rec["name"]), None)

                if recipe_info and recipe_info.get("image"):
                    img_path = recipe_info["image"]
                    if img_path and os.path.exists(img_path):
                        st.image(use_column_width=True)

                st.markdown(f"### {rec['name']}")
                st.write(f"**보유 재료:** {', '.join(rec['have']) if rec['have'] else '없음'}")
                if rec["missing"]:
                    st.write(f"**부족 재료:** {', '.join(rec['missing'])}")
                else:
                    st.success("필요한 재료가 모두 있어요! 바로 조리 가능합니다 :짠:")
                with st.expander("조리 단계 보기"):
                    for i, step in enumerate(rec["steps"], start=1):
                        st.write(f"{i}. {step}")
                st.markdown("---")

        if st.button("← 이전"):
            st.session_state.step = 1
            st.rerun()

    # main 실행 흐름
    def main():
        # 초기 단계
        if "step" not in st.session_state:
            st.session_state.step = 1   #첫실행시 초기화
        if st.session_state.step == 1:  #재료고르기
            page_select()
        elif st.session_state.step == 2:    #결과페이지 실행
            page_result()
        else:
            st.session_state.step = 1
            page_select()
    if __name__ == "__main__":
        main()


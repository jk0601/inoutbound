import io
import openpyxl
import streamlit as st

from parser import parse_orders, get_period
from generator import generate_수불부

# ── 페이지 설정 ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="소모품 수불부 자동 생성",
    page_icon="📋",
    layout="centered",
)

st.title("📋 소모품 수불부 자동 생성")
st.caption("주문내역 엑셀 파일을 업로드하면 수불부를 자동으로 만들어 드립니다.")
st.divider()

# ── 파일 업로드 ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "주문내역 엑셀 파일 선택 (.xlsx)",
    type=["xlsx"],
    help="첫 번째 시트에 주문내역이 있는 파일을 올려주세요.",
)

if not uploaded:
    st.info("👆 파일을 업로드하면 시작됩니다.")
    st.stop()

# ── 파일 파싱 ─────────────────────────────────────────────────────────────────
raw_bytes = uploaded.read()
wb = openpyxl.load_workbook(io.BytesIO(raw_bytes))
ws = wb.worksheets[0]

orders = parse_orders(ws)

if not orders:
    st.error("주문내역 데이터를 읽지 못했습니다. 파일 형식을 확인해 주세요.")
    st.stop()

period = get_period(ws)

st.success(f"✅ **{period}** 주문내역  |  총 **{len(orders)}건** 파싱 완료")

# ── 주문내역 미리보기 ─────────────────────────────────────────────────────────
with st.expander("📂 주문내역 미리보기", expanded=True):
    # 분류별 집계 요약
    from collections import Counter
    cat_counts = Counter(o["분류"] for o in orders)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**분류별 품목 수**")
        for cat, cnt in cat_counts.items():
            st.markdown(f"- {cat}: **{cnt}건**")
    with col2:
        st.markdown("**전체 금액 합계**")
        total = sum(o["금액"] for o in orders)
        st.markdown(f"합계: **{total:,}원**")

    st.markdown("---")
    st.markdown("**전체 목록**")
    # 표 형태로 표시 (pandas 없이)
    rows_display = [
        {
            "No": i + 1,
            "분류": o["분류"],
            "품목명": o["품목명"],
            "규격": o["규격"],
            "단위": o["단위"],
            "일자": o["일자"],
            "수량": o["수량"],
            "금액(원)": f"{o['금액']:,}",
        }
        for i, o in enumerate(orders)
    ]
    st.dataframe(rows_display, use_container_width=True, hide_index=True)

st.divider()

# ── 수불부 생성 ───────────────────────────────────────────────────────────────
st.subheader("수불부 생성")

if st.button("⚙️ 수불부 생성하기", type="primary", use_container_width=True):
    with st.spinner("수불부를 생성하는 중..."):
        # 원본 워크북에 수불부 시트 추가
        wb_out = openpyxl.load_workbook(io.BytesIO(raw_bytes))
        orders_fresh = parse_orders(wb_out.worksheets[0])
        generate_수불부(wb_out, orders_fresh, period)

        buf = io.BytesIO()
        wb_out.save(buf)
        buf.seek(0)

    st.success("✅ 수불부 생성 완료!")

    # 생성된 수불부 요약 표시
    with st.expander("📊 생성된 수불부 요약", expanded=True):
        from collections import OrderedDict

        def _group(orders):
            g = OrderedDict()
            for o in orders:
                g.setdefault(o["분류"], []).append(o)
            return g

        groups = _group(orders_fresh)
        summary_rows = []
        for cat, items in groups.items():
            summary_rows.append({
                "분류": cat,
                "입고 건수": len(items),
                "입고 수량": sum(i["수량"] for i in items),
                "입고 금액(원)": f"{sum(i['금액'] for i in items):,}",
            })
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)

    # 다운로드 버튼
    output_filename = f"{period}_수불부.xlsx"
    st.download_button(
        label="📥 수불부 다운로드",
        data=buf,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

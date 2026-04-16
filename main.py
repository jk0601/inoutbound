"""
수불부 자동 생성 도구

사용법:
    python main.py <주문내역_엑셀파일>
    python main.py <주문내역_엑셀파일> [출력파일]

예:
    python main.py example.xlsx
    python main.py 2026_04월주문내역.xlsx 2026_04월수불부.xlsx

동작:
    1. 입력 파일의 첫 번째 시트에서 주문내역을 파싱
    2. 수불부 시트를 자동 생성하여 입력 파일의 모든 시트와 함께 별도 파일로 저장
    3. 출력 파일명을 지정하지 않으면 '{원본파일명}_수불부.xlsx' 로 저장

출고 수량 수정 방법:
    생성된 수불부의 K열(출고수량)은 기본적으로 입고수량(H열)과 동일한 수식(=H{행})으로
    설정됩니다. 실제 출고량이 다를 경우 해당 셀에 숫자를 직접 입력하면
    L열(출고금액), M열(재고)이 자동으로 재계산됩니다.
"""
import sys
from pathlib import Path

import openpyxl

from parser import parse_orders, get_period
from generator import generate_수불부


def main():
    if len(sys.argv) < 2:
        print("사용법: python main.py <주문내역_엑셀파일> [출력파일]")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"오류: 파일을 찾을 수 없습니다 — {input_path}")
        sys.exit(1)

    output_path = (
        Path(sys.argv[2])
        if len(sys.argv) >= 3
        else input_path.parent / f"{input_path.stem}_수불부{input_path.suffix}"
    )

    print(f"입력 파일: {input_path}")

    wb = openpyxl.load_workbook(input_path)
    ws = wb.worksheets[0]  # 첫 번째 시트 = 주문내역

    orders = parse_orders(ws)
    if not orders:
        print("오류: 주문내역 데이터를 찾을 수 없습니다. 날짜(datetime) 형식을 확인하세요.")
        sys.exit(1)

    period = get_period(ws)
    print(f"관리기간: {period}  /  총 {len(orders)}건 파싱 완료")

    generate_수불부(wb, orders, period)

    wb.save(output_path)
    print(f"저장 완료: {output_path}")


if __name__ == '__main__':
    main()

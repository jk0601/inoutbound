"""
주문내역 엑셀 시트 파싱 모듈

컬럼 순서 (0-based):
  0: 일자, 1: 주문명, 2: 프로그램명, 3: 대분류, 4: 중분류,
  5: 구분, 6: 상품내역, 7: 수량, 8: 단위, 9: 공급가액, 10: 부가세, 11: 합계
"""
import re
from datetime import datetime

# 규격에서 제외할 보관조건 키워드
STORAGE_KEYWORDS = {'상온', '냉장', '냉동', '실온', '냉장보관', '냉동보관'}


def strip_category_prefix(category: str) -> str:
    """[비], [식] 등 접두어 제거.  예: '[비]세제류' → '세제류'"""
    return re.sub(r'^\[.\]', '', str(category)).strip()


def parse_product(상품내역: str):
    """
    상품내역 문자열을 (품목명, 규격)으로 분리.
    괄호 안의 쉼표를 보호하여 분리한다.

    예:
      '식기세척기용세제,18L,상온,하이코리아' → ('식기세척기용세제', '18L, 하이코리아')
      '니트릴장갑(블루,M),100매,중국산'      → ('니트릴장갑(블루,M)', '100매, 중국산')
    """
    parts = re.split(r',(?![^(]*\))', 상품내역.strip())
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) == 1:
        return parts[0], ''

    품목명 = parts[0]
    규격_parts = [p for p in parts[1:] if p not in STORAGE_KEYWORDS]
    규격 = ', '.join(규격_parts)
    return 품목명, 규격


def _parse_date(value):
    """datetime 객체 또는 날짜 문자열을 datetime으로 변환. 실패 시 None."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y'):
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                continue
    return None


def parse_orders(ws) -> list:
    """
    주문내역 시트에서 주문 목록 파싱.
    날짜(datetime 또는 날짜 문자열) 값이 있는 행만 처리.
    """
    orders = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        date = _parse_date(row[0])
        if not date:
            continue
        try:
            중분류 = row[4]
            상품내역 = row[6]
            if not 중분류 or not 상품내역:
                continue

            분류 = strip_category_prefix(str(중분류))
            품목명, 규격 = parse_product(str(상품내역))

            orders.append({
                '분류': 분류,
                '품목명': 품목명,
                '규격': 규격,
                '단위': str(row[8]) if row[8] else '',
                '일자': date.strftime('%m/%d'),
                '수량': int(row[7]) if row[7] else 0,
                '금액': int(row[11]) if row[11] else 0,
            })
        except Exception:
            continue
    return orders


def get_period(ws) -> str:
    """주문내역의 첫 번째 날짜로부터 관리기간 추출.  예: '2026년 3월'"""
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        date = _parse_date(row[0])
        if date:
            return f'{date.year}년 {date.month}월'
    return ''

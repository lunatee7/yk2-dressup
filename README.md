# 개요

용과 같이: 극 2 (Yakuza Kiwami 2)의 컨텐츠인 물장사 아일랜드의 캐스트 드레스 세팅을 도와줍니다.

# 실행 방법

pandas와 tqdm 패키지의 설치가 필요합니다.

```
pip install pandas
pip install tqdm
python yk2_dressup.py
```

# 사용 방법

코드를 실행 후 생성 된 YK2_Dressup.xlsx 파일을 엑셀 또는 Google 스프레드시트를 이용하여 불러온 후, 필터 기능을 통해 자신이 원하는 드레스 세팅을 합니다.

# 참고 사항

- 데이터가 너무 많으면 데이터를 전부 불러오지 못하는 엑셀의 문제 때문에 머리와 드레스 목록은 기본만을 지원합니다.
- yk2_dressup.py 파일이 위치한 곳에 **dressup.xlsx** 또한 존재해야 합니다.

# 업데이트 내역

- 0.1.0
  - 첫 릴리즈

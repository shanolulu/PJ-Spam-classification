# 텍스트 마이닝 기반 스팸메일 분석기

> 텍스트 마이닝 기반 스팸메일 분석기술 연구 (정보보호학회 하계학술대회 논문집 Vol.29, No. 1)

> 메일의 스팸 여부를 판단하는 프로그램
네이버 메일을 스크레이핑하여 SVM알고리즘으로 학습하된 모델을 통해 스팸여부 판단 및 스팸처리된 메일을 네이버 메일함에서 휴지통으로 이동시키는 기능

1. 프로그램이 실행되면 tkinter을 통해 gui가 생성
2. 회원 아이디와 비밀번호를 입력
3. selenium의 webdriver을 이용하여 크롬브라우저를 사용
4. 네이버에 로그인해 메일을 스크레이핑
5. 메일의 스팸여부 분석
6. 선택에 따라 스팸의 삭제 가능

! webdriver를 사용하는 브라우저의 버전에 맞도록 다운받아야 한다.
> 크롬 브라우저를 사용했기 떄문에 chromedriver를 다운받아 사용
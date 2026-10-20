# ============================================================
# PMS2 v0.6.28 | Ollama 인공지능의견 / 모바일 / 다중 PC 접속
# ============================================================
# [실행] 이 파일을 app.R 로 저장하고 pms.sqlite 와 같은 폴더에 둡니다.
#        RStudio에서 Run App 클릭 또는 R 콘솔: shiny::runApp(".")
# [서버] 기본 설정: 0.0.0.0:3838 (다른 PC에서도 접속 허용).
#        서버 PC 접속: http://localhost:3838
#        같은 LAN의 다른 PC/스마트폰: http://서버PC의사설IP:3838
#        서버 IP 확인(Windows): cmd 실행 후 ipconfig -> IPv4 주소 확인.
#        Windows 방화벽에서 TCP 3838 인바운드 허용이 필요할 수 있습니다.
#        다른 네트워크/인터넷 접속은 공유기 포트포워딩 또는 VPN이 별도 필요합니다.
# [화면] PC: 기존 탭/화면 유지, 스마트폰(폭 768px 이하): 하단 메뉴/전체메뉴.
# [주의] 사용자 인증 없이 0.0.0.0으로 열면 해당 네트워크의 다른 사람이
#        금융정보를 볼 수 있습니다. 인터넷에 포트를 무방비로 공개하지 마세요.
# [주의] 서버 PC가 켜져 있고 R Shiny가 실행 중이어야 접속할 수 있습니다.
# ============================================================

# PMS2 v0.6.28: future 직접 실행 + Shiny 완료 폴링으로 promises 콜백 의존 제거.
# PMS2 v0.6.31: AI 응답 길이 슬라이더 (짧게/중간/길게) 및 요청량 연동.
# PMS2 v0.6.30: AI 분석 요청 경량화 (상위 12종목, 생성 250토큰, 사용자 메모/질문 750자).
# PMS2 v0.6.32: AI 의견 Markdown(**굵게**, 제목, 목록) 화면 렌더링.
# PMS2 v0.6.26: AI 버튼 클릭/서버 접수 상태 분리, 호출 진단 로그, 비동기 실행.
# PMS2 v0.6.25: AI API를 future_promise 비동기 호출로 분리; 다른 탭/초기 화면 유지.
# 필요 패키지: install.packages(c("future", "promises"))
# PMS2 v0.6.18: add in-cell allocation data bars to board portfolio-weight column.
# PMS2 v0.6.15: unified header totals and remove duplicate auto-status.
# PMS2 v0.6.14: header portfolio NAV and prior-close changes from board valuation rows.
# PMS2 v0.6.12: fix spurious tab created by tags$head nested in tabsetPanel; v0.6.11 KRW formatting retained.
# PMS2 v0.6.09: legible board clock in dark mode; labeled horizontal allocation bars.
# PMS2 v0.6.08 - reliable latest fetch timestamp via direct browser message; horizontal allocation bars.
# PMS2 v0.6.07: independent SQLite quote fetch clock; two 100% stacked diversification bars.
# PMS2 v0.6.06: diversification auto-classification module companion; v0.6.05:  show SQLite prices.fetched_at KST alongside market quote time.
# PMS2 v0.6.01: native DT row selection repairs holdings editing, preserves DB IDs; broker choice stability.
# PMS2 v0.6.04: delegated row click edit, restored clock; v0.6.02: explicit edit buttons and ID selector; holdings selection repair; dividend 20M max; timestamp; diversification.
# PMS2 v0.5.96 (2026-10-18): 보고서 탭 3분 한국어 음성 브리핑(Web Speech API, 재생/일시정지/중지).
# PMS2 v0.5.94 (2026-10-18): fix startup spinner staying on screen; jQuery Shiny events + rendered DOM detection + 15s failsafe.
# PMS2 v0.5.91 (2026-10-18): 엑셀 다운로드 상태표시 및 Shiny 출력 활성화 보강.
# PMS2 v0.5.90 (2026-10-18): 보유종목 Excel 다운로드·검증 미리보기·트랜잭션 일괄 업로드.
# PMS2 v0.5.89 (2026-10-18): 지금FIRE를한다면 탭 · 인출률별 세금 및 건보료 추정.
# PMS2 v0.5.88 (2026-10-18): 투자성과분석 평가액지수/낙폭 축 단위 수정; 상세 DT 손익 부호별 색상.
# PMS2 v0.5.87 (2026-10-18): Chrome 새 창 최대화 실행 및 Quarto PDF 벤치마크 표 고정폭(Sharpe/Sortino/Calmar 줄바꿈 방지).
# PMS2 v0.5.86 (2026-10-18): 전일대비전광판 총평가액 내림차순, 보유수량 반영 평가증감 동시 표시.
# PMS2 v0.5.85 (2026-10-18): 보고서 벤치마크 표 분리로 Sortino/Calmar 강제 줄바꿈 방지.
# PMS2 v0.5.85 (2026-10-18): 전일대비전광판 DT 천단위 구분 및 한국식 등락 색상.
# PMS2 v0.5.83 (2026-10-18): Quarto HTML/PDF 금융 통계표 수치 열 오른쪽 정렬.
# PMS2 v0.5.82 (2026-10-18): 보고서 생성 중 즉시 표시 스피너 및 완료/오류 자동 종료.
# PMS2 v0.5.81 (2026-10-18): PDF XeLaTeX/Malgun Gothic Korean font and report number formatting fix.
# ============================================================
# PMS2 - 개인금융자산 통계연구소
# Version : v0.5.88
# Date    : 2026-10-18
# 변경사항:
# v0.5.84: 전일대비전광판 DT 수량/금액 천단위 콤마, 상승 빨강·하락 파랑·보합 중립색.
# v0.5.83: 보고서 표의 금액·수익률·수량·통계 수치 우측 정렬, 이름·설명 왼쪽 정렬.
# v0.5.82: Quarto 보고서 생성 버튼 클릭 즉시 회전 애니메이션, 완료·오류 시 자동 숨김.
# v0.5.80: 동일기간 비교표 Calmar 지수 및 공식(연환산 Return / |MDD|) 추가.
# v0.5.79: 동일조건 투자성과 비교에 Return·MDD·Sharpe 동일관측일 지표 표 추가.
# v0.5.78: 홈 동일 초기금액·동일 추정 현금흐름의 S&P500 ETF 가상계좌 비교 그래프 추가.
# v0.5.77: ETF 전체기간 재수집 버튼, Yahoo→quantmod 이중 조회, 저장기간과 수집 실패 상태 표시.
# v0.5.76: 홈 전체기간 KODEX 미국S&P500·원달러 환율 누락 기간 자동 보충 및 SQLite 재사용.
# v0.5.75: 홈 총평가액·ETF·환율 비교 그래프에 매월 세로선 및 가독성에 맞춘 월 눈금 추가.
# v0.5.74: 보고서 PDF 다운로드(Quarto/LaTeX 필요), 전일대비전광판 아래 보유수량 기준 평가변화 DT 기본 100건.
# v0.5.73: 홈 월 예상 배당금 명시; 전체 평가액 + KODEX 미국S&P500 및 USD/KRW 보조축 시계열.
# v0.5.71: 홈 최근 1년 월평균 순매수 추정 및 평가액 증가액 추가.
# v0.5.70: 명칭 변경, 위험관리 지표 상세 설명/해석, 홈 전체 기록, 최근 1년 월평균 매수 추정금액.
# v0.5.69: 원본 PMS risk_module.R 기반 위험관리 탭: 추정 현금흐름 보정 성과, VaR/CVaR, 위험경보, Monte Carlo, 적립/인출 시나리오.
# v0.5.68: 절세계좌 세제 안내 및 개인 납입 목표 막대, Quarto 보고서 절세·최근 1년 매매 추정 통계 추가.
# v0.5.67: 연금저축·IRP·ISA 실제 납입액 기록 및 연간 한도 대비 막대그래프(수동 입력, SQLite 누적).
# v0.5.66: 이관자료/앱 스냅샷의 계좌 식별체계 차이로 발생하는 가짜 매도 방지, 배당 전용 탭 분리.
# v0.5.62: 투자성과분석 별도 탭, 평가액 기반 낙폭/변동성/관측 간 비율 참고치와 경고 추가.
# v0.5.60: 보고서 탭 최우측 이동, 금융소득 세후/원천징수 색상 그래프 및 세후 월평균·툴팁.
# v0.5.59: 일별 스냅샷을 SQLite 3개 테이블에만 누적(날짜 폴더/CSV 자동 생성 제거).
# v0.5.58: 자동 일별 SQLite+CSV 스냅샷, 보고서 탭 내 HTML/마크다운 미리보기.
# v0.5.57: 상단 탭 스크롤 고정, Quarto 종합 보고서(.qmd / HTML) 생성 탭.
# v0.5.56: 홈 section(title, ...) 다중 UI 인자 지원으로 unused argument 오류 수정.
# v0.5.55: 내종목뉴스 및 네트워크 조회 제거; SQLite 기반 홈 요약 대시보드 첫 탭 추가.
# v0.5.53: 네이버 curl 22 진단 개선, 모바일 종목별 뉴스 API 우선 및 기존 금융 뉴스 HTML 재시도.
# v0.5.52: 내종목뉴스 탭 진입 시 국내 개별주식 최근 3일 뉴스 자동 조회, 수동 새로고침 유지.
# v0.5.51: 국내주식 분류 OTHER 허용, ETF 상품명 접두어 판별, 영숫자 혼합 티커 처리 보완.
# v0.5.50: 국내 개별주식만 최근 3일 뉴스 제목 표시, ETF/해외 제외, 원문 링크 제공.
# v0.5.49: 내종목뉴스 표 정렬 오류 수정(order decreasing 길이 오류); 종목순서 오름차순·게시일 내림차순.
# v0.5.48: 내종목뉴스 탭: SQLite 보유 종목 자동 목록, 네이버 금융 뉴스 조회, 클릭 원문, 오류시 네이버 뉴스 검색 링크.
# v0.5.47: 월별 배당 막대 hover 상위 5종목 요약, 클릭 시 모든 종목을 스크롤 가능한 Shiny 모달 상세표로 표시.
# v0.5.46: 저장 시세 실제 기준일 노후화 진단(상장폐지 단정 금지), 월별 배당 막대 종목별 hover 풍선 도움말.
# v0.5.45: 소스 애니메이션 JS 선택자의 R 문자열 큰따옴표 충돌 수정 및 onStop 블록 명료화.
# v0.5.44: 수집 중 공급원별 애니메이션 및 시세/배당 소스별 성공 집계, 결과표 표시.
# v0.5.43: quantmod 시세/배당 우선 수집, 오류 시 기존 네이버/Yahoo 방식으로 폴백(quantmod 미설치 허용).
# v0.5.42: 일반계좌 배당금/2,000만원 기준선 직접 표시, 월별 로딩 메시지 제거.
# v0.5.41: 과거 12개월 월별 배당 추정액 직접 렌더링(집계표 + 월별 누적 시각화), 실패시 오류 표시.
# v0.5.40: 저장된 과거 배당 이벤트 상세를 SQLite에서 직접 출력하여 DT 미표시 문제 우회, 상태/오류 표시.
# v0.5.38: 배당 달력 Shiny 동적 출력 실패 우회: 상시 자리표시자 + insertUI 직접 렌더링 및 실시간 진단.
# v0.5.37: 예상 배당 달력 항상 렌더링, 저장 이벤트와 향후 30일 후보 건수 진단 및 오류 설명.
# v0.5.36: 저장 시세·배당 재평가 버튼 클릭/서버 수신/계산 결과 단계별 상태 표시. 서버 예외 메시지 노출 및 UI 감시.
# v0.5.35: 저장된 배당금정보로 평가 버튼 추가. SQLite 기존 배당·환율·보유수량으로 요약/계좌/월별/달력/세금 기준 통계 재계산(API 호출 없음).
# v0.5.34: SQLite 저장 시세를 탭 진입 시 자동 집계하여 카드 3종을 한 UI에서 출력; 빈칸 대신 오류 표시.
# v0.5.33: 앱 시작 및 탭 진입 시 SQLite의 시세·환율·배당정보 자동 표시(외부 API 미호출).
# v0.5.31: 배당 저장자료 불러오기 버튼 추가 및 예상 배당 달력 종목·금액 한 줄 표시.
# v0.5.30: 합계 카드 renderUI를 안정적인 renderText/textOutput으로 변경, 강제 갱신 및 실패 메시지 제공.
# v0.5.29: 전체 평가액/평가손익 카드의 값 출력 경로 수정(평가표 원자료 직접 집계, 항상 상태 표시).
# v0.5.28: 평가 요약 카드 및 CSV 스냅샷 저장.
# v0.5.27: 숫자 오른쪽 정렬, 항목별 소수점 자릿수 통일, 정렬/검색은 원본 숫자로 유지.
# v0.5.26: 자산 평가표 기본 정렬을 비중(%) 내림차순으로 변경(사용자 정렬 가능).
# v0.5.25: 종목별 총 평가액 대비 비중(%) 추가, 음수 평가액/평가손익/수익률을 빨간 글씨로 표시.
# v0.5.24: 화면 배율 버튼의 브라우저 즉시 적용 및 Shiny 메시지 동기화 보강.
# v0.5.23: SQLite 소수점 필드 마이그레이션의 벡터 조건문 오류(&&→&) 수정.
# v0.5.22: 한국·미국 보유종목 소수점 주식 여부 필드 추가; 일반/소수점 동일종목 별도 보유 허용; 기존 DB 무손실 마이그레이션.
# v0.5.21: 보유종목 저장 시 Shiny의 DT 교체 및 브라우저 전체 스크롤까지 반복 복원(이전 복원 타이밍 오류 수정).
# v0.5.20: 보유종목 저장 시 DT 목록 세로/가로 스크롤 및 페이지 위치 복원; 선택한 편집 종목 유지.
# v0.5.19: 계좌번호 기존 목록 선택과 별도 직접 입력 병행; 직접 입력 우선 저장, 편집 전환 시 입력값 초기화.
# v0.5.18: 증권사별로 한국/미국 보유분을 한곳에 묶도록 증권사 우선 정렬 변경(증권사→국가→계좌→평가액).
# v0.5.17: 보유종목 표의 증권사명 한글 표기, 종목명 옆 평가액 배치, 그룹별 평가액 정렬 안내.
# v0.5.16: 계좌번호 콤보박스 SQL 문자열 따옴표 오류 수정(app.R sourcing 실패 대응).
# v0.5.15: 보유종목 편집표 국가→증권사→계좌유형 그룹, 그룹별 평가액 내림차순(결측값은 마지막).
# v0.5.14: 기존 계좌번호 증권사별 선택 및 새 계좌번호 직접 입력 지원.
# v0.5.13: 보유종목 기본 id 오름차순; 평균 매입환율 및 원화 취득원가 도움말 추가.
# v0.5.12: 과세유형 미국지수ETF(KR_US_INDEX_ETF) 추가; 보유편집표 한글 표시/분류 안내; 기존 분류 및 DB 보존.
# v0.5.11: 보유종목 편집표 기본 정렬에 증권사(가나다/알파벳순) 추가; ID 설명 보완.
# v0.5.10: 보유종목 편집표 기본 정렬(한국/미국, 일반/ISA/IRP/기타, 평가액 내림차순) 및 평가액 표시.
# v0.5.9: Total Return 일별 수익률 비율에서 -1 누락 수정; Sharpe, Sortino 정상화.
# v0.5.8: Total Return 낙폭/60거래일 이동 Sharpe/동일 지표 해설을 가격수익률과 같은 구성으로 추가.
# v0.5.7: 총수익률(TR) 백테스트, 배당 수집 상태 검증 및 분석지표 도움말.
# v0.5.6: 낙폭·60일 이동 Sharpe 그래프 및 Sortino/Calmar 비교.
# v0.5.5: 종목명 병기, Sharpe 내림차순, 내 포트폴리오 굵은 빨간선.
# v0.5.4: 벤치마크 탭 최하단 4개 대상 수익률·MDD·Sharpe 비교표(동일 기간, 무위험수익률 0%).
# 1. 최근 시세수집 실패 이력 초기화 버튼 추가
# 2. 최근 365일 종목별 배당 이벤트 수집 및 SQLite 저장
# 3. 최근 1년 USD/KRW 일별 환율 수집·저장, 미국 배당락일별 원화 환산
# 4. 배당 수집 상태 / 건별 배당 내역 표 추가
# 5. 시세 및 배당 전체 조회 버튼 추가 (기존 개별 버튼 유지)
# 6. 국내 종목명 및 국내 배당의 주당 원화표시 보완
# 7. 연간 주당 배당금 합계 명칭 및 보유수량 기준 계좌별 배당 추정집계
# 8. 종목별 수집·평가·기존 보유종목 스키마는 그대로 유지
# 9. 오늘 수집한 저장시세 확인 및 인터넷 접속 없는 즉시 재평가
# 10. 배당 추정액 3자리 콤마, 일반계좌 금융소득 기준선 시각화
# 11. 종목별 최근 365일 세전 배당수익률 (마지막 저장 시세 기준)
# 12. 시세/금액 천 단위 쉼표 표시 통일, 저장 시세 평가 버튼 상단 배치
# 13. 배당 상세표 주당배당금_원화 열 연노랑 강조
# 14. 크롬 우선 실행 런처(run_pms2.R) 및 최근 12개월 계좌별 월별 배당 누적막대그래프
# 15. 월별 누적막대 위 배당금 총합계 표시(천 단위 콤마) 및 Chrome 실행 안내 강화
# 16. 앱 실행 중 10분 간격 시세 자동 수집 (기본 ON, 배당 수집 제외)
# 17. Chrome 외부 실행용 런처 개선 / RStudio Viewer 사용 방지
# 18. 공급원 진단 조회값 상세 안내 및 테스트 수신값 컬럼명
# 19. Windows Chrome 전용 배치 실행기와 Viewer 비사용 실행 경로
# 20. 최근 1년 배당락일 이력 기반 향후 30일 예상 배당 일정·작년 배당 금액
# 21. 최근 1년 배당 이벤트 상세/해당배당일환율 명칭 개선 및 30일 예상 배당 달력
# 23. 벤치마크 저장 시세 재사용, 전체 초기화 후 재수집, 종목별 진행률 표시
# 24. 벤치마크 1년 이력 부족 종목 제외 / 4개 고정 비교곡선 / 제외 사유 표시
# 22. 벤치마크 탭: KODEX 3종 및 보유종목 과거 1년 일별 종가 수집, SQLite 재사용, 원화기준 비교
# ============================================================
library(shiny)
library(DT)

# 시세/금액: 데이터는 숫자로 유지하고 화면의 표시만 3자리 쉼표로 변환한다.
# 정렬/검색 시에는 원래 값 사용. 종목코드·계좌번호에는 적용하지 않는다.
pms_money_defs <- function(df) {
  # Always keep numeric data numeric for DataTables sorting and filtering.
  # Only 'display' is formatted; 'sort' and 'type' get the original numeric value.
  if (!nrow(df) && !ncol(df)) return(list())
  cols <- names(df)[vapply(df,is.numeric,logical(1))]
  cols <- setdiff(cols,grep('^(ID|id)$|코드|일자|날짜|시각|연도|거래일수|건수|횟수|^Sharpe$|^Sortino$|^Calmar$|^MDD$',cols,value=TRUE))
  if (!length(cols)) return(list())
  render <- DT::JS("function(data,type,row,meta) {
    if (type !== 'display') return data;
    if (data === null || data === undefined || data === '') return '';
    var n=Number(String(data).replace(/,/g,''));
    if (!Number.isFinite(n)) return data;
    var key=meta.settings.aoColumns[meta.col].sTitle || '';
    key=key.replace(/<[^>]*>/g,'');
    var decimals=0;
    if (/수량|quantity/i.test(key)) {
      decimals=6;
    } else if (/환율|fx_rate|수익률|퍼센트|비중|yield|return_pct/i.test(key)) {
      decimals=2;
    } else if (/주당배당금|배당금_주당|배당_주당|per.share/i.test(key)) {
      decimals=6;
    } else if (/현재가|시세|가격|avg_price/i.test(key)) {
      var curIdx=meta.settings.aoColumns.findIndex(function(c){return /^(통화|currency)$/.test((c.sTitle||'').replace(/<[^>]*>/g,''));});
      var cur=curIdx>=0 ? row[curIdx] : '';
      decimals=(String(cur).toUpperCase()==='USD')?2:0;
    }
    var opts={minimumFractionDigits:decimals===6?0:decimals,maximumFractionDigits:decimals};
    return n.toLocaleString('en-US',opts);
  }")
  return(list(list(targets=as.integer(match(cols,names(df))-1L),
                   className='dt-right pms-numeric',render=render)))
}

# PMS2 Telegram 개인 알림 설정 (소스 외부 공유 전 토큰 삭제)
# BotFather에서 노출된 토큰은 /revoke 후 새 토큰으로 교체하세요.
PMS_TG_BOT_TOKEN <- "새_봇_토큰_입력"
PMS_TG_CHAT_ID   <- "개인_chat_id_입력"

PMS_TG_BOT_TOKEN <- "8990278369:AAFtjphXAxUhNii4bJ1gjGo6QmfX7Fn_Qr4"

PMS_TG_CHAT_ID   <- "7589805373"
pms_tg_auto_busy <- FALSE
pms_tg_configured <- function() {
  nzchar(PMS_TG_BOT_TOKEN) && nzchar(PMS_TG_CHAT_ID) &&
    !grepl("입력|여기에|<", PMS_TG_BOT_TOKEN) &&
    !grepl("입력|여기에|<", PMS_TG_CHAT_ID)
}
pms_tg_send <- function(message) {
  if (!pms_tg_configured()) stop("app.R 상단에 새 Bot Token과 개인 chat_id를 입력하세요.")
  if (!requireNamespace("httr", quietly=TRUE)) stop("httr 패키지가 필요합니다: install.packages('httr')")
  token <- trimws(PMS_TG_BOT_TOKEN)
  chat_id <- trimws(as.character(PMS_TG_CHAT_ID))
  if (!nzchar(chat_id)) stop("텔레그램 Chat ID가 비어 있습니다.")
  if (!nzchar(as.character(message))) stop("텔레그램 메시지가 비어 있습니다.")
  url <- paste0("https://api.telegram.org/bot", token, "/sendMessage")
  resp <- httr::POST(url,
    body=list(chat_id=chat_id, text=enc2utf8(as.character(message))),
    encode="form", httr::timeout(12))
  # Parse the Telegram error body before checking HTTP status so 400/403
  # can identify invalid chat IDs, blocked bots and missing permissions.
  status <- httr::status_code(resp)
  ans <- tryCatch(httr::content(resp, as="parsed", type="application/json", encoding="UTF-8"),
                  error=function(e) NULL)
  if (!identical(status, 200L) || !isTRUE(ans$ok)) {
    description <- if (is.list(ans) && !is.null(ans$description))
      as.character(ans$description) else "텔레그램 응답 내용을 확인할 수 없습니다."
    stop(sprintf("Telegram HTTP %d: %s", status, description), call.=FALSE)
  }
  invisible(TRUE)
}
pms_tg_setting_init <- function(db) {
  DBI::dbExecute(db, "CREATE TABLE IF NOT EXISTS pms_tg_settings (setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL)")
  DBI::dbExecute(db, "INSERT OR IGNORE INTO pms_tg_settings(setting_key,setting_value) VALUES ('auto_16','1')")
  DBI::dbExecute(db, "CREATE TABLE IF NOT EXISTS pms_tg_sent (send_date TEXT PRIMARY KEY, sent_at TEXT NOT NULL, total REAL NOT NULL, delta REAL NOT NULL)")
}
pms_tg_setting_get <- function(db) {
  x <- DBI::dbGetQuery(db,"SELECT setting_value FROM pms_tg_settings WHERE setting_key='auto_16'")
  !nrow(x) || identical(x$setting_value[[1]], "1")
}
pms_tg_setting_put <- function(db, enabled) {
  DBI::dbExecute(db, "INSERT OR REPLACE INTO pms_tg_settings(setting_key,setting_value) VALUES ('auto_16',?)",
                 params=list(if (isTRUE(enabled)) "1" else "0"))
}

source(file.path('R', 'db.R'), local = TRUE)
source(file.path('R', 'holdings_excel.R'), local = TRUE)
source(file.path('R', 'prices.R'), local = TRUE)
source(file.path('R', 'valuation.R'), local = TRUE)
source(file.path('R', 'diversification.R'), local = TRUE)
source(file.path('R', 'dividends.R'), local = TRUE)
source(file.path('R', 'quote_board.R'), local = TRUE)
source(file.path('R', 'board_ticker.R'), local = TRUE)
source(file.path('R', 'benchmark.R'), local = TRUE)
source(file.path('R', 'report.R'), local = TRUE)
source(file.path('R', 'voice_brief.R'), local = TRUE)
source(file.path('R', 'snapshot.R'), local = TRUE)
source(file.path('R', 'trajectory.R'), local = TRUE)
source(file.path('R', 'equal_cashflow.R'), local = TRUE)
source(file.path('R', 'home_market_backfill.R'), local = TRUE)
source(file.path('R', 'performance.R'), local = TRUE)
source(file.path('R', 'risk_analysis.R'), local = TRUE)
source(file.path('R', 'trades.R'), local = TRUE)
source(file.path('R', 'contributions.R'), local = TRUE)
source(file.path('R', 'fire_now.R'), local = TRUE)
source(file.path('R', 'fire_planning.R'), local = TRUE)

con <- open_pms_db()
pms_contrib_init(con)
pms_tg_setting_init(con)
onStop(function() {
  if (DBI::dbIsValid(con)) DBI::dbDisconnect(con)
})

choices <- list(
  country = c('한국'='KR','미국'='US'),
  currency = c('원화'='KRW','미국 달러'='USD'),
  asset_class = c('주식'='STOCK','채권'='BOND','금'='GOLD','현금성'='CASH','기타'='OTHER'),
  tax_type = c('국내 일반주식'='KR_STOCK','국내 주식형 ETF'='KR_EQUITY_ETF',
    '국내 기타 ETF'='KR_OTHER_ETF','미국지수ETF'='KR_US_INDEX_ETF','미국 주식'='US_STOCK','미국 ETF'='US_ETF',
    'KRX 금현물(수동환산)'='KRX_GOLD','기타'='OTHER'),
  broker = c('삼성증권'='SAMSUNG','NH투자증권'='NH','미래에셋증권'='MIRAE',
    '한국투자증권'='KOREA','키움증권'='KIWOOM','KB증권'='KB',
    '대신증권'='DAISHIN','카카오페이증권'='KAKAOPAY','기타'='OTHER'),
  account_type = c('일반'='GENERAL','ISA'='ISA','IRP'='IRP','연금저축'='PENSION')
)

# 같은 R Shiny 프로세스에 접속 중인 브라우저 세션 수를 공유합니다.
# 각 브라우저 탭이 1세션이며, 접속/종료 시 모든 화면에 즉시 반영됩니다.
pms_active_sessions <- shiny::reactiveVal(0L)

ui <- fluidPage(
  # 스마트폰 전용 하단 메뉴. 실제 탭은 기존 pms_tabs 하나만 사용합니다.
  tags$head(
    tags$meta(name='viewport', content='width=device-width, initial-scale=1, viewport-fit=cover'),
    tags$style(HTML("
      .pms-mobile-dock, .pms-mobile-sheet { display:none; }
      @media (max-width:768px) {
        body { padding-bottom:calc(78px + env(safe-area-inset-bottom, 0px)); }
        body > .container-fluid { padding-left:9px; padding-right:9px; max-width:100%; }
        /* 원래 탭은 DOM에서 유지하여 Shiny 입력/출력을 보존합니다. */
        #pms_tabs.nav { display:none !important; }
        .pms-mobile-dock { display:grid; grid-template-columns:repeat(4,minmax(0,1fr));
          position:fixed; left:0; right:0; bottom:0; z-index:1200;
          padding:7px 4px calc(8px + env(safe-area-inset-bottom, 0px));
          background:#fff; border-top:1px solid #d9e2ec;
          box-shadow:0 -3px 14px rgba(25,45,75,.12); }
        .pms-mobile-dock button { background:none; border:0; color:#516178;
          display:flex; flex-direction:column; align-items:center; gap:2px;
          padding:7px 2px; font-size:12px; min-height:49px; border-radius:9px; }
        .pms-mobile-dock button .pms-mobile-icon {font-size:21px;line-height:1.1;}
        .pms-mobile-dock button[aria-current='page'] {color:#1266a6;background:#e9f3fa;font-weight:700;}
        .pms-mobile-sheet { position:fixed; left:0; right:0; bottom:calc(72px + env(safe-area-inset-bottom, 0px));
          z-index:1199; background:#fff; border-radius:15px 15px 0 0; max-height:67vh;
          overflow-y:auto; padding:14px 12px; box-shadow:0 -6px 30px rgba(0,0,0,.18); }
        .pms-mobile-sheet.pms-open {display:block;}
        .pms-mobile-sheet-title {font-size:16px;font-weight:700;margin:0 0 9px 4px;}
        .pms-mobile-sheet-grid {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;}
        .pms-mobile-sheet-grid button {border:1px solid #dce4ed;border-radius:9px;background:#f8fafc;
          text-align:left;padding:11px 9px;font-size:13px;min-height:44px;overflow-wrap:anywhere;}
        .pms-mobile-sheet-grid button.pms-selected {border-color:#1980b8;background:#eaf5fc;}
        .pms-mobile-sheet-close {float:right;border:0;background:none;font-size:20px;}
        .tab-content {max-width:100%;min-width:0;}
        .tab-pane {max-width:100%;min-width:0;}
        h1,h2,h3 {overflow-wrap:anywhere;}
        h3 {font-size:19px;} h4 {font-size:16px;}
        .row {margin-left:-5px;margin-right:-5px;}
        .row > [class*='col-'] {padding-left:5px;padding-right:5px;}
        .well,.panel,.card {max-width:100%;}
        .form-control, .selectize-input {max-width:100%;}
        .btn,.form-control {min-height:40px;}
        .dataTables_wrapper {max-width:100%; overflow-x:auto; -webkit-overflow-scrolling:touch;}
        .dataTables_wrapper table {width:max-content !important; min-width:100%;}
        .table-responsive {overflow-x:auto; -webkit-overflow-scrolling:touch;}
        .plotly, .plot-container, .html-widget, svg {max-width:100%;}
      }
      body.pms-dark .pms-mobile-dock, body.pms-dark .pms-mobile-sheet {
        background:#1b2938;color:#eef3f9;border-color:#44576b;
      }
      body.pms-dark .pms-mobile-sheet-grid button {background:#283b4e;color:#eef3f9;border-color:#51657b;}
    ")),
    tags$script(HTML("
      (function(){
        function initMobile(){
          var nav=document.getElementById('pms_tabs');
          var sheet=document.getElementById('pms_mobile_sheet');
          if(!nav || !sheet || !window.jQuery)return;
          var list=sheet.querySelector('.pms-mobile-sheet-grid');
          if(!list || list.getAttribute('data-ready'))return;
          var anchors=Array.prototype.slice.call(nav.querySelectorAll('a[data-toggle=tab]'));
          // Bootstrap Shiny tabsetPanel uses data-value for each direct tab.
          if(!anchors.length)anchors=Array.prototype.slice.call(nav.querySelectorAll('a[href]'));
          anchors.forEach(function(a){
            var name=a.getAttribute('data-value')||a.textContent.trim();
            if(!name)return;
            var b=document.createElement('button'); b.type='button';b.textContent=name;
            b.setAttribute('data-pms-tab',name);list.appendChild(b);
          });
          list.setAttribute('data-ready','1');
          function activate(name){
            var target=anchors.find(function(a){return (a.getAttribute('data-value')||a.textContent.trim())===name;});
            if(target){jQuery(target).tab('show'); sheet.classList.remove('pms-open');
              var more=document.getElementById('pms_mobile_more');if(more)more.setAttribute('aria-expanded','false');}
          }
          document.addEventListener('click',function(e){
            var btn=e.target.closest('[data-pms-tab]');
            if(btn){activate(btn.getAttribute('data-pms-tab'));return;}
            var more=e.target.closest('#pms_mobile_more');
            if(more){var open=sheet.classList.toggle('pms-open');more.setAttribute('aria-expanded',open?'true':'false');return;}
            if(e.target.closest('#pms_mobile_close')){sheet.classList.remove('pms-open');
              document.getElementById('pms_mobile_more').setAttribute('aria-expanded','false');}
          });
          function sync(){
            var active=nav.querySelector('li.active > a');
            var name=active&&(active.getAttribute('data-value')||active.textContent.trim());
            document.querySelectorAll('.pms-mobile-dock [data-pms-tab]').forEach(function(b){
              if(b.getAttribute('data-pms-tab')===name)b.setAttribute('aria-current','page');
              else b.removeAttribute('aria-current');
            });
            list.querySelectorAll('button').forEach(function(b){b.classList.toggle('pms-selected',b.getAttribute('data-pms-tab')===name);});
          }
          jQuery(nav).on('shown.bs.tab', 'a[data-toggle=tab]',sync);sync();
        }
        document.addEventListener('DOMContentLoaded',initMobile);
        if(window.jQuery)jQuery(document).on('shiny:connected',initMobile);
      })();
    "))
  ),
  # 모든 탭에서 보이는 접속 세션 표시 (모바일/PC 공통).
  tags$head(tags$style(HTML("
    .pms-session-strip {display:flex;justify-content:flex-end;align-items:center;
      padding:3px 8px;margin:0 0 5px;min-height:25px;font-size:12px;color:#52687b;}
    .pms-session-strip .shiny-text-output {display:inline-block;font-weight:700;
      color:#1266a6;margin-left:5px;}
    body.pms-dark .pms-session-strip {color:#c4d4e5;}
    body.pms-dark .pms-session-strip .shiny-text-output {color:#9ad1ff;}
    @media(max-width:768px){.pms-session-strip{font-size:11px;padding:2px 4px;}}
  "))),
  tags$div(class='pms-session-strip',
    tags$span('● 접속 중 '),textOutput('pms_session_count',inline=TRUE)),
  tags$div(class='pms-mobile-dock', role='navigation', 'aria-label'='모바일 빠른 메뉴',
    tags$button(type='button', `data-pms-tab`='홈', tags$span(class='pms-mobile-icon','⌂'),'홈'),
    tags$button(type='button', `data-pms-tab`='보유종목 편집', tags$span(class='pms-mobile-icon','▤'),'보유종목'),
    tags$button(type='button', `data-pms-tab`='분산투자', tags$span(class='pms-mobile-icon','◉'),'분산투자'),
    tags$button(id='pms_mobile_more',type='button',`aria-expanded`='false',
      tags$span(class='pms-mobile-icon','☰'),'전체메뉴')
  ),
  tags$div(id='pms_mobile_sheet',class='pms-mobile-sheet',role='dialog',`aria-label`='전체 메뉴',
    tags$button(id='pms_mobile_close',type='button',class='pms-mobile-sheet-close',`aria-label`='닫기','×'),
    tags$div(class='pms-mobile-sheet-title','전체 메뉴'),
    tags$div(class='pms-mobile-sheet-grid')
  ),

  # v0.6.11: all DataTables: comma-group KRW monetary display in columns
  # not using pms_money_defs(). Manipulates visible cells only; DT's raw
  # numeric data remain unchanged for sorting/filtering/calculation.
  tags$head(tags$script(HTML("\n    (function(){\n      function pmsFormatKRWCells(table){\n        if(!table || !window.jQuery || !jQuery.fn.dataTable ||\n           !jQuery.fn.dataTable.isDataTable(table)) return;\n        var dt=jQuery(table).DataTable();\n        var cols=dt.columns().header().toArray();\n        var indices=[];\n        cols.forEach(function(th,i){\n          var h=(th.textContent||'').replace(/\\s+/g,'').toLowerCase();\n          // Monetary columns only, not ratios, prices in dollars, IDs or quantities.\n          var money=/(평가액|평가손익|취득원가|매입원가|원화|총자산|총금액|평가증감|평가변화|투자금액|매수금액|매도금액|배당금|월납입|원금|총납입|자산금액|원$|_원$|krw|금액)/.test(h);\n          var excluded=/(코드|계좌번호|id$|비중|수익률|퍼센트|환율|수량|주당|단가|현재가|전일종가|가격|일자|시각|날짜|횟수)/.test(h);\n          if(money && !excluded) indices.push(i);\n        });\n        if(!indices.length)return;\n        dt.rows({page:'current'}).every(function(){\n          var cells=this.node() && this.node().cells;\n          if(!cells)return;\n          indices.forEach(function(i){\n            var cell=cells[i];if(!cell)return;\n            // Do not overwrite markup, buttons, spans or annotated content.\n            if(cell.children.length)return;\n            var val=(cell.textContent||'').trim();\n            if(!/^[+-]?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?$/.test(val))return;\n            var pieces=val.replace(/,/g,'').split('.');\n            pieces[0]=pieces[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g,',');\n            var fmt=pieces.join('.');\n            if(fmt!==val)cell.textContent=fmt;\n          });\n        });\n      }\n      jQuery(document).on('init.dt draw.dt',function(ev){\n        var table=ev.target;\n        // Defer until DataTables initialization is complete.\n        window.setTimeout(function(){pmsFormatKRWCells(table);},0);\n      });\n    })();\n  "))),
  tags$head(tags$style(HTML("
    .pms-board-totals{color:#20364c;}
    body.pms-dark .pms-board-totals{color:#eef4ff;}
    #board_total_change.pms-board-change-up{color:#cf3535!important;}
    #board_total_change.pms-board-change-down{color:#267dc5!important;}
    body.pms-dark #board_total_change.pms-board-change-up{color:#ff9292!important;}
    body.pms-dark #board_total_change.pms-board-change-down{color:#8cc4ff!important;}
  ")),tags$script(HTML("
    Shiny.addCustomMessageHandler('pms_board_total_change_color',function(x){
      var el=document.getElementById('board_total_change');
      if(!el)return;
      el.classList.remove('pms-board-change-up','pms-board-change-down');
      if(x.color==='up')el.classList.add('pms-board-change-up');
      if(x.color==='down')el.classList.add('pms-board-change-down');
    });
  "))),
  tags$head(tags$style(HTML("
    #pms_zoom_content .nav-tabs > li > a:hover,
    #pms_zoom_content .nav-tabs > li.active > a {font-weight:700!important;}
    body.pms-dark {background:#0f141b!important;color:#edf2f8!important;}
    body.pms-dark #pms_zoom_content,body.pms-dark .container-fluid {color:#edf2f8;}
    body.pms-dark #pms_zoom_content > .tabbable > .nav-tabs {background:#141d29!important;border-bottom-color:#415369;}
    body.pms-dark .nav-tabs>li>a {color:#c7d8ec;background:#182432;border-color:#435268;}
    body.pms-dark .nav-tabs>li.active>a,body.pms-dark .nav-tabs>li>a:hover {background:#304256!important;color:white!important;border-color:#60758b;}
    body.pms-dark .well,body.pms-dark .panel,body.pms-dark .panel-body,body.pms-dark .card,
    body.pms-dark .pms-home,body.pms-dark .pms-flow-panel,
    body.pms-dark .pms-calendar-day,body.pms-dark .pms-calendar-dow {background:#202d3b!important;color:#e6edf7!important;border-color:#506176!important;}
    body.pms-dark .pms-calendar-day.outside {background:#192330!important;}
    body.pms-dark .alert-info,body.pms-dark .alert-warning {background:#263d50!important;color:#f1eecf!important;border-color:#547187!important;}
    body.pms-dark .text-muted,body.pms-dark small,body.pms-dark .help-block {color:#b3c3d7!important;}
    body.pms-dark table,body.pms-dark table.dataTable,body.pms-dark table.dataTable tbody tr,
    body.pms-dark .dataTables_wrapper,body.pms-dark .dataTables_info,body.pms-dark .dataTables_paginate {background:#1d2a38!important;color:#edf3fc!important;}
    body.pms-dark table.dataTable tbody td {border-color:#43556a!important;background:#1d2a38;}
    body.pms-dark table.dataTable thead th,body.pms-dark .dataTables_filter input {background:#2b3b4e!important;color:#edf3fc!important;}
    body.pms-dark input,body.pms-dark select,body.pms-dark textarea,body.pms-dark .form-control {background:#273547!important;color:#f2f7ff!important;border-color:#5b6b80!important;}
    body.pms-dark .btn-default {background:#283b50!important;color:#f4f7ff!important;border-color:#61758b!important;}
    body.pms-dark .pms-calendar-day .event {background:#334256!important;color:#f4f7ff!important;}
    body.pms-dark .pms-calendar-day .daytotal {background:#635225!important;color:#fff3b0!important;}
    body.pms-dark .pms-calendar-explain {color:#bac7d9!important;}
    body.pms-dark .pms-trade-day,body.pms-dark .pms-trade-dow {background:#243348!important;color:#e6eef9!important;border-color:#52667e!important;}
    .pms-calendar-grid > .pms-calendar-dow:nth-child(7n+6),
    .pms-calendar-grid > .pms-calendar-day:nth-child(7n+6) .dayno {color:#2563eb!important;}
    .pms-calendar-grid > .pms-calendar-dow:nth-child(7n+7),
    .pms-calendar-grid > .pms-calendar-day:nth-child(7n+7) .dayno {color:#dc2626!important;}
    body.pms-dark .pms-calendar-grid > .pms-calendar-dow:nth-child(7n+6),
    body.pms-dark .pms-calendar-grid > .pms-calendar-day:nth-child(7n+6) .dayno {color:#83b8ff!important;}
    body.pms-dark .pms-calendar-grid > .pms-calendar-dow:nth-child(7n+7),
    body.pms-dark .pms-calendar-grid > .pms-calendar-day:nth-child(7n+7) .dayno {color:#ff9393!important;}
  ")),tags$script(HTML("
    (function(){
      window.pmsSetDarkMode=function(enabled){
        document.body.classList.toggle('pms-dark',!!enabled);
        localStorage.setItem('pms2-dark-mode',enabled?'1':'0');
        var b=document.getElementById('pms_dark_toggle');
        if(b){b.textContent=enabled?'☀ 라이트모드':'☾ 다크모드';b.setAttribute('aria-pressed',enabled?'true':'false');}
      };
      function initTheme(){window.pmsSetDarkMode(localStorage.getItem('pms2-dark-mode')==='1');}
      if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initTheme);else initTheme();
    })();
  "))),
  tags$head(tags$script(HTML("
    $(document).on('click', '.pms-edit-holding', function(e){
      e.preventDefault(); e.stopPropagation();
      var id=Number(this.getAttribute('data-id'));
      if(Number.isInteger(id) && window.Shiny)
        Shiny.setInputValue('pms_edit_holding_id',{id:id,nonce:Date.now()},{priority:'event'});
    });
  "))),
  # v0.6.04: delegated click handler survives DataTables redraw/search/page changes.
  tags$head(tags$script(HTML("
    (function(){
      if(window.pmsRowEditDelegatedInstalled) return;
      window.pmsRowEditDelegatedInstalled=true;
      document.addEventListener('click',function(e){
        var tr=e.target.closest && e.target.closest('#tbl table tbody tr');
        if(!tr || e.target.closest('button, a, input, select, textarea')) return;
        var first=tr.querySelector('td:first-child');
        if(!first) return;
        var id=Number((first.textContent||'').trim().replace(/,/g,''));
        if(!Number.isSafeInteger(id) || id<1) return;
        if(window.Shiny && typeof Shiny.setInputValue==='function'){
          Shiny.setInputValue('pms_click_edit_id', {id:id,nonce:Date.now()}, {priority:'event'});
        }
      },true);
    })();
  "))),
  # v0.6.04: independent local clock shown even when quote data is missing.
  tags$head(tags$script(HTML("
    (function(){
      function stamp(){
        var el=document.getElementById('pms_board_live_clock');
        if(!el)return;
        el.textContent=new Intl.DateTimeFormat('ko-KR',{
          timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit',
          weekday:'short',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false
        }).format(new Date())+' (한국시간)';
      }
      if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',stamp);
      else stamp();
      window.setInterval(stamp,1000);
    })();
  "))),
  tags$head(tags$style(HTML("
    .pms-ticker-shell{position:fixed;bottom:0;left:0;right:0;z-index:9998;height:49px;
      display:flex;align-items:center;background:#162331;border-top:2px solid #4d657c;
      box-shadow:0 -3px 12px #0003;color:#f8fafc;overflow:hidden;font-size:14px;}
    .pms-ticker-caption{z-index:2;flex:0 0 auto;padding:0 15px;font-size:12px;font-weight:700;
      color:#dbeafe;background:#223750;height:100%;display:flex;align-items:center;white-space:nowrap;}
    .pms-ticker-window{overflow:hidden;flex:1;min-width:0;height:100%;display:flex;align-items:center;}
    .pms-ticker-runner{display:flex;width:max-content;animation:pmsTickerMove 140s linear infinite;}
    .pms-summary-shell{bottom:49px;background:#26384b;border-top:1px solid #7790a7;z-index:9998;}
    .pms-summary-shell .pms-ticker-caption{background:#354e68;color:#f8fafc;}
    .pms-summary-shell .pms-ticker-item b{color:#e7f3ff;}
    .pms-summary-shell .pms-ticker-item strong{color:#f5df9b;}
    .pms-summary-shell .pms-ticker-up strong{color:#ff8585;}
    .pms-summary-shell .pms-ticker-down strong{color:#7bb7ff;}
    .pms-loading-shell{position:fixed;bottom:0;left:0;right:0;height:98px;z-index:9997;background:#162331;color:#dbeafe;display:flex;flex-direction:column;}
    .pms-loading-row{height:49px;display:flex;align-items:center;overflow:hidden;border-top:1px solid #46607a;white-space:nowrap;}
    .pms-loading-row:first-child{background:#26384b;}
    .pms-loading-text{display:inline-block;padding-left:100vw;animation:pmsLoadingMove 20s linear infinite;}
    @keyframes pmsLoadingMove{from{transform:translateX(0)}to{transform:translateX(-100%);}}
    body:has(#board_bottom_ticker .pms-ticker-shell) .pms-loading-shell{display:none;}
    .pms-ticker-sequence{display:flex;flex:0 0 auto;white-space:nowrap;align-items:center;}
    .pms-ticker-item{display:inline-flex;gap:5px;align-items:center;padding:0 17px;
      border-right:1px solid #506074;white-space:nowrap;}
    .pms-ticker-item b{color:#e7edf3;}
    .pms-ticker-item strong{font-weight:750;}
    .pms-ticker-up strong{color:#ff7272;}.pms-ticker-down strong{color:#65aaff;}
    .pms-ticker-flat strong{color:#d1d5db;}.pms-ticker-none strong{color:#a9b3c3;}
    .pms-ticker-item small{color:#93a4b5;font-size:10px;}
    @keyframes pmsTickerMove{from{transform:translateX(0)}to{transform:translateX(-50%)}}
    .pms-ticker-shell:hover .pms-ticker-runner{animation-play-state:paused;}
    body:has(.pms-ticker-shell),body:has(.pms-loading-shell:not([hidden])) {padding-bottom:108px;}
    @media(prefers-reduced-motion:reduce){.pms-ticker-runner,.pms-loading-text{animation:none;}}
  "))),
  tags$head(tags$script(HTML("
    (function(){
      function updateTickerSpeeds(){
        document.querySelectorAll('.pms-ticker-shell').forEach(function(shell){
          var runner=shell.querySelector('.pms-ticker-runner');
          var seq=runner && runner.querySelector('.pms-ticker-sequence');
          if(!runner || !seq)return;
          var pixelsPerSecond=shell.classList.contains('pms-summary-shell') ? 34 : 48;
          var width=seq.getBoundingClientRect().width;
          if(width>0)runner.style.animationDuration=Math.max(8,width/pixelsPerSecond).toFixed(2)+'s';
        });
      }
      document.addEventListener('DOMContentLoaded',function(){
        var root=document.getElementById('board_bottom_ticker');
        if(root){new MutationObserver(function(){requestAnimationFrame(updateTickerSpeeds);})
          .observe(root,{childList:true,subtree:true});}
        window.addEventListener('resize',updateTickerSpeeds);
        updateTickerSpeeds();
      });
    })();
  "))),
  tags$head(tags$style(HTML("
    #pms-initial-loading {position:fixed;inset:0;z-index:99999;background:rgba(243,247,251,.94);display:flex;align-items:center;justify-content:center;flex-direction:column;gap:18px;font-size:17px;color:#294962;}
    #pms-initial-loading .ring {width:58px;height:58px;border:6px solid #c8d5df;border-top-color:#2678b9;border-radius:50%;animation:pmsStartupSpin .9s linear infinite;}
    @keyframes pmsStartupSpin {to{transform:rotate(360deg)}}
    #pms-initial-loading .hint{font-size:13px;color:#596d7d;}
  ")),tags$script(HTML("
    (function(){
      var finished=false;
      function done(reason){
        if(finished) return;
        finished=true;
        var el=document.getElementById('pms-initial-loading');
        if(el){el.style.display='none';el.setAttribute('aria-hidden','true');}
        if(reason==='timeout'){
          var hint=document.getElementById('pms-startup-warning');
          if(hint){hint.style.display='block';hint.textContent='초기 화면 응답이 지연되어 대기 화면을 해제했습니다. 내용이 비어 있다면 RStudio Console의 오류를 확인하세요.';}
        }
        if(window.pmsStartupObserver){window.pmsStartupObserver.disconnect();window.pmsStartupObserver=null;}
      }
      function homeReady(){
        var home=document.getElementById('home_dashboard');
        return !!(home && home.children.length && home.textContent.trim().length>15 && !home.querySelector('.shiny-output-error'));
      }
      function bind(){
        // Shiny emits shiny:value as a jQuery event, not consistently as a native DOM event.
        if(window.jQuery){
          window.jQuery(document).on('shiny:value.pmsStartup shiny:error.pmsStartup',function(event){
            if(event.name==='home_dashboard')done(event.type);
          });
          window.jQuery(document).on('shiny:idle.pmsStartup',function(){if(homeReady())done('idle');});
        }
        if(homeReady()){done('ready');return;}
        var root=document.getElementById('home_dashboard') || document.body;
        if(window.MutationObserver){
          window.pmsStartupObserver=new MutationObserver(function(){if(homeReady())done('rendered');});
          window.pmsStartupObserver.observe(root,{childList:true,subtree:true});
        }
      }
      if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',bind);
      else bind();
      // Never leave an opaque overlay blocking the entire application indefinitely.
      window.setTimeout(function(){done('timeout');},15000);
    })();
  "))),
  tags$div(id='pms-initial-loading',tags$div(class='ring'),tags$strong('PMS2 데이터 조회 중…'),tags$span(class='hint','저장된 자산 자료로 첫 화면을 준비하고 있습니다.')),
  tags$div(id='pms-startup-warning',style='display:none;background:#fff5dc;border:1px solid #e0c579;border-radius:6px;padding:8px 12px;margin:8px 0;color:#715211;font-size:13px;'),
  tags$head(tags$style(HTML("
  /* Sticky navigation stays visible while the report and analysis pages scroll. */
  #pms_zoom_content > .tabbable > .nav-tabs { position: sticky; top: 0; z-index: 1050; background: #fff; padding-top: 7px; box-shadow: 0 2px 5px rgba(0,0,0,.09); }
  #pms_zoom_content .nav-tabs { scroll-margin-top: 0; }
  ")),tags$script(HTML("
    (function(){
      function status(msg,kind){
        var el=document.getElementById('pms_saved_eval_diagnostic');
        if(!el) return;
        el.textContent=msg;
        el.style.color=kind==='err'?'#b71c1c':'#144577';
      }
      document.addEventListener('click',function(e){
        var btn=e.target.closest&&e.target.closest('#refresh_valuation, #evaluate_saved_dividends');
        if(!btn) return;
        var name=btn.id==='refresh_valuation'?'시세':'배당';
        status(name+' 평가 버튼 클릭 감지. Shiny 서버 응답을 기다립니다… ('+new Date().toLocaleTimeString()+')');
      },true);
      document.addEventListener('shiny:connected',function(){status('Shiny 서버 연결됨. 평가 버튼을 누르면 결과가 표시됩니다.');});
      document.addEventListener('shiny:disconnected',function(){status('Shiny 서버 연결이 끊어졌습니다. R 콘솔을 확인하세요.','err');});
      if(window.Shiny && typeof Shiny.addCustomMessageHandler==='function'){
        Shiny.addCustomMessageHandler('pms_eval_ack',function(x){status(x.text,x.error?'err':'ok');});
      }
    })();
  "))),
  tags$head(tags$script(HTML("Shiny.addCustomMessageHandler('pms_profit_sign', function(x) { var el=document.getElementById('pms_profit_card_value'); if(el) el.style.color=x.negative?'#c62828':'#222222'; });"))),
  tags$head(tags$style(HTML("
    table.dataTable td.pms-numeric, table.dataTable td.dt-right { text-align:right !important; font-variant-numeric:tabular-nums; font-feature-settings:'tnum'; white-space:nowrap; }

    .pms-calendar {margin:10px 0 12px;}
    .pms-calendar-grid {display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:5px;}
    .pms-calendar-dow {font-weight:700;background:#eaf0f5;padding:8px 3px;text-align:center;border-radius:5px;}
    .pms-calendar-day {min-height:106px;padding:6px;border:1px solid #d6dfe6;border-radius:7px;background:#fff;overflow-wrap:anywhere;}
    .pms-calendar-day.outside {background:#f5f7f9;color:#8e979e;}
    .pms-calendar-day.today {border:2px solid #3578b9;}
    .pms-calendar-day .dayno {font-weight:700;margin-bottom:5px;}
    .pms-calendar-day .daytotal {font-size:12px;font-weight:700;color:#8c5600;background:#fff2a8;border-radius:4px;padding:3px 4px;}
    .pms-calendar-day .event {font-size:11px;line-height:1.4;border-left:3px solid #e2b54b;padding:3px 4px;margin-top:4px;background:#fff9e6;display:flex;align-items:baseline;gap:4px;white-space:nowrap;overflow:hidden;}
    .pms-calendar-day .event-name {min-width:0;flex:1 1 auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
    .pms-calendar-day .event-amount {flex:0 0 auto;white-space:nowrap;font-variant-numeric:tabular-nums;font-weight:700;}
    .pms-calendar-day .event:hover {overflow:visible;position:relative;z-index:2;box-shadow:0 1px 5px #c8c8c8;} 
    .pms-calendar-explain {font-size:12px;color:#697780;margin-bottom:10px;}
    @media(max-width:900px){.pms-calendar-grid{gap:2px}.pms-calendar-day{min-height:90px;padding:3px}.pms-calendar-day .event{font-size:10px}}
    .pms-flow-panel {border:1px solid #d7e1e8;border-radius:12px;padding:14px 16px;margin:12px 0 18px;background:#f7fafc;}
    .pms-flow-header {display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px;}
    .pms-flow-steps {display:flex;gap:8px;flex-wrap:wrap;}
    .pms-flow-step {flex:1 1 135px;border:1px solid #d4dde4;border-radius:7px;padding:10px 8px;background:#edf1f4;color:#596975;text-align:center;transition:background .25s,border-color .25s;}
    .pms-step-done {background:#e1f5e9;border-color:#5bbd83;color:#246b40;}
    .pms-step-error {background:#ffebeb;border-color:#d66b6b;color:#a02525;}
    .pms-step-active {background:#deefff;border-color:#3186c6;color:#185783;animation:pmsPulse 1.1s ease-in-out infinite;}
    .pms-flow-marker {padding:4px 11px;border-radius:15px;font-size:12px;background:#e5e9ed;}
    .pms-state-running {background:#dcefff;color:#215e99;}
    .pms-state-success {background:#d9f2df;color:#206738;}
    .pms-state-failed {background:#ffe1e1;color:#a12323;}
    .pms-flow-detail {margin:12px 0 5px;overflow-wrap:anywhere;}
    @keyframes pmsPulse {0%,100%{box-shadow:0 0 0 0 rgba(43,127,195,.28)}50%{box-shadow:0 0 0 7px rgba(43,127,195,0)}}
    @media(prefers-reduced-motion:reduce){.pms-step-active{animation:none}}
    /* v0.5.44: 실제 응답에 근거한 공급원 현황 */
    .pms-src-panel{border:1px solid #ccddea;border-radius:11px;padding:12px;background:#f8fbff;margin:12px 0;}
    .pms-src-grid{display:flex;flex-wrap:wrap;gap:9px;margin-top:8px;}
    .pms-src-card{flex:1 1 155px;padding:10px 11px;border-radius:9px;background:white;border:1px solid #d7e0e8;position:relative;min-width:140px;}
    .pms-src-card .pms-src-num{font-size:19px;font-weight:700;display:block;margin:3px 0;}
    .pms-src-card .pms-src-label{font-size:12px;color:#62758b;}
    .pms-src-card.pms-src-flash{border-color:#3486c4;animation:pmsSourceFlash .85s ease-out;}
    .pms-src-dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:#9db0c0;margin-right:6px;vertical-align:middle;}
    .pms-src-running .pms-src-dot{background:#3188c8;animation:pmsSourceBounce .8s ease-in-out infinite;}
    .pms-src-card.pms-src-flash .pms-src-dot{background:#35a56c;}
    .pms-src-meta{font-size:12px;color:#42556a;margin-top:9px;overflow-wrap:anywhere;}
    .pms-src-chip{display:inline-block;border-radius:9px;background:#e1f1ff;padding:2px 7px;font-weight:600;white-space:nowrap;}
    @keyframes pmsSourceBounce{0%,100%{transform:translateY(0);opacity:.65}50%{transform:translateY(-5px);opacity:1}}
    @keyframes pmsSourceFlash{0%{box-shadow:0 0 0 0 rgba(47,135,200,.42)}100%{box-shadow:0 0 0 12px rgba(47,135,200,0)}}
    @media(prefers-reduced-motion:reduce){.pms-src-running .pms-src-dot,.pms-src-card.pms-src-flash{animation:none}}

  ")),tags$script(HTML("
    // 평균매수가/매입환율 입력창에서 포커스를 벗어날 때 콤마 표시
    $(document).on('blur', '#f_avg_price, #f_avg_fx_rate', function() {
      var raw = this.value.replace(/,/g, '').trim();
      if (!raw || !Number.isFinite(Number(raw))) return;
      var parts = raw.split('.');
      parts[0] = parts[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',');
      this.value = parts.join('.');
      $(this).trigger('change');
    });
    // 수집 상태는 서버에서 각 종목을 처리할 때 별도 메시지로 전송합니다.
    // 애니메이션은 서버에서 전송한 실제 상태만 완료로 표시합니다.
    const stageLabels = ['HTTPS 연결','SSL 인증서','HTTP 응답','JSON 해석','가격 확인'];
    function pmsDrawFlow(msg) {
      const root = document.getElementById('pms_flow_steps');
      const title = document.getElementById('pms_flow_title');
      const desc = document.getElementById('pms_flow_detail');
      const marker = document.getElementById('pms_flow_indicator');
      if (!root) return;
      const stage = Math.max(0,Math.min(5,Number(msg.stage || 0)));
      const state = msg.state || 'running';
      if (title) title.textContent = msg.title || '수집 대기';
      if (desc) desc.textContent = msg.detail || '연결 진단이나 시세 수집을 시작하세요.';
      if (marker) { marker.textContent = state === 'success' ? '정상' : state === 'failed' ? '오류' : state === 'waiting' ? '대기' : '진행 중'; marker.className = 'pms-flow-marker pms-state-' + state; }
      root.innerHTML = '';
      stageLabels.forEach(function(label,idx) {
        let box = document.createElement('div'); box.className = 'pms-flow-step';
        if (state === 'success' || idx < stage - 1) box.classList.add('pms-step-done');
        else if (idx === stage - 1 && state === 'failed') box.classList.add('pms-step-error');
        else if (idx === stage - 1 && state === 'running') box.classList.add('pms-step-active');
        box.textContent = (idx+1) + '. ' + label;
        root.appendChild(box);
      });
    }
    Shiny.addCustomMessageHandler('pms_flow', pmsDrawFlow);
    $(function(){pmsDrawFlow({state:'waiting',stage:0,title:'연결 확인 대기',detail:'연결 진단 또는 시세 수집을 누르면 과정이 표시됩니다.'});});
    Shiny.addCustomMessageHandler('pms_collection_reset', function(msg) {
      const el = document.getElementById('pms_collection_rows');
      if (el) el.innerHTML = '';
      const count = document.getElementById('pms_collection_count');
      if (count) count.textContent = '0 / ' + msg.total + '종목 · 성공 0 · 실패 0';
    });
    Shiny.addCustomMessageHandler('pms_collection_row', function(msg) {
      const el = document.getElementById('pms_collection_rows');
      if (!el) return;
      const tr = document.createElement('tr');
      const values = [msg.index, msg.country, msg.symbol, msg.price, msg.source, msg.status, msg.message];
      values.forEach(function(value) {
        const td = document.createElement('td');
        td.textContent = value == null ? '' : String(value);
        td.style.padding = '5px 8px'; td.style.borderBottom = '1px solid #ddd';
        tr.appendChild(td);
      });
      if (msg.status === 'FAILED') tr.style.backgroundColor = '#fff1f0';
      el.appendChild(tr);
      const count = document.getElementById('pms_collection_count');
      if (count) count.textContent = msg.index + ' / ' + msg.total + '종목 · 성공 ' + msg.ok + ' · 실패 ' + msg.failed;
      const sc = document.getElementById('pms_collection_scroll');
      if (sc) sc.scrollTop = sc.scrollHeight;
    });

    // v0.5.44: backend on_result after actual successful collection owns source truth.
    function pmsSourceGroup(source){
      source=String(source||'').toUpperCase();
      if(source.indexOf('QUANTMOD')>=0) return 'quantmod';
      if(source.indexOf('NAVER')>=0) return 'naver';
      if(source.indexOf('YAHOO')>=0) return 'yahoo';
      return 'unknown';
    }
    const pmsSourceCounts={prices:{quantmod:0,naver:0,yahoo:0,unknown:0,failed:0},dividends:{quantmod:0,naver:0,yahoo:0,unknown:0,failed:0}};
    function pmsSourceUpdate(msg){
      const kind=msg.kind==='dividends'?'dividends':'prices';
      const counts=pmsSourceCounts[kind];
      const root=document.getElementById('pms_sources_'+kind);
      if(!root)return;
      if(msg.event==='reset'){Object.keys(counts).forEach(k=>counts[k]=0);}
      if(msg.event==='result'){
        const key=msg.ok?pmsSourceGroup(msg.source):'failed';counts[key]++;
        const card=root.querySelector('[data-provider='+key+']');
        if(card){card.classList.remove('pms-src-flash');void card.offsetWidth;card.classList.add('pms-src-flash');}
      }
      root.querySelectorAll('[data-provider]').forEach(function(card){
        const key=card.getAttribute('data-provider');
        const num=card.querySelector('.pms-src-num');if(num)num.textContent=counts[key]||0;
      });
      const status=root.querySelector('.pms-src-meta');
      if(status){
        if(msg.event==='reset') status.textContent='조회 준비 · quantmod 우선, 실패 시 기존 공급원으로 재시도';
        else if(msg.event==='start') status.textContent='조회 중 · '+(msg.index||0)+'/'+(msg.total||'?')+' · '+(msg.symbol||'')+' (첫 시도: quantmod)';
        else if(msg.event==='result') status.textContent=(msg.index||0)+'/'+(msg.total||'?')+' · '+(msg.symbol||'')+' → '+(msg.ok?(msg.source||'성공 출처 미확인'):'조회 실패')+(msg.events!=null?' · 이벤트 '+msg.events+'건':'');
        else if(msg.event==='done')status.textContent='수집 완료 · 성공 출처별 건수는 실제 서버 응답 기준';
      }
    }
    Shiny.addCustomMessageHandler('pms_source_visual',pmsSourceUpdate);
    Shiny.addCustomMessageHandler('pms_dividend_source_reset', function(){const el=document.getElementById('pms_dividend_source_rows');if(el)el.innerHTML='';});
    Shiny.addCustomMessageHandler('pms_dividend_source_row', function(msg){
      const el=document.getElementById('pms_dividend_source_rows');if(!el)return;
      const tr=document.createElement('tr');
      [msg.index,msg.symbol,msg.status,msg.source,msg.events].forEach(function(v){const td=document.createElement('td');td.textContent=v==null?'':String(v);td.style.padding='5px 8px';tr.appendChild(td);});
      if(msg.status!=='OK')tr.style.background='#fff1f0';el.appendChild(tr);
    });
    // v0.5.21: persist BOTH the browser viewport and DT scroll during output replacement.
    // A DT draw event alone is not reliable: Shiny may replace #tbl after that event.
    var pmsHoldingScrollSnapshot = null;
    var pmsHoldingRestoreTimer = null;
    function pmsHoldingsBody() {
      var root = document.getElementById('tbl');
      return root ? root.querySelector('.dataTables_scrollBody') : null;
    }
    function pmsRememberHoldingsPosition() {
      var body = pmsHoldingsBody();
      var table = window.jQuery ? jQuery('#tbl table.dataTable') : null;
      var dt = table && table.length && jQuery.fn.dataTable.isDataTable(table[0])
        ? table.DataTable() : null;
      var root = document.getElementById('tbl');
      pmsHoldingScrollSnapshot = {
        y: body ? body.scrollTop : 0,
        x: body ? body.scrollLeft : 0,
        page: dt ? dt.page() : 0,
        windowY: window.scrollY || document.documentElement.scrollTop || 0,
        windowX: window.scrollX || document.documentElement.scrollLeft || 0,
        // Keep relative position if the outer document changes height.
        anchorTop: root ? root.getBoundingClientRect().top : null,
        expires: Date.now() + 4500
      };
    }
    function pmsRestoreHoldingsPosition() {
      var s = pmsHoldingScrollSnapshot;
      if (!s || Date.now() > s.expires) return;
      var body = pmsHoldingsBody();
      if (body) {
        var table = jQuery('#tbl table.dataTable');
        if (table.length && jQuery.fn.dataTable.isDataTable(table[0])) {
          var dt = table.DataTable();
          var info = dt.page.info();
          if (info && s.page < info.pages && dt.page() !== s.page) dt.page(s.page).draw('page');
        }
        body.scrollTop = s.y;
        body.scrollLeft = s.x;
      }
      window.scrollTo(s.windowX, s.windowY);
    }
    function pmsRestoreHoldingAfterSave() {
      if (!pmsHoldingScrollSnapshot) return;
      if (pmsHoldingRestoreTimer) window.clearInterval(pmsHoldingRestoreTimer);
      var attempts = 0;
      // Repeated restoration is intentional: Shiny DT replacement and CSS layout
      // can finish asynchronously after the first paint.
      pmsHoldingRestoreTimer = window.setInterval(function() {
        attempts++;
        pmsRestoreHoldingsPosition();
        if (attempts >= 18) {
          window.clearInterval(pmsHoldingRestoreTimer);
          pmsHoldingRestoreTimer = null;
          pmsHoldingScrollSnapshot = null;
        }
      }, 120);
    }
    if (window.jQuery) {
      jQuery(document).on('mousedown touchstart', '#save_record', pmsRememberHoldingsPosition);
      jQuery(document).on('click', '#save_record', function() {
        if (!pmsHoldingScrollSnapshot) pmsRememberHoldingsPosition();
        pmsRestoreHoldingAfterSave();
      });
      // Restore on both Shiny output replacement and DT initialization.
      jQuery(document).on('shiny:value', function(e) {
        if (e.name === 'tbl' && pmsHoldingScrollSnapshot) pmsRestoreHoldingAfterSave();
      });
      jQuery(document).on('draw.dt init.dt', function(e) {
        if (!pmsHoldingScrollSnapshot) return;
        var root = document.getElementById('tbl');
        if (root && root.contains(e.target)) window.setTimeout(pmsRestoreHoldingsPosition, 0);
      });
    }
    Shiny.addCustomMessageHandler('pms_holdings_preserve_position', function() {
      if (!pmsHoldingScrollSnapshot) pmsRememberHoldingsPosition();
      pmsRestoreHoldingAfterSave();
    });
    Shiny.addCustomMessageHandler('pms_set_zoom', function(value) {
      // Server sync; local click handler in a separate script also applies zoom instantly.
      if (typeof window.pmsApplyZoom === 'function') window.pmsApplyZoom(value);
    });
  "))),
  # 독립 스크립트로 등록해 다른 UI 로직의 지연과 관계없이 즉시 반응합니다.
  tags$script(HTML("
    (function() {
      window.pmsZoomCurrent = 100;
      window.pmsApplyZoom = function(value) {
        var num = Number(value);
        if (!isFinite(num)) return;
        num = Math.max(70, Math.min(150, Math.round(num / 10) * 10));
        window.pmsZoomCurrent = num;
        var el = document.getElementById('pms_zoom_content');
        if (el) {
          el.style.setProperty('zoom', (num / 100).toString(), 'important');
          el.setAttribute('data-pms-zoom', String(num));
        }
        var label = document.getElementById('pms_zoom_local_label');
        if (label) label.textContent = '화면 배율: ' + num + '%';
      };
      window.pmsZoomClick = function(direction) {
        var current = window.pmsZoomCurrent || 100;
        window.pmsApplyZoom(direction === 0 ? 100 : current + direction * 10);
      };
      // Apply after DOM ready. The server remains the source of truth on reconnect.
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() { window.pmsApplyZoom(window.pmsZoomCurrent); });
      } else { window.pmsApplyZoom(window.pmsZoomCurrent); }
    })();
  ")),
  # 확대/축소 버튼은 배율 적용 영역 밖에 두어 크기를 일정하게 유지합니다.
  tags$div(style='display:flex;align-items:center;gap:8px;margin:12px 0 8px 0;',
    actionButton('zoom_out', label='− 10%', class='btn-default', onclick='window.pmsZoomClick(-1)'),
    actionButton('zoom_reset', label='100%', class='btn-default', onclick='window.pmsZoomClick(0)'),
    actionButton('zoom_in', label='+ 10%', class='btn-default', onclick='window.pmsZoomClick(1)'),
    tags$span(id='pms_zoom_local_label',style='margin-left:8px;font-weight:bold;', '화면 배율: 100%'),
    tags$button(id='pms_dark_toggle',type='button',class='btn btn-default',style='margin-left:12px;',onclick='window.pmsSetDarkMode(!document.body.classList.contains(\"pms-dark\"))','☾ 다크모드')
  ),
  tags$div(id='pms_zoom_content',
  titlePanel('PMS2 | 개인금융자산 통계연구소'),
  tags$head(tags$script(HTML("
    (function(){
      function registerFetchClock(){
        if(!window.Shiny || typeof Shiny.addCustomMessageHandler!=='function') return;
        Shiny.addCustomMessageHandler('pms_board_fetch_time_v608',function(message){
          var el=document.getElementById('pms_board_last_fetch_text');
          if(el) el.textContent=(message && message.text) || '최근 시세 수집: 기록 없음';
        });
      }
      registerFetchClock();
      document.addEventListener('DOMContentLoaded',registerFetchClock);
      if(window.jQuery) jQuery(document).on('shiny:connected',registerFetchClock);
    })();
  "))),
  tags$head(tags$style(HTML("
    /* v0.6.09: enforce legible board header in both theme modes.
       Bootstrap 3 does not define --bs-body-bg. */
    .pms-board-header {background:#eef3f8!important;color:#182b40!important;
      border:1px solid #cbd7e3;border-radius:4px;}
    .pms-board-header h3,.pms-board-header span {color:inherit!important;}
    body.pms-dark .pms-board-header {background:#192839!important;color:#f3f7ff!important;
      border-color:#4f657c!important;}
    body.pms-dark .pms-board-header h3,body.pms-dark .pms-board-header span {color:#f3f7ff!important;}
    #pms_board_last_fetch_text {font-size:13px;font-weight:650;}
    .pms-board-header .pms-move-up, body.pms-dark .pms-board-header .pms-move-up {color:#d43838!important;}
    .pms-board-header .pms-move-down, body.pms-dark .pms-board-header .pms-move-down {color:#267dc5!important;}
    .pms-board-header .pms-move-flat, body.pms-dark .pms-board-header .pms-move-flat {color:#667788!important;}
    .pms-price-roll {display:inline-block;overflow:hidden;vertical-align:middle;line-height:1.35;}
    .pms-price-roll-track {display:flex;flex-direction:column;white-space:nowrap;animation:pms-price-roll-up .65s cubic-bezier(.2,.75,.25,1) both;}
    .pms-price-roll-track > span {display:block;line-height:1.35;}
    @keyframes pms-price-roll-up {from {transform:translateY(0);} to {transform:translateY(-50%);}}
    @media (prefers-reduced-motion:reduce) {.pms-price-roll-track {animation:none;transform:translateY(-50%);}}
  "))),
  tags$head(tags$style(HTML("
    .pms-ai-page {max-width:1100px;margin:0 auto;padding-bottom:32px;}
    .pms-ai-card {border:1px solid #dce5ee;background:#f9fbff;border-radius:12px;padding:17px;margin:14px 0;}
    .pms-ai-card .shiny-text-output {max-width:100%;} 
    body.pms-dark .pms-ai-card {background:#1e2e3e;color:#ecf4ff;border-color:#41586b;}
    @media(max-width:768px){.pms-ai-card{padding:11px;margin:9px 0;} .pms-ai-page h3{font-size:19px;}}
  "))),
  tabsetPanel(id='pms_tabs', type='tabs',
  tabPanel('홈',
    tags$div(class='pms-home',
      h3('내 자산 한눈에 보기'),
      tags$p(class='text-muted','SQLite에 저장된 마지막 시세·배당정보 기준입니다. 홈을 열어도 외부 API를 호출하지 않습니다.'),
      uiOutput('home_dashboard'),
      tags$div(style='margin-top:10px;',actionButton('home_refresh','저장된 정보 새로고침',class='btn-default')),
      tags$p(class='text-muted',style='margin-top:8px;font-size:12px;',
        '현금·예수금은 기존 보유종목 평가액에 포함되지 않습니다. 배당금은 실제 수령액이 아닌 현재 보유수량 기준 세전 추정액입니다.')
    )
  ),
  tabPanel('전일대비전광판',
    tags$div(class='pms-board-header',style='display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;position:sticky;top:47px;z-index:1020;padding:7px 12px;',
      tags$div(style='display:flex;align-items:center;flex-wrap:wrap;gap:14px;',
        h3(style='margin:7px 0;','전일대비전광판 · 보유종목 최신 시세', tags$small(style='font-size:11px;color:#835d18;margin-left:9px;', '진단판 10.20-B')),
        tags$div(class='pms-board-totals',style='display:flex;align-items:center;flex-wrap:wrap;gap:11px;font-size:15px;font-weight:700;',
          uiOutput('board_header_totals',inline=TRUE))),
      tags$div(style='display:flex;flex-wrap:wrap;justify-content:flex-end;align-items:center;gap:12px;',
        tags$span(id='pms_board_live_clock',style='font-size:14px;font-weight:700;white-space:nowrap;',
                  '한국시간 확인 중…'),
        tags$span(style='font-size:13px;font-weight:650;white-space:nowrap;',
                  textOutput('pms_board_rendered_at',inline=TRUE)),
        tags$span(style='font-size:13px;font-weight:650;white-space:nowrap;',
                  textOutput('pms_board_auto_detail',inline=TRUE)),
        tags$span(style='font-size:14px;font-weight:700;white-space:nowrap;',
                  tags$span(id='pms_board_last_fetch_text', '최근 시세 수집: 조회 중…')))),
    uiOutput('board_missing_detail'),
    tags$p(class='text-muted','현재가와 직전 거래일 종가 비교 · 10분 자동수집 옵션에 따라 새 저장값 반영'),
    tags$div(style='display:flex;gap:10px;align-items:center;',
      actionButton('board_refresh','저장된 시세로 지금 갱신',class='btn-default')),
    tags$div(style='margin:8px 0 10px;',
      actionButton('board_toggle_cards','종목별 시세 카드 펼치기',class='btn-default')),
    uiOutput('board_cards_container'),
    h4('종목별 총평가액 · 매수금 대비 손익/수익률 · 전체비중 데이터바 · 전일대비 평가증감'),
    tags$p(class='text-muted','동일 종목을 여러 계좌에 보유하면 수량을 합산합니다. 총평가액 = 현재가 × 총보유수량 × 적용환율, 평가증감 = (현재가 − 전일종가) × 총보유수량 × 적용환율. 환율 자체의 변동 효과는 제외합니다. 매수금대비손익 = 총평가액 − 총취득원가, 매수금대비수익률 = 손익 ÷ 총취득원가 × 100입니다. 취득원가가 누락되면 산출 불가로 표시합니다. 금액 산출 불가 종목은 맨 아래에 표시합니다.'),
    DTOutput('board_value_change_table'),
    conditionalPanel(condition="input.pms_tabs == '전일대비전광판'",
      tags$div(class='pms-loading-shell',role='status',
        tags$div(class='pms-loading-row',tags$div(class='pms-loading-text','자산 요약 데이터 조회중입니다. 잠시만 기다려 주세요...')),
        tags$div(class='pms-loading-row',tags$div(class='pms-loading-text','시장지수 및 보유종목 데이터 조회중입니다. 잠시만 기다려 주세요...'))),
      uiOutput('board_bottom_ticker'))
  ),
  tabPanel('일별스냅샷',
    h3('일별 포트폴리오 스냅샷'),
    tags$p('앱을 시작하면 당일 첫 스냅샷을 자동 저장하고, 시세 재수집이 완료되면 당일 스냅샷을 최신 값으로 다시 저장합니다. 과거 날짜는 덮어쓰지 않습니다.'),
    tags$p('기준은 SQLite에 저장된 현재 보유종목과 마지막 시세이며, 현금·예수금과 실제 외부 입출금은 포함되지 않습니다.'),
    actionButton('snapshot_save_now','현재 자료로 오늘 스냅샷 저장',class='btn-primary'),
    tags$div(style='margin:10px 0;',textOutput('snapshot_save_status')),
    tags$p('저장 위치: data/pms.sqlite + snapshots/YYYY-MM-DD.csv · 같은 날짜의 CSV는 덮어써서 1개만 유지합니다.'),
    DTOutput('snapshot_history_table'),
    tags$hr(),
    h4('종목별 평가액 비중 트리맵'),
    tags$p(class='text-muted','현재 SQLite에 저장된 보유종목·마지막 시세 기준입니다. 과거 날짜별 보유종목 구성은 표시하지 않습니다. 사각형 면적은 종목별 원화 평가액에 비례합니다.'),
    plotOutput('snapshot_symbol_treemap',height='520px'),
    h4('자산군별 평가액 비중 트리맵'),
    tags$p(class='text-muted','위 종목별 트리맵과 동일한 평가액을 자산군(주식·채권·금·현금성·기타) 기준으로 합산합니다.'),
    plotOutput('snapshot_asset_treemap',height='300px'),
    textOutput('snapshot_treemap_note')
  ),
  tabPanel('평가액추이',
    h3('평가액·평가손익 시계열 분석'),
    tags$p(class='text-muted','2024-09-17부터 과거 PMS 전체 집계자료와 이후 일별 SQLite 스냅샷을 연결합니다. 과거 자료에는 종목별 상세 내역이 없으며, 실제 입출금이 기록되지 않아 운용수익률(TWR)로 해석할 수 없습니다.'),
    fluidRow(column(4,selectInput('trajectory_period','조회 기간',choices=c('전체 기간'='all','최근 1년'='365','최근 6개월'='180','최근 3개월'='90','최근 1개월'='30'),selected='all')),
             column(8,tags$p(class='text-muted',style='margin-top:30px;',textOutput('trajectory_source_note')))),
    h4('총평가액 · 평가손익 통합 추이 (왼쪽/오른쪽 독립 Y축)'),
    tags$p(class='text-muted','왼쪽 파란 눈금은 총평가액(억원), 오른쪽 주황 눈금은 평가손익(억원)입니다. 각각 독립된 범위로 표시되므로 선의 높이를 서로 직접 비교하지 마세요.'),
    plotOutput('trajectory_dual_axis',height='460px'),
    h4('선형회귀 추세선 · 일평균 기울기'),
    tags$p(class='text-muted','선택 기간의 관측 날짜를 실제 일수로 변환해 최소제곱 직선을 계산합니다. 기울기는 하루당 원화 변화량이며, 추가 납입·인출의 영향을 포함하므로 투자수익률이 아닙니다.'),
    fluidRow(
      column(6,textOutput('trajectory_value_slope')),
      column(6,textOutput('trajectory_profit_slope'))
    ),
    h4('전량 매도 시 추정 세후 평가손익'),
    fluidRow(column(4,numericInput('trajectory_domestic_tax','국내상장 미국지수 ETF 차익 적용세율 (%)',value=38,min=0,max=100,step=0.5)),
      column(8,uiOutput('trajectory_aftertax_result'))),
    tags$p(class='text-muted','시뮬레이션 가정: 국내 상장 미국지수 ETF(tax_type=KR_US_INDEX_ETF)의 양(+)의 평가차익 합계에만 입력 세율을 적용합니다. 그 밖의 일반 국내주식과 국내 보유분은 본 단순 계산에서 세금을 0원으로 가정합니다. 해외 보유분은 양(+)의 평가차익 합계에서 20,000,000원을 차감한 초과분에 22%를 적용합니다. 손실에는 세금을 부과하지 않습니다. 실제 과세는 종목 유형·계좌 유형·거주자 여부·연간 실현손익·공제 규정에 따라 다르므로 신고 예상세액이 아닙니다. 과거 평가액 추이와 별개로 현재 보유분의 저장 시세를 사용합니다.'),
    h4('날짜별 평가 기록'),
    DTOutput('trajectory_table')
  ),
  tabPanel('투자성과분석',
    h3('투자성과분석 · 평가액 및 위험 추이'),
    tags$p(class='text-muted','과거 520일 총평가액·미실현손익 기록과 SQLite 일별 스냅샷을 이용합니다. 외부 API를 호출하지 않습니다.'),
    fluidRow(column(4,selectInput('performance_period','분석 기간',
       choices=c('전체 기간'='all','최근 1년'='365','최근 6개월'='180','최근 3개월'='90','최근 1개월'='30'),selected='all')),
      column(8,tags$p(class='text-muted',style='margin-top:30px;',textOutput('performance_source_note')))),
    uiOutput('performance_dashboard'),
    h4('최대하락폭(MDD) 및 평가액 변화 상세'),
    DTOutput('performance_detail')
  ),
  tabPanel('위험관리',
    h3('위험관리 · 기존 PMS 리스크 모듈 기반'),
    tags$p(class='text-muted','기존 PMS의 현금흐름 추정 공식, 몬테카를로, VaR·CVaR·변동성 경보 개념을 SQLite 기록에 연결했습니다. 저장값만 사용하며 외부 API를 호출하지 않습니다.'),
    fluidRow(
      column(4,selectInput('risk_period','분석 기간',choices=c('전체 기간'='all','최근 1년'='365','최근 6개월'='180','최근 3개월'='90'),selected='all')),
      column(4,selectInput('risk_confidence','VaR/CVaR 신뢰수준',choices=c('95%'='0.95','99%'='0.99'),selected='0.95')),
      column(4,tags$p(style='margin-top:29px;color:#65758c;','현금흐름은 추정값 · 무위험수익률 0%'))),
    uiOutput('risk_cards'),
    tags$details(style='margin:8px 0 15px;',tags$summary(tags$b('변동성 경보 NORMAL / ALERT / CRITICAL 판정 기준 보기')),
      tags$p('실제 계산식: 최근 60개 유효 일별 추정 수익률에 가중치 0.94를 적용한 EWMA 변동성 ÷ 전체 유효 수익률의 역사적 표준편차.'),
      tags$p('NORMAL: 비율 1.2 이하 · ALERT: 1.2 초과 1.8 이하 · CRITICAL: 1.8 초과 · 측정 불가: 계산 자료 부족.'),
      tags$p('현재 판정값: ',textOutput('risk_volatility_explanation',inline=TRUE)),
      tags$p('NORMAL은 손실이 없다는 뜻이 아니며, 추정 입출금 기반 지표이므로 참고용입니다.')),
    uiOutput('risk_explanations'),
    fluidRow(column(6,plotOutput('risk_nav_chart',height='255px')),
             column(6,plotOutput('risk_dd_chart',height='255px'))),
    h4('미래 위험 시나리오 · Monte Carlo'),
    tags$p(class='text-muted','정규분포 로그수익률 기반 예시 시뮬레이션입니다. 예측이 아닌 가정 계산이며, 과거 추정수익률에 편향이 있으면 결과도 왜곡됩니다.'),
    fluidRow(
      column(3,selectInput('risk_mc_mode','시나리오',choices=c('적립기'='save','은퇴 인출기'='withdraw'))),
      column(2,numericInput('risk_mc_years','기간(년)',10,min=1,max=40,step=1)),
      column(3,numericInput('risk_mc_monthly','월 적립금(원)',1000000,min=0,step=100000)),
      column(3,numericInput('risk_mc_withdraw','연간 인출액(원)',48000000,min=0,step=1000000)),
      column(1,tags$div(style='margin-top:24px;',actionButton('risk_run_mc','실행',class='btn-primary')))),
    uiOutput('risk_mc_summary'),
    h4('상세 위험 통계'),DTOutput('risk_stats_table'),
    tags$div(class='alert alert-info',style='margin-top:16px;',
      tags$b('선택적 고급 분석: '),'GARCH(1,1)는 rugarch 패키지와 충분한 관측수가, PCA·팩터 회귀·과거 위기 리플레이는 별도의 자산별 가격·팩터·목표비중 자료가 필요합니다. 현재 기록만으로 결과를 꾸며내지 않으며 별도 연동 전까지는 제외합니다.')
  ),
  tabPanel('매매내역분석',
    h3('보유수량 변화 달력 · 매매 추정'),
    tags$p(class='text-muted','과거 엑셀 기록과 신규 SQLite 스냅샷은 계좌 식별 방식이 다릅니다. 두 출처가 바뀌는 날짜 사이를 매매로 판정하지 않습니다. 같은 출처의 연속 보유기록만 비교하며, 실제 체결내역이 아닙니다.'),
    fluidRow(column(4,selectInput('trade_month','조회 기간',choices=NULL)),
             column(8,tags$p(style='margin-top:30px;',textOutput('trade_stats')))),
    uiOutput('trade_calendar'),
    tags$hr(),h4('수량 변화 상세'),DTOutput('trade_changes_table')
  ),
  tabPanel('보유종목 편집',
  tags$p('현재 정렬: 증권사별 묶음 → 국가(한국 → 미국) → 계좌유형(일반 → ISA → IRP → 연금저축 → 기타) → 같은 그룹 내 평가금액 내림차순. 증권사명은 한글로 표시되며, 평가액은 마지막 저장 시세 기준입니다. ID는 DB 고유번호입니다.'),
  tags$div(class='well pms-excel-panel',
    h4('엑셀로 일괄 편집'),
    tags$p('현재 보유종목 전체를 엑셀로 내려받아 수정하세요. 기존 ID는 유지하고, 신규 행은 ID와 updated_at을 비워 두세요. 엑셀에서 삭제한 행은 DB에서 삭제되지 않습니다.'),
    downloadButton('holdings_xlsx_download','보유종목 엑셀 다운로드'),
    tags$span(style='margin-left:12px;color:#46627a;font-size:12px;',textOutput('holdings_xlsx_download_ready',inline=TRUE)),
    fileInput('holdings_xlsx_upload','편집한 엑셀 업로드 (.xlsx)',accept=c('.xlsx')),
    tags$p(class='text-muted','업로드 후 검증 결과를 확인한 다음 적용 버튼을 눌러야 저장됩니다. 계좌번호가 포함되니 파일 보관에 유의하세요.'),
    textOutput('holdings_xlsx_status'),
    DTOutput('holdings_xlsx_preview'),
    uiOutput('holdings_xlsx_confirm_ui')
  ),
  tags$head(tags$style(HTML("
    .pms-excel-panel{padding:9px 13px!important;margin-bottom:10px!important;}
    .pms-excel-panel h4{margin:0 0 4px!important;font-size:15px;}
    .pms-excel-panel p{margin:3px 0!important;font-size:12px;line-height:1.4;}
    .pms-excel-panel .form-group{margin-bottom:3px!important;}
    .pms-excel-panel .form-group>label{margin-bottom:2px!important;font-size:12px;}
    .pms-excel-panel .shiny-input-container{margin-bottom:3px!important;}
    .pms-excel-panel .input-group{margin-bottom:0!important;}
    .pms-excel-panel .progress{margin-bottom:0!important;}
  "))),
  tags$p('왼쪽 표에서 종목을 클릭하면 오른쪽에서 자동으로 수정할 수 있습니다.'),
  fluidRow(
    column(8,
      checkboxInput('show_zero','수량 0 종목 표시',FALSE),
      fluidRow(column(9, selectInput('holding_pick_id','종목 직접 선택 (행 클릭이 안 될 때)', choices=c('선택하세요'=''))),
               column(3, tags$div(style='margin-top:25px;',actionButton('holding_edit_selected','선택 종목 편집',class='btn-primary')))),
      tags$div(style='display:flex;align-items:end;gap:8px;margin:3px 0 7px;flex-wrap:wrap;',
        tags$div(style='width:145px;',numericInput('holding_id_manual','DB ID 직접 입력',value=NA,min=1,step=1,width='145px')),
        actionButton('holding_edit_by_id','ID로 편집',class='btn-warning'),
        tags$span(style='font-size:12px;color:#64748b;','행 클릭이 안 될 때 ID 열의 번호를 입력하세요.')),
      DTOutput('tbl'),
      tags$p('행을 클릭하면 오른쪽 입력 폼이 자동으로 바뀝니다. 전량매도는 수량 0으로 저장하세요.')
    ),
    column(4,
      wellPanel(
        h4(textOutput('form_title', inline = TRUE)),
        actionButton('new','신규 입력'),
        tags$hr(),
        fluidRow(column(6,selectInput('f_country','국가',choices$country)),
                 column(6,selectInput('f_currency','통화',choices$currency))),
        textInput('f_symbol','종목코드 *'),
        textInput('f_name','종목명'),
        fluidRow(column(6,selectInput('f_asset_class','자산군',choices$asset_class)),
                 column(6,selectInput('f_tax_type','과세유형',choices$tax_type))),
        tags$p(style='font-size:12px;color:#606f7b;',
               '미국지수ETF: 국내 상장 미국 S&P500·나스닥100·미국배당다우존스(SCHD형) ETF에 사용하는 사용자 분류입니다. 일반적인 국내 주식형 ETF와 구분해 기록합니다. 이 항목만으로 실제 세금이 자동 계산되지는 않습니다.'),
        fluidRow(column(6,selectInput('f_broker','증권사',choices$broker)),
                 column(6,selectInput('f_account_type','계좌유형',choices$account_type))),
        selectInput('f_account_number','기존 계좌번호 선택',choices=c('새 계좌번호 직접 입력'=''),selected=''),
        textInput('f_account_number_manual','새 계좌번호 직접 입력 (선택)',value='',
                  placeholder='목록에 없으면 여기에 입력하세요'),
        tags$p(class='text-muted',style='font-size:12px;',
          '기존 계좌번호를 선택하거나 새 번호를 직접 입력하세요. 직접 입력한 번호가 있으면 목록 선택보다 우선 저장됩니다.'),
        selectInput('f_fractional_share','소수점 주식 여부 (한국·미국 공통)',
                    choices=c('아니오 (일반주식)'='0','예 (소수점주식)'='1'),selected='0'),
        tags$p(class='text-muted',style='font-size:12px;',
          '같은 종목·증권사·계좌라도 일반주식과 소수점주식을 서로 다른 보유내역으로 등록할 수 있습니다.'),
        fluidRow(column(6,numericInput('f_quantity','보유수량',0,min=0,step=0.000001)),
                 column(6,textInput('f_avg_price','평균매수가','0'))),
        textInput('f_avg_fx_rate','평균 매입환율 (한국 주식은 1 입력)','1'),
        textInput('f_cost_krw','원화 취득원가 (전체 보유수량의 원화 매입원금, 없으면 공란)'),
        tags$p(class='text-muted',style='font-size:12px;',
          '증권사에서 확인한 해당 보유내역의 실제 원화 매입원금이 있으면 입력합니다. 공란이면 보유수량 × 평균매수가 × 평균매입환율로 추정합니다. 거래 수수료 포함 여부는 증권사 기록을 확인하세요.'),
        tags$div(tags$strong('마지막 수정일시 (자동 기록)'), textOutput('updated_at_display')),
        textAreaInput('f_memo','메모',rows=3),
        actionButton('save_record','저장',class='btn-primary'),
        actionButton('delete_record','선택 행 삭제',class='btn-danger'),
        tags$hr(),
        strong('처리 결과'),
        verbatimTextOutput('notice')
      )
    )
  )
  ), # 보유종목 편집 탭
  tabPanel('시세수집 · 자산평가',
  h3('시세 및 보유자산 평가'),
  tags$div(id='pms_saved_eval_diagnostic',class='alert alert-info',style='padding:9px 12px;margin-bottom:10px;', '평가 버튼 동작 확인 대기 중입니다.'),
  textOutput('saved_eval_server_status'),
  tags$p('앱 실행 시 SQLite에 저장된 시세·환율·배당을 자동으로 표시합니다. 저장 데이터 확인에는 인터넷을 사용하지 않습니다. 시세는 거래소 실시간 체결가가 아닐 수 있습니다.'),
  h4('저장된 시세 기준 전체 평가 현황'),
  uiOutput('valuation_dashboard'),
  tags$p(class='text-muted',style='font-size:12px;', '앱을 열거나 이 탭을 선택하면 SQLite의 최근 저장값을 이용해 자동 계산합니다. 시세와 배당의 신규 조회는 수행하지 않습니다.'),
  tags$div(class='alert alert-success',style='margin-top:6px;', textOutput('cached_data_status')),
  h4('연결 및 시세 수신 과정'),
  tags$div(class='pms-flow-panel',role='status',`aria-live`='polite',
    tags$div(class='pms-flow-header',tags$strong(id='pms_flow_title','연결 확인 대기'),
             tags$span(id='pms_flow_indicator',class='pms-flow-marker pms-state-waiting','대기')),
    tags$div(id='pms_flow_steps',class='pms-flow-steps'),
    tags$p(id='pms_flow_detail',class='pms-flow-detail','연결 진단 또는 시세 수집을 누르면 과정이 표시됩니다.'),
    tags$small('회전 표시는 연결 작업 중임을 의미하며, 서버 응답이 확인되기 전에는 성공으로 표시하지 않습니다.')
  ),
  h4('시세 공급원 연결 진단'),
  tags$p('전체 수집 전에 4개 공급원 연결을 각각 시험합니다. 연결 실패와 SSL 인증서 오류, HTTP 429를 구별합니다.'),
  tags$div(class='alert alert-info',style='margin-bottom:12px;',
    tags$strong('테스트 수신값(주가·환율)이란?'),
    tags$p(style='margin-top:6px;margin-bottom:6px;',
      '각 공급원이 실제로 가격 또는 환율 숫자 하나를 전달할 수 있는지 확인한 값입니다. 보유하지 않은 종목도 통신 시험용으로 조회됩니다.'),
    tags$ul(style='margin-bottom:6px;',
      tags$li(tags$b('NAVER KR / SK하이닉스(000660): '),'국내 주식 한 주 가격 (원)'),
      tags$li(tags$b('NAVER US / 엔비디아(NVDA): '),'미국 주식 한 주 가격 (달러)'),
      tags$li(tags$b('YAHOO / 애플(AAPL): '),'미국 주식 한 주 가격 (달러)'),
      tags$li(tags$b('NAVER FX / USD/KRW: '),'미국 달러 1달러를 원화로 바꿀 때의 기준 환율 (원/달러)')),
    tags$small('HTTP 200은 통신 성공을 뜻하지만 모든 보유종목의 조회 성공을 보장하지 않습니다. 표시된 숫자는 진단 실행 시점에 받은 데이터입니다.')
  ),
  tags$div(class='alert alert-info',style='margin-top:12px;margin-bottom:10px;',
    tags$strong('SSL 연결 오류, 왜 생기나요?'),
    tags$p(style='margin-top:7px;',
      '시세를 받기 전에 웹서버의 디지털 신분증(SSL 인증서)을 검사합니다. 일부 네트워크에서는 인증서가 폐기됐는지 확인하는 곳에 연결하지 못해 시세 수집이 멈출 수 있습니다.'),
    tags$p('아래 옵션을 선택하면 인증서의 폐기 여부 검사만 생략하고 다시 연결합니다. HTTPS 암호화 및 다른 인증서 검사는 유지됩니다.'),
    tags$p(style='margin-bottom:0;',
      tags$strong('안내: '),
      '기본적으로 체크하지 마세요. SSL 폐기 확인 오류가 확인되고 네트워크 보안정책상 허용될 때만 선택하세요. 이 설정에서는 폐기된 인증서를 놓칠 위험이 있습니다.')
  ),
  checkboxInput('ssl_revocation_skip',
    'SSL 인증서 폐기 확인 오류 시 폐기 여부 검사 생략 (--ssl-no-revoke)',
    FALSE),
  tags$small(style='display:block;margin-top:-5px;margin-bottom:12px;color:#666;',
    '기본 체크 해제. 폐기 확인 실패가 명확히 감지되면 해당 요청만 자동 재시도하며, 나머지 인증서 검증은 유지합니다.'),
  tags$div(style='margin:4px 0 8px;color:#935b08;',textOutput('pms_ssl_auto_retry_notice')),
  actionButton('diagnose_sources','공급원 연결 진단',class='btn-info'),
  textOutput('diagnosis_notice'),
  DTOutput('diagnosis_table'),
  tags$hr(),
  tags$div(class='well',style='margin-bottom:14px;',
    tags$strong('저장된 시세로 즉시 평가 (API 접속 없음)'),
    tags$p('오늘 이미 수집한 시세가 있으면 다시 조회할 필요가 없습니다. 기존 가격·환율로 평가금과 손익을 다시 계산합니다.'),
    actionButton('refresh_valuation','저장된 시세로 평가',class='btn-success'),
    tags$div(style='margin-top:8px;',textOutput('cached_quote_status')),
    tags$small('오늘 수집한 가격이 일부 없더라도 마지막 저장가격으로 계산합니다. 기준일을 확인하세요. 배당은 다시 수집하지 않습니다.')
  ),
  tags$hr(),
  actionButton('collect_all','시세 및 배당 전체 조회',class='btn-success'),
  tags$span(style='margin-left:8px;', '시세 → 일별 USD/KRW 환율 → 최근 1년 배당을 순서대로 조회합니다.'),
  verbatimTextOutput('all_collect_notice'),
  tags$hr(),
  actionButton('collect_prices','시세 새로 수집',class='btn-primary'),
  tags$div(style='margin-top:12px;margin-bottom:8px;padding:10px 12px;background:#f3f8fe;border:1px solid #d8e6f4;border-radius:7px;',
    checkboxInput('auto_prices_10m','10분마다 시세 자동 조회 (기본 켜짐)',value=TRUE),
    tags$small('체크: 앱을 실행하는 동안 10분 간격으로 시세·현재 USD/KRW를 갱신합니다. 배당 이력은 다시 받지 않습니다. 체크 해제: 자동 갱신 중지. 수동 조회는 계속 가능합니다.'),
    tags$div(style='margin-top:5px;',textOutput('auto_prices_status'))
  ),
  checkboxInput('enable_yahoo_fallback','주 시세원 실패 시 대체 조회 (KR→Yahoo, US→Naver)',TRUE),
  checkboxInput('valuation_show_zero','평가 목록에 수량 0 표시',FALSE),
  verbatimTextOutput('collect_notice'),

  tags$div(id='pms_sources_prices',class='pms-src-panel',
    tags$strong('시세 · 수집 출처 실시간 현황'),
    tags$div(class='pms-src-grid',
      lapply(list(c('quantmod','quantmod / Yahoo'),c('naver','네이버 API'),c('yahoo','Yahoo JSON'),c('failed','조회 실패')),
        function(item) tags$div(class='pms-src-card',`data-provider`=item[1],
          tags$span(class='pms-src-label',tags$span(class='pms-src-dot'),item[2]),
          tags$span(class='pms-src-num','0'),tags$small(if(item[1]=='quantmod') '1차 성공' else if(item[1]=='failed') '실패 건수' else '대체 조회 성공')))),
    tags$div(class='pms-src-meta','수집 대기 · 성공 출처별 건수를 보여줍니다.')
  ),
  h4('종목별 시세 수집 현황 (이번 실행)'),
  tags$p(id='pms_collection_count', '시세수집 버튼을 누르면 각 종목의 결과가 표시됩니다.'),
  tags$div(id='pms_collection_scroll',style='max-height:350px;overflow:auto;border:1px solid #ddd;margin-bottom:14px;',
    tags$table(class='table table-condensed table-striped',style='margin:0;width:100%;font-size:12px;',
      tags$thead(tags$tr(lapply(c('순번','국가','종목코드','조회가격','수집처','상태','실패 원인'),tags$th))),
      tags$tbody(id='pms_collection_rows')
    )
  ),

  h4('최근 수집 실패 원인'),
  actionButton('clear_collection_failures','최근 실패 이력 지우기',class='btn-warning'),
  tags$small('실패 기록만 삭제합니다. 보유종목·정상 시세·평가 데이터는 그대로 보존됩니다.'),
  DTOutput('collection_failures'),
  tags$hr(),
  tags$p('비중(%)은 현재 평가 가능한 종목의 총 평가액을 100%로 계산합니다. 시세·환율이 없는 종목은 계산에서 제외됩니다. 빨간 숫자는 마이너스 금액이나 손실률입니다.'),
  h4('시세 신선도 점검 (SQLite 저장 시세 기준)'),
  tags$p('마지막 시세 기준일을 확인합니다. 오래된 시세는 거래정지·상장폐지·API 지연 등의 가능성을 알리는 경고이며 상장폐지 확정 판정이 아닙니다.'),
  tags$div(id='pms_price_freshness_anchor'),
  DTOutput('valuation_table'),
  h4('전체 합계 (수량 0 제외)'),
  tableOutput('valuation_summary'),
  textOutput('valuation_warning'),
  tags$p('평가액이 없는 종목은 0원으로 간주하지 않고 제외합니다. 미국 주식은 실시간 USD/KRW 시세가 없으면 평가불가로 표시합니다.'),
  tags$p('주의: KRX 금현물을 ACE ETF 상당 수량으로 입력했다면 평가값도 ETF 대체 추정치입니다.')
  ), # 시세수집 탭
  tabPanel('배당내역',
    h3('배당내역 · 예상 달력 · 배당 통계'),
    tags$p(class='text-muted','저장된 SQLite 배당자료로 통계를 표시합니다. 새 배당 데이터가 필요할 때만 수집 버튼을 누르세요.'),
  h4('최근 1년 종목별 현금배당 정보'),
  tags$p('Yahoo Finance에서 최근 365일 배당락일과 주당 배당금을 조회합니다. 미국 종목은 최근 1년 USD/KRW 일별 종가를 함께 저장해 배당락일 환율로 원화 환산합니다. 휴장일에는 최근 7일 이내 직전 거래일 환율을 사용하며, 없으면 원화 금액을 비워 둡니다.'),
  tags$p('주의: 배당락일 기준 세전 주당 배당 정보이며 실제 지급일 환율·원천징수 세액·입금액과 다릅니다. 국내 주식의 주당 원화 배당금은 원래 원화 금액을 그대로 표시합니다. 일부 국내 ETF의 배당 데이터는 누락될 수 있습니다.'),
  tags$p('SQLite에 이미 저장된 배당정보로 다시 통계를 계산하려면 아래 녹색 버튼을 누르세요. 인터넷 연결은 필요하지 않습니다.'),
  actionButton('evaluate_saved_dividends','저장된 배당금정보로 평가',class='btn-success'),
  tags$span(style='margin-left:8px;', '기존 배당·환율·현재 보유수량으로 모든 배당 통계 및 달력을 갱신합니다.'),
  tags$br(), tags$br(),
  actionButton('collect_dividends','최근 1년 배당정보 수집',class='btn-info'),

  tags$div(id='pms_sources_dividends',class='pms-src-panel',
    tags$strong('배당 · 수집 출처 실시간 현황'),
    tags$div(class='pms-src-grid',
      lapply(list(c('quantmod','quantmod / Yahoo'),c('naver','네이버 API'),c('yahoo','Yahoo JSON'),c('failed','조회 실패')),
        function(item) tags$div(class='pms-src-card',`data-provider`=item[1],
          tags$span(class='pms-src-label',tags$span(class='pms-src-dot'),item[2]),
          tags$span(class='pms-src-num','0'),tags$small(if(item[1]=='quantmod') '1차 성공' else if(item[1]=='failed') '실패 건수' else '대체 조회 성공')))),
    tags$div(class='pms-src-meta','수집 대기 · 성공 출처별 건수를 보여줍니다.')
  ),
  tags$details(tags$summary('종목별 배당 수집 출처 확인'),tags$div(style='max-height:280px;overflow:auto;',tags$table(class='table table-condensed table-striped',tags$thead(tags$tr(lapply(c('순번','종목코드','상태','출처','이벤트 수'),tags$th))),tags$tbody(id='pms_dividend_source_rows')))),
  textOutput('dividend_notice'),
  tags$p('최근 365일 세전 배당수익률 = 최근 1년 주당 배당금 합계 ÷ 마지막 저장 시세 × 100. 미국 종목은 달러 배당금과 달러 주가를 비교합니다. 시세/배당 조회 실패 시 수익률은 표시하지 않습니다.'),
  DTOutput('dividend_summary'),
  h4('향후 1개월 예상 배당일정 (작년 배당락일 기준)'),
  tags$p('작년 같은 시기의 배당락일을 1년 뒤로 옮겨 예상한 참고 일정입니다. 확정된 배당 공시나 지급일이 아닙니다. 작년 주당 배당금과 현재 보유수량 기준 원화 추정액을 함께 표시합니다. 미래 실제 배당일·금액·보유수량은 달라질 수 있습니다.'),
  textOutput('dividend_calendar_note'),
  # This is always visible even if the Shiny output binding fails.
  tags$div(id='pms_calendar_anchor',style='padding:8px;border:1px solid #d7e5f1;background:#f6faff;margin-bottom:8px;',
    '예상 배당 달력: SQLite 자료를 확인하고 있습니다…'),
  # Direct calendar inserted after anchor; legacy output retained server-side.
  tags$details(tags$summary('예상 배당일정 상세표 보기'), DTOutput('dividend_calendar')),
  h4('최근 1년 배당 이벤트 상세'),
  tags$div(id='pms_dividend_events_anchor',style='padding:8px;border:1px solid #d7e5f1;background:#f6faff;margin-bottom:8px;',
    'SQLite 배당 이벤트 상세 자료를 불러오는 중입니다…'),
  tags$details(tags$summary('기존 DataTables 상세표 보기'), DTOutput('dividend_events')),
  h4('계좌별 최근 1년 배당금 (현재 보유수량 기준 추정)'),
  tags$p('최근 365일 세전 주당 배당금 × 현재 보유수량으로 계산합니다. 실제 과거 보유수량, 지급일 환율, 세금 및 실수령액은 반영되지 않습니다. 미수집/환율 누락이 있으면 해당 금액을 0원으로 가정하지 않고 미완료건수로 표시합니다.'),
  DTOutput('dividend_account_totals'),
  h4('최근 12개월 월별 배당금 추정액 (계좌유형별)'),
  tags$p('배당락일 기준 월별 세전 추정액입니다. 각 배당 이벤트의 주당 원화금액 × 현재 보유수량으로 계산합니다. 막대 색은 일반계좌·IRP·ISA·연금저축·기타를 나타냅니다. 배당 당시의 실제 보유수량이나 실입금액과 다를 수 있습니다.'),
  tags$div(id='pms_monthly_dividend_anchor'),
  tags$details(tags$summary('기존 월별 누적막대그래프 보기'),
    plotOutput('dividend_monthly_chart',height='360px')),
  textOutput('dividend_monthly_note'),
  h4('일반계좌 배당금과 금융소득종합과세 기준선 비교'),
  tags$p('최근 365일 일반계좌 세전 배당 추정액을 연간 금융소득 2,000만 원 기준선과 비교합니다. 실제 과세 여부는 1월~12월의 이자·배당소득 전체, 과세 구분 및 실제 지급내역을 기준으로 판단해야 합니다. IRP·ISA·연금저축 집계액은 이 그래프에 합산하지 않습니다.'),
  tags$div(id='pms_tax_threshold_anchor'),
  tags$details(tags$summary('기존 금융소득 기준선 그래프 보기'),
    plotOutput('general_income_threshold',height='180px')),
  textOutput('general_income_threshold_note'),
  ), # 배당내역 탭
  tabPanel('절세 납입관리',
    h3('IRP · 연금저축 · ISA 납입 진도'),
    tags$div(class='alert alert-info', style='line-height:1.75;',
      tags$h4(style='margin-top:0;', '연금 세제 핵심 안내 — 납입한도와 세액공제 한도는 다릅니다'),
      tags$p(tags$b('① 연간 납입한도: '), '연금저축·IRP 등 개인 추가납입 합산 최대 1,800만 원(일반적인 기준). 이는 세액공제 한도와 다릅니다.'),
      tags$p(tags$b('② 세액공제 한도: '), '연금저축 단독 최대 600만 원, IRP를 포함한 연금계좌 합산 최대 900만 원입니다. 각 계좌에 900만 원씩 넣어도 1,800만 원 전부 세액공제되는 것은 아닙니다.'),
      tags$p(tags$b('③ 개인 납입관리 목표: '), 'IRP 900만 원 + 연금저축 900만 원 + ISA 2,000만 원을 각각 추적합니다. 개인 목표 금액은 세법상 독립된 공제한도가 아닙니다.'),
      tags$p(tags$b('④ ISA: '), '기본 연간 납입한도 2,000만 원이며 미사용 납입한도는 이월될 수 있습니다. ISA 납입액은 IRP·연금저축 세액공제 900만 원에 포함되지 않습니다.'),
      tags$small('참고: 세액공제는 실제 납부세액, 소득, 납입 유형 등에 따라 달라질 수 있습니다. 퇴직금·회사부담금 등은 개인 추가납입과 구분해야 합니다. 법령 개정 및 개별 계좌 조건은 금융기관·국세청에서 확인하세요.')
    ),
    tags$p('보유종목 평가액과 별개로 실제 외부 납입액을 등록합니다. 입력하지 않은 납입 내역은 계산할 수 없습니다.'),
    selectInput('contrib_year','조회 연도',choices=as.character(seq(as.integer(format(Sys.Date(),'%Y'))-4L,as.integer(format(Sys.Date(),'%Y')))),selected=format(Sys.Date(),'%Y')),
    uiOutput('contrib_bars'),
    tags$hr(),h4('실제 납입내역 등록'),
    fluidRow(
      column(3,dateInput('contrib_date','납입일',value=Sys.Date())),
      column(3,selectInput('contrib_type','계좌종류',choices=c('IRP'='IRP','연금저축'='PENSION','ISA'='ISA'))),
      column(3,numericInput('contrib_amount','납입금액(원)',value=1000000,min=1,step=100000)),
      column(3,textInput('contrib_broker','증권사(선택)',''))),
    fluidRow(column(4,textInput('contrib_account','계좌번호(선택)','')),column(8,textInput('contrib_memo','메모(선택)',''))),
    actionButton('contrib_add','납입내역 저장',class='btn-primary'),
    tags$hr(),h4('ISA 미사용한도 이월 설정'),
    tags$p('ISA의 이월금액은 계좌 개설연도와 과거 납입액에 따라 달라집니다. 증권사에 표시된 해당 연도 이월 가능 금액만 직접 입력하세요.'),
    numericInput('isa_carry','선택 연도 ISA 이월 한도(원)',value=0,min=0,step=1000000),
    actionButton('isa_carry_save','이월 한도 저장'),
    textOutput('contrib_notice'),
    h4('올해 기록된 납입 내역'),
    DTOutput('contrib_table'),
    tags$p('잘못된 항목은 표의 ID를 아래 입력란에 넣어 삭제할 수 있습니다.'),
    numericInput('contrib_delete_id','삭제할 납입 ID',value=NA,min=1,step=1),
    actionButton('contrib_delete','선택 ID 삭제',class='btn-warning')
  ),
  tabPanel('지금FIRE를한다면',
    h3('지금 FIRE를 한다면 · 금융자산 인출 시나리오'),
    tags$div(class='alert alert-warning',
      tags$b('인출액 ≠ 과세소득. '),
      '보유자산 매도대금에는 원금 회수가 포함됩니다. 이자·배당, 해외주식 실현차익과 실제 외부소득을 구분한 단순 추정입니다. ',
      '퇴직연금·국민연금 과세, 배우자 소득, 배당세액공제, 해외 납부세액 및 건강보험 재산·세대 정보를 완전하게 반영하지 않습니다.'),
    tags$p('현재 금융자산 평가액에 3%부터 8%까지 0.5%p 단위로 적용합니다. 인출률은 보장 수익률이나 평생 유지 가능한 비율이 아닙니다.'),
    fluidRow(
      column(4,numericInput('fire_finance_income','연간 과세 이자·배당소득(원)',value=19000000,min=0,step=1000000)),
      column(4,numericInput('fire_foreign_gain_pct','인출액 중 해외주식 실현차익 비율(%)',value=0,min=0,max=100,step=1)),
      column(4,numericInput('fire_other_tax_base','다른 종합소득 과세표준(원)',value=0,min=0,step=1000000))),
    fluidRow(
      column(4,checkboxInput('fire_regional_health','퇴직 후 지역가입자 건강보험 적용',value=TRUE)),
      column(4,numericInput('fire_property_health','재산점수분 월 건강보험료(원)',value=0,min=0,step=10000))),
    tags$p(class='text-muted',style='font-size:12px;',
      '2026년 가정: 이자·배당 합산 2,000만 원 초과 시 간이 비교과세, 해외주식 연 250만 원 양도차익 공제 후 지방세 포함 22%, ',
      '지역가입자 소득분 7.19%, 장기요양보험 추가. 입력하지 않은 재산분 보험료는 0원으로 처리되어 실제 보험료보다 작을 수 있습니다.'),
    h4('나의 FIRE 지속가능성 계산기 · 예상 생활비와 자산 고갈 시점'),
    tags$p(class='text-muted','금융자산만으로 은퇴한다고 가정한 단순 장기 시뮬레이션입니다. 생활비는 물가만큼 증가하고, 수익률은 시나리오별로 매년 일정하다고 가정합니다. 시장 예측이나 인출 보장이 아닙니다.'),
    fluidRow(
      column(4,numericInput('fire_plan_age','은퇴 시작 나이(세)',value=53,min=18,max=100,step=1)),
      column(4,numericInput('fire_plan_end_age','계획 종료 나이(세)',value=95,min=40,max=110,step=1)),
      column(4,numericInput('fire_plan_spend','은퇴 첫해 월 생활비(원)',value=4000000,min=0,step=100000)),
      column(4,numericInput('fire_plan_inflation','연간 생활비 물가상승률(%)',value=2.5,min=0,max=30,step=.5)),
      column(4,numericInput('fire_plan_pension_age','공적연금 개시 나이(세)',value=65,min=18,max=110,step=1)),
      column(4,numericInput('fire_plan_pension','공적연금 예상 월 수령액(원)',value=1900000,min=0,step=100000)),
      column(4,numericInput('fire_plan_fee','연간 투자비용·보수(%)',value=.3,min=0,max=10,step=.1)),
      column(4,numericInput('fire_plan_crash','은퇴 첫해 급락 시나리오(%)',value=-25,min=-90,max=-1,step=1))),
    tags$p(class='text-muted',style='font-size:12px;',
      '초깃값 53세·월 400만원·연금 월 190만원 등은 예시입니다. 연금은 해당 나이부터 명목금액 고정으로 지급한다고 가정하며 국민연금 실제 예상수령액과 다를 수 있습니다.'),
    uiOutput('fire_plan_summary'),
    h4('장기 잔고 경로 · 3% / 5% / 7% 및 은퇴 초기 급락'),
    plotOutput('fire_plan_plot',height='360px'),
    DTOutput('fire_plan_result'),
    h4('5년 단위 현금흐름과 잔고 · 연 5% 기준'),
    DTOutput('fire_plan_checkpoints'),
    tags$details(tags$summary('장기 시뮬레이션 계산 방식과 한계'),
      tags$p('해당 연도 인출 필요액 = 물가 반영 연간 생활비 − 연금 수령액 (최소 0원).'),
      tags$p('연말 잔고 = max(0, 연초 잔고 × (1 + 가정 연수익률 − 연간 비용) − 연간 필요 인출액).'),
      tags$p('초기 급락: 은퇴 첫해 -25%(입력값), 이후 매년 연 5% 수익률. 다른 시나리오는 매년 고정 수익률.'),
      tags$p('시뮬레이션은 확률적 몬테카를로가 아닙니다. 세금·건보료·연금 과세·실제 현금흐름과 수익률 변동을 반영하지 않아 잔고가 충분해 보여도 안전한 은퇴를 보장하지 않습니다.')),
    tags$hr(),
    h4('기존 인출률별 세후 생활비 비교'),
    uiOutput('fire_summary'),
    h4('향후 10년 물가 반영 · 현재 구매력 기준 월 실질가용액'),
    tags$p(class='text-muted','한국은행 소비자물가 통계에 수록된 2015·2025년 연평균 CPI를 이용한 최근 10년 기하평균 인플레이션(연복리) 적용. 미래 물가는 과거 추세와 다를 수 있습니다.'),
    uiOutput('fire_inflation_note'),
    DTOutput('fire_real_table'),
    plotOutput('fire_real_plot',height='310px'),
    h4('인출률별 월 세전·세후 금액 및 세금·건강보험료'),
    DTOutput('fire_table'),
    plotOutput('fire_plot',height='320px'),
    tags$details(tags$summary('계산 공식과 주의사항'),
      tags$p('월 세전 인출액 = 총평가액 × 인출률 ÷ 12.'),
      tags$p('월 세후 추정액 = 월 세전 인출액 − 연간 금융소득세·해외양도세 ÷ 12 − 월 건보료 − 월 장기요양보험료.'),
      tags$p('해외주식 실현차익 비율은 매도대금이 아니라 양도차익 금액 비율입니다. 실제 세금에는 취득환율·손익통산·비용 등이 중요합니다.'),
       tags$p('인출이 오래 지속 가능한지는 과거 수익률만으로 결정할 수 없습니다. 물가상승·급락·연금 개시와 자산고갈 위험을 별도 분석하세요.')),
     tags$hr(),
     tags$div(class='well',style='margin-top:15px;margin-bottom:18px;',
       h4('은퇴 후 장기 투자수익률 가정 · 물가 반영 해설'),
       tags$p(tags$b('기준 시나리오: 연 실질수익률 3.5%. '),
         'S&P500 중심의 장기 투자 포트폴리오를 가정할 때 참고할 수 있는 계획용 수치이며, 예상 수익을 보장하지 않습니다.'),
       tags$p('실질수익률은 물가상승에 따른 구매력 하락을 제외한 수익률입니다. ',
         '연 물가상승률을 2.5%로 가정하면 실질 3.5%에 해당하는 명목수익률은 ',
         '(1 + 0.035) × (1 + 0.025) − 1 = 약 6.09%입니다.'),
       tags$p(tags$b('권장 민감도 분석: '),
         '보수적 실질 2%, 기준 실질 3.5%, 낙관적 실질 5%를 비교하고, ',
         '은퇴 초기 급락 시나리오도 함께 확인하세요.'),
       tags$p(tags$b('현재 FIRE 계산기의 해석 주의: '),
         '위 장기 잔고 경로의 연 3% / 5% / 7%는 코드에 설정된 명목수익률 시나리오이며, ',
         '생활비 물가상승률은 별도로 반영합니다. 따라서 실질 3.5%를 명목수익률 입력값처럼 취급하면 ',
         '물가가 중복 차감될 수 있습니다. 투자비용·보수도 별도로 계산됩니다.'),
       tags$p(class='text-muted',style='font-size:12px;margin-bottom:0;',
         '참고: 실질수익률 = (1 + 명목수익률) / (1 + 물가상승률) − 1. ',
         '단순 고정수익률 모형은 실제 시장 변동, 인출 순서 위험, 환율과 세금 변동을 충분히 반영하지 못합니다.')
     )
  ),
  tabPanel('벤치마크 · 1년 백테스트',
    h3('최근 1년 벤치마크 / 보유종목 가상 백테스트'),
    tags$p('비교 대상: KODEX 미국S&P500(379800), KODEX 미국나스닥100(379810), KODEX 200(069500) 및 현재 보유종목입니다.'),
    tags$div(class='alert alert-warning',
      strong('해석 안내: '),
      '과거 일별 종가 기준 가격수익률 비교입니다. 배당·분배금 재투자, 거래비용, 세금과 실제 매매 시점은 반영하지 않습니다. ',
      '포트폴리오는 현재 보유수량·구성을 기준으로 계산한 가상 백테스트이지 실제 운용성과가 아닙니다. ',
      '미국 종목은 저장된 과거 일별 USD/KRW로 원화 환산합니다.'),
    actionButton('bench_collect','저장 데이터 활용 · 부족한 기간만 수집',class='btn-primary'),
    actionButton('bench_reset_request','전체 초기화 후 처음부터 다시 수집',class='btn-danger'),
    tags$div(style='margin-top:12px;margin-bottom:10px;',
      tags$div(class='progress',style='height:25px;',
        tags$div(id='pms_bench_progress_bar',class='progress-bar progress-bar-striped active',role='progressbar',style='width:0%;min-width:0%;', '0%'))),
    tags$div(id='pms_bench_progress_text','대기 중 · 기존 데이터는 SQLite에 유지됩니다.'),
    tags$script(HTML("Shiny.addCustomMessageHandler('pms_bench_progress',function(x){var bar=document.getElementById('pms_bench_progress_bar'),txt=document.getElementById('pms_bench_progress_text');if(bar){bar.style.width=x.percent+'%';bar.textContent=x.percent+'%';}if(txt){txt.textContent=x.message;}});")),
    actionButton('bench_recalc','저장된 과거 주가로 재분석',class='btn-success'),
    tags$p('과거 환율이 없다면 기존 [최근 1년 배당정보 수집] 기능으로 USD/KRW 일별 환율을 먼저 저장하세요.'),
    textOutput('bench_notice'),
    h4('벤치마크 상대수익률 추이 (시작일=100)'),
    plotOutput('bench_plot',height='440px'),
    h4('고점 대비 낙폭 (Drawdown, %)'),
    tags$p('각 날짜까지의 최고 지수 대비 하락률입니다. 0%는 이전 최고점 수준이며, 음수가 클수록 하락 폭이 큽니다.'),
    plotOutput('bench_drawdown',height='310px'),
    h4('60거래일 이동 Sharpe Ratio'),
    tags$p('최근 60거래일의 일별 수익률로 매일 계산합니다. 무위험수익률 연 0%, 252거래일 연환산. 시작 약 60거래일은 표시되지 않습니다.'),
    plotOutput('bench_rolling_sharpe',height='310px'),
    h4('내 포트폴리오 구성 비중 (현재 보유수량 기준)'),
    tags$p('현재 보유수량 × 비교 종료일의 종가(미국 종목은 원화 환산)로 비중을 계산합니다. 이력 부족으로 제외한 종목은 비중 계산에도 포함하지 않습니다.'),
    DTOutput('bench_weights'),
    h4('1년 이력 부족으로 포트폴리오 백테스트에서 제외된 종목'),
    DTOutput('bench_excluded'),
    h4('일별 주가 수집 현황'),
    DTOutput('bench_fetch_status'),
    tags$p('상장 이력 등으로 1년치 가격이 충분하지 않은 보유종목은 백테스트에서 제외하고 위 표에 사유를 남깁니다. 3개 ETF는 필수 비교기준이므로 ETF 시세가 부족하면 임의의 그래프를 그리지 않습니다.'),
    tags$hr(),
    h4('최근 1년 성과 비교 · 내 포트폴리오 vs 벤치마크 ETF 3종'),
    tags$p('동일 비교기간·원화 기준 일별 종가로 계산합니다. 수익률은 공통기간 누적 가격수익률, MDD는 최대낙폭입니다. Sharpe는 일별 단순수익률, 무위험수익률 0%, 연 252거래일 기준입니다. Sortino는 하방편차 기준, Calmar는 연환산 복리수익률(CAGR) / MDD 절댓값입니다. 배당·세금·수수료는 제외합니다.'),
    DTOutput('bench_stats'),
    tags$div(style='margin:12px 0;padding:12px;background:#f5f8fb;border-left:4px solid #4b86a8;',
      tags$b('판단지표 읽는 방법'),
      tags$p(tags$b('최근 1년 수익률: '),'시작일 대비 종료일의 가격 상승률입니다. 가격수익률은 배당·분배금을 제외합니다.'),
      tags$p(tags$b('MDD: '),'기간 중 최고점에서 최저점까지 가장 크게 떨어진 비율입니다. -10%가 -30%보다 방어력이 좋습니다.'),
      tags$p(tags$b('Sharpe: '),'가격 변동 전체(상승·하락)를 위험으로 보고 1단위 위험당 초과수익을 비교합니다. 높을수록 좋습니다.'),
      tags$p(tags$b('Sortino: '),'하락 변동성만 위험으로 계산합니다. 높을수록 하락 위험 대비 수익성이 좋습니다.'),
      tags$p(tags$b('Calmar: '),'연환산 수익률을 MDD의 절댓값으로 나눈 수치입니다. 높을수록 큰 낙폭 대비 보상이 좋습니다.'),
      tags$p('위 지표는 같은 비교 기간과 동일한 무위험수익률(연 0%) 가정에서 비교해야 합니다.')),
    tags$hr(),
    h3('Total Return · 세전 배당 재투자 기준 백테스트'),
    tags$p('가격수익률뿐 아니라 배당·ETF 분배금을 배당락일 종가에 해당 종목으로 전액 재투자한 가상 총수익률입니다. 세금·거래비용·실제 과거 보유수량은 반영되지 않습니다.'),
    actionButton('bench_tr_collect','Total Return용 배당이력 보충 조회',class='btn-primary'),
    textOutput('bench_tr_collect_notice'),
    textOutput('bench_tr_notice'),
    h4('Total Return 상대수익률 추이 (시작일=100)'),
    plotOutput('bench_tr_plot',height='440px'),
    h4('Total Return 고점 대비 낙폭 (Drawdown, %)'),
    tags$p('배당·분배금을 재투자한 총수익 지수가 이전 최고점에서 얼마나 하락했는지를 나타냅니다. 0%는 최고점이며, 음수가 클수록 하락 폭이 큽니다.'),
    plotOutput('bench_tr_drawdown',height='310px'),
    h4('Total Return 60거래일 이동 Sharpe Ratio'),
    tags$p('배당 재투자 기준 일별 총수익률을 이용해 최근 60거래일의 위험 대비 성과를 매일 계산합니다. 무위험수익률 연 0%, 연 252거래일 가정입니다.'),
    plotOutput('bench_tr_rolling_sharpe',height='310px'),
    h4('Total Return 최근 1년 성과 비교 · 내 포트폴리오 vs 벤치마크 ETF 3종'),
    tags$p('네 대상 모두 같은 비교기간·원화 기준으로 비교합니다. 세전 배당·분배금 재투자, 거래비용·세금·실제 거래내역 미반영. Sharpe 내림차순.'),
    DTOutput('bench_tr_stats'),
    tags$div(style='margin:12px 0;padding:12px;background:#f5f8fb;border-left:4px solid #4b86a8;',
      tags$b('Total Return 지표 읽는 방법'),
      tags$p(tags$b('최근 1년 총수익률: '),'주가 변동과 배당·분배금 재투자를 합친 수익률입니다. 높을수록 기간 성과가 좋습니다.'),
      tags$p(tags$b('MDD: '),'총수익 지수가 과거 최고점에서 가장 깊게 하락한 비율입니다. -10%가 -30%보다 방어력이 좋습니다.'),
      tags$p(tags$b('Sharpe: '),'상승·하락을 모두 포함한 전체 변동성 대비 수익성입니다. 높을수록 동일 위험에서 성과가 좋았습니다.'),
      tags$p(tags$b('Sortino: '),'하락 위험만 고려한 수익성입니다. 높을수록 하락 위험 대비 성과가 좋습니다.'),
      tags$p(tags$b('Calmar: '),'연환산 복리 총수익률을 최대낙폭 절댓값으로 나눈 지표입니다. 높을수록 깊은 손실에 대한 보상이 좋습니다.'),
      tags$p('두 성과표는 동일한 수익률 산식·공통 거래일·무위험수익률 0% 가정을 사용합니다. 배당 이력이 누락됐거나 종목이 제외된 경우 실제 운용성과로 해석하지 마세요.')),
    h4('종목별 배당이력 확보 상태'),
    DTOutput('bench_tr_coverage'),
    tags$hr(),
    tags$details(open='open',tags$summary(tags$b('투자성과 해석 가이드: 수익률과 최대하락폭(MDD)의 관계')),
      tags$p('수익률(CAGR)은 자산의 성장 정도를 나타내지만 중간에 겪은 손실과 회복기간은 충분히 설명하지 못합니다.'),
      tags$p('MDD는 관측기간의 최고점에서 최저점까지 최대 손실폭입니다. -10%와 -30%는 과거 낙폭의 차이이지 미래 손실 상한선이 아닙니다.'),
      tags$p('높은 수익률을 추구하면 하락 위험이 커질 수 있지만, 분산투자로 수익률을 상당 부분 유지하면서 MDD를 낮출 수도 있습니다.'),
      tags$p('어떤 한 지표도 절대적으로 신뢰할 수 없습니다. Sharpe·Sortino·Calmar 역시 기간과 계산 가정에 따라 평가가 달라집니다.'),
      tags$p('실제 입출금·배당재투자·환율·세금·수수료·분석기간·생존편향 여부를 함께 살펴보세요. PMS2의 과거 평가액 기반 수익률에는 추정 현금흐름이 포함되어 실제 TWR과 다를 수 있습니다.'),
      tags$b('결론: 수익률이 가장 높은 전략보다 자신이 감내할 수 있는 낙폭 안에서 오랫동안 지속할 수 있는 전략을 선택하는 것이 중요합니다.'))
  ), # 벤치마크 탭
  tabPanel('분산투자',
    h3('분산투자 · 7개 투자전략 그룹 분석'),
    tags$p(class='text-muted','총자산 대비 현재 비중과 목표 비중을 비교합니다. 종목별 전략분류는 사용자가 확인한 값을 사용하며 자동 매매는 하지 않습니다.'),
    uiOutput('diversification_summary'),
    h4('현재 비중과 목표 비중'),
    plotOutput('diversification_chart',height='315px'),
    DTOutput('diversification_table'),
    tags$hr(),
    h4('신규 투자금 배분 제안'),
    numericInput('diversification_new_cash','신규 투자금 (원)',value=3000000,min=0,step=100000),
    uiOutput('diversification_buy_advice'),
    tags$hr(),
    h4('종목별 투자그룹 관리'),
    tags$p('아래에서 종목별 투자그룹을 확인·변경하고 저장하세요. 기타/미분류 종목은 매수 권고 대상에서 제외합니다.'),
    DTOutput('diversification_holdings'),
    selectInput('diversification_holding_id','수정할 보유종목 ID',choices=NULL),
    selectInput('diversification_group','목표 투자그룹',choices=pms_strategy_choices()),
    actionButton('diversification_save_group','종목 투자그룹 저장',class='btn-primary'),
    textOutput('diversification_status')
  ),
  tabPanel('인공지능의견',
    tags$div(class='pms-ai-page',
      tags$h3('인공지능의견 · 로컬 Ollama'),
      tags$div(class='alert alert-info',style='margin:10px 0 15px;',
        tags$strong('필수 설치 안내: '),
        '이 기능을 사용하려면 PMS2를 실행하는 서버 PC에 Ollama가 별도로 설치되어 실행 중이어야 합니다. ',
        '모델 qwen3.5:4b도 미리 다운로드해야 합니다. ',
        tags$code('ollama pull qwen3.5:4b'),
        ' 실행 후 아래의 [연결 확인] 버튼으로 상태를 점검하세요. ',
        'Ollama가 없더라도 PMS2의 다른 탭은 사용할 수 있습니다.'),
      tags$p(class='text-muted','서버 PC에서 실행하는 Ollama(qwen3.5:4b)가 저장된 종목·시세 데이터를 해석합니다. 버튼을 눌렀을 때만 분석하며 새로운 인터넷 시황이나 뉴스는 자동 조회하지 않습니다.'),
      tags$div(class='pms-ai-card',
        tags$div(style='display:flex;gap:10px;align-items:center;flex-wrap:wrap;',
          tags$strong('AI 연결:'), textOutput('pms_ai_connection',inline=TRUE),
          actionButton('pms_ai_check','연결 확인',class='btn-default')),
        tags$p(style='margin:9px 0 0;color:#607184;', '로컬 API: http://127.0.0.1:11434 · 모델: qwen3.5:4b')),
      tags$div(class='pms-ai-card',
        tags$h4('분석에 사용할 저장 데이터'),
        verbatimTextOutput('pms_ai_data_info'),
        tags$p(class='text-muted','당일 시황: 저장된 종목별 최근 시세와 기준시각만 자동 참조합니다. 시황 뉴스가 필요하면 아래에 직접 적어주세요.'),
        textAreaInput('pms_ai_market_note','오늘의 시장 상황 메모 (선택)',value='',rows=3,
          placeholder='예: 코스피 상승, 나스닥 하락 등 직접 확인한 사실과 출처/시각을 입력'),
        textAreaInput('pms_ai_question','추가로 묻고 싶은 내용 (선택)',
          value='현재 보유종목의 집중위험, 자산배분, 리밸런싱 시 고려사항을 분석해 주세요.',rows=3),
        sliderInput('pms_ai_length','AI 분석 분량',min=1,max=3,value=2,step=1,
          ticks=FALSE,sep='',width='360px'),
        tags$div(style='font-size:12px;color:#607184;margin:-7px 0 10px;',
          '1 · 짧게 (빠름)　　2 · 중간 (기본)　　3 · 길게 (상세)'),
        textOutput('pms_ai_length_info'),
        actionButton('pms_ai_analyze','현재 저장 데이터로 AI 의견 생성',class='btn-primary'),
        tags$div(id='pms_ai_click_debug',role='status',style='margin-top:10px;font-weight:600;color:#245b92;',
                 'AI 버튼을 누르면 이곳에 클릭 감지 여부가 표시됩니다.'),
        tags$script(HTML("
          (function(){
            if(window.pmsAiDiagnosticInstalled)return;
            window.pmsAiDiagnosticInstalled=true;
            document.addEventListener('click',function(e){
              var b=e.target.closest && e.target.closest('#pms_ai_analyze');
              if(!b)return;
              var d=document.getElementById('pms_ai_click_debug');
              if(d)d.textContent='① 클릭 감지: Shiny 연결 상태 확인 중...';
              var send=function(){
                var connected=!!(window.Shiny && Shiny.shinyapp && Shiny.shinyapp.$socket && Shiny.shinyapp.$socket.readyState===1);
                if(d)d.textContent=connected?'① 서버 연결 확인 · 클릭 이벤트 전송 중...':'① 서버 연결 상태 불명 · 이벤트 전송 시도 중...';
                if(window.Shiny && typeof Shiny.setInputValue==='function'){
                  Shiny.setInputValue('pms_ai_explicit_click',Date.now()+'-'+Math.random(),{priority:'event'});
                  window.setTimeout(function(){
                    if(d && d.textContent.indexOf('①')===0)d.textContent='① 이벤트 전송 후 서버 응답 지연(5초). RStudio Console 확인 · 서버가 다른 계산 중일 수 있음.';
                  },5000);
                }else if(d)d.textContent='① Shiny JavaScript 바인딩을 찾을 수 없습니다. 브라우저를 새로고침하세요.';
              };
              send();
            },true);
            if(window.Shiny && Shiny.addCustomMessageHandler){
              Shiny.addCustomMessageHandler('pms_ai_debug',function(m){
                var d=document.getElementById('pms_ai_click_debug');
                if(d)d.textContent=m.text || '';
              });
            }else{
              document.addEventListener('shiny:connected',function(){
                Shiny.addCustomMessageHandler('pms_ai_debug',function(m){
                  var d=document.getElementById('pms_ai_click_debug');
                  if(d)d.textContent=m.text || '';
                });
              },{once:true});
            }
          })();
        "))) ,
      tags$div(class='pms-ai-card',id='pms_ai_result_card',
        tags$h4('AI 투자분석 의견'),
        textOutput('pms_ai_status'),
        tags$div(class='pms-ai-markdown',style='overflow-wrap:anywhere;line-height:1.7;',uiOutput('pms_ai_answer')),
        tags$p(class='text-muted',style='margin-top:15px;', 'AI 의견은 참고자료이며 투자수익을 보장하지 않습니다. 수치와 세금은 반드시 PMS2 계산 및 현행 규정으로 확인하세요.'))
    )
  ),
  tabPanel('보고서',
    h3('PMS2 종합 투자 보고서 · Quarto'),
    tags$p('현재 SQLite의 저장값을 읽어 자산규모, 자산 배분, 종목 집중도, 배당, 세금 참고지표, 과거 위험지표와 데이터 품질을 자세히 설명합니다. 외부 API는 호출하지 않습니다.'),
    tags$div(class='alert alert-info',
      'Quarto 문서(.qmd)는 Quarto가 설치되지 않아도 생성·다운로드할 수 있습니다. HTML 출력은 PC에 Quarto CLI가 설치된 경우에만 가능합니다.'),
    checkboxInput('report_benchmark','저장된 벤치마크 캐시를 보고서에 포함',TRUE),
    tags$style(HTML('
      #pms_report_busy {display:none;align-items:center;gap:12px;margin:12px 0;padding:14px 16px;border:1px solid #c5dcee;background:#f4f9fe;border-radius:8px;color:#244a68;}
      #pms_report_busy.pms-report-running {display:flex;}
      .pms-report-spinner {width:28px;height:28px;flex-shrink:0;border:4px solid #c9dbe9;border-top-color:#2476b9;border-radius:50%;animation:pmsReportSpin .85s linear infinite;}
      @keyframes pmsReportSpin {to {transform:rotate(360deg);}}
      @media (prefers-reduced-motion:reduce) {.pms-report-spinner {animation-duration:2.5s;}}
    ')),
    tags$script(HTML('
      (function(){
        function toggleReportBusy(running) {
          var box=document.getElementById("pms_report_busy");
          var btn=document.getElementById("report_generate");
          if(box){box.classList.toggle("pms-report-running",!!running);box.setAttribute("aria-hidden",running?"false":"true");}
          if(btn){btn.disabled=!!running;btn.setAttribute("aria-busy",running?"true":"false");}
        }
        $(document).on("click", "#report_generate", function(){toggleReportBusy(true);});
        function registerReportSpinner(){
          if(window.Shiny && Shiny.addCustomMessageHandler){
            Shiny.addCustomMessageHandler("pms_report_busy",function(msg){toggleReportBusy(!!msg.running);});
          }
        }
        $(document).on("shiny:connected",registerReportSpinner);
        if(window.Shiny){registerReportSpinner();}
        $(document).on("shiny:disconnected",function(){toggleReportBusy(false);});
      })();
    ')),
    actionButton('report_generate','현재 데이터로 종합 보고서 생성',class='btn-primary'),
    tags$div(id='pms_report_busy',role='status',`aria-live`='polite',`aria-hidden`='true',
      tags$span(class='pms-report-spinner',`aria-hidden`='true'),
      tags$div(tags$strong('보고서를 생성하고 있습니다.'),
        tags$div('SQLite 집계 → Quarto HTML → PDF 변환 중입니다. 완료될 때까지 기다려 주세요.'))),
    tags$div(style='margin:12px 0;',textOutput('report_status')),
    downloadButton('report_download_qmd','Quarto 원본(.qmd) 다운로드'),
    downloadButton('report_download_html','HTML 보고서 다운로드'),
    downloadButton('report_download_pdf','PDF 보고서 다운로드'),
    tags$hr(),
    h4('3분 금융자산 음성 브리핑'),
    tags$p(class='text-muted','현재 저장된 자료로 작성한 한국어 해설을 크롬이 읽습니다. 실제 재생시간은 글자 수와 음성 속도에 따라 달라지며 약 2~4분입니다. 버튼을 누를 때의 브리핑 원고를 재생합니다.'),
    tags$div(style='display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:12px;',
      tags$button(id='voice_play',type='button',class='btn btn-primary','3분 브리핑 음성듣기'),
      tags$button(id='voice_pause',type='button',class='btn btn-default','일시정지 / 계속'),
      tags$button(id='voice_stop',type='button',class='btn btn-default','중지'),
      tags$label(`for`='voice_rate',style='margin:0;','속도'),
      tags$select(id='voice_rate',class='form-control',style='width:100px;',
        tags$option(value='0.9','0.9배'),tags$option(value='1.0',selected='selected','1.0배'),tags$option(value='1.2','1.2배'))),
    tags$p(id='voice_status',role='status',`aria-live`='polite','음성 재생 준비됨. 브라우저에서 한국어 음성 지원이 필요합니다.'),
    tags$details(tags$summary('읽을 브리핑 원고 보기'),
      tags$pre(id='voice_brief_copy',style='white-space:pre-wrap;max-height:320px;overflow:auto;',textOutput('voice_brief_text'))),
    tags$script(HTML('
      (function(){
        var speech=window.speechSynthesis, segments=[], index=0, playing=false, cancelled=false;
        function status(msg){var el=document.getElementById("voice_status");if(el)el.textContent=msg;}
        function stop(){cancelled=true;playing=false;segments=[];index=0;if(speech)speech.cancel();status("음성 재생이 중지되었습니다.");}
        function next(){
          if(cancelled || !playing || index>=segments.length){
            if(!cancelled){playing=false;status("음성 브리핑이 완료되었습니다.");}
            return;
          }
          var message=segments[index++];
          var u=new SpeechSynthesisUtterance(message);
          u.lang="ko-KR";
          u.rate=Number(document.getElementById("voice_rate").value)||1;
          var korean=speech.getVoices().filter(function(v){return /^ko[-_]/i.test(v.lang);});
          if(korean.length)u.voice=korean[0];
          u.onend=function(){if(playing&&!cancelled)next();};
          u.onerror=function(e){if(e.error==="interrupted"||e.error==="canceled")return;
            playing=false;status("음성 오류: "+e.error+". 크롬 음성 설정을 확인하세요.");};
          speech.speak(u);
          status("브리핑 재생 중 ("+index+" / "+segments.length+")");
        }
        $(document).on("click","#voice_play",function(){
          if(!speech){status("이 브라우저는 Web Speech API를 지원하지 않습니다.");return;}
          // Read the server-validated plain text from Shiny, never arbitrary HTML.
          var el=document.getElementById("voice_brief_copy");
          var txt=el?el.textContent.trim():"";
          if(!txt){status("브리핑 문장이 아직 준비되지 않았습니다. 잠시 후 다시 눌러주세요.");return;}
          stop();cancelled=false;playing=true;
          var sentences=txt.match(/[^.!?]+[.!?]?/g)||[txt];
          var block="";
          sentences.forEach(function(sentence){
            if((block+sentence).length>170 && block){segments.push(block.trim());block="";}
            block+=sentence+" ";
          });
          if(block.trim())segments.push(block.trim());
          next();
        });
        $(document).on("click","#voice_pause",function(){
          if(!speech||!playing)return;
          if(speech.paused){speech.resume();status("브리핑 다시 재생 중입니다.");}
          else{speech.pause();status("음성 브리핑 일시정지 중입니다.");}
        });
        $(document).on("click","#voice_stop",stop);
        $(window).on("beforeunload",function(){if(speech)speech.cancel();});
        if(speech)speech.getVoices();
      })();
    ')),
    uiOutput('report_summary_ui'),
    tags$h4('보고서 본문 미리보기'),
    uiOutput('report_inline_view'),
    tags$hr(),
    tags$h4('텔레그램 일일 자산 보고'),
    tags$p(class='text-muted','전광판과 동일한 보유자산 총평가액 및 평가증감(당일 환율 고정 기준)을 발송합니다. 시세 갱신시각을 확인하세요.'),
    checkboxInput('pms_tg_auto_16','16:00 자동보고 사용 (기본 ON)',value=TRUE),
    actionButton('pms_tg_send_test','텔레그램 테스트 메시지 보내기',class='btn-default'),
    actionButton('pms_tg_send_manual','텔레그램 수동 전송',class='btn-primary'),
    tags$div(style='margin-top:7px;',textOutput('pms_tg_missing_quote_status')),
    tags$div(style='margin-top:10px;',textOutput('pms_tg_status')),
    tags$p(class='text-muted','보고서 생성은 현재 저장된 값의 일회성 스냅샷이며, 새 시세·배당 조회나 자동 재생성은 하지 않습니다. HTML 파일은 독립된 웹페이지로 열 수 있습니다.')
  )
  ) # tabsetPanel
  ) # pms_zoom_content
)

server <- function(input, output, session) {
  # 전역 reactiveVal은 하나의 R 프로세스에 연결된 세션들이 공유합니다.
  # 종료 콜백은 해당 세션이 연결 해제될 때 한 번만 실행됩니다.
  pms_active_sessions(shiny::isolate(pms_active_sessions()) + 1L)
  session$onSessionEnded(function() {
    pms_active_sessions(max(0L, shiny::isolate(pms_active_sessions()) - 1L))
  })
  output$pms_session_count <- shiny::renderText({
    paste0(pms_active_sessions(), '세션')
  })

  # AI 분석: 버튼 입력 시에만 SQLite를 읽고 로컬 Ollama에 전달합니다.
  # 인터넷 뉴스/당일 지수 변동은 임의로 생성하지 않으며, 저장 기준시각을 포함합니다.
  pms_ai_status <- reactiveVal('분석 실행 버튼을 누르면 로컬 AI 의견을 표시합니다.')
  pms_ai_answer <- reactiveVal('아직 분석 결과가 없습니다.')
  pms_ai_connection <- reactiveVal('연결 확인 전')
  output$pms_ai_status <- renderText(pms_ai_status())
  # Ollama 답변은 Markdown 문법을 사용하므로 일반 textOutput 대신 HTML로 표시합니다.
  # HTML 태그는 먼저 이스케이프하여 AI 출력 내 임의 HTML 실행을 방지합니다.
  output$pms_ai_answer <- renderUI({
    txt <- pms_ai_answer()
    safe <- htmltools::htmlEscape(txt)
    if (requireNamespace('commonmark', quietly=TRUE)) {
      html <- commonmark::markdown_html(safe, hardbreaks=TRUE)
    } else if (requireNamespace('markdown', quietly=TRUE)) {
      html <- markdown::markdownToHTML(text=safe, fragment.only=TRUE)
    } else {
      # 별도 Markdown 패키지가 없다면 **굵은 글씨**만이라도 표시합니다.
      escaped <- gsub('\\*\\*([^*\\n]+)\\*\\*', '<strong>\\1</strong>', safe, perl=TRUE)
      html <- gsub('\n','<br>',escaped,fixed=TRUE)
    }
    shiny::HTML(html)
  })
  output$pms_ai_connection <- renderText(pms_ai_connection())

  pms_ai_length_settings <- function(value) {
    level <- suppressWarnings(as.integer(value))
    if (length(level)!=1L || is.na(level) || !(level %in% 1:3)) level <- 2L
    switch(as.character(level),
      '1'=list(label='짧게',top=12L,tokens=450L,ctx=3072L,input_chars=750L,
               instruction='핵심 위험 2가지와 점검사항 2가지를 매우 간결하게 정리하세요.'),
      '2'=list(label='중간',top=25L,tokens=1000L,ctx=4096L,input_chars=1500L,
               instruction='핵심 위험 3가지, 자산배분 및 리밸런싱 점검사항 3가지를 설명하세요.'),
      '3'=list(label='길게',top=40L,tokens=1800L,ctx=6144L,input_chars=2000L,
               instruction='포트폴리오 집중 위험, 자산군 분산, 환율·세금, 리밸런싱 유의점과 실행 전 확인할 사항을 근거 중심으로 자세히 설명하세요.')
    )
  }
  output$pms_ai_length_info <- renderText({
    cfg <- pms_ai_length_settings(input$pms_ai_length)
    sprintf('선택: %s · 상위 %d종목 · 최대 %d토큰 (길수록 분석시간 증가)',
            cfg$label,cfg$top,cfg$tokens)
  })
  pms_ai_context <- function(top_n=12L) {
    d <- pms_valuation(con,include_zero=FALSE)
    if(!nrow(d)) stop('보유종목 데이터가 없습니다.')
    d <- d[!is.na(d$quantity) & d$quantity > 0,,drop=FALSE]
    if(!nrow(d)) stop('수량이 양수인 보유종목이 없습니다.')
    numeric_value <- function(v) {v <- suppressWarnings(as.numeric(v));v[!is.finite(v)] <- NA_real_;v}
    val <- numeric_value(d$market_value_krw)
    cost <- if('cost_basis_krw' %in% names(d)) numeric_value(d$cost_basis_krw) else rep(NA_real_,nrow(d))
    profit <- if('profit_krw' %in% names(d)) numeric_value(d$profit_krw) else rep(NA_real_,nrow(d))
    n_valid <- sum(is.finite(val))
    nav <- if(n_valid) sum(val,na.rm=TRUE) else NA_real_
    csum <- if(any(is.finite(cost)))sum(cost,na.rm=TRUE) else NA_real_
    psum <- if(any(is.finite(profit)))sum(profit,na.rm=TRUE) else NA_real_
    fmt <- function(x) if(length(x)!=1 || !is.finite(x)) '자료 없음' else format(round(x),big.mark=',',scientific=FALSE,trim=TRUE)
    price_stamp <- tryCatch(DBI::dbGetQuery(con,
      'SELECT MAX(fetched_at) AS v FROM prices')$v[1],error=function(e) NA_character_)
    fx <- tryCatch(DBI::dbGetQuery(con,
      "SELECT pair, rate, quoted_at, fetched_at FROM fx_rates WHERE pair='USD/KRW' LIMIT 1"),error=function(e) data.frame())
    meta <- c(paste0('분석 시각(한국시간): ',format(Sys.time(),'%Y-%m-%d %H:%M:%S',tz='Asia/Seoul')),
      paste0('최근 시세 수집시각(원본 DB 기록): ',if(is.na(price_stamp))'기록 없음' else price_stamp),
      paste0('활성 보유내역 수: ',nrow(d),' / 원화평가액 산출 가능: ',n_valid),
      paste0('원화 평가총액(산출 가능한 종목만): ',fmt(nav),' 원'),
      paste0('원화 취득원가(산출 가능한 종목만): ',fmt(csum),' 원'),
      paste0('평가손익(산출 가능한 종목만): ',fmt(psum),' 원'),
      paste0('평가불가 내역 수: ',sum(!is.finite(val))))
    if(nrow(fx))meta <- c(meta,paste0('USD/KRW: ',fmt(as.numeric(fx$rate[1])), ' / 기준 ',fx$quoted_at[1]))
    for(nm in c('country','asset_class','account_type')){
      if(nm %in% names(d) && n_valid){
        agg <- tapply(val[is.finite(val)],as.character(d[[nm]][is.finite(val)]),sum)
        meta <- c(meta,paste0(nm,'별 평가액: ',paste(paste(names(agg),vapply(agg,fmt,character(1)),sep='='),collapse='; ')))
      }
    }
    # 같은 종목이 여러 계좌에 있는 경우 한 종목으로 합산하여 중복 집중도 왜곡 방지
    good <- which(is.finite(val))
    lines <- character()
    if(length(good)){
      z <- aggregate(val[good],by=list(country=as.character(d$country[good]),
        symbol=as.character(d$symbol[good]),name=as.character(d$name[good])),FUN=sum)
      z <- z[order(z$x,decreasing=TRUE),,drop=FALSE]
      z <- head(z,as.integer(top_n))
      lines <- sprintf('%s %s (%s): %s원%s',z$country,z$name,z$symbol,
        vapply(z$x,fmt,character(1)),if(is.finite(nav) && nav>0) paste0(' / ',sprintf('%.2f',100*z$x/nav),'%') else '')
    }
    context <- paste(c('아래는 PMS2 SQLite와 R 계산 결과이며 뉴스/전체 시장 상황은 포함하지 않음:',
       meta,sprintf('평가액 상위 보유종목(종목별 계좌 합산, 최대 %d종목):',as.integer(top_n)),lines),collapse='\n')
    list(text=context, info=paste(meta,collapse='\n'))
  }
  output$pms_ai_data_info <- renderText({
    # 버튼 클릭 또는 탭 진입 때 다시 조회
    input$pms_tabs
    tryCatch(pms_ai_context()$info,error=function(e) paste('데이터 확인 실패:',conditionMessage(e)))
  })
  observeEvent(input$pms_ai_check,{
    pms_ai_connection('연결 확인 중...')
    tryCatch({
      res <- httr::GET('http://127.0.0.1:11434/api/tags',httr::timeout(5))
      httr::stop_for_status(res)
      models <- jsonlite::fromJSON(httr::content(res,as='text',encoding='UTF-8'))$models$name
      pms_ai_connection(if('qwen3.5:4b' %in% models)'정상 · qwen3.5:4b 사용 가능' else 'Ollama 응답 정상 · qwen3.5:4b 모델을 설치하세요')
    },error=function(e)pms_ai_connection(paste('연결 실패:',conditionMessage(e))))
  })
  # AI 호출은 별도 R 프로세스로 실행하여 Shiny 홈/기타 세션의 UI를 막지 않습니다.
  # 버튼 중복 클릭 시 결과가 서로 덮어쓰지 않도록 세션별 실행 잠금.
  pms_ai_running <- reactiveVal(FALSE)
  pms_ai_job <- reactiveVal(NULL)
  pms_ai_started <- reactiveVal(NULL)
  pms_ai_stream_path <- reactiveVal(NULL)
  observeEvent(input$pms_ai_explicit_click,{
    message('PMS2 AI DEBUG: EXPLICIT click received by R server; ',Sys.time())
    session$sendCustomMessage('pms_ai_debug',list(text='② Shiny 서버 명시적 이벤트 수신 완료. SQLite 분석 자료를 준비합니다.'))
    if (isTRUE(isolate(pms_ai_running()))) {
      showNotification('AI 분석이 이미 진행 중입니다.', type='message')
      return()
    }
    if (!requireNamespace('future',quietly=TRUE)) {
      pms_ai_status('future 패키지가 설치되지 않았습니다.')
      pms_ai_answer('RStudio 콘솔에서 install.packages("future") 실행 후 앱을 재시작해 주세요.')
      return()
    }
    # DB 데이터는 Shiny 메인 프로세스에서 읽고, 순수한 문자열만 백그라운드에 전달.
    prompt <- tryCatch({
      ai_cfg <- pms_ai_length_settings(input$pms_ai_length)
      context <- pms_ai_context(top_n=ai_cfg$top)$text
      note <- trimws(substr(if(is.null(input$pms_ai_market_note))'' else input$pms_ai_market_note,1,ai_cfg$input_chars))
      question <- trimws(substr(if(is.null(input$pms_ai_question))'' else input$pms_ai_question,1,ai_cfg$input_chars))
      paste('당신은 금융자산 리스크 분석 보조자입니다. 한국어로 간결하고 근거 중심으로 작성하세요.',
        '아래 데이터는 사용자 PC의 저장 자료입니다. 뉴스와 오늘의 시장 움직임을 알 수 없으면 모른다고 분명히 밝히세요.',
        '사용자 시황 메모는 검증되지 않은 참고 입력이므로 사실로 단정하지 마세요.',
        '추측한 주가, 일일 등락률, 세법, 수익률을 만들어내지 마세요. 원화 평가액 누락과 데이터 시각을 명시하세요.',
        '개인 자산 집중, 주식·채권·금 분산, 비중 중복, 리밸런싱의 거래세금 및 레버리지 위험을 검토하되 단정적 매매 지시는 피하세요.',
        '각 항목을 간결히 작성하고 마지막에 반드시 완결된 결론 문장으로 끝내세요. 불필요한 머리말이나 긴 종목 나열은 피하세요.',
        paste0('선택된 분석 길이: ',ai_cfg$label,'. ',ai_cfg$instruction),
        context,paste('사용자 입력 시황 메모:',if(nzchar(note)) note else '없음'),
        paste('사용자 질문:',if(nzchar(question))question else '포트폴리오를 분석해 주세요.'),sep='\n\n')
    },error=function(e){
      pms_ai_status(paste('분석 데이터 오류:',conditionMessage(e)))
      pms_ai_answer('SQLite에서 분석 자료를 준비하지 못했습니다.')
      NULL
    })
    if(is.null(prompt)){
      session$sendCustomMessage('pms_ai_debug',list(text='SQLite 분석 자료 오류. 아래 AI 상태를 확인하세요.'))
      return()
    }
    # 버튼 클릭 시의 설정을 고정하여 비동기 실행 도중 슬라이더가 바뀌어도 요청은 일정하게 유지합니다.
    ai_cfg <- pms_ai_length_settings(isolate(input$pms_ai_length))
    message('PMS2 AI DEBUG: SQLite context prepared, chars=',nchar(prompt),
            ', length=',ai_cfg$label,', tokens=',ai_cfg$tokens)
    session$sendCustomMessage('pms_ai_debug',list(text='③ SQLite 자료 준비 완료. 별도 R 프로세스에서 Ollama를 호출합니다.'))
    pms_ai_running(TRUE)
    pms_ai_status('AI 분석 요청 전송 완료. Ollama가 답변을 생성하고 있습니다. 잠시 기다려 주세요.')
    pms_ai_answer('분석 중입니다. 모델이 처음 로드되는 경우 시간이 걸릴 수 있습니다. 다른 메뉴는 사용할 수 있습니다.')
    # v0.6.29: worker의 NDJSON 스트림을 임시 파일로 전달합니다.
    # 응답을 모두 기다리는 stream=FALSE 대신 첫 토큰부터 화면에 표시합니다.
    stream_file <- tempfile(pattern='pms_ai_stream_',fileext='.txt')
    file.create(stream_file)
    pms_ai_stream_path(stream_file)
    pending <- tryCatch({
      if (inherits(future::plan(), 'sequential')) {
        future::plan(future::multisession, workers=2L)
      }
      prompt_to_send <- as.character(prompt)
      output_path <- stream_file
      max_tokens <- as.integer(ai_cfg$tokens)
      context_window <- as.integer(ai_cfg$ctx)
      future::future({
        if (!requireNamespace('curl',quietly=TRUE)) stop('curl 패키지를 설치해 주세요.')
        if (!requireNamespace('jsonlite',quietly=TRUE)) stop('jsonlite 패키지를 설치해 주세요.')
        payload <- jsonlite::toJSON(list(
          model='qwen3.5:4b',prompt=prompt_to_send,
          stream=TRUE,think=FALSE,keep_alive='10m',
          options=list(num_predict=max_tokens,num_ctx=context_window,temperature=0.3)
        ),auto_unbox=TRUE,null='null')
        h <- curl::new_handle()
        curl::handle_setopt(h,post=TRUE,postfields=as.character(payload),
          timeout=600L,connecttimeout=10L)
        curl::handle_setheaders(h,'Content-Type'='application/json')
        buff <- ''
        all_text <- ''
        err <- NULL
        done <- FALSE
        truncated <- FALSE
        sink <- file(output_path,open='ab')
        on.exit(close(sink),add=TRUE)
        process_line <- function(line){
          if(!nzchar(trimws(line)))return(invisible(NULL))
          x <- tryCatch(jsonlite::fromJSON(line,simplifyVector=FALSE),error=function(e)NULL)
          if(is.null(x))return(invisible(NULL))
          if(!is.null(x$error))stop(as.character(x$error))
          if(!is.null(x$response) && nzchar(x$response)){
            chunk <- as.character(x$response)
            writeBin(charToRaw(enc2utf8(chunk)),sink)
            flush(sink)
            all_text <<- paste0(all_text,chunk)
          }
          if(isTRUE(x$done)) {
            done <<- TRUE
            if(!is.null(x$done_reason) && identical(as.character(x$done_reason),'length')) truncated <<- TRUE
          }
          invisible(NULL)
        }
        response <- curl::curl_fetch_stream('http://127.0.0.1:11434/api/generate',
          function(bytes){
            buff <<- paste0(buff,rawToChar(bytes))
            parts <- strsplit(buff,'\n',fixed=TRUE)[[1]]
            if(endsWith(buff,'\n')){
              buff <<- ''
              lines <- parts
            }else{
              buff <<- parts[length(parts)]
              lines <- head(parts,-1L)
            }
            for(line in lines)process_line(line)
          },handle=h)
        if(nzchar(trimws(buff)))process_line(buff)
        if(response$status_code != 200L)
          stop(sprintf('Ollama HTTP %s',response$status_code))
        if(!done)stop('Ollama 스트림이 정상 완료 표시 없이 종료됐습니다.')
        if(!nzchar(trimws(all_text)))stop('AI 응답 본문이 비어 있습니다.')
        if(truncated) {
          paste0(all_text,'\n\n[안내: 생성 토큰 제한으로 답변이 중간에 종료됐습니다. AI 분석 분량을 한 단계 높여 다시 요청해 주세요.]')
        } else all_text
      })
    },error=function(e)e)
    if(inherits(pending,'error')){
      pms_ai_running(FALSE)
      pms_ai_status(paste('AI 실행 시작 오류:',conditionMessage(pending)))
      pms_ai_answer('분석을 시작하지 못했습니다. RStudio Console을 확인하세요.')
      session$sendCustomMessage('pms_ai_debug',list(text=paste('④ 실행 시작 오류:',conditionMessage(pending))))
      unlink(stream_file)
      pms_ai_stream_path(NULL)
      return()
    }
    pms_ai_job(pending)
    pms_ai_started(Sys.time())
    message('PMS2 AI DEBUG: streaming future dispatched')
    session$sendCustomMessage('pms_ai_debug',list(text='④ Ollama 스트리밍 연결 중 — 첫 응답을 기다립니다.'))
  })

  # Shiny 세션이 살아 있는 동안만 1초 주기로 결과 파일을 읽습니다.
  observe({
    if (!isTRUE(pms_ai_running())) return()
    shiny::invalidateLater(1000, session)
    job <- isolate(pms_ai_job())
    if(is.null(job))return()
    path <- isolate(pms_ai_stream_path())
    partial <- if(!is.null(path) && file.exists(path))tryCatch({
      sz <- file.info(path)$size
      if(is.finite(sz) && sz>0){
        r <- readBin(path,what='raw',n=sz)
        iconv(rawToChar(r),from='UTF-8',to='UTF-8',sub='')
      }else ''
    },error=function(e)'') else ''
    if(nzchar(partial)){
      pms_ai_answer(partial)
    }
    ready <- tryCatch(future::resolved(job),error=function(e)e)
    if(inherits(ready,'error')){
      msg <- conditionMessage(ready)
      pms_ai_answer(paste('비동기 상태 확인 실패:',msg,'\n',partial))
      pms_ai_status('AI 분석 실패')
      session$sendCustomMessage('pms_ai_debug',list(text=paste('④ 상태 확인 실패:',msg)))
      pms_ai_job(NULL);pms_ai_running(FALSE)
      return()
    }
    if(!isTRUE(ready)){
      elapsed <- as.numeric(difftime(Sys.time(),isolate(pms_ai_started()),units='secs'))
      if(is.finite(elapsed)){
        status <- if(nzchar(partial))'AI 답변 수신 중' else '첫 AI 응답 대기 중'
        pms_ai_status(sprintf('%s · 경과 %.0f초 · 최대 600초',status,elapsed))
        if(nzchar(partial))session$sendCustomMessage('pms_ai_debug',list(text=sprintf('④ AI 답변 수신 중 · %d자',nchar(partial))))
      }
      return()
    }
    answer <- tryCatch(future::value(job),error=function(e)e)
    pms_ai_job(NULL);pms_ai_running(FALSE)
    if(!is.null(path) && file.exists(path))unlink(path)
    pms_ai_stream_path(NULL)
    if(inherits(answer,'error')){
      msg <- conditionMessage(answer)
      message('PMS2 AI ERROR: ',msg)
      if(!nzchar(partial))pms_ai_answer(paste('AI 분석 실패:',msg))
      else pms_ai_answer(paste0(partial,'\n\n[응답 도중 오류: ',msg,']'))
      pms_ai_status('Ollama 분석 오류')
      session$sendCustomMessage('pms_ai_debug',list(text=paste('④ AI 호출 실패:',msg)))
    }else{
      pms_ai_answer(answer)
      pms_ai_status(paste0('분석 완료 · ',format(Sys.time(),'%Y-%m-%d %H:%M:%S',tz='Asia/Seoul'),' KST · 저장 시세 기반'))
      session$sendCustomMessage('pms_ai_debug',list(text='④ AI 답변 도착. 결과 표시 완료.'))
    }
  })
  session$onSessionEnded(function(){
    path <- isolate(pms_ai_stream_path())
    if(!is.null(path) && file.exists(path))unlink(path)
  })

  saved_eval_status <- reactiveVal('서버 준비 완료. 저장된 시세·배당 평가 버튼을 누르세요.')
  output$saved_eval_server_status <- renderText(saved_eval_status())
  saved_eval_ack <- function(message, error=FALSE) {
    saved_eval_status(message)
    session$sendCustomMessage('pms_eval_ack',list(text=message,error=isTRUE(error)))
  }
  flow_update <- function(state='running',stage=1L,title='',detail='') {
    session$sendCustomMessage('pms_flow',list(state=state,stage=as.integer(stage),title=title,detail=detail))
  }

  init_market_tables(con)
  pms_init_dividend_tables(con)
  pms_init_benchmark(con)
  contrib_refresh <- reactiveVal(0L)
  contrib_notice <- reactiveVal('납입액을 직접 등록하면 그래프와 표에 반영됩니다.')
  output$contrib_notice <- renderText(contrib_notice())
  contrib_data <- reactive({contrib_refresh();pms_contrib_summary(con,as.integer(input$contrib_year))})
  output$contrib_bars <- renderUI(pms_contrib_panel(contrib_data()))
  output$contrib_table <- renderDT({contrib_refresh();x<-pms_contrib_rows(con,as.integer(input$contrib_year));DT::datatable(x,rownames=FALSE,options=list(pageLength=12,scrollX=TRUE),escape=TRUE)})
  observeEvent(input$contrib_year, {
    contrib_refresh(isolate(contrib_refresh())+1L)
    d<-DBI::dbGetQuery(con,'SELECT carryover_krw FROM isa_allowance WHERE year=?',params=list(as.integer(input$contrib_year)))
    updateNumericInput(session,'isa_carry',value=if(nrow(d)) d$carryover_krw[1] else 0)
  },ignoreInit=FALSE)
  observeEvent(input$contrib_add, {
    tryCatch({
      d<-as.character(input$contrib_date)
      if(!identical(nchar(d),10L)||is.na(as.Date(d))) stop('올바른 납입일을 입력해 주세요.')
      a<-as.numeric(input$contrib_amount)
      if(length(a)!=1L||!is.finite(a)||a<=0) stop('납입금액은 0보다 커야 합니다.')
      DBI::dbExecute(con,'INSERT INTO account_contributions(contribution_date,account_type,broker,account_number,amount_krw,memo) VALUES (?,?,?,?,?,?)',params=list(d,input$contrib_type,trimws(input$contrib_broker),trimws(input$contrib_account),a,trimws(input$contrib_memo)))
      updateSelectInput(session,'contrib_year',selected=substr(d,1,4))
      contrib_refresh(isolate(contrib_refresh())+1L)
      contrib_notice('납입내역을 SQLite에 저장했습니다.')
    },error=function(e) contrib_notice(paste('저장 오류:',conditionMessage(e))))
  })
  observeEvent(input$isa_carry_save, {
    tryCatch({
      a<-as.numeric(input$isa_carry)
      if(length(a)!=1L||!is.finite(a)||a<0) stop('이월한도는 0 이상의 금액이어야 합니다.')
      DBI::dbExecute(con,'INSERT INTO isa_allowance(year,carryover_krw,updated_at) VALUES (?,?,CURRENT_TIMESTAMP) ON CONFLICT(year) DO UPDATE SET carryover_krw=excluded.carryover_krw, updated_at=CURRENT_TIMESTAMP',params=list(as.integer(input$contrib_year),a))
      contrib_refresh(isolate(contrib_refresh())+1L)
      contrib_notice('ISA 이월 한도를 저장했습니다.')
    },error=function(e) contrib_notice(paste('저장 오류:',conditionMessage(e))))
  })
  observeEvent(input$contrib_delete, {
    id<-suppressWarnings(as.integer(input$contrib_delete_id))
    if(length(id)!=1L||is.na(id)||id<=0) {contrib_notice('삭제할 올바른 ID를 입력해 주세요.');return()}
    n<-DBI::dbExecute(con,'DELETE FROM account_contributions WHERE id=?',params=list(id))
    contrib_refresh(isolate(contrib_refresh())+1L)
    contrib_notice(if(n>0) paste0('ID ',id,' 삭제 완료') else '해당 ID가 없습니다.')
  })

  fire_inputs <- reactive({
    market_refresh()
    d <- pms_valuation(con,include_zero=FALSE)
    active <- if(nrow(d)) !is.na(d$quantity) & d$quantity>0 else logical()
    known <- active & is.finite(d$market_value_krw)
    nav <- if(any(known)) sum(d$market_value_krw[known]) else NA_real_
    missing <- sum(active & !known)
    x <- pms_fire_scenarios(nav,finance_income=input$fire_finance_income,
      foreign_gain_share=input$fire_foreign_gain_pct,other_tax_base=input$fire_other_tax_base,
      property_health_month=input$fire_property_health,
      include_health=isTRUE(input$fire_regional_health))
    list(nav=nav,missing=missing,table=x)
  })
  fire_plan_data <- reactive({
    z <- fire_inputs()
    pms_fire_plan(z$nav, age=as.numeric(input$fire_plan_age),
      end_age=as.numeric(input$fire_plan_end_age),
      monthly_spend=as.numeric(input$fire_plan_spend),
      inflation=as.numeric(input$fire_plan_inflation)/100,
      pension_age=as.numeric(input$fire_plan_pension_age),
      pension_monthly=as.numeric(input$fire_plan_pension),
      fee=as.numeric(input$fire_plan_fee)/100,
      shock=as.numeric(input$fire_plan_crash)/100)
  })
  output$fire_plan_summary <- renderUI({
    pms_fire_plan_ui(fire_plan_data(),fire_inputs()$nav,
      as.numeric(input$fire_plan_spend),as.numeric(input$fire_plan_age),
      as.numeric(input$fire_plan_pension_age),as.numeric(input$fire_plan_pension))
  })
  output$fire_plan_plot <- renderPlot({
    x <- fire_plan_data()$paths
    if(!nrow(x)) {plot.new();text(.5,.5,'입력 조건 또는 자산 시세를 확인해 주세요.');return(invisible(NULL))}
    labels <- unique(x$시나리오)
    cols <- c('#3167ac','#35996a','#c45d55','#8a63ac')
    rng <- range(x$연말잔고/1e8,0,finite=TRUE)
    plot(NA,xlim=range(x$나이),ylim=rng,xlab='나이(세)',ylab='연말 잔고(억 원)',
      main='은퇴 이후 금융자산 예상 잔고',panel.first=grid(col='#e7e7e7'))
    for(i in seq_along(labels)) {
      d <- x[x$시나리오==labels[i],,drop=FALSE]
      lines(d$나이,d$연말잔고/1e8,col=cols[i],lwd=2,lty=if(i==4)2 else 1)
    }
    legend('topright',legend=labels,col=cols,lwd=2,lty=c(1,1,1,2),cex=.75,bg='white')
  })
  output$fire_plan_result <- DT::renderDT({
    x <- fire_plan_data()$summary
    if(!nrow(x))return(DT::datatable(data.frame(안내='시나리오 데이터 없음'),options=list(dom='t')))
    DT::datatable(x,rownames=FALSE,options=list(dom='t',scrollX=TRUE,
      columnDefs=pms_money_defs(x)),escape=TRUE)
  })
  output$fire_plan_checkpoints <- DT::renderDT({
    x <- fire_plan_data()$paths
    if(!nrow(x))return(DT::datatable(data.frame(안내='시나리오 데이터 없음'),options=list(dom='t')))
    d <- x[x$시나리오=='기준 · 연 5%' & (x$경과연수==1L | x$경과연수%%5L==0L |
      x$경과연수==max(x$경과연수)),c('나이','연간생활비','연간연금','연간필요인출','연말잔고'),drop=FALSE]
    DT::datatable(d,rownames=FALSE,options=list(pageLength=15,scrollX=TRUE,
      columnDefs=pms_money_defs(d)),escape=TRUE)
  })
  output$fire_summary <- renderUI({z<-fire_inputs();pms_fire_scenario_ui(z$table,z$nav,z$missing)})
  output$fire_table <- renderDT({
    x<-fire_inputs()$table
    if(!nrow(x)) return(DT::datatable(data.frame(안내='평가 가능한 금융자산이 없습니다.'),options=list(dom='t')))
    DT::datatable(x,rownames=FALSE,options=list(pageLength=11,ordering=FALSE,scrollX=TRUE,
      columnDefs=pms_money_defs(x)),escape=TRUE)
  })
  # 2015/2025 annual-average CPI (2020=100), sourced from Korean official CPI
  # time series disseminated in the Bank of Korea ECOS. Historical extrapolation only.
  pms_fire_cpi_2015 <- 94.86
  pms_fire_cpi_2025 <- 116.61
  pms_fire_inf_rate <- (pms_fire_cpi_2025 / pms_fire_cpi_2015)^(1/10) - 1
  output$fire_inflation_note <- renderUI({
    tags$div(class='alert alert-info',
      sprintf('2015년 CPI %.2f → 2025년 CPI %.2f (2020=100) · 최근 10년 연평균 물가상승률(복리): %.2f%%',
         pms_fire_cpi_2015,pms_fire_cpi_2025,100*pms_fire_inf_rate),
      tags$br(),
      '명목 월 세후 인출액은 매년 일정하다고 가정하며, 세금·건강보험료와 인출률 자체의 장래 변화는 예측하지 않습니다. ',
      '출처: 한국은행 ECOS 소비자물가지수(국가데이터처 작성 통계), 연평균 2015~2025.')
  })
  fire_real_data <- reactive({
    z<-fire_inputs()$table
    if(!nrow(z))return(data.frame())
    sel<-z[c(1L,3L,5L,7L,9L,11L),,drop=FALSE]
    years<-0:10
    long<-do.call(rbind,lapply(seq_len(nrow(sel)),function(i){
      base<-as.numeric(sel$월_세후_가용액_추정[i]);
      data.frame(인출률=sel$인출률[i],경과연수=years,
        월_명목_가용액_원=round(base),
        월_실질_가용액_원=round(base/(1+pms_fire_inf_rate)^years),check.names=FALSE)
    }))
    rownames(long)<-NULL;long
  })
  output$fire_real_table <- DT::renderDT({
    z<-fire_real_data()
    if(!nrow(z))return(DT::datatable(data.frame(안내='평가가능 자산 없음'),options=list(dom='t')))
    z<-z[z$경과연수 %in% c(0,1,3,5,7,10),,drop=FALSE]
    DT::datatable(z,rownames=FALSE,options=list(pageLength=25,scrollX=TRUE,
      columnDefs=list(list(targets=c(2,3),className='dt-right'))),filter='top') |>
      DT::formatCurrency(columns=c('월_명목_가용액_원','월_실질_가용액_원'),currency='',digits=0,mark=',',interval=3)
  })
  output$fire_real_plot <- renderPlot({
    z<-fire_real_data()
    if(!nrow(z)){plot.new();text(.5,.5,'자산 평가 필요');return(invisible(NULL))}
    opts<-unique(z$인출률)
    ylim<-range(z$월_실질_가용액_원/10000,finite=TRUE)
    plot(0:10,z$월_실질_가용액_원[z$인출률==opts[1]]/1e4,type='n',
       xlab='향후 경과 연수',ylab='월 실질가용액 (현재 만원 구매력)',ylim=ylim,
       main='동일 명목 인출액 가정: 물가상승에 따른 구매력 감소')
    grid(col='gray88')
    cols<-grDevices::hcl.colors(length(opts),'Dark 3')
    for(i in seq_along(opts))lines(0:10,z$월_실질_가용액_원[z$인출률==opts[i]]/1e4,
      col=cols[i],lwd=2)
    legend('topright',opts,col=cols,lty=1,lwd=2,bty='n',cex=.8,ncol=3)
  })

  output$fire_plot <- renderPlot({
    x<-fire_inputs()$table
    if(!nrow(x)) {plot.new();text(.5,.5,'평가액 데이터가 필요합니다.');return(invisible(NULL))}
    v<-x$월_세전_인출액/1e4;n<-x$월_세후_가용액_추정/1e4
    plot(seq_along(v),v,type='b',pch=16,lwd=2,col='#c85050',xaxt='n',
      xlab='연간 인출률',ylab='월 금액 (만원)',ylim=range(c(0,v,n)),main='월 세전·세후 가용액 비교')
    axis(1,at=seq_along(v),labels=x$인출률)
    lines(seq_along(n),n,type='b',pch=16,lwd=2,col='#2769a6')
    grid(col='gray88');legend('topleft',c('월 세전','월 세후 추정'),lty=1,pch=16,
      col=c('#c85050','#2769a6'),bty='n')
  })

  bench_refresh <- reactiveVal(0L)
  bench_notice <- reactiveVal('과거 시세를 처음 조회할 때는 [최근 1년 일별 주가 수집]을 눌러주세요.')
  output$bench_notice <- renderText({bench_refresh();paste(bench_notice(),pms_benchmark_series(con)$notice)})
  output$bench_fetch_status <- renderDT({bench_refresh();x<-pms_history_fetch_table(con);DT::datatable(x,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE))})
  benchmark_data <- reactive({bench_refresh();pms_benchmark_series(con)})
  output$bench_stats <- renderDT({x<-benchmark_data()$statistics;if(nrow(x) && 'Sharpe' %in% names(x)) x<-x[order(-x$Sharpe,na.last=TRUE),,drop=FALSE];DT::datatable(x,rownames=FALSE,options=list(pageLength=8,scrollX=TRUE,ordering=FALSE,columnDefs=pms_money_defs(x)))})
  tr_refresh <- reactiveVal(0L)
  total_return_data <- reactive({bench_refresh();tr_refresh();pms_benchmark_total_return(con,benchmark_data())})
  output$bench_tr_notice <- renderText({total_return_data()$notice})
  output$bench_tr_collect_notice <- renderText({tr_collect_notice()})
  tr_collect_notice<-reactiveVal('저장된 배당정보가 불충분할 경우에만 보충 조회하세요. (기존 데이터는 재사용)')
  output$bench_tr_stats <- renderDT({x<-total_return_data()$statistics;if(nrow(x)) x<-x[order(-x$Sharpe,na.last=TRUE),,drop=FALSE];DT::datatable(x,rownames=FALSE,options=list(pageLength=8,scrollX=TRUE,columnDefs=pms_money_defs(x)))})
  output$bench_tr_coverage <- renderDT({DT::datatable(total_return_data()$coverage,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE))})
  output$bench_tr_plot <- renderPlot({
    x<-total_return_data()$series
    if(!nrow(x)) {plot.new();text(.5,.5,'총수익률 데이터가 부족합니다. 아래 안내와 배당이력 상태를 확인하세요.');return(invisible(NULL))}
    y<-as.matrix(x[-1]);cols<-c('#d7191c','#2479ad','#53a36c','#9e65b2');widths<-c(4.5,1.8,1.8,1.8)
    matplot(x$날짜,y,type='l',lty=1,col=cols,lwd=widths,xlab='거래일',ylab='총수익 지수 (시작=100)',main='Total Return · 배당 재투자 포함')
    abline(h=100,col='gray70',lty=3);lines(x$날짜,y[,1],col=cols[1],lwd=4.5)
    legend('topleft',legend=names(x)[-1],col=cols,lty=1,lwd=widths,cex=.8,bty='n')
  })
  output$bench_tr_drawdown <- renderPlot({
    z <- pms_benchmark_risk_curves(total_return_data()$series)$drawdown
    if(!nrow(z)) {plot.new();text(.5,.5,'총수익률 데이터를 확보한 뒤 낙폭을 표시합니다.');return(invisible(NULL))}
    cols <- c('#d7191c','#2479ad','#53a36c','#9e65b2')
    matplot(z$날짜,as.matrix(z[-1]),type='l',lty=1,col=cols,lwd=c(4,1.5,1.5,1.5),
            xlab='거래일',ylab='고점 대비 낙폭 (%)',main='Total Return Drawdown 비교')
    abline(h=0,col='gray75',lty=3)
    lines(z$날짜,z[[2]],col=cols[1],lwd=4)
    legend('bottomleft',legend=names(z)[-1],col=cols,lty=1,lwd=c(4,1.5,1.5,1.5),cex=.7,bty='n')
  })
  output$bench_tr_rolling_sharpe <- renderPlot({
    z <- pms_benchmark_risk_curves(total_return_data()$series,60L)$rolling
    if(!nrow(z) || !any(is.finite(as.matrix(z[-1])))) {
      plot.new();text(.5,.5,'Total Return 60거래일 이동 Sharpe 계산에 충분한 데이터가 없습니다.');return(invisible(NULL))
    }
    cols <- c('#d7191c','#2479ad','#53a36c','#9e65b2')
    matplot(z$날짜,as.matrix(z[-1]),type='l',lty=1,col=cols,lwd=c(4,1.5,1.5,1.5),
            xlab='거래일',ylab='60거래일 연환산 Sharpe',main='Total Return Rolling Sharpe 비교')
    abline(h=0,col='gray65',lty=3)
    lines(z$날짜,z[[2]],col=cols[1],lwd=4)
    legend('topleft',legend=names(z)[-1],col=cols,lty=1,lwd=c(4,1.5,1.5,1.5),cex=.7,bty='n')
  })
  observeEvent(input$bench_tr_collect,{
    tr_collect_notice('배당이력을 확인하고 있습니다...')
    ans<-tryCatch(shiny::withProgress(message='Total Return 배당 이력 확인',value=0,{
      pms_bench_collect_dividend_coverage(con,on_result=function(i,n,cn,sy,status){shiny::setProgress(value=i/max(1,n),detail=sprintf('%d/%d · %s %s: %s',i,n,cn,sy,status))})
    }),error=function(e)e)
    if(inherits(ans,'error')) tr_collect_notice(paste('배당 조회 오류:',conditionMessage(ans))) else tr_collect_notice(sprintf('배당 조회 완료: 성공 %d, 캐시 %d, 실패 %d',sum(ans$status=='OK'),sum(ans$status=='CACHED'),sum(ans$status=='FAILED')))
    tr_refresh(isolate(tr_refresh())+1L)
  })
  output$bench_weights <- renderDT({x<-benchmark_data()$weights;DT::datatable(x,rownames=FALSE,options=list(pageLength=12,scrollX=TRUE,columnDefs=pms_money_defs(x)))})
  output$bench_excluded <- renderDT({x<-benchmark_data()$excluded;DT::datatable(x,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE))})
  output$bench_drawdown <- renderPlot({
    z <- pms_benchmark_risk_curves(benchmark_data()$series)$drawdown
    if(!nrow(z)) {plot.new();text(.5,.5,'저장된 공통 가격 데이터가 없습니다.');return(invisible(NULL))}
    cols<-c('#d7191c','#2479ad','#53a36c','#9e65b2')
    matplot(z$날짜,as.matrix(z[-1]),type='l',lty=1,col=cols,lwd=c(4,1.5,1.5,1.5),
            xlab='거래일',ylab='고점 대비 낙폭 (%)',main='Drawdown 비교')
    abline(h=0,col='gray75',lty=3)
    lines(z$날짜,z[[2]],col=cols[1],lwd=4)
    legend('bottomleft',legend=names(z)[-1],col=cols,lty=1,lwd=c(4,1.5,1.5,1.5),cex=.7,bty='n')
  })
  output$bench_rolling_sharpe <- renderPlot({
    z <- pms_benchmark_risk_curves(benchmark_data()$series,60L)$rolling
    if(!nrow(z) || !any(is.finite(as.matrix(z[-1])))) {
      plot.new();text(.5,.5,'60거래일 이동 Sharpe 계산에 충분한 데이터가 없습니다.');return(invisible(NULL))
    }
    cols<-c('#d7191c','#2479ad','#53a36c','#9e65b2')
    matplot(z$날짜,as.matrix(z[-1]),type='l',lty=1,col=cols,lwd=c(4,1.5,1.5,1.5),
            xlab='거래일',ylab='60거래일 연환산 Sharpe',main='Rolling Sharpe 비교')
    abline(h=0,col='gray65',lty=3)
    lines(z$날짜,z[[2]],col=cols[1],lwd=4)
    legend('topleft',legend=names(z)[-1],col=cols,lty=1,lwd=c(4,1.5,1.5,1.5),cex=.7,bty='n')
  })
  output$bench_plot <- renderPlot({
    x<-benchmark_data()$series
    if(!nrow(x)) {plot.new();text(.5,.5,'비교 가능한 과거 일별 주가가 없습니다. 수집 현황을 확인하세요.');return(invisible(NULL))}
    y<-as.matrix(x[,-1,drop=FALSE]);cols<-c('#d7191c','#2479ad','#53a36c','#9e65b2')[seq_len(ncol(y))]
    widths<-ifelse(colnames(y)=='내 포트폴리오',4.5,1.8)
    matplot(x$날짜,y,type='l',lty=1,lwd=widths,col=cols,xlab='거래일',ylab='기준지수 (시작=100)',main='최근 1년 벤치마크 비교')
    abline(h=100,col='gray65',lty=3)
    # Plot again last so the portfolio line stays visible above benchmark curves.
    if('내 포트폴리오' %in% colnames(y)) lines(x$날짜,y[,'내 포트폴리오'],col='#d7191c',lwd=4.5)
    legend('topleft',legend=colnames(y),col=cols,lty=1,lwd=widths,bty='n',cex=.8)
  })
  observeEvent(input$bench_recalc,{bench_refresh(isolate(bench_refresh())+1L);bench_notice('저장된 데이터로 다시 계산했습니다. (인터넷 재조회 없음)')})
  bench_busy <- reactiveVal(FALSE)
  bench_progress <- function(percent,message) {
    session$sendCustomMessage('pms_bench_progress',list(percent=as.integer(percent),message=as.character(message)))
  }
  bench_run <- function(force=FALSE) {
    if(isolate(bench_busy())) return(invisible(NULL))
    bench_busy(TRUE)
    on.exit(bench_busy(FALSE),add=TRUE)
    bench_progress(0L,'작업 준비 중...')
    bench_notice(if(force)'과거 시세를 초기화하고 전체 재수집합니다.' else '저장된 시세를 확인하고 부족한 구간만 수집합니다.')
    res <- tryCatch(shiny::withProgress(message=if(force)'전체 초기화 후 일별 주가 재수집' else '과거 주가 캐시 검사 및 보충', value=0, {
      if(force) pms_reset_history(con)
      pms_collect_history(con,force=force,on_result=function(row,i,n,message){
        completed <- if(is.null(row)) i-1L else i
        pct<-floor(100*completed/max(1,n))
        bench_progress(pct,sprintf('%d / %d 종목 · %s',completed,n,message))
        shiny::setProgress(value=completed/max(1,n),detail=sprintf('%d/%d · %s',completed,n,message))
      })
    }),error=function(e)e)
    if(inherits(res,'error')) {
      bench_notice(paste('과거 주가 조회 오류:',conditionMessage(res)))
      bench_progress(0L,paste('오류:',conditionMessage(res)))
    } else {
      ok<-sum(res$상태=='성공'); reused<-sum(res$상태=='재사용'); failed<-sum(res$상태=='실패')
      bench_notice(sprintf('조회 완료: 새 수집 %d · 캐시 재사용 %d · 실패 %d 종목. 실패한 종목의 저장된 과거 가격은 유지됩니다.',ok,reused,failed))
      bench_progress(100L,sprintf('완료 · 수집 %d / 재사용 %d / 실패 %d',ok,reused,failed))
    }
    bench_refresh(isolate(bench_refresh())+1L)
  }
  observeEvent(input$bench_collect,{bench_run(FALSE)})
  observeEvent(input$bench_reset_request,{
    showModal(modalDialog(title='과거 주가 전체 초기화 확인',
      '벤치마크 및 보유종목의 저장된 과거 일별 주가와 수집 상태만 삭제한 후 전부 다시 조회합니다. 현재 시세, 배당, 환율 및 보유종목은 삭제되지 않습니다. 네트워크 실패 시 과거 가격을 복구할 수 없으므로 주의하세요.',
      footer=tagList(modalButton('취소'),actionButton('bench_reset_confirm','초기화하고 다시 수집',class='btn-danger')),easyClose=TRUE))
  })
  observeEvent(input$bench_reset_confirm,{
    removeModal()
    bench_run(TRUE)
  })
  dividend_refresh <- reactiveVal(0L)
  dividend_message <- reactiveVal('저장된 배당정보를 자동으로 표시합니다. 신규 배당정보가 필요할 때만 수집 버튼을 누르세요.')
  # 수동 재평가는 DB를 읽기만 한다. 종목가격, 배당, 환율 API 호출이나 DB 삭제/갱신 없음.
  observeEvent(input$evaluate_saved_dividends, {
    saved_eval_ack(paste('배당 버튼 요청을 Shiny 서버에서 수신했습니다. SQLite 계산 중…',format(Sys.time(),'%H:%M:%S')))
    tryCatch({
      dividend_refresh(isolate(dividend_refresh()) + 1L)
      market_refresh(isolate(market_refresh()) + 1L)
      events_n <- DBI::dbGetQuery(con,'SELECT COUNT(*) AS n FROM dividend_events')$n[1]
      account_totals <- pms_dividend_account_totals(con)
      summary_rows <- nrow(pms_dividend_summary(con))
      msg <- sprintf('배당 재평가 완료: 저장 이벤트 %d건 / 종목별 통계 %d건 / 계좌별 통계 %d건 / %s (API 호출 없음)',
        as.integer(events_n),summary_rows,nrow(account_totals),format(Sys.time(),'%H:%M:%S'))
      dividend_message(msg)
      saved_eval_ack(msg)
    },error=function(e) {
      msg <- paste('저장된 배당금 평가 오류:',conditionMessage(e))
      dividend_message(msg)
      saved_eval_ack(msg,TRUE)
    })
  })
  output$dividend_notice <- renderText(dividend_message())
  output$dividend_summary <- renderDT({
    dividend_refresh()
    market_refresh()
    { x <- pms_dividend_summary(con); DT::datatable(x,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE,columnDefs=pms_money_defs(x)),escape=TRUE) }
  })
  dividend_account_data <- reactive({
    dividend_refresh()
    pms_dividend_account_totals(con)
  })
  output$dividend_account_totals <- renderDT({
    x <- dividend_account_data()
    amount_col <- '최근1년_배당금_원화_추정합계'
    # Keep raw numeric values until rendering; commas are display-only.
    tab <- DT::datatable(x,rownames=FALSE,options=list(pageLength=6,dom='t',scrollX=TRUE),escape=TRUE)
    DT::formatCurrency(tab,columns=amount_col,currency='',digits=0,mark=',',interval=3,dec.mark='.')
  })
  monthly_dividends <- reactive({
    dividend_refresh()
    pms_dividend_monthly_accounts(con)
  })
  # Independent HTML visualization: renders even on clients where Shiny plotOutput
  # or DataTables outputs never bind. Reads stored SQLite rows only, without APIs.
  # Per-symbol breakdown uses the same EX-date and current-holdings method
  # as pms_dividend_monthly_accounts(). Summed across accounts for the tooltip.
  pms_dividend_month_symbol_breakdown <- function() {
    first_month <- format(seq(as.Date(format(Sys.Date(),'%Y-%m-01')),by='-1 month',length.out=12)[12],'%Y-%m-01')
    raw <- DBI::dbGetQuery(con,"SELECT substr(d.ex_date,1,7) AS month,
       d.country,d.symbol,MIN(h.name) AS name,SUM(h.quantity) AS qty,d.amount,
       (SELECT f.rate FROM usdkrw_daily f WHERE f.rate_date<=d.ex_date
          AND f.rate_date>=date(d.ex_date,'-7 days') ORDER BY f.rate_date DESC LIMIT 1) AS fx_rate
       FROM dividend_events d JOIN holdings h ON h.country=d.country AND h.symbol=d.symbol
       WHERE h.quantity>0 AND h.country IN ('KR','US')
         AND d.ex_date>=? AND d.ex_date<=date('now')
       GROUP BY d.country,d.symbol,d.ex_date,d.amount,d.source",
       params=list(first_month))
    if(!nrow(raw)) return(data.frame(month=character(),name=character(),amount=numeric()))
    krw <- ifelse(raw$country=='KR',raw$amount,raw$amount*raw$fx_rate)*raw$qty
    keep <- is.finite(krw)
    if(!any(keep)) return(data.frame(month=character(),name=character(),amount=numeric()))
    df <- data.frame(month=raw$month[keep],name=paste0(raw$name[keep],' (',raw$symbol[keep],')'),amount=krw[keep])
    stats::aggregate(amount~month+name,df,sum)
  }
  pms_monthly_direct_build <- function() {
    tryCatch({
      df <- pms_dividend_monthly_accounts(con)
      types <- c('일반계좌','IRP','ISA','연금저축','기타/미분류')
      colors <- c('일반계좌'='#3679c6','IRP'='#e7993a','ISA'='#3aaf80',
                  '연금저축'='#a380cf','기타/미분류'='#8493a3')
      months <- unique(as.character(df$월))
      security_breakdown <- pms_dividend_month_symbol_breakdown()
      money <- function(x) format(round(x),big.mark=',',scientific=FALSE,trim=TRUE)
      if(!length(months)) stop('월별 집계 결과에 기준월 정보가 없습니다.')
      mx <- matrix(0,nrow=length(months),ncol=length(types),
                   dimnames=list(months,types))
      for(i in seq_len(nrow(df))) {
        mm <- as.character(df$월[i]); typ <- as.character(df$계좌유형[i])
        if(mm %in% months && typ %in% types && is.finite(df$금액[i]))
          mx[mm,typ] <- mx[mm,typ] + as.numeric(df$금액[i])
      }
      totals <- rowSums(mx)
      ceiling_val <- max(1,totals,na.rm=TRUE)
      bars <- lapply(seq_along(months),function(i) {
        mm <- months[i]; amt <- totals[i]
        # Each colored segment shares the same total pixel width, so its area
        # represents the account-type contribution to the monthly amount.
        segments <- lapply(types,function(typ) {
          width <- 100*mx[mm,typ]/ceiling_val
          if(!is.finite(width) || width<=0) return(NULL)
          tags$div(style=sprintf('display:inline-block;height:18px;width:%.5f%%;background:%s;',
                                 width,colors[[typ]]),
                   title=paste(typ,money(mx[mm,typ]),'원'))
        })
        by_sym <- security_breakdown[security_breakdown$month==mm,,drop=FALSE]
        if(nrow(by_sym)) by_sym <- by_sym[order(-by_sym$amount),,drop=FALSE]
        shown <- head(by_sym,5L)
        bubble <- if(nrow(shown)) tags$div(
          tags$strong(paste0(mm,' · 상위 ',nrow(shown),'종목')),
          lapply(seq_len(nrow(shown)),function(j) tags$div(
            style='display:flex;justify-content:space-between;gap:12px;',
            tags$span(style='overflow:hidden;text-overflow:ellipsis;white-space:nowrap;',shown$name[j]),
            tags$strong(style='white-space:nowrap;',paste0(money(shown$amount[j]),'원')))),
          if(nrow(by_sym)>5L) tags$div(style='opacity:.8;margin-top:5px;',
             sprintf('외 %d종목 · 막대를 클릭하면 전체 내역',nrow(by_sym)-5L))
          else tags$div(style='opacity:.8;margin-top:5px;','막대를 클릭하면 상세 내역'),
          tags$div(style='margin-top:5px;border-top:1px solid #697386;padding-top:5px;',
            paste0('월 합계: ',money(amt),'원'))
        ) else tags$div('월별 종목 데이터 없음 (배당/환율 누락 가능)')
        tags$tr(
          tags$td(mm,style='width:82px;white-space:nowrap;font-weight:600;'),
          tags$td(style='min-width:180px;width:58%;',
            tags$div(class='pms-month-tooltip-wrap',
              tags$button(type='button',class='pms-month-bar-button',
                `aria-label`=paste0(mm,' 배당 종목별 전체내역 열기'),
                onclick=sprintf("Shiny.setInputValue('pms_month_detail_clicked',{month:'%s',stamp:Date.now()},{priority:'event'});",mm),
                tags$div(style='white-space:nowrap;width:100%;height:18px;background:#f4f6f8;',segments)),
              tags$div(class='pms-month-tooltip',bubble))),
          tags$td(paste0(money(amt),'원'),style='text-align:right;white-space:nowrap;'))
      })
      detail_rows <- lapply(seq_along(months),function(i) {
        mm <- months[i]
        tags$tr(c(list(tags$td(mm,style='white-space:nowrap;')),
          lapply(types,function(typ)tags$td(money(mx[mm,typ]),style='text-align:right;')),
          list(tags$td(money(totals[i]),style='text-align:right;font-weight:700;'))))
      })
      tags$div(id='pms_monthly_dividend_direct',
        tags$style(HTML('.pms-month-tooltip-wrap{position:relative;} .pms-month-bar-button{border:0;background:transparent;padding:0;display:block;width:100%;text-align:left;cursor:pointer;} .pms-month-bar-button:focus-visible{outline:2px solid #327db8;outline-offset:2px;} .pms-month-tooltip{pointer-events:none;display:none;position:absolute;left:12px;bottom:24px;z-index:99999;width:330px;max-width:90vw;padding:10px 13px;background:#1f2937;color:white;border-radius:8px;box-shadow:0 5px 20px rgba(0,0,0,.25);font-size:12px;white-space:normal;} .pms-month-tooltip-wrap:hover .pms-month-tooltip,.pms-month-tooltip-wrap:focus-within .pms-month-tooltip{display:block;}')),
        tags$div(style='padding:9px;border:1px solid #bcd7bc;background:#f1faf1;margin-bottom:10px;',
          sprintf('SQLite 월별 배당 추정액 표시 완료 · 12개월 중 배당 있는 달 %d개월 · 총 %s원 · %s',
            sum(totals>0),money(sum(totals)),format(Sys.time(),'%H:%M:%S'))),
        tags$div(style='display:flex;flex-wrap:wrap;gap:12px;margin-bottom:8px;',
          lapply(types,function(typ)tags$span(
            tags$span(style=sprintf('display:inline-block;width:11px;height:11px;background:%s;margin-right:4px;',colors[[typ]])),typ))),
        tags$table(class='table table-condensed',style='max-width:1050px;width:100%;',
          tags$thead(tags$tr(tags$th('배당락 월'),tags$th('월별 누적 배당 추정액 (계좌유형별)'),
                              tags$th('월 합계',style='text-align:right;'))),tags$tbody(bars)),
        tags$details(tags$summary('12개월 계좌유형별 상세 금액표 펼치기'),
          tags$div(style='overflow-x:auto;',
            tags$table(class='table table-bordered table-condensed',style='width:100%;max-width:1050px;',
              tags$thead(tags$tr(c(list(tags$th('월')),
                lapply(types,tags$th),list(tags$th('합계'))))),tags$tbody(detail_rows)))),
        tags$p(style='font-size:12px;color:#666;margin-top:6px;',
          '막대에 마우스를 올리면 상위 5종목 요약, 클릭하면 전체 종목별 배당금 상세창을 볼 수 있습니다. 배당락일 기준 세전 추정치이며 실제 지급일·실수령액과 다릅니다.'))
    },error=function(e) {
      tags$div(id='pms_monthly_dividend_direct',class='alert alert-danger',
        paste('월별 배당금 통계 생성 오류:',conditionMessage(e)))
    })
  }
  observe({
    dividend_refresh()
    shiny::removeUI(selector='#pms_monthly_dividend_direct',multiple=TRUE,
                    immediate=TRUE,session=session)
    shiny::insertUI(selector='#pms_monthly_dividend_anchor',where='afterEnd',
                    ui=pms_monthly_direct_build(),immediate=TRUE,session=session)
  })
  # The hover summary is intentionally non-scrollable. A user click opens a
  # persistent Shiny modal; this also supports keyboard and touch interaction.
  observeEvent(input$pms_month_detail_clicked, {
    mm <- input$pms_month_detail_clicked$month
    if (!is.character(mm) || length(mm)!=1L || is.na(mm) ||
        !grepl('^[0-9]{4}-[0-9]{2}$',mm)) return(invisible(NULL))
    tryCatch({
      by_sym <- pms_dividend_month_symbol_breakdown()
      by_sym <- by_sym[by_sym$month==mm,,drop=FALSE]
      by_sym <- by_sym[order(-by_sym$amount,by_sym$name),,drop=FALSE]
      money <- function(x) format(round(x),big.mark=',',scientific=FALSE,trim=TRUE)
      df_month <- pms_dividend_monthly_accounts(con)
      total_month <- sum(df_month$금액[as.character(df_month$월)==mm],na.rm=TRUE)
      rows <- lapply(seq_len(nrow(by_sym)),function(i) {
        tags$tr(tags$td(by_sym$name[i]),
                tags$td(paste0(money(by_sym$amount[i]),'원'),
                        style='text-align:right;white-space:nowrap;'))
      })
      body <- if(nrow(by_sym)) tags$div(
        style='max-height:55vh;overflow-y:auto;overscroll-behavior:contain;',
        tags$table(class='table table-striped table-condensed',
          tags$thead(tags$tr(tags$th('종목명 (종목코드)'),
                              tags$th('세전 배당 추정액',style='text-align:right;'))),
          tags$tbody(rows)))
      else tags$p('해당 월의 종목별 배당 상세가 없습니다. 배당/환율 자료가 누락됐을 수 있습니다.')
      showModal(modalDialog(
        title=paste0(mm,' 종목별 배당금 상세'),
        tags$p(tags$strong(paste0('월별 배당 추정액: ',money(total_month),'원')),
               ' · ',nrow(by_sym),'종목'),
        body,
        tags$p(style='font-size:12px;color:#666;margin-top:8px;',
          '현재 보유수량과 과거 배당락일 기준 세전 추정액입니다. 실제 입금일 및 수령액과 다를 수 있습니다.'),
        footer=modalButton('닫기'),easyClose=TRUE,size='l'
      ))
    },error=function(e) {
      showModal(modalDialog(title='배당 상세 조회 오류',
        paste('SQLite 종목별 배당금 조회 실패:',conditionMessage(e)),
        easyClose=TRUE,footer=modalButton('닫기')))
    })
  })
  output$dividend_monthly_chart <- renderPlot({
    df <- monthly_dividends()
    types <- c('일반계좌','IRP','ISA','연금저축','기타/미분류')
    palette <- c('일반계좌'='#3679c6','IRP'='#e7993a','ISA'='#3aaf80',
                 '연금저축'='#a380cf','기타/미분류'='#8493a3')
    months <- unique(df$월)
    vals <- matrix(0,nrow=length(types),ncol=length(months),
                   dimnames=list(types,months))
    for(i in seq_len(nrow(df))) {
      if (is.finite(df$금액[i])) vals[df$계좌유형[i],df$월[i]] <- df$금액[i]
    }
    monthly_total <- colSums(vals,na.rm=TRUE)
    ytop <- max(1,monthly_total)
    # Above-bar totals need dedicated headroom and must not overlap the legend.
    par(mar=c(6.0,6.5,4.8,1.5),xpd=FALSE)
    bp <- barplot(vals,col=palette[types],border=NA,
                  names.arg=substr(months,3L,7L),las=2,cex.names=.85,
                  ylim=c(0,ytop*1.28),ylab='월별 세전 배당 추정액 (원)',
                  yaxt='n',main='최근 12개월 월별 배당금 (계좌별 누적)')
    yt <- pretty(c(0,ytop),n=5)
    yt <- yt[yt>=0 & yt<=ytop*1.12]
    axis(2,at=yt,labels=format(yt,big.mark=',',scientific=FALSE,trim=TRUE),las=1,cex.axis=.78)
    abline(h=yt,col='#e8edf2',lty=3)
    # Total KRW estimate for each calendar month, all account types combined.
    labels <- format(round(monthly_total),big.mark=',',scientific=FALSE,trim=TRUE)
    text(x=bp,y=monthly_total+ytop*.035,labels=labels,cex=.75,font=2,pos=3)
    legend('top',legend=types,fill=unname(palette[types]),
           horiz=TRUE,bty='n',cex=.77,inset=c(0,-.005),xpd=TRUE)
  })
  output$dividend_monthly_note <- renderText({
    df <- monthly_dividends()
    nmiss <- sum(df$환율누락건수,na.rm=TRUE)
    bad <- DBI::dbGetQuery(con,"SELECT COUNT(*) AS n FROM dividend_fetch_status WHERE status!='OK'")$n[1]
    paste0('기준: 최근 12개 달력월(이번 달 포함), 배당락일 기준·세전·현재 보유수량 적용. ',
           '환율 누락 이벤트 ',nmiss,'건, 최근 배당 조회 실패 종목 ',bad,'개. ',
           '조회 실패 종목의 배당은 그래프에서 빠질 수 있습니다. 월별 0원은 실제 무배당 확정이 아닙니다.')
  })
  general_income <- reactive({
    x <- dividend_account_data()
    g <- x[x$계좌유형=='일반계좌',,drop=FALSE]
    if(!nrow(g)) return(list(amount=NA_real_,missing=0L,complete=FALSE,holdings=0L))
    val <- suppressWarnings(as.numeric(g$최근1년_배당금_원화_추정합계[1]))
    list(amount=val, missing=as.integer(g$미완료건수[1]),
         complete=as.integer(g$미완료건수[1])==0L, holdings=as.integer(g$보유건수[1]))
  })
  # SQLite 기반 독립 HTML 시각화: Shiny plotOutput이 연결되지 않는 환경에서도 보임.
  pms_tax_direct_build <- function() {
    tryCatch({
      d <- general_income()
      lim <- 20000000
      money <- function(x) format(round(x),big.mark=',',scientific=FALSE,trim=TRUE)
      if (!is.finite(d$amount)) {
        return(tags$div(id='pms_tax_threshold_direct',class='alert alert-info',
          '일반계좌의 평가 가능한 배당금이 없습니다. 저장된 배당 이벤트와 계좌유형을 확인하세요.'))
      }
      amt <- max(0,as.numeric(d$amount))
      pct <- amt/lim*100
      maxval <- lim
      barpct <- min(100,amt/maxval*100)
      marker <- 100
      remaining <- lim-amt
      withheld <- amt*0.154 # 일반계좌 15.4% 원천징수 단순 추정
      net_amount <- amt-withheld
      monthly_net <- net_amount/12
      tip <- sprintf('세전 추정: %s원 | 원천징수 15.4%%: %s원 | 세후 추정: %s원 | 월평균 세후: %s원',money(amt),money(withheld),money(net_amount),money(monthly_net))
      tags$div(id='pms_tax_threshold_direct',style='max-width:1100px;margin:8px 0 18px;',
        tags$div(style='display:flex;flex-wrap:wrap;gap:22px;margin-bottom:12px;',
          tags$div(tags$div(style='font-size:12px;color:#666;', '일반계좌 세전 배당 추정액 (최근 365일)'),
            tags$strong(style='font-size:24px;',paste0(money(amt),'원'))),
          tags$div(tags$div(style='font-size:12px;color:#666;','비교 기준선'),
            tags$strong(style='font-size:24px;','20,000,000원')),
          tags$div(tags$div(style='font-size:12px;color:#666;','기준선 대비'),
            tags$strong(style='font-size:24px;',sprintf('%.1f%%',pct))),
          tags$div(tags$div(style='font-size:12px;color:#666;','세후 실수령 추정액'),
            tags$strong(style='font-size:24px;color:#167b62;',paste0(money(net_amount),'원'))),
          tags$div(tags$div(style='font-size:12px;color:#666;','세후 월평균 (÷12)'),
            tags$strong(style='font-size:24px;',paste0(money(monthly_net),'원')))),
        tags$div(style='position:relative;height:32px;background:#edf1f4;border-radius:4px;margin:10px 0 8px;',
          tags$div(title=tip,style=sprintf('display:flex;height:100%%;width:%.4f%%;border-radius:4px;overflow:hidden;cursor:help;',barpct),
            tags$div(title=tip,style='height:100%;width:84.6%;background:#4385c7;'),
            tags$div(title=tip,style='height:100%;width:15.4%;background:#e4a047;')),
          tags$div(style=sprintf('position:absolute;left:%.4f%%;top:-7px;height:46px;border-left:2px dashed #c23b3b;',marker),
            title='2,000만 원 기준선')),
        tags$div(style='display:flex;align-items:center;gap:18px;flex-wrap:wrap;font-size:12px;margin:8px 0;',
          tags$span(tags$span(style='color:#4385c7;','■ '),'세후 실수령 추정액: ',tags$b(paste0(money(net_amount),'원'))),
          tags$span(tags$span(style='color:#e4a047;','■ '),'15.4% 원천징수 추정액: ',tags$b(paste0(money(withheld),'원'))),
          tags$span(title=tip,style='cursor:help;text-decoration:underline dotted;', '막대에 마우스를 올리면 상세 금액이 표시됩니다.')),
        tags$div(style='display:flex;justify-content:space-between;font-size:12px;color:#666;',
          tags$span('0원'),tags$span('차트 상한 20,000,000원 (종합과세 비교 기준)')),
        if(amt>lim) tags$p(style='color:#c23b3b;font-weight:700;',paste0('상한 초과: 실제 추정액 ',money(amt),'원 (막대는 2,000만 원까지만 표시)')),
        tags$p(style='margin-top:12px;font-weight:600;',
          if(remaining>=0) paste0('기준선까지 ',money(remaining),'원 (',sprintf('%.1f',pct),'% 도달)')
          else paste0('기준선 초과 ',money(abs(remaining)),'원 (',sprintf('%.1f',pct),'% 도달)')),
        if(isTRUE(d$missing>0L)) tags$div(class='alert alert-warning',
          sprintf('배당/환율 미완료 %d건: 실제 배당금이 더 클 수 있습니다.',as.integer(d$missing))),
        tags$p(style='font-size:12px;color:#666;',
          '15.4%는 일반계좌 국내 원천징수의 단순 가정입니다. 해외배당의 외국 납부세액, 추가세금 및 종합과세 결과에 따라 실제 수령액은 다릅니다. 월평균은 최근 365일 세후 추정액을 12로 나눈 값으로 월별 실제 입금액을 뜻하지 않습니다. 최근 365일 배당락일 및 현재 보유수량을 적용한 세전 추정액이며, 실제 과세기간(1~12월)의 이자·배당 전체 금액과는 다릅니다. ISA·IRP·연금저축 금액은 제외됩니다.'),
        tags$div(style='font-size:12px;color:#607286;',
          paste('SQLite 배당 통계 갱신:',format(Sys.time(),'%H:%M:%S'))))
    },error=function(e) {
      tags$div(id='pms_tax_threshold_direct',class='alert alert-danger',
        paste('금융소득 비교 자료 표시 오류:',conditionMessage(e)))
    })
  }
  observe({
    dividend_refresh()
    shiny::removeUI(selector='#pms_tax_threshold_direct',multiple=TRUE,
                    immediate=TRUE,session=session)
    shiny::insertUI(selector='#pms_tax_threshold_anchor',where='afterEnd',
                    ui=pms_tax_direct_build(),immediate=TRUE,session=session)
  })
  output$general_income_threshold <- renderPlot({
    d <- general_income()
    lim <- 20000000
    plot(c(0,lim),c(0,1),type='n',
         axes=FALSE,xlab='',ylab='',ylim=c(0,1),bty='n')
    abline(v=lim,col='#c23b3b',lty=2,lwd=2)
    if(is.finite(d$amount)) {
      rect(0,0.25,min(lim,max(0,d$amount)),0.75,col='#4385c7',border=NA)
      text(min(lim,max(0,d$amount)),0.84,labels=paste0(format(round(d$amount),big.mark=',',scientific=FALSE,trim=TRUE),'원'),pos=2,cex=.9)
    } else text(lim*.45,.5,'확인된 배당 추정액 없음',cex=1)
    axis(1,at=seq(0,lim,by=5000000),
         labels=c('0','500','1,000','1,500','2,000'),cex.axis=.85)
    mtext('단위: 만 원 / 빨간 점선: 2,000만 원',side=1,line=2.6,cex=.8)
  })
  output$general_income_threshold_note <- renderText({
    d <- general_income()
    if(!is.finite(d$amount)) return('일반계좌의 확인 가능한 최근 365일 배당금이 없습니다. 배당정보를 조회해 주세요.')
    pct <- 100*d$amount/20000000
    paste0('일반계좌 추정 배당금 ',format(round(d$amount),big.mark=',',scientific=FALSE,trim=TRUE),
           '원 / 2,000만 원 기준선의 ',sprintf('%.1f',pct),'%. ',
           if(d$missing>0L) paste0('미완료 ',d$missing,'건이 있어 실제 합계는 더 클 수 있습니다. ') else '',
           '금융소득종합과세 판정에는 같은 과세연도의 다른 이자·배당도 포함됩니다.')
  })
  # 배당 이벤트 수와 예측 대상 건수는 서로 다르다. 예상 후보가 0건이어도 달력을 표시한다.
  calendar_data <- reactive({
    dividend_refresh()
    tryCatch(pms_dividend_forecast_30d(con), error=function(e) {
      structure(data.frame(), calendar_error=conditionMessage(e))
    })
  })
  output$dividend_calendar_view <- renderUI({
    x <- calendar_data()
    problem <- attr(x,'calendar_error',exact=TRUE)
    today <- Sys.Date()
    first <- today + 1L
    last <- today + 30L
    weekday_offset <- (as.POSIXlt(first)$wday + 6L) %% 7L
    grid_start <- first - weekday_offset
    num <- as.integer(last-grid_start)+1L
    day_dates <- seq(grid_start,by='day',length.out=ceiling(num/7L)*7L)
    fmt_money <- function(v) {
      if(length(v)!=1L || !is.finite(v)) return('금액미확인')
      paste0(format(round(v),big.mark=',',scientific=FALSE,trim=TRUE),'원')
    }
    tiles <- lapply(day_dates,function(d){
      d <- as.Date(d,origin='1970-01-01')
      day <- format(d,'%Y-%m-%d')
      inside <- d>=first && d<=last
      inds <- if(inside && nrow(x) && '예상_배당락일' %in% names(x)) {
        which(as.character(x[['예상_배당락일']])==day)
      } else integer(0)
      event_tags <- lapply(inds,function(i){
        amount <- suppressWarnings(as.numeric(x[['작년_배당금_원화_보유량기준']][i]))
        nm <- as.character(x[['종목명']][i])
        tags$div(class='event',title=nm,
          tags$span(class='event-name',title=nm,nm),
          tags$span(class='event-amount',fmt_money(amount)))
      })
      nums <- if(length(inds)) suppressWarnings(as.numeric(x[['작년_배당금_원화_보유량기준']][inds])) else numeric(0)
      tags$div(class=paste('pms-calendar-day',if(!inside) 'outside' else ''),
        tags$div(class='dayno',format(d,'%m/%d')),
        event_tags,
        if(length(inds)) tags$div(class='daytotal',paste0('일별 합계 ',
          if(all(is.finite(nums))) fmt_money(sum(nums)) else '일부 금액미확인')))
    })
    tags$div(class='pms-calendar',
      if(!is.null(problem)) tags$div(class='alert alert-danger',
        paste('예상 배당 달력 조회 오류:',problem)),
      if(is.null(problem) && !nrow(x)) tags$div(class='alert alert-info',
        '향후 30일에 해당하는 예상 배당락일이 없습니다. 아래 날짜 달력은 정상적으로 표시되며, 저장 배당 이벤트 291건 등의 연간 통계와 향후 30일 예상 건수는 다를 수 있습니다.'),
      tags$div(class='pms-calendar-explain',
        '예상 배당락일은 작년 같은 날짜를 1년 이동한 참고 일정입니다. 확정 지급일이 아닙니다.'),
      tags$div(class='pms-calendar-grid',
        lapply(c('월','화','수','목','금','토','일'),function(w) tags$div(class='pms-calendar-dow',w)),
        tiles))
  })
  # Direct DOM insertion fallback. This does not rely on the regular renderUI
  # output binding and visibly reports SQL/render errors to help diagnosis.
  pms_calendar_direct_build <- function() {
    tryCatch({
      x <- pms_dividend_forecast_30d(con)
      today <- Sys.Date()
      start <- today + 1L
      finish <- today + 30L
      offset <- (as.POSIXlt(start)$wday + 6L) %% 7L
      grid_start <- start - offset
      dates <- seq(grid_start,by='day',length.out=ceiling((as.integer(finish-grid_start)+1L)/7L)*7L)
      fmt <- function(v) {
        if(length(v)!=1L || !is.finite(v)) return('금액 미확인')
        paste0(format(round(v),big.mark=',',scientific=FALSE,trim=TRUE),'원')
      }
      headers <- lapply(c('월','화','수','목','금','토','일'),function(z)tags$div(class='pms-calendar-dow',z))
      tiles <- lapply(dates,function(d){
        d <- as.Date(d,origin='1970-01-01')
        key <- format(d,'%Y-%m-%d')
        inside <- d >= start && d <= finish
        ii <- if(inside && nrow(x) && '예상_배당락일' %in% names(x)) which(as.character(x$예상_배당락일)==key) else integer(0)
        items <- lapply(ii,function(i){
          name <- as.character(x$종목명[i])
          val <- suppressWarnings(as.numeric(x[['작년_배당금_원화_보유량기준']][i]))
          tags$div(class='event',title=name,
            tags$span(class='event-name',title=name,name),
            tags$span(class='event-amount',fmt(val)))
        })
        tags$div(class=paste('pms-calendar-day',if(!inside)'outside' else ''),
          tags$div(class='dayno',format(d,'%m/%d')),
          items,
          if(length(ii)){
            amounts<-suppressWarnings(as.numeric(x[['작년_배당금_원화_보유량기준']][ii]));
            tags$div(class='daytotal',paste0('일별 합계 ',if(all(is.finite(amounts))) fmt(sum(amounts)) else '일부 금액미확인'))
          })
      })
      total_events <- as.integer(DBI::dbGetQuery(con,'SELECT COUNT(*) n FROM dividend_events')$n[1])
      tags$div(id='pms_calendar_direct',
        tags$div(style='padding:9px;border:1px solid #bcd7bc;background:#f1faf1;margin-bottom:8px;',
          sprintf('SQLite 배당 이벤트 %d건 · 향후 30일 예상 %d건 · 달력 갱신 %s',
            total_events,nrow(x),format(Sys.time(),'%H:%M:%S')),
          if(!nrow(x)) tags$span(' · 해당 기간 예상 일정이 없어 날짜만 표시합니다.')),
        tags$div(class='pms-calendar',
          tags$div(class='pms-calendar-explain','예상 배당락일: 작년 배당락일 + 1년 (확정 지급일 아님)'),
          tags$div(class='pms-calendar-grid',headers,tiles)))
    },error=function(e) {
      tags$div(id='pms_calendar_direct',class='alert alert-danger',
        paste('SQLite 달력 직접 렌더링 오류:',conditionMessage(e)))
    })
  }
  observe({
    dividend_refresh()
    # No output$ binding: insert rendered tags beside the permanent anchor.
    # Remove the previous version before reinserting to avoid duplicate calendars.
    shiny::removeUI(selector='#pms_calendar_direct',multiple=TRUE,immediate=TRUE,session=session)
    shiny::insertUI(selector='#pms_calendar_anchor',where='afterEnd',
      ui=pms_calendar_direct_build(),immediate=TRUE,session=session)
  })
  # Explicitly refresh on tab visit if the original UI had been suspended.
  observeEvent(input$pms_tabs, {
    if(identical(input$pms_tabs,'시세수집 · 자산평가')) {
      dividend_refresh(isolate(dividend_refresh())+1L)
    }
  },ignoreInit=TRUE)
  outputOptions(output,'dividend_calendar_view',suspendWhenHidden=FALSE)
  output$dividend_calendar <- renderDT({
    x <- calendar_data()
    DT::datatable(x,rownames=FALSE,options=list(pageLength=15,scrollX=TRUE,
      columnDefs=pms_money_defs(x)),escape=TRUE)
  })
  output$dividend_calendar_note <- renderText({
    x <- calendar_data()
    problem <- attr(x,'calendar_error',exact=TRUE)
    n_all <- tryCatch(as.integer(DBI::dbGetQuery(con,
      'SELECT COUNT(*) AS n FROM dividend_events')$n[1]),error=function(e) NA_integer_)
    n_hist <- tryCatch(as.integer(DBI::dbGetQuery(con,
      'SELECT COUNT(*) AS n FROM dividend_events WHERE ex_date>=? AND ex_date<=?',
      params=list(format(Sys.Date()-373L,'%Y-%m-%d'),format(Sys.Date()-335L,'%Y-%m-%d')))$n[1]),
      error=function(e) NA_integer_)
    paste0(format(Sys.Date()+1,'%Y-%m-%d'),' ~ ',format(Sys.Date()+30,'%Y-%m-%d'),
      ' | SQLite 전체 배당 이벤트 ',n_all,'건',
      ' | 작년 대응기간 이벤트 ',n_hist,'건',
      ' | 현재 보유종목에 대응하는 30일 예상 ',nrow(x),'건',
      if(!is.null(problem)) paste0(' | 오류: ',problem) else if(!nrow(x))
        ' | 예상 일정 없음(저장된 배당통계와 별개)' else '',
      '. 실제 지급일이나 확정 공시 날짜가 아닙니다.')
  })
  # Shiny requires the output binding to exist before outputOptions().
  outputOptions(output,'dividend_calendar_note',suspendWhenHidden=FALSE)
  # Fail-safe historical dividend table: the calendar uses insertUI because
  # the regular reactive output was blank on some Windows Shiny sessions.
  # Use an independent, server-rendered HTML table for historical events too.
  pms_dividend_events_direct_build <- function() {
    tryCatch({
      df <- pms_dividend_events(con)
      total <- nrow(df)
      shown <- utils::head(df, 100L)
      money <- function(v, digits=2L) {
        z <- suppressWarnings(as.numeric(v))
        if(length(z)!=1L || !is.finite(z)) return('미확인')
        format(round(z,digits),big.mark=',',scientific=FALSE,trim=TRUE,nsmall=digits)
      }
      selected <- c('배당락일','종목명','종목코드','국가','주당배당금','통화',
        '해당배당일환율','주당배당금_원화','환산기준','배당출처')
      selected <- selected[selected %in% names(shown)]
      rows <- lapply(seq_len(nrow(shown)), function(i) {
        cells <- lapply(selected, function(nm) {
          val <- shown[[nm]][i]
          if(nm %in% c('주당배당금','주당배당금_원화')) {
            val <- money(val,if(nm=='주당배당금') 6L else 2L)
          } else if(nm=='해당배당일환율') {
            val <- money(val,2L)
          } else if(length(val)!=1L || is.na(val)) val <- '—'
          tags$td(as.character(val),style=if(nm %in% c('주당배당금','주당배당금_원화','해당배당일환율'))
            'text-align:right;white-space:nowrap;' else 'white-space:nowrap;')
        })
        tags$tr(cells)
      })
      tags$div(id='pms_dividend_events_direct',
        tags$div(style='padding:9px;border:1px solid #bcd7bc;background:#f1faf1;margin-bottom:8px;',
          sprintf('저장된 최근 365일 배당 이벤트 %d건 · 화면 표시 %d건 · 새로고침 %s (SQLite만 조회)',
            total,nrow(shown),format(Sys.time(),'%H:%M:%S')),
          if(total>100L) tags$span(' · 전체 내역은 기존 DataTables 상세표에서 조회할 수 있습니다.')),
        if(!total) tags$div(class='alert alert-info',
          '최근 365일의 배당 이벤트가 없습니다. 저장된 전체 배당 이력과 기간 조건은 다를 수 있습니다.'),
        if(total) tags$div(style='overflow-x:auto;max-height:500px;overflow-y:auto;border:1px solid #dbe4eb;',
          tags$table(class='table table-condensed table-striped',
            style='font-size:12px;margin-bottom:0;width:100%;',
            tags$thead(tags$tr(lapply(selected,function(nm)tags$th(nm,
              style='position:sticky;top:0;background:#eef3f8;white-space:nowrap;')))),
            tags$tbody(rows))))
    },error=function(e) {
      tags$div(id='pms_dividend_events_direct',class='alert alert-danger',
        paste('저장된 과거 배당 이벤트 표시 오류:',conditionMessage(e)))
    })
  }
  observe({
    dividend_refresh()
    shiny::removeUI(selector='#pms_dividend_events_direct',multiple=TRUE,
      immediate=TRUE,session=session)
    shiny::insertUI(selector='#pms_dividend_events_anchor',where='afterEnd',
      ui=pms_dividend_events_direct_build(),immediate=TRUE,session=session)
  })
  output$dividend_events <- renderDT({
    dividend_refresh()
    { x <- pms_dividend_events(con); tab <- DT::datatable(x,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE,columnDefs=pms_money_defs(x)),escape=TRUE); if('주당배당금_원화' %in% names(x)) tab <- DT::formatStyle(tab,'주당배당금_원화',backgroundColor='#fff2a8',color='#443300',fontWeight='bold'); tab }
  })
  run_dividend_collection <- function() {
    Sys.setenv(PMS_SSL_NO_REVOKE=if(isTRUE(input$ssl_revocation_skip)) '1' else '0')
    dividend_message('USD/KRW 일별 환율 및 배당 이력을 조회 중입니다...')
    session$sendCustomMessage('pms_source_visual',list(kind='dividends',event='reset'))
    session$sendCustomMessage('pms_dividend_source_reset',list())
    tryCatch({
      withProgress(message='최근 365일 배당 이력·일별 USD/KRW 수집',value=0,{
        fx_result <- tryCatch(pms_collect_usdkrw_daily(con),error=function(e)e)
        if(inherits(fx_result,'error')) {
          dividend_message(paste('일별 USD/KRW 조회 실패:',conditionMessage(fx_result),'· 기존 저장된 환율만 사용합니다. 환율이 없는 날짜는 원화 금액을 표시하지 않습니다.'))
        }
        x <- pms_collect_dividends(con,on_result=function(rec,i,n){
          ok <- identical(as.character(rec$status[1]),'OK')
          source <- if(ok) as.character(rec$detail[1]) else ''
          session$sendCustomMessage('pms_source_visual',list(kind='dividends',event='result',
            index=i,total=n,symbol=as.character(rec$symbol[1]),ok=ok,source=source,
            events=if(ok) as.integer(rec$count[1]) else NA_integer_))
          session$sendCustomMessage('pms_dividend_source_row',list(index=i,
            symbol=as.character(rec$symbol[1]),status=as.character(rec$status[1]),
            source=if(ok) source else '실패',events=if(ok) as.integer(rec$count[1]) else NA_integer_))
          incProgress(1/max(n,1),detail=paste(i,'/',n,rec$symbol[1],if(ok) source else '실패'))
        })
        session$sendCustomMessage('pms_source_visual',list(kind='dividends',event='done'))
        dividend_message(sprintf('배당 조회 완료: 성공 %d종목, 실패 %d종목. 일별 USD/KRW: %s. 실패한 종목은 이전 성공 이력을 보존합니다.',sum(x$status=='OK'),sum(x$status=='FAILED'),if(inherits(fx_result,'error')) '조회 실패(기존 저장값 사용)' else paste(fx_result,'거래일 저장')))
      })
    },error=function(e){ dividend_message(paste('배당 조회 오류:',conditionMessage(e))) })
    dividend_refresh(isolate(dividend_refresh())+1L)
    }
  observeEvent(input$collect_dividends, { run_dividend_collection() })
  all_collection_message <- reactiveVal('전체 조회 버튼을 누르면 시세와 배당을 차례대로 수집합니다.')
  output$all_collect_notice <- renderText(all_collection_message())
  market_refresh <- reactiveVal(0L)
  board_manual <- reactiveVal(0L)
  observeEvent(input$board_refresh,{board_manual(isolate(board_manual())+1L)})
  board_prices <- reactive({
    market_refresh()
    board_manual()
    invalidateLater(600000,session)
    pms_board_data(con)
  })
  pms_ticker_init(con)
  ticker_update <- reactiveVal(0L)
  ticker_note <- reactiveVal('저장된 시장지수를 표시합니다.')
  observeEvent(input$pms_tabs, {
    if(!identical(input$pms_tabs,'전일대비전광판')) return()
    # First collection on entering tab; 10-minute periodic checks below.
    ticker_note(pms_ticker_collect_indices(con))
    ticker_update(isolate(ticker_update())+1L)
  },ignoreInit=TRUE)
  observe({
    invalidateLater(600000,session)
    if(!identical(input$pms_tabs,'전일대비전광판')) return()
    if(!isTRUE(input$auto_prices_10m)) return()
    ticker_note(pms_ticker_collect_indices(con))
    ticker_update(isolate(ticker_update())+1L)
  })
  output$board_bottom_ticker <- renderUI({
    ticker_update()
    pms_ticker_ui(pms_ticker_entries(con,board_prices()),pms_summary_ticker_entries(con,board_prices()))
  })
  # Independent polling: market data can change without triggering board_prices().
  # The prices table stores fetched_at as UTC; the display is always KST.
  # v0.6.08: send directly to a fixed DOM span. Avoid dynamic Shiny textOutput
  # suspension / invisible output binding in the sticky ticker header.
  observe({
    shiny::invalidateLater(10000,session)
    q <- tryCatch(DBI::dbGetQuery(con,
      "SELECT MAX(fetched_at) AS fetched_at FROM prices WHERE fetched_at IS NOT NULL AND TRIM(fetched_at) <> ''"),
      error=function(e) NULL)
    stamp <- '최근 시세 수집: 기록 없음'
    if(!is.null(q) && nrow(q) && !is.na(q$fetched_at[1]) && nzchar(q$fetched_at[1])) {
      dt <- suppressWarnings(as.POSIXct(as.character(q$fetched_at[1]),
                               format='%Y-%m-%d %H:%M:%S',tz='UTC'))
      if(length(dt)==1L && !is.na(dt)) {
        stamp <- paste0('최근 시세 수집: ',
          format(dt,'%Y-%m-%d %H:%M:%S',tz='Asia/Seoul'),' KST')
      } else stamp <- '최근 시세 수집: 시각 해석 오류'
    }
    session$sendCustomMessage('pms_board_fetch_time_v608',list(text=stamp))
  })
  board_render_stamp <- reactiveVal(NULL)
  output$pms_board_rendered_at <- renderText({
    t <- board_render_stamp()
    if(is.null(t)) '화면 갱신: 대기 중' else paste0('화면 갱신: ',format(t,'%H:%M:%S',tz='Asia/Seoul'))
  })
  # The detailed table below remains visible regardless of card state.
  board_cards_open <- reactiveVal(FALSE)
  observeEvent(input$board_toggle_cards, {
    opened <- !isTRUE(isolate(board_cards_open()))
    board_cards_open(opened)
    updateActionButton(session,'board_toggle_cards',
      label=if(opened) '종목별 시세 카드 접기' else '종목별 시세 카드 펼치기')
  })
  output$board_cards_container <- renderUI({
    if(!isTRUE(board_cards_open())) return(NULL)
    uiOutput('board_panel')
  })
  output$board_panel <- renderUI({
    d <- board_prices()
    panel <- pms_board_panel(d)
    board_render_stamp(Sys.time())
    panel
  })
  # 보유수량 0주 종목은 DB의 계좌별 보유수량 합계를 기준으로 전광판 및 전송 검증에서 제외.
  # NULL/미확인 보유수량은 오류 감지를 위해 유지합니다. 매매이력과 저장 시세는 변경하지 않습니다.
  pms_board_active_valuation <- function() {
    z <- pms_board_valuation_changes(board_prices(), con)
    if (!nrow(z)) return(z)
    if (!all(c('국가','종목코드') %in% names(z))) return(z)
    h <- DBI::dbGetQuery(con, paste0(
      'SELECT UPPER(TRIM(country)) AS country, UPPER(TRIM(symbol)) AS symbol, ',
      'SUM(CASE WHEN quantity IS NOT NULL THEN quantity ELSE 0 END) AS total_qty, ',
      'SUM(CASE WHEN quantity IS NULL THEN 1 ELSE 0 END) AS unknown_qty ',
      'FROM holdings GROUP BY UPPER(TRIM(country)), UPPER(TRIM(symbol))'))
    key <- paste(toupper(trimws(as.character(z[["국가"]]))),
                 toupper(trimws(as.character(z[["종목코드"]]))),sep='::')
    hk <- paste(h$country,h$symbol,sep='::')
    idx <- match(key,hk)
    # DB에서 명확히 0주인 종목만 제외 (보유 미확인이나 양수/음수 이상값은 유지)
    excluded <- !is.na(idx) & is.finite(h$total_qty[idx]) &
      h$total_qty[idx] == 0 & h$unknown_qty[idx] == 0
    excluded[is.na(excluded)] <- FALSE
    z[!excluded, , drop=FALSE]
  }
  # Summary uses the exact same per-holding valuation rows as the board below.
  # Never treat missing quotes as zero movements; show partial coverage explicitly.
  board_totals <- reactive({
    z <- pms_board_active_valuation()
    get_sum <- function(col) {
      if(!col %in% names(z)) return(list(value=NA_real_,valid=0L,total=nrow(z)))
      v <- suppressWarnings(as.numeric(gsub(',', '', as.character(z[[col]]),fixed=TRUE)))
      good <- is.finite(v)
      list(value=if(any(good)) sum(v[good]) else NA_real_,valid=sum(good),total=length(v))
    }
    list(nav=get_sum('총평가액(원)'),change=get_sum('평가증감(원)'))
  })
  # Show incomplete calculation details right on the ticker board. The entire
  # portfolio return is intentionally withheld when any active ticker lacks a
  # comparable daily change; do not silently treat missing values as zero.
  output$board_missing_detail <- renderUI({
    z <- pms_board_active_valuation()
    if (!nrow(z)) return(NULL)
    num <- function(col) {
      if (!col %in% names(z)) return(rep(NA_real_, nrow(z)))
      suppressWarnings(as.numeric(gsub(',', '', as.character(z[[col]]), fixed=TRUE)))
    }
    nv <- num('총평가액(원)')
    dv <- num('평가증감(원)')
    missing_nav <- !is.finite(nv)
    missing_change <- !is.finite(dv)
    bad <- missing_nav | missing_change
    if (!any(bad)) return(NULL)
    codes <- if ('종목코드' %in% names(z)) as.character(z[['종목코드']]) else rep('코드 미확인', nrow(z))
    qty <- num('보유수량')
    details <- vapply(which(bad), function(i) {
      reason <- if (missing_nav[i] && missing_change[i]) '평가액·전일증감 모두 미산출' else
        if (missing_nav[i]) '평가액 미산출' else '전일증감 미산출'
      paste0(codes[i], ' (보유수량: ',
             if (is.finite(qty[i])) format(qty[i], scientific=FALSE, big.mark=',', trim=TRUE) else '미확인',
             ', ', reason, ')')
    }, character(1))
    tags$div(style='padding:7px 11px;margin:5px 0 8px;border:1px solid #e4c588;border-radius:5px;background:#fff9ea;color:#765313;font-size:13px;',
      tags$strong(sprintf('전광판 산출 누락: %d/%d종목 — ',sum(bad),nrow(z))),
      paste(details,collapse=' · '),
      tags$span(style='margin-left:7px;', '누락 원인을 확인할 때까지 전체 등락률은 표시하지 않습니다.'))
  })
  # Telegram: 0주 제외 후 전광판과 동일한 집계값 사용. 실제 보유종목 누락 시 전송 차단.
  tg_status <- reactiveVal('대기 중 · 상단 Bot Token / chat_id 입력 필요')
  output$pms_tg_status <- renderText(tg_status())
  output$pms_tg_missing_quote_status <- renderText({
    z <- pms_board_active_valuation()
    if (!nrow(z)) return('평가 대상 종목 없음')
    get_numeric <- function(col) {
      if(!col %in% names(z)) return(rep(NA_real_,nrow(z)))
      suppressWarnings(as.numeric(gsub(',','',as.character(z[[col]]),fixed=TRUE)))
    }
    bad <- !is.finite(get_numeric('총평가액(원)')) |
           !is.finite(get_numeric('평가증감(원)'))
    if(!any(bad)) return(sprintf('평가 및 전일 대비 집계 정상: %d/%d종목',nrow(z),nrow(z)))
    paste0('평가/전일 대비 계산 누락: ',sum(bad),'종목 — ',
      paste(paste(z[['종목코드']][bad],collapse=', ')),
      ' (해당 종목의 시세·보유수량을 확인하세요)')
  })
  observeEvent(input$pms_tg_send_test, {
    tryCatch({
      # 연결 확인 전용: NAV/등락률/시세 충족 여부와 무관하게 발송
      msg <- paste0('✅ PMS2 텔레그램 연결 테스트\n',
        '발송시각: ',format(Sys.time(),'%Y-%m-%d %H:%M:%S',tz='Asia/Seoul'),
        ' KST\n', 'PMS2와 텔레그램 봇의 연결이 정상입니다.')
      pms_tg_send(msg)
      tg_status(paste('테스트 메시지 발송 성공:',format(Sys.time(),'%Y-%m-%d %H:%M:%S')))
      showNotification('텔레그램 테스트 메시지 발송 성공',type='message',duration=8)
    },error=function(e){
      tg_status(paste('테스트 메시지 발송 실패:',conditionMessage(e)))
      showNotification(paste('텔레그램 테스트 실패:',conditionMessage(e)),type='error',duration=12)
    })
  })

  observeEvent(TRUE, {
    updateCheckboxInput(session, 'pms_tg_auto_16', value=pms_tg_setting_get(con))
  }, once=TRUE)
  observeEvent(input$pms_tg_auto_16, {
    pms_tg_setting_put(con, isTRUE(input$pms_tg_auto_16))
  }, ignoreInit=TRUE)
  # 수동 보고는 누락된 값을 0으로 가정하지 않고 부분 집계로 발송한다.
  # 16시 자동 보고는 기존과 동일하게 NAV/등락 모두 100% 계산돼야 발송한다.
  tg_message <- function(allow_partial_change=FALSE) {
    z <- isolate(board_totals())
    nav <- z$nav; delta <- z$change
    complete_nav <- is.finite(nav$value) && nav$total > 0L && nav$valid == nav$total
    complete_change <- is.finite(delta$value) && delta$total > 0L &&
      delta$valid == delta$total && delta$total == nav$total
    if (!isTRUE(allow_partial_change) && !complete_nav)
      stop(sprintf('총평가액 미산출: %d/%d종목. 자동 보고를 중단합니다.',nav$valid,nav$total))
    if (!isTRUE(allow_partial_change) && !complete_change)
      stop(sprintf('전일 대비 평가증감 미산출: %d/%d종목. 자동 보고를 중단합니다.',delta$valid,delta$total))
    if (nav$total < 1L) stop('현재 보유종목이 없어 보고하지 않습니다.')
    fmt <- function(v) format(round(v),big.mark=',',scientific=FALSE,trim=TRUE)
    head <- paste0('📊 PMS2 일일 자산 보고\n',
      '기준시각: ',format(Sys.time(),'%Y-%m-%d %H:%M',tz='Asia/Seoul'),' KST\n')
    if (complete_nav) {
      head <- paste0(head,'총평가액: ',fmt(nav$value),'원 (',nav$valid,'/',nav$total,'종목)\n')
    } else {
      head <- paste0(head,'⚠️ 평가액 부분 합계: ',
        if(is.finite(nav$value)) paste0(fmt(nav$value),'원') else '산출 불가',
        ' (',nav$valid,'/',nav$total,'종목)\n',
        '전체 총평가액: 산출 불가 (누락 종목의 평가액 미포함)\n')
    }
    if (complete_change && complete_nav) {
      change <- delta$value
      pct <- if (is.finite(nav$value-change) && (nav$value-change)>0)
        100*change/(nav$value-change) else NA_real_
      sgn <- if(change>0) '▲' else if(change<0) '▼' else '－'
      details <- paste0('전일 종가 대비 평가증감: ',sgn,' ',fmt(abs(change)),'원',
                       if(is.finite(pct)) sprintf(' (%+.2f%%)',pct) else '', '\n')
    } else {
      details <- paste0('⚠️ 전일 대비 부분 집계: ',
        if(is.finite(delta$value)) paste0(if(delta$value>0) '+' else '',fmt(delta$value),'원') else '산출 불가',
        ' (',delta$valid,'/',delta$total,'종목)\n',
        '전체 포트폴리오 등락률: 산출 불가\n')
    }
    if (!complete_nav || !complete_change) {
      rows <- isolate(pms_board_active_valuation())
      get_num <- function(col) {
        if(!col %in% names(rows)) return(rep(NA_real_,nrow(rows)))
        suppressWarnings(as.numeric(gsub(',','',as.character(rows[[col]]),fixed=TRUE)))
      }
      if(nrow(rows) && '종목코드' %in% names(rows)) {
        nv <- get_num('총평가액(원)')
        cv <- get_num('평가증감(원)')
        bad <- !is.finite(nv) | !is.finite(cv)
        if(any(bad)) {
          reasons <- ifelse(!is.finite(nv[bad]) & !is.finite(cv[bad]),'평가액·증감 누락',
                      ifelse(!is.finite(nv[bad]),'평가액 누락','증감 누락'))
          details <- paste0(details,'미산출 종목: ',
            paste0(as.character(rows[['종목코드']][bad]),'(',reasons,')',collapse=', '),'\n')
        }
      }
      details <- paste0(details,'※ 부분 집계는 전체 자산/일일 손익이 아닙니다.\n')
    }
    paste0(head,details,
      '※ 전광판 기준. 보유수량/현재 적용환율 고정, 환율변동·입출금 효과 제외.\n',
      '※ 16시 기준 해외주식은 전일 또는 최근 저장 시세일 수 있습니다.')
  }
  observeEvent(input$pms_tg_send_manual, {
    tryCatch({
      pms_tg_send(tg_message(allow_partial_change=TRUE))
      tg_status(paste('수동 전송 성공:',format(Sys.time(),'%Y-%m-%d %H:%M:%S')))
      showNotification('텔레그램 수동 전송 성공',type='message')
    },error=function(e){
      tg_status(paste('수동 전송 실패:',conditionMessage(e)))
      showNotification(conditionMessage(e),type='error',duration=10)
    })
  })
  # DB 날짜 PRIMARY KEY로 동일 날짜 자동발송 중복 방지.
  # Shiny 세션이 살아 있고 국내 거래일(월~금)인 날 16:00~16:59에 한 번 시도합니다.
  observe({
    invalidateLater(30000,session)
    now <- as.POSIXct(Sys.time(),tz='Asia/Seoul')
    today <- format(now,'%Y-%m-%d',tz='Asia/Seoul')
    hour <- as.integer(format(now,'%H',tz='Asia/Seoul'))
    weekday <- as.integer(format(now,'%u',tz='Asia/Seoul'))
    if (hour!=16L || weekday>5L || !isTRUE(pms_tg_setting_get(con)) || !pms_tg_configured())
      return(invisible(NULL))
    done <- DBI::dbGetQuery(con,'SELECT send_date FROM pms_tg_sent WHERE send_date=?',params=list(today))
    if (nrow(done)) return(invisible(NULL))
    # 복수 브라우저 세션을 위한 프로세스 수준 발송 잠금
    if (isTRUE(pms_tg_auto_busy)) return(invisible(NULL))
    pms_tg_auto_busy <<- TRUE
    on.exit({pms_tg_auto_busy <<- FALSE},add=TRUE)
    tryCatch({
      msg <- tg_message()
      z <- isolate(board_totals())
      pms_tg_send(msg)
      DBI::dbExecute(con,
        'INSERT OR IGNORE INTO pms_tg_sent(send_date,sent_at,total,delta) VALUES (?,?,?,?)',
        params=list(today,format(now,'%Y-%m-%d %H:%M:%S',tz='Asia/Seoul'),
                    z$nav$value,z$change$value))
      tg_status(paste('16시 자동보고 성공:',today))
    },error=function(e){
      tg_status(paste('16시 자동보고 실패:',conditionMessage(e)))
    })
  })
  # FX quote in the board header: current saved USD/KRW versus prior trading date.
  # Read-only SQLite access; no new schema and no network call on rendering.
  pms_board_fx_header <- function() {
    market_refresh()
    board_manual()
    invalidateLater(600000,session)
    tryCatch({
      now <- DBI::dbGetQuery(con,
        "SELECT rate,quoted_at FROM fx_rates WHERE pair='USD/KRW' LIMIT 1")
      if(!nrow(now)) return(tags$span('USD/KRW: 환율 미수집'))
      rate <- suppressWarnings(as.numeric(now$rate[1]))
      if(!is.finite(rate) || rate <= 0) return(tags$span('USD/KRW: 환율 산출 불가'))
      quote_date <- substr(as.character(now$quoted_at[1]),1L,10L)
      if(is.na(quote_date) || !grepl('^[0-9]{4}-[0-9]{2}-[0-9]{2}$',quote_date)) {
        quote_date <- format(Sys.Date(),'%Y-%m-%d')
      }
      title <- sprintf('USD/KRW %s원',format(round(rate,2),big.mark=',',nsmall=2,trim=TRUE))
      prev <- NULL
      available <- DBI::dbGetQuery(con,
        "SELECT name FROM sqlite_master WHERE type='table' AND name='usdkrw_daily'")
      if(nrow(available)) {
        prev <- tryCatch(DBI::dbGetQuery(con,
          "SELECT rate,rate_date FROM usdkrw_daily WHERE rate_date < ? AND rate > 0 ORDER BY rate_date DESC LIMIT 1",
          params=list(quote_date)),error=function(e) NULL)
      }
      if(is.null(prev) || !nrow(prev)) {
        return(tags$span(style='white-space:nowrap;',title,
          tags$small(style='font-weight:400;margin-left:4px;','(전일 자료 없음)')))
      }
      prior <- suppressWarnings(as.numeric(prev$rate[1]))
      if(!is.finite(prior) || prior <= 0) return(tags$span(title,' (전일 자료 없음)'))
      change <- rate-prior
      pct <- change/prior*100
      cls <- if(change > 0) 'pms-move-up' else if(change < 0) 'pms-move-down' else 'pms-move-flat'
      tags$span(class=cls,style='white-space:nowrap;',
        tags$b('USD/KRW '),
        sprintf('%s원',format(round(rate,2),big.mark=',',nsmall=2,trim=TRUE)),
        tags$span(style='font-weight:750;margin-left:6px;',
          sprintf('%+.2f원 (%+.2f%%)',change,pct)))
    },error=function(e) tags$span('USD/KRW: 확인 필요'))
  }
  board_amount_text <- function(x, show_sign=FALSE) {
    if(!is.finite(x)) return('산출 불가')
    paste0(if(show_sign && x>0) '+' else '',
           format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),'원')
  }
  # One UI binding makes both totals appear together in the same sticky header.
  # Failure is visible rather than silently leaving the entire header blank.
  output$board_header_totals <- renderUI({
    tryCatch({
      x <- board_totals()
      val <- function(a, signed=FALSE) {
        suffix <- if(a$valid < a$total) sprintf(' · %d/%d종목',a$valid,a$total) else ''
        paste0(board_amount_text(a$value,signed),suffix)
      }
       delta <- x$change$value
      # Previous close portfolio value = current value - daily price-driven change.
      # If some prices are missing, do not present a misleading portfolio rate.
      prev_value <- x$nav$value - delta
      complete <- x$nav$total > 0L && x$nav$valid == x$nav$total &&
                  x$change$valid == x$change$total && x$nav$total == x$change$total
      daily_pct <- if(complete && is.finite(prev_value) && prev_value > 0 && is.finite(delta))
        100 * delta / prev_value else NA_real_
      pct_label <- if(is.finite(daily_pct))
        sprintf(' (%+.2f%%)',daily_pct) else ' (등락률 산출 불가)'
      # Show failed security directly in the headline; separate renderUI diagnostic can be hidden by layout/CSS.
      if (!is.finite(daily_pct)) {
        zz <- pms_board_active_valuation()
        if (nrow(zz)) {
          readnum <- function(nm) {
            if (!nm %in% names(zz)) return(rep(NA_real_, nrow(zz)))
            suppressWarnings(as.numeric(gsub(',', '', as.character(zz[[nm]]), fixed=TRUE)))
          }
          bad_rows <- which(!is.finite(readnum('총평가액(원)')) |
                            !is.finite(readnum('평가증감(원)')))
          if (length(bad_rows)) {
            codes <- if ('종목코드' %in% names(zz)) as.character(zz[['종목코드']][bad_rows]) else rep('종목코드 없음',length(bad_rows))
            qty <- readnum('보유수량')[bad_rows]
            desc <- paste0(codes, '[', ifelse(is.finite(qty),format(qty,trim=TRUE,scientific=FALSE),'?'),'주]')
            pct_label <- paste0(pct_label,' · 미산출: ',paste(head(desc,5L),collapse=', '),
              if(length(desc)>5L) sprintf(' 외 %d종목',length(desc)-5L) else '')
          }
        }
      }
      delta_cls <- if(!is.finite(delta)) 'pms-move-flat' else if(delta > 0) 'pms-move-up' else if(delta < 0) 'pms-move-down' else 'pms-move-flat'
      tags$span(style='display:inline-flex;flex-wrap:wrap;align-items:center;gap:12px;',
        tags$span(tags$b('총평가액 '),val(x$nav)),
        tags$span(class=delta_cls,style='font-weight:750;',
          tags$b('전일대비 '),val(x$change,TRUE),pct_label),
        pms_board_fx_header())
    },error=function(e) {
      tags$span(style='font-size:12px;color:#b35b13;',
        paste0('합계 계산 확인 필요: ',conditionMessage(e)))
    })
  })
  outputOptions(output,'board_header_totals',suspendWhenHidden=FALSE)
  # Aggregate actual purchase costs for each security (across broker accounts).
  # A missing cost must not be treated as a zero cost or an invented profit.
  pms_board_purchase_profit <- function(z) {
    new_col <- '매수금대비손익(원)'
    return_col <- '매수금대비수익률(%)'
    z[[new_col]] <- rep(NA_real_,nrow(z))
    z[[return_col]] <- rep(NA_real_,nrow(z))
    if (!nrow(z) || !all(c('국가','종목코드','총평가액(원)') %in% names(z))) return(z)
    h <- DBI::dbGetQuery(con, paste0(
      'SELECT country,symbol,currency,quantity,avg_price,avg_fx_rate,cost_krw ',
      'FROM holdings WHERE quantity > 0'))
    if (!nrow(h)) return(z)
    numeric_col <- function(x) suppressWarnings(as.numeric(gsub(',', '',as.character(x),fixed=TRUE)))
    h$quantity <- numeric_col(h$quantity)
    h$avg_price <- numeric_col(h$avg_price)
    h$avg_fx_rate <- numeric_col(h$avg_fx_rate)
    h$cost_krw <- numeric_col(h$cost_krw)
    # Persisted KRW total cost takes precedence. Otherwise use acquisition-price
    # times quantity, converting USD purchases at the original acquisition FX.
    explicit <- is.finite(h$cost_krw) & h$cost_krw >= 0
    fx <- ifelse(toupper(h$currency)=='KRW', 1, h$avg_fx_rate)
    fallback <- h$quantity * h$avg_price * fx
    good_fallback <- is.finite(fallback) & fallback >= 0 &
      is.finite(h$avg_price) & h$avg_price > 0 & is.finite(fx) & fx > 0
    h$computed_cost <- ifelse(explicit,h$cost_krw,ifelse(good_fallback,fallback,NA_real_))
    key <- paste(toupper(trimws(h$country)),toupper(trimws(h$symbol)),sep='::')
    z_key <- paste(toupper(trimws(z[['국가']])),toupper(trimws(z[['종목코드']])),sep='::')
    by_key <- split(seq_len(nrow(h)),key)
    nav <- numeric_col(z[['총평가액(원)']])
    for (i in seq_len(nrow(z))) {
      idx <- by_key[[z_key[i]]]
      if (is.null(idx) || !length(idx) || !is.finite(nav[i])) next
      costs <- h$computed_cost[idx]
      if (any(!is.finite(costs))) next
      total_cost <- sum(costs)
      z[[new_col]][i] <- nav[i] - total_cost
      if (is.finite(total_cost) && total_cost > 0) {
        z[[return_col]][i] <- (nav[i] / total_cost - 1) * 100
      }
    }
    # Position exactly after the existing total market valuation column.
    cols <- names(z)
    cols <- setdiff(cols,c(new_col,return_col))
    at <- match('총평가액(원)',cols)
    z[,append(cols,c(new_col,return_col),after=at),drop=FALSE]
  }
  output$board_value_change_table <- DT::renderDT({
    z <- pms_board_purchase_profit(pms_board_active_valuation())
    # Determine monetary and fractional columns from their names, never fixed
    # numeric positions: the purchase-profit column is inserted mid-table.
    formatted_cols <- intersect(c('보유수량','전일종가','현재가','등락률(%)',
      '총평가액(원)','매수금대비손익(원)','매수금대비수익률(%)','전체비중(%)','평가증감(원)'), names(z))
    js_renderer <- DT::JS("function(data,type,row,meta) {
      if(type !== 'display') return data;
      if(data === null || data === undefined || data === '' || !isFinite(Number(data))) return '-';
      var title=String(meta.settings.aoColumns[meta.col].sTitle||'').replace(/<[^>]+>/g,'');
      var value=Number(data);
      var digits = /수량/.test(title) ? 6 : (/등락률|비중|수익률/.test(title) ? 2 :
        (/종가|현재가/.test(title) ? (String(row[0])==='US' ? 2 : 0) : 0));
      var v = value.toLocaleString('en-US',{minimumFractionDigits:digits===6?0:digits,maximumFractionDigits:digits});
      if(/등락률|평가증감|매수금대비손익|매수금대비수익률/.test(title) && value > 0) v = '+' + v;
      // Excel-style in-cell data bar: 100% means full cell width.
      // Display-only HTML: original numeric value remains available for DT sorting/search.
      if(title.indexOf('전체비중') >= 0) {
        var pct=Math.max(0,Math.min(100,value));
        return '<div style=\"position:relative;min-width:94px;height:20px;line-height:20px;text-align:right;padding-right:5px;overflow:hidden;border-radius:2px;\">'
          + '<span aria-hidden=\"true\" style=\"position:absolute;left:0;top:2px;bottom:2px;width:'+pct+'%;background:rgba(56,142,213,.42);border-radius:2px;\"></span>'
          + '<span style=\"position:relative;font-weight:600;\">'+v+'%</span></div>';
      }
      return v;
    }")
    numeric_targets <- as.integer(match(formatted_cols,names(z))-1L)
    nav_idx <- match('총평가액(원)',names(z))-1L
    # Animate only actual price-driven numeric changes; sorting/search redraws are inert.
    roll_columns <- intersect(c('현재가','등락률(%)','총평가액(원)','평가증감(원)'), names(z))
    roll_indices <- as.integer(match(roll_columns,names(z))-1L)
    roll_indices_js <- paste(roll_indices,collapse=',')
    roll_callback <- DT::JS(paste0("function(settings) {
      var api = this.api();
      var tableId = 'pms_board_roll_v627';
      var state = window[tableId] || (window[tableId] = {values:{},ready:false});
      var columns = [",roll_indices_js,"];
      api.rows({page:'current'}).every(function(){
        var row = this.data(), node = this.node();
        if(!row || !node) return;
        var key = String(row[0]) + '|' + String(row[1]);
        columns.forEach(function(col){
          var td = node.cells[col];
          if(!td) return;
          var value = Number(row[col]);
          if(!isFinite(value)) return;
          var k = key + '|' + col;
          var previous = state.values[k];
          var label = td.textContent.trim();
          if(state.ready && previous && isFinite(previous.value) && previous.value !== value){
            var oldLabel = previous.label;
            var outer = document.createElement('span'); outer.className='pms-price-roll';
            var track = document.createElement('span'); track.className='pms-price-roll-track';
            var a=document.createElement('span'), b=document.createElement('span');
            a.textContent=oldLabel; b.textContent=label;
            track.appendChild(a);track.appendChild(b);outer.appendChild(track);
            td.textContent='';td.appendChild(outer);
          }
          state.values[k]={value:value,label:label};
        });
      });
      state.ready=true;
    }"))
    tab <- DT::datatable(z,rownames=FALSE,
      options=list(pageLength=100,lengthMenu=c(25,50,100,200),scrollX=TRUE,autoWidth=TRUE,
        order=list(list(nav_idx,'desc')),dom='ltip',
        columnDefs=list(list(targets=numeric_targets,className='dt-right',render=js_renderer)),
        drawCallback=roll_callback),
      filter='top',escape=TRUE)
    if(nrow(z)) {
      tab <- DT::formatStyle(tab,columns='총평가액(원)',fontWeight='bold')
      tab <- DT::formatStyle(tab, columns=c('매수금대비손익(원)','매수금대비수익률(%)'),
        valueColumns='매수금대비손익(원)',
        color=DT::styleInterval(c(-1e-10,1e-10),c('#267dc5','#526171','#d43838')),
        fontWeight='bold')
      tab <- DT::formatStyle(tab,
        columns=c('현재가','등락률(%)','평가증감(원)'),
        valueColumns='등락률(%)',
        color=DT::styleInterval(c(-1e-10,1e-10),
          c('#267dc5','#526171','#d43838')),
        fontWeight='bold')
    }
    tab
  },server=FALSE)
  output$pms_board_auto_detail <- renderText({
    invalidateLater(10000,session)
    if(!isTRUE(input$auto_prices_10m)) return('10분 자동조회: 꺼짐')
    last <- auto_last_at()
    next_at <- if(inherits(last,'POSIXt')) format(last+600,'%H:%M:%S',tz='Asia/Seoul') else '미정'
    note <- auto_last_note()
    # State is already displayed in this line; do not repeat '자동 조회 켜짐'.
    if(grepl('^자동 조회 켜짐',note)) note <- '첫 자동조회 대기 중'
    if(grepl('^자동 조회 꺼짐',note)) note <- '대기 중'
    paste0('10분 자동조회: 켜짐 · ',note,' · 다음 예정 ',next_at)
  })
  output$board_auto_status <- renderText({
    if(isTRUE(input$auto_prices_10m)) '10분 자동조회: 켜짐' else '10분 자동조회: 꺼짐'
  })
  pms_snapshot_init(con)
  pms_trades_seed(con)
  pms_trajectory_seed(con)  # 과거 CSV 520행을 최초 1회 SQLite에 이관 (기존 기록 보존)
  # Backfill only the two home-overlay series, never the entire holdings universe.
  home_overlay_fetch_note <- reactiveVal('ETF·환율 과거시세 확인 중')
  output$home_overlay_status <- renderText(home_overlay_fetch_note())
  observeEvent(input$home_etf_retry, {
    home_overlay_fetch_note('ETF 전체기간 재수집 중...')
    msg <- tryCatch(pms_home_market_backfill(con,force=TRUE,only_etf=TRUE),error=function(e) paste('ETF 재수집 실패:',conditionMessage(e)))
    home_overlay_fetch_note(as.character(msg))
    market_refresh(isolate(market_refresh())+1L)
  })
  observeEvent(TRUE, {
    msg <- tryCatch(pms_home_market_backfill(con), error=function(e) paste('과거 시세 보충 실패:',conditionMessage(e)))
    home_overlay_fetch_note(as.character(msg))
    market_refresh(isolate(market_refresh())+1L)
  }, once=TRUE)
  trade_refreshed <- reactiveVal(0L)
  trade_records <- reactive({ trade_refreshed(); pms_trades_changes(con) })
  observe({
    x <- trade_records()
    months <- if(nrow(x)) sort(unique(substr(x$date,1,7)),decreasing=TRUE) else format(Sys.Date(),'%Y-%m')
    choices <- c('최근 1년 · 월별 내역'='year',setNames(months,months))
    current <- isolate(input$trade_month)
    chosen <- if(!is.null(current) && current %in% unname(choices)) current else months[1]
    shiny::updateSelectInput(session,'trade_month',choices=choices,selected=chosen)
  })
  trade_selected <- reactive({
    x <- trade_records()
    if(!nrow(x) || is.null(input$trade_month)) return(data.frame())
    if(identical(input$trade_month,'year')) {
      last <- Sys.Date()
      return(x[as.Date(x$date) >= last-365L & as.Date(x$date)<=last,,drop=FALSE])
    }
    x[substr(x$date,1,7)==input$trade_month,,drop=FALSE]
  })
  output$trade_stats <- renderText({
    x <- trade_selected()
    buys <- if(nrow(x)) sum(x$kind=='증가(매수 추정)') else 0L
    sells <- if(nrow(x)) sum(x$kind=='감소(매도 추정)') else 0L
    conversions <- if(nrow(x)) sum(grepl('온주전환',x$kind)) else 0L
    sprintf('보유수량 변화 %d건 · 매수 추정 %d건 · 매도 추정 %d건 · 온주전환 추정 %d건 (체결내역 아님)',nrow(x),buys,sells,conversions)
  })
  output$trade_calendar <- renderUI({
    req(input$trade_month)
    x <- trade_selected()
    if(identical(input$trade_month,'year')) {
      if(!nrow(x)) return(tags$div(class='alert alert-info','최근 1년간 비교 가능한 수량 변화가 없습니다.'))
      months <- sort(unique(substr(x$date,1,7)),decreasing=TRUE)
      return(tags$div(lapply(months,function(mm) {
        z <- x[substr(x$date,1,7)==mm,,drop=FALSE]
        tags$div(style='margin-bottom:24px;',
          tags$h4(mm,style='padding:9px 12px;background:#edf3f9;border-left:4px solid #4385c7;'),
          pms_trades_calendar(z,mm))
      })))
    }
    pms_trades_calendar(x,input$trade_month)
  })
  output$trade_changes_table <- DT::renderDT({
    x <- trade_selected()
    if(!nrow(x)) return(DT::datatable(data.frame(안내='선택 기간에 감지된 보유수량 변화가 없습니다.'),rownames=FALSE))
    y <- data.frame(날짜=x$date,국가=x$country,종목=x$name,코드=x$symbol,증권사=x$broker,
      이전수량=x$previous_quantity,현재수량=x$quantity,증감수량=x$change,구분=x$kind,
      기록간격_일=x$gap_days,check.names=FALSE)
    DT::datatable(y,rownames=FALSE,options=list(pageLength=20,scrollX=TRUE))
  })
  observeEvent(input$pms_trade_day,{
    x <- trade_records()
    z <- x[x$date==input$pms_trade_day,,drop=FALSE]
    if(!nrow(z)) return()
    # Modal: use a native HTML table. DT inside a modal can fail to initialize.
    rows <- lapply(seq_len(nrow(z)),function(i) {
      r <- z[i,,drop=FALSE]
      tags$tr(
        tags$td(as.character(r$country)),tags$td(as.character(r$name)),
        tags$td(as.character(r$broker)),
        tags$td(format(r$previous_quantity,trim=TRUE,scientific=FALSE)),
        tags$td(format(r$quantity,trim=TRUE,scientific=FALSE)),
        tags$td(format(r$change,trim=TRUE,scientific=FALSE)),
        tags$td(as.character(r$kind)))
    })
    showModal(modalDialog(title=paste('보유수량 변화 ·',input$pms_trade_day),
      tags$p('실제 체결내역이 아닌 날짜 사이 보유수량 변동 추정입니다. 소수점→온주 전환은 같은 종목·증권사에서 증가분과 감소분이 일치하는 경우만 표시합니다.'),
      tags$div(style='overflow-x:auto;max-height:65vh;overflow-y:auto;',
        tags$table(class='table table-striped table-condensed',
          tags$thead(tags$tr(lapply(c('국가','종목','증권사','이전수량','현재수량','증감수량','구분'),tags$th))),
          tags$tbody(rows))),
      size='l',easyClose=TRUE,footer=modalButton('닫기')))
  })
  snapshot_refresh <- reactiveVal(0L)
  snapshot_save_status <- reactiveVal('오늘 기록 대기 중')
  output$snapshot_save_status <- renderText(snapshot_save_status())
  output$snapshot_history_table <- DT::renderDT({
    snapshot_refresh()
    x <- pms_snapshot_history(con)
    DT::datatable(x,rownames=FALSE,options=list(pageLength=15,scrollX=TRUE,order=list(list(0,'desc'))))
  })
  # v0.6.28: snapshot tab treemaps; read-only (current holdings and saved prices).
  snapshot_treemap_data <- reactive({
    snapshot_refresh()
    refresh()
    market_refresh()
    d <- pms_valuation(con,include_zero=FALSE)
    if(!nrow(d)) return(data.frame())
    need <- c('symbol','name','asset_class','market_value_krw')
    if(!all(need %in% names(d))) stop('트리맵에 필요한 평가액 필드가 없습니다: ',paste(setdiff(need,names(d)),collapse=', '))
    vals <- suppressWarnings(as.numeric(d$market_value_krw))
    ok <- is.finite(vals) & vals > 0
    if(!any(ok)) return(data.frame())
    z <- d[ok,,drop=FALSE]
    z$value <- vals[ok]
    z$symbol <- as.character(z$symbol)
    z$name <- as.character(z$name)
    z$asset_class <- as.character(z$asset_class)
    z$symbol[is.na(z$symbol) | !nzchar(z$symbol)] <- '미확인 종목'
    z$name[is.na(z$name) | !nzchar(z$name)] <- z$symbol[is.na(z$name) | !nzchar(z$name)]
    z$asset_class[is.na(z$asset_class) | !nzchar(z$asset_class)] <- 'OTHER'
    z
  })
  snapshot_treemap_draw <- function(z,group_col) {
    if(!requireNamespace('treemap',quietly=TRUE)) {
      plot.new();text(.5,.5,'트리맵 패키지가 필요합니다: install.packages("treemap")',cex=1.1)
      return(invisible(NULL))
    }
    if(!nrow(z)) {
      plot.new();text(.5,.5,'표시할 평가 가능 보유종목이 없습니다.',cex=1.1)
      return(invisible(NULL))
    }
    sums <- stats::aggregate(z$value,list(label=z[[group_col]]),sum)
    names(sums)[2] <- 'value'
    sums <- sums[is.finite(sums$value) & sums$value>0,,drop=FALSE]
    total <- sum(sums$value)
    sums$label <- paste0(sums$label,'  ',sprintf('%.1f%%',100*sums$value/total))
    treemap::treemap(sums,index='label',vSize='value',type='index',
      fontsize.labels=if(nrow(sums)>30L) 8 else 12,
      fontcolor.labels='white',fontface.labels=2,
      border.col='white',border.lwds=2,inflate.labels=FALSE,
      bg.labels=0,align.labels=list(c('center','center')),
      title='',palette='Set3')
  }
  output$snapshot_symbol_treemap <- renderPlot({
    z <- snapshot_treemap_data()
    if(nrow(z)) {
      z$label <- paste0(z$name,' (',z$symbol,')')
      snapshot_treemap_draw(z,'label')
    } else snapshot_treemap_draw(z,'symbol')
  },res=110)
  output$snapshot_asset_treemap <- renderPlot({
    z <- snapshot_treemap_data()
    if(nrow(z)) {
      labels <- c(STOCK='주식',BOND='채권',GOLD='금',CASH='현금성',OTHER='기타')
      z$asset_label <- unname(labels[z$asset_class])
      z$asset_label[is.na(z$asset_label)] <- '기타'
      snapshot_treemap_draw(z,'asset_label')
    } else snapshot_treemap_draw(z,'asset_class')
  },res=110)
  output$snapshot_treemap_note <- renderText({
    z <- snapshot_treemap_data()
    if(!nrow(z)) return('현재 저장된 시세로 평가 가능한 종목이 없습니다.')
    sprintf('현재 평가 가능 %d개 보유행 · 합계 %s원 · 종목별/자산군별 동일 기준 (미평가 종목 제외)',
      nrow(z),format(round(sum(z$value)),big.mark=',',scientific=FALSE,trim=TRUE))
  })
  collection_message <- reactiveVal('SQLite의 마지막 저장 시세로 자동 평가합니다. 새 시세가 필요할 때만 수집 버튼을 누르세요.')
  # 앱 시작 시 보유 종목 평가 및 배당 분석은 저장된 SQLite 자료만 읽는다.
  # 어떤 공급원에도 접속하지 않는다. 기존 10분 자동 시세수집 기능은 별도 옵션이다.
  output$cached_data_status <- renderText({
    market_refresh()
    dividend_refresh()
    refresh()
    n_prices <- tryCatch(DBI::dbGetQuery(con,'SELECT COUNT(*) AS n FROM prices')$n[1],error=function(e) 0L)
    n_dividends <- tryCatch(DBI::dbGetQuery(con,'SELECT COUNT(*) AS n FROM dividend_events')$n[1],error=function(e) 0L)
    n_fx <- tryCatch(DBI::dbGetQuery(con,"SELECT COUNT(*) AS n FROM fx_rates WHERE pair='USD/KRW'")$n[1],error=function(e) 0L)
    sprintf('SQLite 저장 데이터 자동 적용: 시세 %d종목 · 배당 이벤트 %d건 · 현재 USD/KRW %s (조회 API 호출 없음)',
            n_prices,n_dividends,if(n_fx>0L) '저장됨' else '없음')
  })

  diagnosis <- reactiveVal(NULL)
  output$pms_ssl_auto_retry_notice <- renderText({
    invalidateLater(10000,session)
    stamp <- getOption('pms_ssl_auto_retry_at',NULL)
    if(is.null(stamp)) 'SSL 자동검사: 기본 검증 사용 · 폐기 확인 오류인 경우에만 자동 재시도'
    else paste0('SSL 자동검사: 폐기 확인 오류 감지 → 해당 요청만 재시도 성공 (',stamp,') · 기본 체크는 해제 유지')
  })
  diagnosis_msg <- reactiveVal('연결 진단 버튼을 누르면 국내·미국·환율의 공급원을 개별 검사합니다.')
  output$diagnosis_notice <- renderText(diagnosis_msg())
  output$diagnosis_table <- renderDT({
    x <- diagnosis()
    if(is.null(x)) return(DT::datatable(data.frame(대기='진단 전'),rownames=FALSE,options=list(dom='t')))
    DT::datatable(x,rownames=FALSE,options=list(dom='t',scrollX=TRUE,pageLength=5,columnDefs=pms_money_defs(x)),
      escape=TRUE)
  })
  observeEvent(input$diagnose_sources, {
    # 진단과 실제 수집에서 동일한 SSL 설정을 사용
    Sys.setenv(PMS_SSL_NO_REVOKE=if(isTRUE(input$ssl_revocation_skip)) '1' else '0')
    diagnosis_msg('진단 진행 중 (공급원별 최대 8초)...')
    flow_update('running',1L,'공급원 연결 진단','4개 공급원을 순차 검사합니다. SSL 설정과 응답을 확인 중입니다.')
    x <- tryCatch(pms_diagnose_sources(isTRUE(input$ssl_revocation_skip),
      on_step=function(z,phase,result=NULL) {
        if(identical(phase,'start')) flow_update('running',1L,paste(z$provider,'연결 중'),paste(z$target,'HTTPS 연결 · SSL 검증 · HTTP 응답 대기'))
        else if(identical(phase,'done')) {
          ok <- identical(as.character(result$결과[1]),'성공')
          flow_update(if(ok) 'success' else 'failed',if(ok) 5L else as.integer(result$실패단계[1]),
            paste(z$provider,if(ok) '정상 수신' else '진단 실패'),
            if(ok) paste('HTTP',result$HTTP상태[1],'· 조회값',result$조회값[1]) else as.character(result$실패원인[1]))
        }
      }),error=function(e)e)
    if(inherits(x,'error')) { diagnosis_msg(paste('진단 실패:',conditionMessage(x))); flow_update('failed',1L,'진단 중단',conditionMessage(x)) } else {
      diag_view <- x[,setdiff(names(x),'실패단계'),drop=FALSE]
      names(diag_view)[names(diag_view)=='조회값'] <- '테스트 수신값(주가·환율)'
      diagnosis(diag_view)
      diagnosis_msg(sprintf('진단 완료: 성공 %d개 / 실패 %d개. 인증서 폐기 확인: %s',
         sum(x$결과=='성공'),sum(x$결과=='실패'),if(isTRUE(input$ssl_revocation_skip)) '생략' else '수행'))
    }
  })
  output$collect_notice <- renderText(collection_message())
  observeEvent(input$clear_collection_failures, {
    showModal(modalDialog(title='최근 실패 기록 삭제',
      '시세수집 실패 이력 전체를 삭제합니다. 정상 시세와 보유종목은 유지됩니다.',
      footer=tagList(modalButton('취소'),actionButton('confirm_clear_failures','실패 이력 삭제',class='btn-danger')),easyClose=TRUE))
  })
  observeEvent(input$confirm_clear_failures, {
    removeModal()
    n <- DBI::dbExecute(con,"DELETE FROM collection_log WHERE status='FAILED'")
    market_refresh(isolate(market_refresh())+1L)
    collection_message('실패 기록을 초기화했습니다. 기존 시세와 정상 수집 기록은 유지됩니다.')
    session$sendCustomMessage('pms_collection_reset',list(total=0L))
    showNotification(paste(n,'건의 실패 기록을 삭제했습니다. 저장된 가격에는 영향이 없습니다.'),type='message')
  })
  output$collection_failures <- renderDT({
    market_refresh()
    x <- DBI::dbGetQuery(con, "SELECT country AS 국가,symbol AS 종목코드,
      source AS 출처,detail AS 오류메시지,run_at AS 수집시각_UTC
      FROM collection_log WHERE status='FAILED'
      ORDER BY id DESC LIMIT 20")
    DT::datatable(x, rownames=FALSE, options=list(pageLength=10,scrollX=TRUE,
                                                searching=TRUE,ordering=FALSE))
  })
  valuation_data <- reactive({
    refresh()
    market_refresh()
    pms_valuation(con,include_zero=isTRUE(input$valuation_show_zero))
  })
  # Show explicitly dated quote freshness; do not infer listing status.
  # If quoted_at is unavailable, use 'unknown', never fetched_at as trade date.
  pms_price_freshness_ui <- function() {
    tryCatch({
      d <- pms_valuation(con,include_zero=FALSE)
      money <- function(x) if(is.finite(x)) paste0(format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),'원') else '미산출'
      if(!nrow(d)) return(tags$div(id='pms_price_freshness_direct',class='alert alert-info','보유 종목이 없습니다.'))
      unique_rows <- d[!duplicated(paste(d$country,d$symbol,sep='|')),,drop=FALSE]
      current_day <- as.Date(Sys.time(),tz='Asia/Seoul')
      get_day <- function(v) {
        if(is.na(v) || !nzchar(trimws(as.character(v)))) return(as.Date(NA))
        # QUANTMOD returns YYYY-MM-DD; Yahoo timestamps use UTC.
        x <- suppressWarnings(as.Date(substr(as.character(v),1L,10L)))
        x
      }
      days <- vapply(unique_rows$price_date,function(v) {
        dt <- get_day(v)
        if(is.na(dt)) return(NA_integer_)
        as.integer(current_day-dt)
      },integer(1))
      state <- ifelse(is.na(days),'기준일 미확인',
                 ifelse(days>10L,'오래된 시세 (10일 초과)',
                   ifelse(days>4L,'주의 (4일 초과)',
                     ifelse(days<0L,'시세 날짜 확인 필요','정상 범위'))))
      state[!is.finite(unique_rows$current_price)] <- '시세 누락'
      concern <- which(state!='정상 범위')
      counts <- sprintf('보유 고유종목 %d개 · 정상 %d개 · 확인필요 %d개',nrow(unique_rows),sum(state=='정상 범위'),length(concern))
      rows <- if(length(concern)) lapply(concern,function(i) {
        r <- unique_rows[i,,drop=FALSE]
        tags$tr(tags$td(r$name),tags$td(r$symbol),
          tags$td(if(is.na(r$price_date)) '미확인' else as.character(r$price_date)),
          tags$td(if(is.na(days[i])) '미확인' else paste0(days[i],'일')),
          tags$td(r$price_source),tags$td(state[i]),
          tags$td(money(r$market_value_krw),style='text-align:right;'))
      }) else list(tags$tr(tags$td(colspan='7','추가 확인이 필요한 종목이 없습니다.')))
      tags$div(id='pms_price_freshness_direct',
        tags$div(class=if(length(concern)) 'alert alert-warning' else 'alert alert-success',
          counts,' · 마지막 시세 기준일의 달력일 차이로 판정(휴장일 포함)'),
        tags$div(style='overflow-x:auto;',tags$table(class='table table-bordered table-condensed',
          tags$thead(tags$tr(lapply(c('종목명','종목코드','시세 기준일','경과일','출처','상태','현재 평가액'),tags$th))),
          tags$tbody(rows))),
        tags$small('거래정지·상장폐지 여부는 공시/증권사 내역으로 확인하세요. 오래된 종가도 기존 평가액 계산에는 유지됩니다.'))
    },error=function(e) tags$div(id='pms_price_freshness_direct',class='alert alert-danger',
                   paste('시세 신선도 검사 오류:',conditionMessage(e))))
  }
  observe({
    market_refresh()
    refresh()
    shiny::removeUI(selector='#pms_price_freshness_direct',multiple=TRUE,immediate=TRUE,session=session)
    shiny::insertUI(selector='#pms_price_freshness_anchor',where='afterEnd',
      ui=pms_price_freshness_ui(),immediate=TRUE,session=session)
  })
  # fetched_at is stored in UTC, whereas the displayed 'today' is Asia/Seoul.
  # This checks cached collection timestamps only; it does not call any provider.
  cached_quote_freshness <- reactive({
    market_refresh()
    refresh()
    d <- DBI::dbGetQuery(con, "SELECT h.country,h.symbol,p.fetched_at
       FROM (SELECT DISTINCT country,symbol FROM holdings WHERE quantity>0
             AND country IN ('KR','US')) h
       LEFT JOIN prices p ON p.country=h.country AND p.symbol=h.symbol")
    local_today <- format(Sys.time(),'%Y-%m-%d',tz='Asia/Seoul')
    is_today <- function(stamps) {
      vapply(stamps, function(x) {
        if(is.na(x) || !nzchar(x)) return(FALSE)
        utc_time <- as.POSIXct(x,format='%Y-%m-%d %H:%M:%S',tz='UTC')
        if(is.na(utc_time)) return(FALSE)
        identical(format(utc_time,'%Y-%m-%d',tz='Asia/Seoul'),local_today)
      },logical(1))
    }
    fx <- DBI::dbGetQuery(con,"SELECT fetched_at,quoted_at,rate FROM fx_rates WHERE pair='USD/KRW'")
    fx_today <- if(nrow(fx)) is_today(fx$fetched_at)[1] else FALSE
    list(today=local_today, total=nrow(d), today_count=sum(is_today(d$fetched_at)),
         cached_count=sum(!is.na(d$fetched_at) & nzchar(d$fetched_at)),
         fx_present=nrow(fx)>0, fx_today=fx_today,
         fx_time=if(nrow(fx)) as.character(fx$fetched_at[1]) else '')
  })
  output$cached_quote_status <- renderText({
    x <- cached_quote_freshness()
    sprintf('오늘(%s) 수집된 종목 %d / %d개 · 저장시세 보유 %d개 · USD/KRW %s',
      x$today,x$today_count,x$total,x$cached_count,
      if(!x$fx_present) '저장 없음' else if(x$fx_today) '오늘 수집' else '과거 수집')
  })
  observeEvent(input$refresh_valuation, {
    saved_eval_ack(paste('시세 버튼 요청을 Shiny 서버에서 수신했습니다. SQLite 계산 중…',format(Sys.time(),'%H:%M:%S')))
    tryCatch({
      market_refresh(isolate(market_refresh()) + 1L)
      d <- pms_valuation(con,include_zero=FALSE)
      active <- is.finite(as.numeric(d$quantity)) & as.numeric(d$quantity)>0
      values <- suppressWarnings(as.numeric(d$market_value_krw))
      profits <- suppressWarnings(as.numeric(d$profit_krw))
      valid <- active & is.finite(values)
      validp <- valid & is.finite(profits)
      total <- if(any(valid)) sum(values[valid]) else NA_real_
      pnl <- if(any(validp)) sum(profits[validp]) else NA_real_
      fm <- function(n) if(is.finite(n)) paste0(format(round(n),big.mark=',',scientific=FALSE,trim=TRUE),'원') else '미산출'
      msg <- sprintf('시세 재평가 완료: %d/%d건 · 전체 평가액 %s · 전체 손익 %s · %s',
        sum(valid),sum(active),fm(total),fm(pnl),format(Sys.time(),'%H:%M:%S'))
      collection_message(msg)
      saved_eval_ack(msg)
    },error=function(e) {
      msg <- paste('저장된 시세 평가 오류:',conditionMessage(e))
      collection_message(msg)
      saved_eval_ack(msg,TRUE)
    })
  })
  run_price_collection <- function() {
    # Applies only to this R process, and only for the explicit user choice.
    Sys.setenv(PMS_SSL_NO_REVOKE=if(isTRUE(input$ssl_revocation_skip)) '1' else '0')

    collection_message('수집 진행 중... 아래 표에 종목별 결과가 기록됩니다.')
    flow_update('running',1L,'시세 수집 시작','보유종목 코드별로 공급원에 연결하고 있습니다.')
    targets_n <- DBI::dbGetQuery(con,"SELECT COUNT(*) AS n FROM (SELECT country,symbol FROM holdings WHERE quantity>0 AND country IN ('KR','US') GROUP BY country,symbol)")$n[1]
    session$sendCustomMessage('pms_collection_reset',list(total=as.integer(targets_n)))
    session$sendCustomMessage('pms_source_visual',list(kind='prices',event='reset'))
    ok_so_far <- 0L
    failed_so_far <- 0L
    tryCatch({
      withProgress(message='국내·미국 네이버 / Yahoo 대체 시세 수집',value=0, {
        result <- pms_collect_prices(
          con,
          allow_yahoo=isTRUE(input$enable_yahoo_fallback),
          on_result=function(rec,i,n) {
            success <- identical(as.character(rec$status[1]),'OK')
            if(success) ok_so_far <<- ok_so_far+1L else failed_so_far <<- failed_so_far+1L
            session$sendCustomMessage('pms_source_visual',list(kind='prices',event='result',
              index=i,total=n,symbol=as.character(rec$symbol[1]),
              ok=success,source=as.character(rec$source[1])))
            flow_update(if(success) 'success' else 'failed',if(success) 5L else 3L,
              paste(rec$symbol[1],if(success) '시세 저장 완료' else '수집 실패'),
              if(success) paste(rec$source[1],'에서',format(rec$price[1],big.mark=',',scientific=FALSE),'수신') else as.character(rec$message[1]))
            price <- rec$price[1]
            session$sendCustomMessage('pms_collection_row',list(
              index=as.integer(i),total=as.integer(n),
              country=as.character(rec$country[1]),symbol=as.character(rec$symbol[1]),
              price=if(is.finite(price)) format(price,big.mark=',',scientific=FALSE,trim=TRUE) else '-',
              source=as.character(rec$source[1]),status=as.character(rec$status[1]),
              message=substr(as.character(rec$message[1]),1,350),
              ok=ok_so_far,failed=failed_so_far
            ))
          },
          progress=function(i,n,label) {
            incProgress(1/max(n,1),detail=paste(i,'/',n,label))
          }
        )
        ok <- sum(result$status=='OK')
        failed <- sum(result$status=='FAILED')
        session$sendCustomMessage('pms_source_visual',list(kind='prices',event='done'))
        collection_message(sprintf('수집 완료: 성공 %d종목, 실패 %d종목. %s',ok,failed,
          if(failed) paste('최근 실패:',paste(head(unique(result$message[result$status=='FAILED']),3),collapse=' / ')) else '모든 종목 조회 성공'))
      })
      market_refresh(isolate(market_refresh())+1L)
    },error=function(e) {
      collection_message(paste('시세수집 오류:',conditionMessage(e)))
      flow_update('failed',3L,'수집 중단',conditionMessage(e))
      market_refresh(isolate(market_refresh())+1L)
      showNotification(conditionMessage(e),type='error',duration=NULL)
    })
    }
  # App startup and successful market refresh: upsert today's archive from stored quotes.
  # An initial no-price situation is recorded as PARTIAL, never silently treated as zero NAV.
  trajectory_all <- reactive({
    snapshot_refresh()
    pms_trajectory_read(con)
  })
  risk_data <- reactive({
    x <- trajectory_all()
    p <- input$risk_period
    if(!is.null(p) && p!='all' && nrow(x)) x <- x[x$date >= max(x$date)-as.integer(p),,drop=FALSE]
    pms_risk_compute(x,confidence=as.numeric(if(is.null(input$risk_confidence)) .95 else input$risk_confidence))
  })
  output$risk_cards <- renderUI({ pms_risk_cards(risk_data()) })
  output$risk_volatility_explanation <- renderText({
    d<-risk_data(); if(!isTRUE(d$ok))return('산출 불가')
    paste0(d$level,' (EWMA/역사적 표준편차 = ',pms_risk_num(d$volratio),'배)')
  })
  output$risk_explanations <- renderUI({ pms_risk_guide() })
  output$risk_nav_chart <- renderPlot({
    d <- risk_data(); if(!isTRUE(d$ok)) return(invisible(NULL))
    dates <- d$z$date[is.finite(d$z$ret)]
    plot(dates,d$curve,type='l',lwd=2,col='#2677bd',xlab='날짜',ylab='보정 추정 누적지수 (초기 1)',main='현금흐름 보정 추정 누적성과')
    grid(col='#e3e9ef')
  })
  output$risk_dd_chart <- renderPlot({
    d <- risk_data();if(!isTRUE(d$ok)) return(invisible(NULL))
    dates <- d$z$date[is.finite(d$z$ret)]
    plot(dates,100*d$dd,type='l',lwd=2,col='#c44847',xlab='날짜',ylab='낙폭 (%)',main='추정 운용수익률 기준 Drawdown')
    abline(h=0,col='#8999ac',lty=2);grid(col='#e3e9ef')
  })
  output$risk_stats_table <- DT::renderDT({
    d<-risk_data()
    if(!isTRUE(d$ok)) return(DT::datatable(data.frame(상태=d$message),options=list(dom='t')))
    z<-data.frame(지표=c('추정 연환산 수익률','연환산 변동성','MDD','현재 DD','Sharpe','Sortino','Calmar','1일 VaR','1일 CVaR','VaR 원화 환산','CVaR 원화 환산','최장 미회복 관측일','최근 미회복 관측일','최고 관측 수익률','최저 관측 수익률','변동성 경보'),
      값=c(pms_risk_pct(d$annual),pms_risk_pct(d$annvol),pms_risk_pct(d$mdd),pms_risk_pct(d$currentdd),pms_risk_num(d$sharpe),pms_risk_num(d$sortino),pms_risk_num(d$calmar),pms_risk_pct(d$var),pms_risk_pct(d$cvar),pms_risk_money(d$varwon),pms_risk_money(d$cvarwon),as.character(d$maxduration),as.character(d$currentduration),pms_risk_pct(d$best),pms_risk_pct(d$worst),d$level))
    DT::datatable(z,rownames=FALSE,options=list(dom='tip',pageLength=20,scrollX=TRUE))
  })
  risk_mc_result <- reactiveVal(NULL)
  observeEvent(input$risk_run_mc,{
    d <- risk_data()
    if(!isTRUE(d$ok)) {showNotification('유효한 수익률 자료가 부족합니다.',type='warning');return()}
    ans <- tryCatch(shiny::withProgress(message='몬테카를로 시뮬레이션 실행',value=0.3,{
      pms_risk_mc(d,years=input$risk_mc_years,monthly=input$risk_mc_monthly,
        mode=input$risk_mc_mode,annual_withdraw=input$risk_mc_withdraw)
    }),error=function(e)e)
    if(inherits(ans,'error')) showNotification(conditionMessage(ans),type='error') else risk_mc_result(ans)
  })
  output$risk_mc_summary <- renderUI({
    d<-risk_mc_result()
    if(is.null(d)) return(tags$p(class='text-muted','위의 실행 버튼을 누르면 시나리오 결과를 계산합니다.'))
    tags$div(class='well',
      tags$h4(sprintf('%s · %d년 · %d개 경로',if(d$mode=='save') '적립기' else '인출기',d$years,d$n)),
      tags$p('종료 자산 하위 10%: ',tags$b(pms_risk_money(d$q[1]))),
      tags$p('종료 자산 중앙값: ',tags$b(pms_risk_money(d$q[2]))),
      tags$p('종료 자산 상위 10%: ',tags$b(pms_risk_money(d$q[3]))),
      tags$p('시뮬레이션 MDD 중앙값: ',tags$b(pms_risk_pct(d$mddq[2]))),
      if(d$mode=='withdraw') tags$p('자산 소진 비율: ',tags$b(sprintf('%.1f%%',100*d$ruin))),
      tags$p(class='text-muted','과거 변동성을 정규분포로 근사한 모의실험으로 실제 결과를 보장하지 않습니다.'))
  })
  trajectory_filtered <- reactive({
    x <- trajectory_all()
    p <- input$trajectory_period
    if(!is.null(p) && p != 'all' && nrow(x)) {
      x <- x[x$date >= max(x$date) - as.integer(p),,drop=FALSE]
    }
    x
  })
  output$trajectory_source_note <- renderText({
    x <- trajectory_all()
    sprintf('기록 %d일 · CSV 이관 %d일 · 앱 스냅샷 %d일 · SQLite 저장',
      nrow(x),sum(x$origin=='IMPORTED'),sum(x$origin!='IMPORTED'))
  })
  output$trajectory_dashboard <- renderUI({
    x <- trajectory_filtered()
    pms_trajectory_panel(x,compact=FALSE)
  })
  # v0.6.29: descriptive least-squares trend (calendar-day slope).
  trajectory_lm_points <- function(col) {
    z <- trajectory_filtered()
    if(!nrow(z) || !all(c('date',col) %in% names(z))) return(data.frame())
    d <- as.Date(z$date)
    y <- suppressWarnings(as.numeric(z[[col]]))
    ok <- !is.na(d) & is.finite(y)
    z <- data.frame(date=d[ok], value=y[ok])
    z <- z[order(z$date),,drop=FALSE]
    z
  }
  trajectory_lm_fit <- function(col) {
    z <- trajectory_lm_points(col)
    if(nrow(z)<2 || length(unique(z$date))<2) return(NULL)
    z$day <- as.numeric(z$date - min(z$date))
    model <- stats::lm(value ~ day,data=z)
    list(data=z,fit=model,slope=unname(stats::coef(model)[2]))
  }
  trajectory_trend_plot <- function(col,label) {
    fit <- trajectory_lm_fit(col)
    if(is.null(fit)) {plot.new();text(.5,.5,'추세선 계산에 필요한 날짜별 자료가 부족합니다.');return(invisible(NULL))}
    z <- fit$data
    plot(z$date,z$value,type='l',col='#2778bb',lwd=2,xlab='날짜',ylab='원',main=label)
    lines(z$date,as.numeric(stats::predict(fit$fit,newdata=z)),col='#db6d22',lty=2,lwd=2)
    legend('topleft',legend=c('실제 기록','선형회귀 추세선'),col=c('#2778bb','#db6d22'),lty=c(1,2),bty='n',cex=.85)
    grid(col='#e3e9ef')
  }
  trajectory_slope_text <- function(col) {
    fit <- trajectory_lm_fit(col)
    if(is.null(fit)) return('일평균 증가분: 산출 불가 (서로 다른 날짜 2개 이상 필요)')
    x <- fit$slope
    paste0('일평균 증가분 (기울기): ',if(x>0) '+' else '',format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),' 원/일')
  }
  # v0.6.32: render both series in one plot with a true independent right-hand scale.
  output$trajectory_dual_axis <- renderPlot({
    z <- trajectory_filtered()
    needed <- c('date','value','profit')
    if(!nrow(z) || !all(needed %in% names(z))) {
      plot.new();text(.5,.5,'평가액·평가손익 기록이 없습니다.');return(invisible(NULL))
    }
    d <- as.Date(z$date)
    v <- suppressWarnings(as.numeric(z$value)) / 1e8
    p <- suppressWarnings(as.numeric(z$profit)) / 1e8
    ok <- !is.na(d) & is.finite(v) & is.finite(p)
    if(!any(ok)) {
      plot.new();text(.5,.5,'양쪽 축에 표시할 유효한 기록이 없습니다.');return(invisible(NULL))
    }
    dat <- data.frame(date=d[ok],value=v[ok],profit=p[ok])
    dat <- dat[order(dat$date),,drop=FALSE]
    if(anyDuplicated(dat$date)) dat <- dat[!duplicated(dat$date,fromLast=TRUE),,drop=FALSE]
    mkfit <- function(y) {
      if(nrow(dat)<2L || length(unique(dat$date))<2L) return(rep(NA_real_,nrow(dat)))
      as.numeric(stats::fitted(stats::lm(y ~ as.numeric(dat$date-min(dat$date)))))
    }
    vfit <- mkfit(dat$value)
    pfit <- mkfit(dat$profit)
    padded <- function(x) {
      r <- range(x[is.finite(x)])
      span <- diff(r)
      if(span < 1e-9) span <- max(abs(r[1])*.03,.1)
      r + c(-.08,.08)*span
    }
    yl <- padded(c(dat$value,vfit))
    yr <- padded(c(dat$profit,pfit))
    # Transform profit to the left-axis coordinate space only for plotting.
    # Values shown on the right axis are always original profit units (100M KRW).
    toleft <- function(y) yl[1] + (y-yr[1])/diff(yr)*diff(yl)
    op <- par(mar=c(5.4,5.5,4.1,5.8)+.1)
    on.exit(par(op),add=TRUE)
    plot(dat$date,dat$value,type='n',ylim=yl,xlab='날짜',ylab='총평가액 (억원)',
         main='총평가액 (왼쪽 축) / 평가손익 (오른쪽 축)',axes=FALSE)
    abline(h=pretty(yl),col='#e8edf2',lty=3)
    axis.Date(1,at=pretty(dat$date),format='%Y-%m-%d',las=2,cex.axis=.78)
    axis(2,at=pretty(yl),las=1,col.axis='#2876bb')
    ticks <- pretty(yr)
    ticks <- ticks[ticks >= yr[1] & ticks <= yr[2]]
    axis(4,at=toleft(ticks),labels=format(ticks,trim=TRUE,scientific=FALSE),
         las=1,col.axis='#d36e2d')
    mtext('평가손익 (억원)',side=4,line=3.8,col='#d36e2d')
    box()
    lines(dat$date,dat$value,col='#2876bb',lwd=2.4)
    lines(dat$date,toleft(dat$profit),col='#d36e2d',lwd=2.4)
    if(all(is.finite(vfit))) lines(dat$date,vfit,col='#2876bb',lty=2,lwd=2)
    if(all(is.finite(pfit))) lines(dat$date,toleft(pfit),col='#d36e2d',lty=2,lwd=2)
    legend('topleft',inset=.015,bty='n',cex=.84,ncol=2,
      legend=c('총평가액','총평가액 추세선','평가손익','평가손익 추세선'),
      col=c('#2876bb','#2876bb','#d36e2d','#d36e2d'),lty=c(1,2,1,2),lwd=2)
  })
  output$trajectory_value_lm <- renderPlot({trajectory_trend_plot('value','총평가액 추세')})
  output$trajectory_profit_lm <- renderPlot({trajectory_trend_plot('profit','평가손익 추세')})
  output$trajectory_value_slope <- renderText({trajectory_slope_text('value')})
  output$trajectory_profit_slope <- renderText({trajectory_slope_text('profit')})
  output$trajectory_aftertax_result <- renderUI({
    tryCatch({
      refresh();market_refresh()
      z <- pms_valuation(con,include_zero=FALSE)
      if(!nrow(z)) return(tags$div('평가 가능한 종목이 없습니다.'))
      must <- c('country','quantity','profit_krw')
      if(!all(must %in% names(z))) return(tags$div('계산 불가: 국가·보유수량·평가손익 필드 확인 필요'))
      active <- is.finite(as.numeric(z$quantity)) & as.numeric(z$quantity)>0
      vals <- suppressWarnings(as.numeric(z$profit_krw))
      if(any(active & !is.finite(vals))) return(tags$div(class='alert alert-warning','시세 또는 취득원가가 누락된 보유종목이 있어 전량 매도 추정치를 계산할 수 없습니다.'))
      country <- toupper(trimws(as.character(z$country)))
      local <- country %in% c('KR','KOR','KOREA','대한민국','한국','국내')
      overseas <- country %in% c('US','USA','UNITED STATES','미국','해외')
      if(any(active & !(local|overseas))) return(tags$div(class='alert alert-warning','국내/해외 구분이 확인되지 않은 종목이 있어 계산할 수 없습니다.'))
      # Use the explicit holdings tax category rather than inferring from names.
      # Ambiguous classifications are not silently taxed or treated as exempt.
      htax <- DBI::dbGetQuery(con, paste(
        'SELECT UPPER(TRIM(country)) AS country, UPPER(TRIM(symbol)) AS symbol,',
        'UPPER(TRIM(tax_type)) AS tax_type FROM holdings WHERE quantity > 0'))
      zkey <- paste(toupper(trimws(as.character(z$country))),toupper(trimws(as.character(z$symbol))),sep='::')
      hkey <- paste(htax$country,htax$symbol,sep='::')
      types <- vapply(zkey,function(key) {
        found <- unique(htax$tax_type[hkey==key])
        if(length(found)!=1L || is.na(found) || !nzchar(found)) return(NA_character_)
        found
      },character(1))
      if(any(active & local & is.na(types))) return(tags$div(class='alert alert-warning',
        '일부 국내 종목의 과세유형이 없거나 서로 다릅니다. 보유종목의 과세유형을 확인해 주세요.'))
      us_index <- active & local & types=='KR_US_INDEX_ETF'
      us_index[is.na(us_index)] <- FALSE
      domestic_profit <- sum(vals[active & local]); foreign_profit <- sum(vals[active & overseas])
      index_profit <- sum(vals[us_index])
      rate <- min(100,max(0,as.numeric(if(is.null(input$trajectory_domestic_tax)) 38 else input$trajectory_domestic_tax)))/100
      domestic_tax <- max(0,index_profit)*rate
      foreign_tax <- max(0,foreign_profit-20000000)*.22
      net <- domestic_profit+foreign_profit-domestic_tax-foreign_tax
      fmt <- function(x) paste0(format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),'원')
      tags$div(class='well',
        tags$div(tags$b('매도 후 추정 평가이익: '),tags$strong(style='font-size:19px;',fmt(net))),
        tags$div('세전 평가손익 ',fmt(domestic_profit+foreign_profit),' · 추정세금 ',fmt(domestic_tax+foreign_tax)),
        tags$small('국내 상장 미국지수 ETF 세금 ',fmt(domestic_tax),' · 그 외 국내종목 세금 0원(가정) · 해외 세금 ',fmt(foreign_tax),' (해외 공제 2,000만원 가정)'))
    },error=function(e)tags$div(class='alert alert-warning',paste('추정 계산 오류:',conditionMessage(e))))
  })
  output$trajectory_table <- DT::renderDT({
    x <- trajectory_filtered()
    if(!nrow(x)) return(DT::datatable(data.frame(안내='해당 기간 데이터 없음'),options=list(dom='t')))
    y <- data.frame(날짜=as.character(rev(x$date)),
                    총평가액_원=rev(x$value),평가손익_원=rev(x$profit),
                    평가액변동_원=rev(x$value_change),손익변동_원=rev(x$profit_change),
                    자료유형=ifelse(rev(x$origin)=='IMPORTED','과거 CSV','일별 자동 스냅샷'),check.names=FALSE)
    DT::datatable(y,rownames=FALSE,options=list(pageLength=15,order=list(list(0,'desc')),scrollX=TRUE)) |>
      DT::formatCurrency(columns=c('총평가액_원','평가손익_원','평가액변동_원','손익변동_원'),currency='',digits=0,interval=3,mark=',') |>
      DT::formatStyle(columns=c('평가액변동_원','손익변동_원'),color=DT::styleInterval(c(-0.00001,0.00001),c('#2563eb','#64748b','#dc2626')),fontWeight='bold')
  })
  performance_filtered <- reactive({
    x <- trajectory_all()
    p <- input$performance_period
    if(!is.null(p) && p!='all' && nrow(x))
      x <- x[x$date >= max(x$date)-as.integer(p),,drop=FALSE]
    x
  })
  performance_result <- reactive({pms_performance_calc(performance_filtered())})
  output$performance_source_note <- renderText({
    x <- performance_filtered()
    sprintf('분석 대상 %d일 · 과거 CSV %d일 · 신규 SQLite 스냅샷 %d일',
      nrow(x),sum(x$origin=='IMPORTED',na.rm=TRUE),sum(x$origin!='IMPORTED',na.rm=TRUE))
  })
  output$performance_dashboard <- renderUI({
    pms_performance_panel(performance_result())
  })
  output$performance_detail <- DT::renderDT({
    x <- performance_result()$data
    if(!nrow(x) || !'drawdown_pct' %in% names(x)) return(DT::datatable(data.frame(안내='분석 기록이 부족합니다.'),options=list(dom='t')))
    y <- data.frame(날짜=as.character(x$date),총평가액_원=x$value,
      미실현손익_원=x$profit,평가액증감률_pct=x$change_pct,
      고점대비하락률_pct=x$drawdown_pct,자료유형=ifelse(x$origin=='IMPORTED','과거 CSV','일별 스냅샷'))
    y <- y[nrow(y):1,,drop=FALSE]
    DT::datatable(y,rownames=FALSE,options=list(pageLength=15,scrollX=TRUE,order=list(list(0,'desc')))) |>
      DT::formatCurrency(c('총평가액_원','미실현손익_원'),currency='',digits=0,interval=3,mark=',') |>
      DT::formatRound(c('평가액증감률_pct','고점대비하락률_pct'),digits=2) |>
      DT::formatStyle(c('미실현손익_원','평가액증감률_pct','고점대비하락률_pct'),
        color=DT::styleInterval(c(-1e-10,1e-10),c('#267dc5','#667788','#d43838')))
  })
  snapshot_write_once <- function() {
    tryCatch({
      x <- pms_snapshot_save(con)
      snapshot_save_status(sprintf('%s SQLite 저장 완료 · 평가 가능 %d/%d건 · 상태 %s · data/pms.sqlite',
        x$date,x$valued,x$holdings,x$status))
      snapshot_refresh(isolate(snapshot_refresh())+1L)
      trade_refreshed(isolate(trade_refreshed())+1L)
      invisible(TRUE)
    },error=function(e) {
      snapshot_save_status(paste('스냅샷 저장 오류:',conditionMessage(e)))
      invisible(FALSE)
    })
  }
  observeEvent(market_refresh(),{snapshot_write_once()},ignoreInit=FALSE)
  observeEvent(input$snapshot_save_now,{snapshot_write_once()},ignoreInit=TRUE)
  # 브라우저 세션에서만 동작하는 10분 자동 조회. 타이머는 R Shiny 세션이
  # 유지되는 동안만 유효하며, 배당이나 과거 일별 환율은 다시 수집하지 않는다.
  auto_last_at <- reactiveVal(Sys.time())
  auto_last_note <- reactiveVal('대기 중 (10분 후 첫 자동 조회)')
  auto_success_info <- reactiveVal('첫 자동 조회 대기')
  auto_running <- reactiveVal(FALSE)
  price_running <- reactiveVal(FALSE)
  # 모든 수동/자동 시세수집의 동시 실행을 차단한다.
  guarded_price_collection <- function() {
    if (isTRUE(isolate(price_running()))) return(invisible(FALSE))
    price_running(TRUE)
    on.exit(price_running(FALSE),add=TRUE)
    run_price_collection()
    invisible(TRUE)
  }
  observeEvent(input$auto_prices_10m, {
    auto_last_at(Sys.time())
    auto_last_note(if(isTRUE(input$auto_prices_10m)) '첫 자동조회 대기 중' else '자동 조회 꺼짐')
  }, ignoreInit=FALSE)
  output$auto_prices_status <- renderText({
    active <- isTRUE(input$auto_prices_10m)
    info <- auto_last_note()
    if (!active) return('자동 조회 중지 · 버튼으로 수동 수집 가능')
    paste0('상태: ', info)
  })
  observe({
    invalidateLater(10000,session)
    if (!isTRUE(input$auto_prices_10m)) return(invisible(NULL))
    if (isTRUE(isolate(price_running()))) return(invisible(NULL))
    elapsed <- as.numeric(difftime(Sys.time(),isolate(auto_last_at()),units='secs'))
    if (!is.finite(elapsed) || elapsed < 600) return(invisible(NULL))
    started <- Sys.time()
    before_log <- tryCatch(DBI::dbGetQuery(con,'SELECT COALESCE(MAX(id),0) AS id FROM collection_log')$id[1],error=function(e) 0)
    auto_last_at(Sys.time()) # 요청 간격은 조회 시작을 기준으로 산정
    auto_last_note(paste0('자동 조회 진행: ',format(Sys.time(),'%Y-%m-%d %H:%M:%S')))
    tryCatch({
      guarded_price_collection()
      counts <- DBI::dbGetQuery(con,"SELECT status,COUNT(*) AS n FROM collection_log WHERE id > ? AND country IN ('KR','US') GROUP BY status",params=list(before_log))
      ok <- if(any(counts$status=='OK')) sum(counts$n[counts$status=='OK']) else 0L
      bad <- if(any(counts$status=='FAILED')) sum(counts$n[counts$status=='FAILED']) else 0L
      state <- if(ok+bad==0) '수집 결과 없음' else if(bad==0) '정상 완료' else '일부 실패'
      auto_last_note(sprintf('마지막 자동조회 %s (%s · 성공 %d / 실패 %d)',format(Sys.time(),'%H:%M:%S'),state,ok,bad))
    },error=function(e) {
      auto_last_note(paste0('자동 조회 오류: ',conditionMessage(e)))
    })
  })
  observeEvent(input$collect_prices, { guarded_price_collection() })
  observeEvent(input$collect_all, {
    # 한 번의 사용자 요청으로 기존 두 수집 작업을 순서대로 실행합니다.
    # 각 수집기는 자체적으로 오류를 표시하므로 배당 단계도 계속 시도합니다.
    all_collection_message('1/2 시세 수집을 시작합니다. 이어서 최근 1년 배당과 일별 환율을 조회합니다.')
    tryCatch(guarded_price_collection(), error=function(e) {
      collection_message(paste('시세 전체조회 오류:', conditionMessage(e)))
    })
    all_collection_message('2/2 최근 1년 배당정보 및 USD/KRW 일별 환율 수집 중입니다...')
    tryCatch(run_dividend_collection(), error=function(e) {
      dividend_message(paste('배당 전체조회 오류:', conditionMessage(e)))
    })
    all_collection_message('전체 조회 작업이 끝났습니다. 시세 및 배당 결과의 성공·실패 건수를 각각 확인해 주세요.')
  })
  # 단일 UI 출력: 카드별 textOutput이 초기 렌더링에서 비는 문제를 방지한다.
  # 사용 데이터는 종목별 평가표와 동일하며, 외부 API를 절대 호출하지 않는다.
  # 비어 있거나 오류가 생기면 카드에 명확한 상태를 출력한다.
  dashboard_data <- reactive({
    d <- valuation_data()
    if (!nrow(d)) return(list(total=NA_real_,profit=NA_real_,missing=0L,
                              active=0L,valued=0L,fx=NA_real_,fx_date=''))
    active <- !is.na(d$quantity) & d$quantity > 0
    value <- suppressWarnings(as.numeric(d$market_value_krw))
    pnl <- suppressWarnings(as.numeric(d$profit_krw))
    good <- active & is.finite(value)
    pnl_good <- good & is.finite(pnl)
    fx <- DBI::dbGetQuery(con,"SELECT rate,quoted_at FROM fx_rates WHERE pair='USD/KRW'")
    list(total=if(any(good)) sum(value[good]) else NA_real_,
         profit=if(any(pnl_good)) sum(pnl[pnl_good]) else NA_real_,
         missing=sum(active & !good), active=sum(active), valued=sum(good),
         fx=if(nrow(fx)) as.numeric(fx$rate[1]) else NA_real_,
         fx_date=if(nrow(fx)) as.character(fx$quoted_at[1]) else '')
  })
  dashboard_krw <- function(x) {
    if(length(x)!=1L || !is.finite(x)) return('미산출 (시세 또는 환율 확인 필요)')
    paste0(format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),' 원')
  }
  output$valuation_dashboard <- renderUI({
    tryCatch({
      d <- dashboard_data()
      note <- if(d$active==0L) '현재 보유수량이 있는 종목이 없습니다.'
              else if(d$missing>0L) sprintf('시세·환율 누락 %d건 제외한 부분 합계',d$missing)
              else sprintf('정상 평가 %d / %d건',d$valued,d$active)
      profit_color <- if(is.finite(d$profit) && d$profit<0) '#c62828' else '#222222'
      fluidRow(
        column(4,wellPanel(h4('전체 평가액'),
          tags$div(style='font-size:25px;font-weight:700;',dashboard_krw(d$total)),
          tags$small(note))),
        column(4,wellPanel(h4('전체 평가손익'),
          tags$div(style=paste0('font-size:25px;font-weight:700;color:',profit_color,';'),dashboard_krw(d$profit)),
          tags$small(if(d$missing>0L) '평가 불가 종목의 손익은 제외' else '저장된 매입원가 기준'))),
        column(4,wellPanel(h4('시세 상태'),
          tags$div(sprintf('평가 완료 %d / %d건',d$valued,d$active)),
          tags$div(if(is.finite(d$fx)) paste0('USD/KRW: ',format(d$fx,big.mark=',',trim=TRUE),
                   ' (',d$fx_date,')') else 'USD/KRW 저장 없음')))
      )
    },error=function(e) {
      tags$div(class='alert alert-danger',
        tags$b('저장된 자산 평가 불러오기 오류: '),conditionMessage(e),
        tags$p('SQLite의 저장된 시세 또는 보유종목 정보를 확인해 주세요.'))
    })
  })
  # 비활성 탭에서도 초기 계산을 준비하고, 탭 진입 시 즉시 표시한다.
  outputOptions(output,'valuation_dashboard',suspendWhenHidden=FALSE)
  output$quote_coverage <- renderText({
    d <- valuation_data()
    if (!nrow(d)) return('평가 대상 없음')
    active <- d$quantity>0
    paste0('정상 ',sum(d$quote_status[active]=='OK'),' / ',sum(active),' 보유 행')
  })
  output$fx_status <- renderText({
    market_refresh()
    x <- DBI::dbGetQuery(con,"SELECT rate,quoted_at FROM fx_rates WHERE pair='USD/KRW'")
    if (!nrow(x)) return('USD/KRW 미수집')
    paste0('USD/KRW: ',format(x$rate[1],big.mark=',',scientific=FALSE),
           ' (기준 ',x$quoted_at[1],')')
  })
  # 합계는 DataTables의 페이지 이동/정렬/검색과 독립적으로 전체 보유내역에서 계산합니다.
  # 조회 실패한 평가금/손익은 0으로 간주하지 않으며 합계에서 제외합니다.
  valuation_totals <- reactive({
    d <- valuation_data()
    if (!nrow(d)) return(list(total=NA_real_, cost=NA_real_, profit=NA_real_,
                              active=0L, evaluated=0L, missing=0L, cost_rows=0L))
    active <- !is.na(d$quantity) & d$quantity > 0
    valid <- active & is.finite(d$market_value_krw)
    valid_cost <- valid & is.finite(d$cost_basis_krw)
    list(
      total=if(any(valid)) sum(d$market_value_krw[valid]) else NA_real_,
      cost=if(any(valid_cost)) sum(d$cost_basis_krw[valid_cost]) else NA_real_,
      profit=if(any(valid_cost)) sum(d$profit_krw[valid_cost],na.rm=TRUE) else NA_real_,
      active=sum(active), evaluated=sum(valid), missing=sum(active & !valid),
      cost_rows=sum(valid_cost)
    )
  })
  fmt_krw <- function(x) {
    if(length(x) != 1L || !is.finite(x)) return('미산출')
    paste0(format(round(x),big.mark=',',scientific=FALSE,trim=TRUE),' 원')
  }
  output$valuation_summary <- renderTable({
    x <- valuation_totals()
    data.frame(
      구분=c('총 평가금액','총 취득원가','총 평가손익','평가 완료 / 보유 행'),
      합계=c(fmt_krw(x$total),fmt_krw(x$cost),fmt_krw(x$profit),
             sprintf('%d / %d건',x$evaluated,x$active)),
      check.names=FALSE
    )
  }, striped=TRUE, bordered=TRUE, spacing='s',rownames=FALSE)
  output$valuation_warning <- renderText({
    x <- valuation_totals()
    if(x$active==0L) return('보유수량이 0보다 큰 종목이 없습니다.')
    if(x$missing>0L) return(sprintf(
      '주의: 보유 %d건 중 %d건은 시세 또는 환율이 없어 총 평가금액에서 제외되었습니다. 위 합계는 전체 자산 총액이 아닙니다.',
      x$active,x$missing))
    if(x$cost_rows<x$evaluated) return('주의: 취득원가 누락 종목은 손익 계산에서 제외되었습니다.')
    '모든 활성 보유내역이 평가에 포함되었습니다.'
  })
  output$valuation_table <- renderDT({
    x <- pms_valuation_view(valuation_data())
    if (!nrow(x)) return(datatable(data.frame(안내='보유종목 또는 저장된 시세가 없습니다.'),rownames=FALSE))
    tab <- datatable(x,rownames=FALSE,filter='top',selection='none',extensions='FixedHeader',
      options=list(pageLength=100,lengthMenu=c(10,25,50,100),scrollX=TRUE,
                   scrollY='56vh',scrollCollapse=TRUE,fixedHeader=TRUE,
                   order=list(list(match('비중_퍼센트', names(x)) - 1L,'desc')),columnDefs=pms_money_defs(x)))
    for(nm in c('평가액_원','평가손익_원','수익률_퍼센트')) {
      if(nm %in% names(x)) {
        tab <- DT::formatStyle(tab,nm,color=DT::styleInterval(0,c('#c62828','inherit')))
      }
    }
    tab
  },server=FALSE)

  zoom_pct <- reactiveVal(100L)
  output$zoom_label <- renderText(paste0('화면 배율: ', zoom_pct(), '%'))
  observeEvent(input$zoom_in, {
    zoom_pct(min(150L, zoom_pct() + 10L))
  })
  observeEvent(input$zoom_out, {
    zoom_pct(max(70L, zoom_pct() - 10L))
  })
  observeEvent(input$zoom_reset, {
    zoom_pct(100L)
  })
  observeEvent(zoom_pct(), {
    session$sendCustomMessage('pms_set_zoom', zoom_pct())
  }, ignoreInit = FALSE)

  # Briefing is constructed from saved quotes / dividend events / historical risk estimates.
  output$voice_brief_text <- renderText({
    refresh()
    tryCatch(pms_voice_brief(con), error=function(e)
      paste('음성 브리핑을 준비하지 못했습니다:',conditionMessage(e)))
  })
  outputOptions(output, 'voice_brief_text', suspendWhenHidden=FALSE)
  # Report creation is explicitly triggered; do not query external endpoints.
  report_paths <- reactiveVal(list(qmd=NULL,html=NULL,pdf=NULL))
  report_status <- reactiveVal('보고서 생성 버튼을 눌러주세요. Quarto CLI 설치 여부: 확인 중')
  output$report_status <- renderText(report_status())
  output$report_summary_ui <- renderUI({
    p <- report_paths()
    if(is.null(p$qmd)) return(tags$p('아직 생성된 보고서가 없습니다.'))
    tags$div(class='alert alert-success',
      tags$p('상세 보고서 원본 생성 완료: ',basename(p$qmd)),
      tags$p(if(!is.null(p$html)) 'HTML 렌더링 완료.' else 'HTML 미생성: Quarto 설치·렌더링 상태를 확인하세요.'),
      tags$p(if(!is.null(p$pdf)) 'PDF 생성 완료(한글 글꼴: 맑은 고딕). PDF 보고서 다운로드 버튼을 이용하세요.' else 'PDF가 없으면 Quarto 및 TinyTeX(LaTeX) 설치 여부를 확인하세요.'))
  })
  output$report_inline_view <- renderUI({
    p <- report_paths()
    if(is.null(p$qmd) || !file.exists(p$qmd))
      return(tags$div(class='alert alert-info','보고서 생성 버튼을 누르면 이 탭 아래에서 본문을 바로 읽을 수 있습니다.'))
    if(!is.null(p$html) && file.exists(p$html)) {
      html_text <- paste(readLines(p$html,warn=FALSE,encoding='UTF-8'),collapse='\n')
      return(tags$iframe(srcdoc=html_text,style='display:block;width:100%;height:1100px;border:1px solid #ccd7e1;border-radius:8px;background:white;',sandbox='allow-scripts allow-popups'))
    }
    qmd_lines <- readLines(p$qmd,warn=FALSE,encoding='UTF-8')
    if(length(qmd_lines)>2L && identical(qmd_lines[1],'---')) {
      ends <- which(qmd_lines=='---'); if(length(ends)>=2) qmd_lines <- qmd_lines[-seq_len(ends[2])]
    }
    txt <- paste(qmd_lines,collapse='\n')
    if(requireNamespace('commonmark',quietly=TRUE))
      return(tags$div(class='pms-report-inline',style='max-height:1200px;overflow-y:auto;padding:24px;border:1px solid #dbe2e9;',HTML(commonmark::markdown_html(txt))))
    if(requireNamespace('markdown',quietly=TRUE))
      return(tags$div(class='pms-report-inline',style='max-height:1200px;overflow-y:auto;padding:24px;border:1px solid #dbe2e9;',HTML(markdown::markdownToHTML(text=txt,fragment.only=TRUE))))
    tags$pre(style='white-space:pre-wrap;max-height:1200px;overflow:auto;padding:22px;',txt)
  })
  observeEvent(input$report_generate, {
    on.exit(session$sendCustomMessage('pms_report_busy', list(running=FALSE)), add=TRUE)
    tryCatch({
      dir <- file.path(tempdir(),'pms2_reports')
      dir.create(dir,recursive=TRUE,showWarnings=FALSE)
      tag <- format(Sys.time(),'%Y%m%d_%H%M%S')
      qmd <- file.path(dir,paste0('PMS2_Report_',tag,'.qmd'))
      report_status('SQLite 자료 집계 및 상세 Quarto 원본 작성 중...')
      pms_report_write(con,qmd,include_benchmark=isTRUE(input$report_benchmark))
      html <- NULL
      pdf <- NULL
      pdf_notice <- ''
      cli <- Sys.which('quarto')
      if(nzchar(cli)) {
        report_status('Quarto HTML 렌더링 중...')
        # Run Quarto in an isolated output folder with no network data dependencies.
        cmdout <- tryCatch(suppressWarnings(system2(cli,c('render',shQuote(qmd),'--to','html'),stdout=TRUE,stderr=TRUE)),error=function(e)e)
        expected <- sub('\\.qmd$','.html',qmd)
        if(file.exists(expected) && file.info(expected)$size>0) html<-expected
        else if(inherits(cmdout,'error')) report_status(paste('HTML 렌더링 오류:',conditionMessage(cmdout)))
        else report_status(paste('HTML 렌더링 실패:',paste(tail(cmdout,3L),collapse=' / ')))
      } else report_status('Quarto CLI가 설치돼 있지 않아 .qmd 원본만 생성했습니다. 설치 후 HTML로 렌더링할 수 있습니다.')
      if(nzchar(cli)) {
        # PDF is independent of HTML; missing TeX must not break report generation.
        report_status('Quarto PDF 렌더링 시도 중 (LaTeX 설치 필요)...')
        pdf_out <- tryCatch(suppressWarnings(system2(cli,
          c('render',shQuote(qmd),'--to','pdf'),stdout=TRUE,stderr=TRUE)),error=function(e)e)
        expected_pdf <- sub('\\.qmd$','.pdf',qmd)
        if(file.exists(expected_pdf) && is.finite(file.info(expected_pdf)$size) && file.info(expected_pdf)$size>0) {
          pdf <- expected_pdf
        } else {
          pdf_notice <- if(inherits(pdf_out,'error')) conditionMessage(pdf_out) else
            paste(tail(as.character(pdf_out),4L),collapse=' / ')
          if(!nzchar(pdf_notice)) pdf_notice <- 'Quarto PDF 렌더링 실패. TinyTeX 등 LaTeX 엔진을 확인하세요.'
        }
      }
      report_paths(list(qmd=qmd,html=html,pdf=pdf))
      report_status(paste0('보고서 생성 완료. HTML: ',if(!is.null(html)) '준비됨' else '미생성',
        ' · PDF: ',if(!is.null(pdf)) '준비됨' else paste0('미생성 (',pdf_notice,')')))

    },error=function(e) {report_paths(list(qmd=NULL,html=NULL,pdf=NULL));report_status(paste('보고서 생성 오류:',conditionMessage(e)))})
  })
  output$report_download_qmd <- downloadHandler(
    filename=function() {p<-report_paths();if(!is.null(p$qmd)) basename(p$qmd) else 'PMS2_Report.qmd'},
    content=function(file) {p<-report_paths();req(!is.null(p$qmd),file.exists(p$qmd));file.copy(p$qmd,file,overwrite=TRUE)}
  )
  output$report_download_pdf <- downloadHandler(
    filename=function() {p<-report_paths();if(!is.null(p$pdf)) basename(p$pdf) else 'PMS2_Report.pdf'},
    content=function(file) {p<-report_paths();req(!is.null(p$pdf),file.exists(p$pdf));file.copy(p$pdf,file,overwrite=TRUE)}
  )
  output$report_download_html <- downloadHandler(
    filename=function() {p<-report_paths();if(!is.null(p$html)) basename(p$html) else 'PMS2_Report.html'},
    content=function(file) {p<-report_paths();req(!is.null(p$html),file.exists(p$html));file.copy(p$html,file,overwrite=TRUE)}
  )
  refresh <- reactiveVal(0L)
  pms_diversification_server(input,output,session,con,refresh)


  # v0.5.55 홈: 모든 숫자는 현재 SQLite의 저장 자료로만 계산한다.
  # reactiveVal을 먼저 참조해 저장/수집 갱신 시 한 번에 다시 렌더링한다.
  output$home_dashboard <- renderUI({
    snapshot_refresh()
    market_refresh()
    refresh()
    dividend_refresh()
    fmt <- function(v,unit='원') {
      if(length(v)!=1L || !is.finite(v)) return('미산출')
      paste0(format(round(v),big.mark=',',scientific=FALSE,trim=TRUE),unit)
    }
    panel <- function(title,value,subtitle='',color='#27384a') {
      tags$div(style=paste0('flex:1 1 220px;min-width:180px;background:#f7f9fb;border:1px solid #dfe6ee;border-radius:9px;padding:16px;'),
        tags$div(style='font-size:13px;color:#586678;',title),
        tags$div(style=paste0('font-size:23px;font-weight:700;color:',color,';margin:7px 0;'),value),
        tags$div(style='font-size:12px;color:#687786;',subtitle))
    }
    section <- function(title, ...) tags$div(style='background:white;border:1px solid #e0e5ec;border-radius:9px;padding:16px;margin:12px 0;',
      tags$h4(style='margin-top:0;font-weight:700;',title), ...)
    tryCatch({
      d <- pms_valuation(con,include_zero=FALSE)
      if(!nrow(d)) return(tags$div(class='alert alert-info','현재 보유종목이 없습니다. 보유종목 편집 탭에서 등록해 주세요.'))
      valid <- is.finite(d$market_value_krw) & d$quantity>0
      validp <- valid & is.finite(d$profit_krw)
      nav <- if(any(valid)) sum(d$market_value_krw[valid]) else NA_real_
      pnl <- if(any(validp)) sum(d$profit_krw[validp]) else NA_real_
      cost <- if(any(validp)) sum(d$cost_basis_krw[validp]) else NA_real_
      ret <- if(is.finite(cost) && cost>0 && is.finite(pnl)) pnl/cost*100 else NA_real_
      nmissing <- sum(d$quantity>0 & !valid)
      nactive <- sum(d$quantity>0)
      div <- tryCatch(pms_dividend_account_totals(con),error=function(e) NULL)
      div_total <- if(!is.null(div) && nrow(div)) as.numeric(div$최근1년_배당금_원화_추정합계[div$계좌유형=='전체 합계'][1]) else NA_real_
      div_general <- if(!is.null(div) && nrow(div)) as.numeric(div$최근1년_배당금_원화_추정합계[div$계좌유형=='일반계좌'][1]) else NA_real_
      div_missing <- if(!is.null(div) && nrow(div)) as.integer(div$미완료건수[div$계좌유형=='전체 합계'][1]) else NA_integer_
      monthly <- tryCatch(pms_dividend_monthly_accounts(con),error=function(e) NULL)
      lastmonth <- if(!is.null(monthly) && nrow(monthly)) max(monthly$월) else ''
      thisdiv <- if(nzchar(lastmonth)) sum(monthly$금액[monthly$월==lastmonth],na.rm=TRUE) else NA_real_
      fx <- DBI::dbGetQuery(con,"SELECT rate,quoted_at FROM fx_rates WHERE pair='USD/KRW'")
      fx_label <- if(nrow(fx) && is.finite(fx$rate[1])) format(round(fx$rate[1],2),big.mark=',',nsmall=2) else '미수집'
      buy_stats <- tryCatch(pms_buy_estimate_12m(con),error=function(e) list(
        total=NA_real_,mean=NA_real_,n=0L,excluded=0L,months=0L,
        note=paste('매수 추정액 집계 오류:',conditionMessage(e))))
      trade_stats <- tryCatch(pms_trade_value_estimate_12m(con),error=function(e) list(
        buy=NA_real_,sell=NA_real_,net=NA_real_,buy_monthly=NA_real_,
        sell_monthly=NA_real_,net_monthly=NA_real_,buy_n=0L,sell_n=0L,excluded=0L,
        note=paste('순매수 추정액 집계 오류:',conditionMessage(e))))
      value_stats <- tryCatch(pms_valuation_change_12m(con),error=function(e) list(
        delta=NA_real_,monthly=NA_real_,start='',end='',span=NA_integer_,
        note=paste('평가액 증감 집계 오류:',conditionMessage(e))))
      basic <- tags$div(style='display:flex;flex-wrap:wrap;gap:12px;',
        panel('증권 평가액',fmt(nav),sprintf('시세 평가 %d / %d건',sum(valid),nactive)),
        panel('평가손익',fmt(pnl),if(is.finite(ret)) sprintf('취득원가 대비 %+.2f%%',ret) else '수익률 미산출',if(is.finite(pnl) && pnl<0) '#cb3434' else '#147a50'),
        panel('최근 365일 배당 추정',fmt(div_total),if(is.finite(div_missing)) sprintf('미완료 보유내역 %d건',div_missing) else '배당 조회 상태 확인'),
        panel('이번 달 예상 배당금 (배당락일 기준)',fmt(thisdiv),paste0(lastmonth,' · 실제 입금일/실수령액과 다를 수 있는 세전 추정')))
      bars <- function(labels,values) {
        if(!length(values)) return(tags$p(class='text-muted','평가 가능한 데이터가 없습니다.'))
        vmax <- max(values,na.rm=TRUE)
        tags$div(lapply(seq_along(labels),function(i) {
          tags$div(style='margin:9px 0;',
            tags$div(style='display:flex;justify-content:space-between;font-size:13px;',tags$span(labels[i]),tags$strong(fmt(values[i]))),
            tags$div(style='background:#edf2f7;border-radius:4px;height:11px;',
              tags$div(style=sprintf('height:11px;width:%.2f%%;background:#4687c6;border-radius:4px;',if(vmax>0) max(0,values[i])/vmax*100 else 0))))
        }))
      }
      vd <- d[valid,,drop=FALSE]
      groups <- if(nrow(vd)) aggregate(vd$market_value_krw,list(account=ifelse(is.na(vd$account_type),'기타',vd$account_type)),sum) else NULL
      if(!is.null(groups)) groups <- groups[order(-groups$x),,drop=FALSE]
      countries <- if(nrow(vd)) aggregate(vd$market_value_krw,list(country=vd$country),sum) else NULL
      if(!is.null(countries)) countries <- countries[order(-countries$x),,drop=FALSE]
      top <- if(nrow(vd)) aggregate(vd$market_value_krw,list(name=vd$name,symbol=vd$symbol),sum) else NULL
      if(!is.null(top)) top <- head(top[order(-top$x),,drop=FALSE],5L)
      fxnote <- if(nrow(fx)) paste0('시세 기준: ',as.character(fx$quoted_at[1])) else '원달러 환율이 없습니다.'
      stale <- if(nrow(vd)) {
        dates <- suppressWarnings(as.Date(substr(as.character(vd$price_date),1L,10L)))
        # Count unique symbols, not brokerage rows. This is a warning, not a delisting verdict.
        sum(!duplicated(paste(vd$country,vd$symbol)) & (is.na(dates) | as.integer(Sys.Date()-dates)>4L))
      } else 0L
      tags$div(
        basic,
        section('최근 1년 월평균 자금 운용과 평가액 변화',
          tags$div(style='display:flex;flex-wrap:wrap;gap:15px;',
            panel('월평균 매수 추정액',fmt(trade_stats$buy_monthly),'증가수량 × 기록 시세'),
            panel('월평균 매도 추정액',fmt(trade_stats$sell_monthly),'감소수량 × 기록 시세'),
            panel('월평균 순매수 추정액',fmt(trade_stats$net_monthly),'매수 − 매도 (온주전환 제외)',
                  if(is.finite(trade_stats$net_monthly) && trade_stats$net_monthly<0) '#cb3434' else '#147a50'),
            panel('월평균 평가액 증가분',fmt(value_stats$monthly),
                  paste0(value_stats$start,' ~ ',value_stats$end),
                  if(is.finite(value_stats$monthly) && value_stats$monthly<0) '#cb3434' else '#147a50')),
          tags$p(style='margin:9px 0 3px;font-size:13px;',
            '최근 1년 순매수 합계: ',tags$strong(fmt(trade_stats$net)),
            ' · 총평가액 증감: ',tags$strong(fmt(value_stats$delta))),
          tags$p(class='text-muted',style='font-size:12px;',trade_stats$note),
          tags$p(class='text-muted',style='font-size:12px;',value_stats$note),
          tags$p(class='alert alert-info',style='font-size:12px;margin-bottom:0;',
            '순매수는 종목별 매수−매도 추정액으로, 실제 외부 불입액이 아닙니다. 매도대금의 재투자·계좌 간 이동과 가격 및 환율 누락이 영향을 줍니다. ',
            '평가액 증가는 자금 순유입과 시장손익이 합쳐진 변화이며 투자수익이나 납입액으로 해석하면 안 됩니다.')),
        pms_trajectory_panel(pms_trajectory_read(con),compact=TRUE),
        pms_home_market_overlay(con,pms_trajectory_read(con)),
        pms_equal_cashflow_panel(con,pms_trajectory_read(con)),
        tags$div(style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:6px;",
          actionButton('home_etf_retry','KODEX 미국S&P500 전체기간 다시 수집',class='btn btn-default btn-sm'),
          tags$span(style="font-size:12px;color:#667789;",textOutput('home_overlay_status',inline=TRUE))),
        if(nmissing>0) tags$div(class='alert alert-warning',style='margin:12px 0;',sprintf('시세 또는 환율이 없어 %d건을 제외한 부분 평가액입니다.',nmissing)),
        tags$div(style='display:flex;flex-wrap:wrap;gap:14px;',
          tags$div(style='flex:1 1 360px;',section('국가별 증권 평가액',if(!is.null(countries)) bars(ifelse(countries$country=='KR','한국','미국'),countries$x) else NULL)),
          tags$div(style='flex:1 1 360px;',section('계좌유형별 증권 평가액',if(!is.null(groups)) bars(groups$account,groups$x) else NULL))),
        tags$div(style='display:flex;flex-wrap:wrap;gap:14px;',
          tags$div(style='flex:1 1 360px;',section('평가액 상위 5종목',if(!is.null(top)) bars(top$name,top$x) else NULL)),
          tags$div(style='flex:1 1 360px;',section('배당·금융소득 현황',
            tags$p('일반계좌 최근 365일 세전 배당 추정: ',tags$strong(fmt(div_general))),
            tags$p('2,000만 원 기준선 대비: ',tags$strong(if(is.finite(div_general)) sprintf('%.1f%%',div_general/20000000*100) else '미산출')),
            tags$p(class='text-muted',style='font-size:12px;', '실제 과세기간의 이자·배당 전체와는 다른 참고치입니다.')))),
        section('시세 및 데이터 품질',
          tags$p('USD/KRW: ',tags$strong(fx_label),' · ',fxnote),
          tags$p(sprintf('저장 시세 기준 4일 초과 또는 기준일 미확인 종목 %d개 (평가 가능 종목 중)',stale)),
          tags$p(sprintf('현재 보유내역 %d건 · 정상 평가 %d건 · 평가 제외 %d건',nactive,sum(valid),nmissing)),
          tags$p(class='text-muted',style='font-size:12px;', '오래된 시세는 거래정지·상장폐지를 의미하지 않으며 휴장일도 포함됩니다.')),
        tags$p(style='font-size:12px;color:#687786;',paste('저장 자료 요약 갱신:',format(Sys.time(),'%Y-%m-%d %H:%M:%S'))))
    },error=function(e) tags$div(class='alert alert-danger',paste('홈 요약 표시 오류:',conditionMessage(e))))
  })
  observeEvent(input$home_refresh, {
    market_refresh(isolate(market_refresh())+1L)
    dividend_refresh(isolate(dividend_refresh())+1L)
    refresh(isolate(refresh())+1L)
  },ignoreInit=TRUE)
  outputOptions(output,'home_dashboard',suspendWhenHidden=FALSE)

  # Excel bulk holdings editing. Export reads full table, including zero-quantity rows.
  excel_plan <- reactiveVal(NULL)
  excel_status <- reactiveVal('엑셀을 업로드하면 변경내용을 검증하고 미리 보여줍니다.')
  output$holdings_xlsx_status <- renderText(excel_status())
  output$holdings_xlsx_download <- downloadHandler(
    filename=function() paste0('PMS2_보유종목_',format(Sys.Date(),'%Y%m%d'),'.xlsx'),
    content=function(file) pms_holdings_excel_export(con,file),
    contentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  )
  # Explicitly keep both download and its diagnostics active across tab switches.
  outputOptions(output, 'holdings_xlsx_download', suspendWhenHidden=FALSE)
  output$holdings_xlsx_download_ready <- renderText({
    if (!requireNamespace('writexl',quietly=TRUE))
      return('엑셀 작성 패키지 미감지: RStudio 실행 R 라이브러리를 확인하세요.')
    paste0('엑셀 생성 준비됨 · writexl ',as.character(utils::packageVersion('writexl')),
           ' · 다운로드가 회색이면 브라우저 새로고침 후 다시 확인')
  })
  outputOptions(output, 'holdings_xlsx_download_ready', suspendWhenHidden=FALSE)
  observeEvent(input$holdings_xlsx_upload, {
    excel_plan(NULL)
    tryCatch({
      req(input$holdings_xlsx_upload$datapath)
      plan <- pms_holdings_excel_validate(con,input$holdings_xlsx_upload$datapath)
      excel_plan(plan)
      excel_status(paste('검증 완료:',plan$summary,'· 적용을 눌러야 DB에 저장됩니다.'))
    },error=function(e) {
      excel_status(paste('엑셀 검증 실패:',conditionMessage(e)))
      showNotification(conditionMessage(e),type='error',duration=12)
    })
  })
  output$holdings_xlsx_preview <- renderDT({
    plan <- excel_plan(); if(is.null(plan)) return(NULL)
    datatable(plan$preview,rownames=FALSE,options=list(pageLength=10,scrollX=TRUE),filter='top')
  })
  output$holdings_xlsx_confirm_ui <- renderUI({
    if(is.null(excel_plan())) return(NULL)
    actionButton('holdings_xlsx_apply','검증된 엑셀 변경사항 DB에 적용',class='btn-primary')
  })
  observeEvent(input$holdings_xlsx_apply, {
    plan <- excel_plan(); if(is.null(plan)) return()
    showModal(modalDialog(title='엑셀 일괄 반영 최종 확인',
      tags$p(plan$summary),tags$p('기존 ID의 값은 수정되고 신규 행은 추가됩니다. 누락된 행은 삭제하지 않습니다.'),
      footer=tagList(modalButton('취소'),actionButton('holdings_xlsx_commit','확인 · 일괄 저장',class='btn-danger'))))
  })
  observeEvent(input$holdings_xlsx_commit, {
    plan <- excel_plan(); if(is.null(plan)) return()
    removeModal()
    tryCatch({
      # Revalidate the original uploaded file, including optimistic timestamp checks.
      req(input$holdings_xlsx_upload$datapath)
      current <- pms_holdings_excel_validate(con,input$holdings_xlsx_upload$datapath)
      pms_holdings_excel_apply(con,current)
      excel_plan(NULL)
      excel_status(paste('일괄 저장 완료:',current$summary))
      refresh(isolate(refresh())+1L)
      showNotification('보유종목 엑셀 일괄 반영 완료',type='message',duration=6)
    },error=function(e){
      excel_status(paste('반영 실패: DB 변경 없이 중단 또는 롤백했습니다.',conditionMessage(e)))
      showNotification(conditionMessage(e),type='error',duration=NULL)
    })
  })

  mode <- reactiveVal('new')
  editing_id <- reactiveVal(NULL)
  selected_id <- reactiveVal(NULL)
  updated_at_display <- reactiveVal('신규 등록 전')
  output$updated_at_display <- renderText(updated_at_display())
  notice <- reactiveVal('준비 완료. 왼쪽 표에서 행을 클릭하면 자동으로 수정 모드가 됩니다.')

  output$notice <- renderText(notice())
  output$form_title <- renderText(if (mode()=='edit') paste0('수정 중 | ID ',editing_id()) else '신규 등록')

  holdings <- reactive({
    refresh()
    df <- list_holdings(con,include_zero=isTRUE(input$show_zero))
    if (!nrow(df)) {
      df$평가액_원 <- numeric(0)
      return(df)
    }
    # 평가액은 DB의 저장 시세/환율만 사용. 외부 API는 호출하지 않는다.
    v <- pms_valuation(con,include_zero=isTRUE(input$show_zero))
    pos <- match(df$id, v$id)
    df$평가액_원 <- as.numeric(v$market_value_krw[pos])
    df$평가액_원[!is.finite(df$평가액_원)] <- NA_real_
    # 기본 정렬: 증권사 > 국가(KR, US, 기타) > 계좌유형 > 평가액 내림차순.
    # 동일 증권사의 한국 및 미국 보유종목을 한 묶음으로 모은다.
    # 같은 그룹 안에서 평가액이 없으면 뒤로 보내고, 동률은 ID순으로 고정한다.
    country <- toupper(trimws(as.character(df$country)))
    country_rank <- ifelse(country %in% c('KR','KOREA'), 1L,
                           ifelse(country %in% c('US','USA'), 2L, 3L))
    broker_key <- toupper(trimws(as.character(df$broker)))
    broker_key[is.na(broker_key)] <- 'ZZZZZZ'
    account <- toupper(trimws(as.character(df$account_type)))
    account_rank <- match(account, c('GENERAL','ISA','IRP','PENSION'))
    account_rank[is.na(account_rank)] <- 5L
    market_value <- as.numeric(df$평가액_원)
    invalid_value <- !is.finite(market_value)
    idx <- order(broker_key, country_rank, account_rank,
                 invalid_value, -ifelse(invalid_value, 0, market_value),
                 as.integer(df$id), na.last=TRUE, method='radix')
    df[idx,,drop=FALSE]
  })

  output$tbl <- renderDT({
    # 화면 전용 사본: DB의 broker 코드와 원본 컬럼 순서는 변경하지 않습니다.
    df <- holdings()
    broker_labels <- setNames(names(choices$broker), unname(choices$broker))
    if ('fractional_share' %in% names(df))
      df$fractional_share <- ifelse(as.integer(df$fractional_share)==1L,'예','아니오')
    if(nrow(df)) {
      broker_code <- as.character(df$broker)
      label <- unname(broker_labels[broker_code])
      df$broker <- ifelse(!is.na(label), label, broker_code)
      n <- nchar(df$account_number)
      df$account_number <- paste0('****',substring(df$account_number,pmax(1,n-3),n))
    }
    # An explicit edit control does not depend on DataTables' row-selection event.
    df$편집 <- sprintf("<button type='button' class='btn btn-primary btn-xs pms-edit-holding' data-id='%d'>편집</button>", as.integer(df$id))
    # 종목명 바로 뒤에 평가액이 오도록 열을 재배치합니다.
    col_names <- names(df)
    col_names <- setdiff(col_names, '평가액_원')
    name_pos <- match('name', col_names)
    if(!is.na(name_pos)) col_names <- append(col_names, '평가액_원', after=name_pos)
    else col_names <- c(col_names, '평가액_원')
    df <- df[, col_names, drop=FALSE]
    asset_col <- as.integer(match('asset_class',names(df))-1L)
    tax_col <- as.integer(match('tax_type', names(df)) - 1L)
    datatable(df,rownames=FALSE,escape=which(names(df)!='편집'),selection=list(mode='single',selected=NULL,target='row'),
      callback=DT::JS("
        table.on('click.pmsEditDirect','tbody tr',function(e){
          if(window.jQuery && jQuery(e.target).closest('button.pms-edit-holding').length) return;
          var row=table.row(this).data();
          if(!row || !row.length) return;
          var id=Number(String(row[0]).replace(/<[^>]*>/g,''));
          if(Number.isInteger(id) && id>0 && window.Shiny && Shiny.setInputValue){
            Shiny.setInputValue('pms_row_edit_id',{id:id,nonce:Date.now()},{priority:'event'});
          }
        });
      "),
      filter='top',extensions='FixedHeader',
      options=list(columnDefs=c(list(list(
        targets=asset_col,
        render=JS("function(data,type,row,meta){
          if(type !== 'display' && type !== 'filter') return data;
          var labels={'STOCK':'주식','BOND':'채권','GOLD':'금','CASH':'현금성','OTHER':'기타'};
          return Object.prototype.hasOwnProperty.call(labels,data)?labels[data]:data;
        }")
      ),list(
        targets=tax_col,
        render=JS("function(data,type,row,meta){
          if (type !== 'display' && type !== 'filter') return data;
          var labels = {'KR_STOCK':'국내 일반주식','KR_EQUITY_ETF':'국내 주식형 ETF',
            'KR_OTHER_ETF':'국내 기타 ETF','KR_US_INDEX_ETF':'미국지수ETF',
            'US_STOCK':'미국 주식','US_ETF':'미국 ETF','KRX_GOLD':'KRX 금현물','OTHER':'기타'};
          return Object.prototype.hasOwnProperty.call(labels,data) ? labels[data] : data;
        }")
      )), pms_money_defs(df)),pageLength=100,lengthMenu=c(10,25,50,100),scrollX=TRUE,
                   scrollY='62vh',scrollCollapse=TRUE,fixedHeader=TRUE,
                   order=list(),stateSave=FALSE))
  },server=FALSE)

  # Prefer an unambiguous database ID, never a sorted DataTable row number.
  edit_holding_id <- function(id) {
    id <- suppressWarnings(as.integer(id))
    if (length(id) != 1L || is.na(id)) {
      notice('편집 ID가 전달되지 않았습니다.')
      return(invisible(NULL))
    }
    r <- get_holding(con, id)
    if (is.null(r)) {
      notice(sprintf('ID %s의 보유종목을 찾지 못했습니다.', id))
      return(invisible(NULL))
    }
    selected_id(id)
    editing_id(id)
    mode('edit')
    set_fields(r)
    notice(sprintf('편집 모드: ID %s (%s). 값을 변경하고 저장하세요.', id, r$symbol[1]))
    showNotification(paste('편집할 종목:', r$symbol[1]), type='message', duration=3)
  }

  observeEvent(holdings(), {
    d <- holdings()
    lab <- if(nrow(d)) sprintf('ID %s | %s | %s | %s', d$id, d$symbol, d$name, d$broker) else character(0)
    choice <- c('선택하세요'='', stats::setNames(as.character(d$id),lab))
    active <- isolate(editing_id())
    updateSelectInput(session,'holding_pick_id',choices=choice,
      selected=if(!is.null(active) && as.character(active) %in% as.character(d$id)) as.character(active) else '')
  })
  observeEvent(input$holding_edit_selected, {
    if (is.null(input$holding_pick_id) || !nzchar(input$holding_pick_id)) {
      notice('먼저 종목을 선택하세요.'); return()
    }
    edit_holding_id(input$holding_pick_id)
  })
  observeEvent(input$holding_edit_by_id, {
    id <- isolate(input$holding_id_manual)
    edit_holding_id(id)
  })
  observeEvent(input$pms_click_edit_id, {
    edit_holding_id(input$pms_click_edit_id$id)
  }, ignoreInit=TRUE)
  observeEvent(input$pms_row_edit_id, {
    edit_holding_id(input$pms_row_edit_id$id)
  })
  observeEvent(input$pms_edit_holding_id, {
    edit_holding_id(input$pms_edit_holding_id$id)
  })
  observeEvent(input$tbl_rows_selected, {
    ix <- input$tbl_rows_selected
    if(length(ix)!=1L || is.na(ix)) return()
    d <- isolate(holdings())
    if(ix>=1L && ix<=nrow(d)) edit_holding_id(d$id[ix])
  }, ignoreInit=TRUE)

  fmt_commas <- function(x) {
    if (is.null(x) || length(x) == 0L || is.na(x)) return('')
    format(as.numeric(x), big.mark=',', scientific=FALSE, trim=TRUE, digits=15)
  }
  parse_commas <- function(x) suppressWarnings(as.numeric(gsub(',', '', trimws(x), fixed=TRUE)))

  set_fields <- function(r=NULL) {
    val <- function(k,default) {
      if(is.null(r)) return(default)
      x <- r[[k]][1]
      if(length(x)==0L || is.na(x)) default else x
    }
    updateSelectInput(session,'f_country',selected=val('country','KR'))
    updateTextInput(session,'f_symbol',value=val('symbol',''))
    updateTextInput(session,'f_name',value=val('name',''))
    updateSelectInput(session,'f_currency',selected=val('currency','KRW'))
    updateSelectInput(session,'f_asset_class',selected=val('asset_class','OTHER'))
    updateSelectInput(session,'f_tax_type',selected=val('tax_type','OTHER'))
    # Old CSV may contain a broker/account type not in the defaults.
    broker <- val('broker','SAMSUNG')
    bchoices <- choices$broker
    if(!broker %in% unname(bchoices)) bchoices <- c(bchoices,setNames(broker,broker))
    updateSelectInput(session,'f_broker',choices=bchoices,selected=broker)
    acc <- val('account_type','GENERAL')
    achoices <- choices$account_type
    if(!acc %in% unname(achoices)) achoices <- c(achoices,setNames(acc,acc))
    updateSelectInput(session,'f_account_type',choices=achoices,selected=acc)
    updateTextInput(session,'f_account_number_manual',value='')
    selected_account <- as.character(val('account_number',''))
    current_accounts <- dbGetQuery(con,
      "SELECT DISTINCT account_number FROM holdings WHERE broker = ? AND TRIM(account_number) <> '' ORDER BY account_number",
      params=list(as.character(broker)))$account_number
    updateSelectInput(session,'f_account_number',
      choices=c('새 계좌번호 직접 입력'='',stats::setNames(unique(c(current_accounts,selected_account[nzchar(selected_account)])),unique(c(current_accounts,selected_account[nzchar(selected_account)])))),selected=selected_account)
    updateSelectInput(session,'f_fractional_share',selected=as.character(val('fractional_share',0)))
    updateNumericInput(session,'f_quantity',value=val('quantity',0))
    updateTextInput(session,'f_avg_price',value=fmt_commas(val('avg_price',0)))
    updateTextInput(session,'f_avg_fx_rate',value=fmt_commas(val('avg_fx_rate',1)))
    updateTextInput(session,'f_cost_krw',value=as.character(val('cost_krw','')))
    updateTextAreaInput(session,'f_memo',value=val('memo',''))
    updated_at_display(as.character(val('updated_at', '기록 없음')))
  }

  # Reuse account numbers from the selected broker. Keep manually typed numbers,
  # and refresh the options after a new holding is saved or removed.
  observeEvent(input$f_broker, {
    broker <- input$f_broker
    if (is.null(broker) || !nzchar(broker)) return()
    accounts <- dbGetQuery(con,
      "SELECT DISTINCT account_number FROM holdings WHERE broker = ? AND TRIM(account_number) <> '' ORDER BY account_number",
      params=list(as.character(broker)))$account_number
    current <- isolate(input$f_account_number)
    if (length(current)!=1L || is.na(current)) current <- ''
    updateSelectInput(session,'f_account_number',
      choices=c('새 계좌번호 직접 입력'='',stats::setNames(unique(c(accounts,current[nzchar(current)])),unique(c(accounts,current[nzchar(current)])))), selected=current)
  }, ignoreInit=FALSE)

  observeEvent(input$new, {
    mode('new'); editing_id(NULL); selected_id(NULL);
    DT::selectRows(DT::dataTableProxy('tbl', session=session), NULL);
    set_fields(); notice('신규 입력 모드입니다.')
  })

  # The save button is always present: no dynamically created modal footer/event binding.
  read_form <- function() {
    required <- c('f_country','f_symbol','f_currency','f_asset_class','f_tax_type',
      'f_broker','f_account_type','f_fractional_share','f_quantity','f_avg_price','f_avg_fx_rate')
    absent <- required[vapply(required,function(k) length(input[[k]])!=1L || is.na(input[[k]]),logical(1))]
    if(length(absent)) stop(paste('필수 입력 누락:',paste(absent,collapse=', ')))
    symbol <- toupper(trimws(input$f_symbol))
    manual_account <- trimws(if(is.null(input$f_account_number_manual)) '' else input$f_account_number_manual)
    selected_account <- trimws(if(is.null(input$f_account_number)) '' else input$f_account_number)
    account <- if(nzchar(manual_account)) manual_account else selected_account
    if(!nzchar(symbol)||!nzchar(account)) stop('종목코드와 계좌번호는 필수입니다.')
    nums <- c(input$f_quantity,parse_commas(input$f_avg_price),parse_commas(input$f_avg_fx_rate))
    if(any(!is.finite(nums)) || any(nums[1:2]<0) || nums[3]<=0) stop('수량/가격/환율을 확인하세요.')
    ctxt <- trimws(if(is.null(input$f_cost_krw)) '' else input$f_cost_krw)
    cost <- if(nzchar(ctxt)) suppressWarnings(as.numeric(gsub(',','',ctxt,fixed=TRUE))) else NA_real_
    if(nzchar(ctxt)&&(!is.finite(cost)||cost<0)) stop('원화 취득원가는 0 이상 숫자 또는 공란이어야 합니다.')
    record <- list(
      country=input$f_country,symbol=symbol,
      name=if(is.null(input$f_name)) '' else trimws(input$f_name),
      currency=input$f_currency,asset_class=input$f_asset_class,tax_type=input$f_tax_type,
      broker=input$f_broker,account_type=input$f_account_type,account_number=account,
      fractional_share=as.integer(input$f_fractional_share),
      quantity=as.numeric(input$f_quantity),avg_price=parse_commas(input$f_avg_price),
      avg_fx_rate=parse_commas(input$f_avg_fx_rate),cost_krw=cost,
      memo=if(is.null(input$f_memo)) '' else input$f_memo,
      updated_at=NA_character_  # R/db.R이 실제 저장 시각을 자동 결정
    )
    stopifnot(identical(names(record),holdings_columns))
    record
  }

  observeEvent(input$save_record, {
    notice(sprintf('저장 버튼 수신: %s / %s',mode(),Sys.time()))
    tryCatch({
      r <- read_form()
      if(identical(mode(),'edit')) {
        id <- editing_id()
        if(is.null(id)) stop('수정 대상 ID가 없습니다. 표에서 해당 행을 다시 선택하세요.')
        n <- update_holding(con,id,r)
        if(n == 0L) {
          if(is.null(get_holding(con,id))) stop('수정 대상이 DB에서 사라졌습니다. 다시 선택하세요.')
          notice(sprintf('변경된 내용이 없습니다: ID %s (%s)',id,r$symbol))
          return()
        }
        if(n!=1L) stop('UPDATE된 레코드가 없습니다. 데이터베이스 내용을 확인하세요.')
        message(sprintf('[PMS2] UPDATE ID=%d, rows=%d',id,n))
        updated_at_display(as.character(get_holding(con,id)$updated_at[1]))
        notice(sprintf('수정 완료: ID %s (%s) / %s',id,r$symbol, get_holding(con,id)$updated_at[1]))
      } else {
        n <- insert_holding(con,r)
        message(sprintf('[PMS2] INSERT rows=%d',n))
        notice(sprintf('등록 완료: %s',r$symbol))
        updated_at_display('저장 완료 (목록에서 행 선택 시 확인)')
      }
      session$sendCustomMessage('pms_holdings_preserve_position', list())
      refresh(refresh()+1L)
      showNotification('저장 완료',type='message',duration=4)
    },error=function(e){
      msg <- paste('저장 실패:',conditionMessage(e))
      message('[PMS2] ',msg)
      notice(msg)
      showNotification(msg,type='error',duration=NULL)
    })
  })

  observeEvent(input$delete_record, {
    id <- selected_id()
    if(is.null(id)) {notice('먼저 표에서 삭제할 행을 선택하세요.');return()}
    r <- get_holding(con,id)
    if(is.null(r)) {notice('삭제 대상이 없습니다.');return()}
    showModal(modalDialog(title='삭제 확인',
      paste(r$name[1], ' / ID ',id,'를 영구 삭제하시겠습니까?'),
      footer=tagList(modalButton('취소'),actionButton('confirm_delete','영구 삭제',class='btn-danger'))))
  })
  observeEvent(input$confirm_delete, {
    id <- selected_id()
    tryCatch({
      if(is.null(id)) stop('삭제 ID가 없습니다.')
      delete_holding(con,id);removeModal();refresh(refresh()+1L)
      selected_id(NULL);mode('new');editing_id(NULL)
      notice(paste('삭제 완료: ID',id));showNotification('삭제 완료',type='message')
    },error=function(e){notice(paste('삭제 실패:',conditionMessage(e)))})
  })

  # 탭 진입 시 SQLite 저장값으로 화면을 다시 갱신한다 (API 접속 없음).
  observeEvent(input$pms_tabs, {
    if (identical(input$pms_tabs, '시세수집 · 자산평가')) {
      market_refresh(isolate(market_refresh()) + 1L)
      dividend_refresh(isolate(dividend_refresh()) + 1L)
    }
  }, ignoreInit=FALSE)

  # 처음 실행 시 저장값을 로드하는 반응형 출력의 버전을 초기화한다.
  # 신규 시세/배당 API 요청은 시작 시 실행하지 않는다.
  observeEvent(TRUE, {
    market_refresh(isolate(market_refresh())+1L)
    dividend_refresh(isolate(dividend_refresh())+1L)
    bench_refresh(isolate(bench_refresh())+1L)
  }, once=TRUE, priority=100)

}
# Windows RStudio Run App: prefer an external maximized Chrome window rather than Viewer.
# The dedicated PMS2_Chrome_실행.bat remains the most reliable startup method.
pms_default_chrome_launch <- function(url) {
  if (.Platform$OS.type!='windows') return(utils::browseURL(url))
  dirs <- c(Sys.getenv('PROGRAMFILES'),Sys.getenv('PROGRAMFILES(X86)'),Sys.getenv('LOCALAPPDATA'))
  chrome <- file.path(dirs,'Google','Chrome','Application','chrome.exe')
  chrome <- chrome[file.exists(chrome)]
  if(!length(chrome))return(utils::browseURL(url))
  tryCatch(system2(chrome[1],c('--new-window','--start-maximized',shQuote(url)),
                   wait=FALSE,stdout=FALSE,stderr=FALSE),error=function(e)utils::browseURL(url))
}
# LAN 내 다른 클라이언트 접속을 기본 허용합니다. 방화벽 허용은 별도 필요.
options(shiny.host='0.0.0.0', shiny.port=3838,
        shiny.launch.browser=pms_default_chrome_launch)
shinyApp(ui,server,options=list(host='0.0.0.0', port=3838,
                                launch.browser=pms_default_chrome_launch))

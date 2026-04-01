---
tags: [project, 발주, 분석, ERP, 재고]
status: active
---
# frequency
> 재고 분석 및 발주 빈도 분석 / 인보이스 생성

## 개요
3개월간 약 300건의 주문 데이터를 분석하여 **반복적으로 주문되는 품목**을 식별하고, 발주 시 선제적으로 대량 생산할 재고 추천 리스트를 제공하는 시스템.

## 분석 방향
- 주문 품목 출현 빈도(frequency) 기반 랭킹
- 일정 주기로 반복되는 패턴 탐지
- 소량 다빈도 vs 대량 소빈도 품목 구분
- **재고 선생산 대상 품목 우선순위** 산정

## 주요 파일
| 파일 | 역할 |
|------|------|
| [make_invoice_v2.py](make_invoice_v2.py) | 인보이스 자동 생성 스크립트 |
| [refresh.py](refresh.py) | 데이터 갱신 |
| [server.py](server.py) | 로컬 서버 |
| [apps_script.js](apps_script.js) | Google Apps Script 연동 |
| [data500.json](data500.json) | 분석 데이터 (500건) |
| [data_wr.json](data_wr.json) | 쓰기용 데이터 |
| [데이터_역할_설명.html](데이터_역할_설명.html) | 데이터 구조 설명 문서 |
| [start.bat](start.bat) | 실행 배치 파일 |

## 관련
- [[BUY]] — 구매 발주서 관리 (이 분석 결과를 활용)
- [[ERP_downloaded_data]] — ERP 원본 데이터

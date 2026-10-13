# Ocean Fleet Atlas

사용자가 제공한 원본 단일 파일 `index.html`의 디자인과 상호작용을 유지하면서, OneDrive의 공개 Excel 통합문서를 GitHub Actions에서 검증된 데이터로 변환해 GitHub Pages에 반영하는 정적 대시보드입니다.

## 제공 기능

- Leaflet 지도에서 200척의 현재 위치와 선택 선박의 예정 항로 표시
- 선명, 선박 ID, MMSI, IMO 및 항로 검색
- 운항 상태와 최종 목적지 필터
- 선박 제원, 운항·환경 제한, 용선 정보 및 연락처 상세 보기
- 원본 3층 차량 적재 시각화와 조작 기능(최초 배치는 Excel의 층별 배정량 반영)
- 공개 OneDrive Excel을 5분마다 동기화하고 데이터 변경 시 GitHub Pages 재배포
- 타임스탬프 쿼리를 통한 브라우저 캐시 우회

## 데이터 흐름

1. `scripts/download_excel.py`가 공개 공유 URL에 `download=1`을 추가하고 SharePoint 익명 쿠키를 유지해 `.xlsx`를 받습니다.
2. `scripts/fleet_sync.py`가 정확히 7개 시트와 열 머리글을 검사합니다.
3. 선박, 실시간 위치, 항차, 항구, 화물, 용선 및 연락처를 키로 결합합니다.
4. 전체 가용량과 3개 데크별 가용량을 배정 화물에서 다시 계산하고 Excel의 캐시된 수식 결과와 대조합니다.
5. 참조 무결성, 좌표, 용량 및 콘텐츠 해시를 검증한 임시 JSON과 브라우저용 JavaScript를 각각 `data/fleet-data.json`, `data/fleet-data.js`로 교체합니다.
6. 데이터가 실제로 달라졌을 때만 커밋하고 `main`에서는 같은 실행 안에서 Pages를 다시 배포합니다.

동기화 중 다운로드, 변환, 테스트 또는 검증이 하나라도 실패하면 교체와 커밋 단계에 도달하지 않으므로 마지막으로 검증된 데이터 파일들이 보존됩니다.

## 로컬 실행

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python scripts/download_excel.py "<공개 Excel URL>" /tmp/fleet.xlsx
python scripts/sync_fleet_data.py /tmp/fleet.xlsx data/fleet-data.json \
  --javascript-output data/fleet-data.js \
  --previous data/fleet-data.json
python -m http.server 8000
```

브라우저에서 `http://localhost:8000`을 엽니다. 페이지는 원본의 인라인 선대 데이터를 대신해 `data/fleet-data.js`를 동기적으로 불러온 뒤 기존 Leaflet·검색·필터·상세·항로·3층 시각화 코드를 그대로 실행합니다.

## 저장소 설정

병합 후 다음 설정이 필요합니다.

1. **Settings → Actions → General → Workflow permissions**에서 `Read and write permissions`를 허용합니다.
2. **Settings → Pages → Build and deployment → Source**를 `GitHub Actions`로 선택합니다.
3. `Sync fleet data`를 수동 실행해 최초 동기화를 확인합니다. 이후 예약 실행은 GitHub 스케줄러가 담당합니다.

GitHub의 예약 실행은 5분 cron으로 등록되어도 부하에 따라 몇 분 지연될 수 있습니다.

## 데이터 계약

페이지는 `data/fleet-data.js?v=<현재 시각>`을 불러와 `window.FLEET_DATA`에 저장합니다. JSON 파일은 검증·감사 용도로 동일한 내용을 보관합니다. 주요 컬렉션은 원본 HTML의 계약과 같습니다.

- `vessel`, `vessel_live`, `voyage`, `cargo`, `port`: Excel의 대응 시트 전체 행
- `maintenance_schedule`, `charter_status`, `fleet_allocation`, `port_compatibility`: Excel 열과 관계에서 생성한 원본 화면 호환 컬렉션
- `meta.content_hash`: 화면 데이터 전체의 결정적 SHA-256

머리글이 바뀌면 묵시적으로 잘못 매핑하지 않고 동기화가 실패하도록 설계했습니다. 스키마를 변경하려면 `SHEET_HEADERS`, 변환 로직과 테스트를 함께 수정해야 합니다.

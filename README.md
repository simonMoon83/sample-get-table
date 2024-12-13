# 데이터베이스 테이블 명세서 생성기

이 프로그램은 데이터베이스 스키마를 읽어서 Excel 형식의 테이블 명세서를 자동으로 생성합니다.

## 지원 데이터베이스

- MySQL
- PostgreSQL
- SQLite
- MSSQL
- Oracle

## 기능

- 데이터베이스의 모든 테이블 정보 추출
- 각 테이블의 컬럼 정보 (이름, 데이터 타입, Nullable 여부 등) 수집
- Primary Key 및 Foreign Key 관계 표시
- 깔끔한 Excel 형식의 출력
- 자동 열 너비 조정
- 생성 일시 기록
- 테이블 코멘트 및 컬럼 코멘트 지원 (Oracle)

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. Oracle 데이터베이스 사용을 위한 추가 설치 (Oracle 사용 시에만 필요):
   1. [Oracle Instant Client](https://www.oracle.com/database/technologies/instant-client/downloads.html) 다운로드
      - Windows x64 Basic Package 또는 Basic Light Package 다운로드
   2. 다운로드 받은 zip 파일을 원하는 경로에 압축 해제 (예: C:\oracle\instantclient_21_13)
   3. 압축 해제한 경로를 시스템 환경 변수 PATH에 추가
   4. 시스템 환경 변수에 TNS_ADMIN 추가 (필요한 경우)
      - 값: Oracle Instant Client 설치 경로 (예: C:\oracle\instantclient_21_13)

## 사용 방법

### 일반 데이터베이스
1. `db_schema_to_excel.py` 파일을 열어 데이터베이스 연결 문자열을 수정합니다.
2. 프로그램 실행:
```bash
python db_schema_to_excel.py
```

### Oracle 데이터베이스
1. `oracle_schema_to_excel.py` 파일을 열어 데이터베이스 연결 정보를 수정합니다.
2. 프로그램 실행:
```bash
python oracle_schema_to_excel.py
```

## 연결 문자열 예시

- MySQL: "mysql+pymysql://username:password@localhost:3306/database_name"
- PostgreSQL: "postgresql://username:password@localhost:5432/database_name"
- SQLite: "sqlite:///database.db"
- MSSQL: "mssql+pyodbc://username:password@server_name/database_name?driver=SQL+Server"
- Oracle: "oracle+oracledb://username:password@hostname:1521/service_name"

## 출력 결과

프로그램은 다음 정보를 포함한 Excel 파일을 생성합니다:

- 테이블명
- 테이블 코멘트 (Oracle)
- 컬럼명
- 컬럼 코멘트 (Oracle)
- 데이터 타입
- Nullable 여부
- Primary Key 여부
- Foreign Key 여부
- Foreign Key 참조 정보

## 주의사항

- Oracle 데이터베이스 사용 시 Oracle Instant Client가 시스템에 설치되어 있어야 합니다.
- Oracle 연결 문자열의 hostname은 tnsnames.ora에 정의된 서비스 이름을 사용할 수도 있습니다.

"""
오라클 데이터베이스 스키마를 Excel 명세서로 추출
"""
import pandas as pd
import oracledb
from db_schema_utils import apply_sheet_style, get_output_file_name
from openpyxl.styles import Font, Alignment

def get_column_info(connection, table_name, owner):
    """Get column information for a specific table using direct SQL queries"""
    column_query = """
        SELECT 
            c.COLUMN_NAME,
            c.DATA_TYPE,
            CASE 
                WHEN c.DATA_TYPE = 'NUMBER' AND c.DATA_PRECISION IS NOT NULL 
                THEN c.DATA_PRECISION || ',' || c.DATA_SCALE
                WHEN c.DATA_TYPE LIKE '%CHAR%' OR c.DATA_TYPE = 'NVARCHAR2'
                THEN c.CHAR_LENGTH
                ELSE NULL 
            END as DATA_LENGTH,
            c.NULLABLE,
            CASE WHEN p.COLUMN_NAME IS NOT NULL THEN 'Y' ELSE 'N' END as IS_PRIMARY_KEY,
            cc.COMMENTS
        FROM ALL_TAB_COLUMNS c
        LEFT JOIN (
            SELECT cols.TABLE_NAME, cols.COLUMN_NAME
            FROM ALL_CONSTRAINTS cons
            JOIN ALL_CONS_COLUMNS cols ON cons.CONSTRAINT_NAME = cols.CONSTRAINT_NAME
            WHERE cons.CONSTRAINT_TYPE = 'P'
            AND cons.OWNER = :owner
        ) p ON c.TABLE_NAME = p.TABLE_NAME AND c.COLUMN_NAME = p.COLUMN_NAME
        LEFT JOIN ALL_COL_COMMENTS cc 
            ON c.TABLE_NAME = cc.TABLE_NAME 
            AND c.COLUMN_NAME = cc.COLUMN_NAME
            AND cc.OWNER = :owner
        WHERE c.TABLE_NAME = :table_name
        AND c.OWNER = :owner
        ORDER BY c.COLUMN_ID
    """
    
    cursor = connection.cursor()
    cursor.execute(column_query, {'table_name': table_name, 'owner': owner})
    columns = cursor.fetchall()
    
    fk_query = """
        SELECT 
            cols.COLUMN_NAME,
            cons.R_OWNER || '.' || cons.R_TABLE_NAME || '.' || rcols.COLUMN_NAME as REFERENCED_TABLE
        FROM ALL_CONSTRAINTS cons
        JOIN ALL_CONS_COLUMNS cols 
            ON cons.CONSTRAINT_NAME = cols.CONSTRAINT_NAME
            AND cons.OWNER = cols.OWNER
        JOIN ALL_CONS_COLUMNS rcols 
            ON cons.R_CONSTRAINT_NAME = rcols.CONSTRAINT_NAME
            AND cons.R_OWNER = rcols.OWNER
        WHERE cons.CONSTRAINT_TYPE = 'R'
        AND cons.OWNER = :owner
        AND cons.TABLE_NAME = :table_name
    """
    
    cursor.execute(fk_query, {'table_name': table_name, 'owner': owner})
    fks = {row[0]: row[1] for row in cursor.fetchall()}
    
    result = []
    for col in columns:
        col_name = col[0]
        data_type = col[1]
        data_length = col[2]
        nullable = col[3]
        is_pk = col[4]
        comments = col[5] or ''
        
        type_with_length = f"{data_type}({data_length})" if data_length else data_type
        fk_ref = fks.get(col_name, '')
        
        result.append({
            '컬럼명': col_name,
            '데이터 타입': type_with_length,
            'Nullable': 'Y' if nullable == 'Y' else 'N',
            'PK': is_pk,
            'FK': 'Y' if fk_ref else 'N',
            'FK 참조': fk_ref,
            '설명': comments
        })
    
    return result

def get_table_comment(connection, table_name, owner):
    """Get table comment"""
    query = """
        SELECT COMMENTS
        FROM ALL_TAB_COMMENTS
        WHERE TABLE_NAME = :table_name
        AND OWNER = :owner
    """
    cursor = connection.cursor()
    cursor.execute(query, {'table_name': table_name, 'owner': owner})
    result = cursor.fetchone()
    return result[0] if result and result[0] else ''

def get_tables(connection, owner):
    """Get all tables for the specified owner"""
    query = """
        SELECT TABLE_NAME
        FROM ALL_TABLES
        WHERE OWNER = :owner
        ORDER BY TABLE_NAME
    """
    cursor = connection.cursor()
    cursor.execute(query, {'owner': owner})
    return [row[0] for row in cursor.fetchall()]

def create_table_specification(username, password, hostname, port, service_name, owner):
    """
    데이터베이스 스키마를 읽어서 Excel 형식의 테이블 명세서를 생성합니다.
    
    Args:
        username (str): 데이터베이스 사용자 이름
        password (str): 데이터베이스 비밀번호
        hostname (str): 호스트 이름
        port (int): 포트 번호
        service_name (str): 서비스 이름
        owner (str): 스키마 소유자
    """
    try:
        # Thin 모드로 연결
        dsn = f"{hostname}:{port}/{service_name}"
        connection = oracledb.connect(user=username,
                                    password=password,
                                    dsn=dsn,
                                    config_dir=None,
                                    lib_dir=None)
        
        tables = get_tables(connection, owner)
        
        # Excel 파일 생성
        output_file = get_output_file_name('oracle_tables')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for table_name in tables:
                # 테이블 정보 수집
                columns = get_column_info(connection, table_name, owner)
                table_comment = get_table_comment(connection, table_name, owner)
                
                # DataFrame 생성
                df = pd.DataFrame(columns)
                
                # 시트 이름으로 사용할 수 있는 길이로 조정
                sheet_name = table_name[:31]  # Excel 시트 이름 제한
                
                # DataFrame을 Excel 시트에 저장
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 시트 스타일 적용
                ws = writer.sheets[sheet_name]
                
                # 목차로 돌아가기 링크 추가
                ws['A1'] = '목차로 돌아가기'
                ws['A1'].hyperlink = '#목차!A1'
                ws['A1'].font = Font(color="0563C1", underline="single")
                ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
                
                # 스타일 적용
                apply_sheet_style(ws, df)
        
        connection.close()
        print(f"오라클 테이블 명세서가 생성되었습니다: {output_file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        print("Help: https://python-oracledb.readthedocs.io/en/latest/user_guide/troubleshooting.html")

if __name__ == "__main__":
    # 오라클 연결 정보
    username = "username"
    password = "password"
    hostname = "hostname"
    port = 1521
    service_name = "service_name"
    owner = "SCHEMA_NAME"  # 스키마 이름을 대문자로 지정
    
    create_table_specification(username, password, hostname, port, service_name, owner)

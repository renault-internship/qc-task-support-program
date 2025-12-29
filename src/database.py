"""
SQLite 데이터베이스 관리 모듈
기업정보 저장 및 조회
"""
import sqlite3
from pathlib import Path
from typing import Optional, Dict, Any, List

# 데이터베이스 파일 경로
DB_PATH = Path("data/companies.db")


def init_database():
    """데이터베이스 초기화 및 테이블 생성"""
    # data 폴더가 없으면 생성
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL UNIQUE,
            mileage_threshold INTEGER DEFAULT 50000,
            warranty_years INTEGER DEFAULT 2,
            sheet_index INTEGER DEFAULT 0,
            header_row INTEGER DEFAULT 3,
            data_start_row INTEGER DEFAULT 4,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    conn.commit()
    conn.close()


def get_company_info(company_name: str) -> Optional[Dict[str, Any]]:
    """기업정보 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT * FROM companies WHERE company_name = ?
    """, (company_name,))
    
    row = cursor.fetchone()
    conn.close()
    
    if row:
        return dict(row)
    return None


def get_all_companies() -> List[str]:
    """모든 기업명 목록 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute("SELECT company_name FROM companies ORDER BY company_name")
    rows = cursor.fetchall()
    conn.close()
    
    return [row[0] for row in rows]


def upsert_company(company_name: str, **kwargs):
    """
    기업정보 저장/업데이트
    
    Args:
        company_name: 기업명
        **kwargs: mileage_threshold, warranty_years, sheet_index, header_row, data_start_row 등
    """
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    # 기존 데이터 확인
    existing = get_company_info(company_name)
    
    if existing:
        # 업데이트
        updates = ", ".join([f"{k} = ?" for k in kwargs.keys()])
        updates += ", updated_at = CURRENT_TIMESTAMP"
        values = list(kwargs.values()) + [company_name]
        
        cursor.execute(f"""
            UPDATE companies 
            SET {updates}
            WHERE company_name = ?
        """, values)
    else:
        # 삽입
        keys = ["company_name"] + list(kwargs.keys())
        placeholders = ", ".join(["?"] * len(keys))
        values = [company_name] + list(kwargs.values())
        
        cursor.execute(f"""
            INSERT INTO companies ({", ".join(keys)})
            VALUES ({placeholders})
        """, values)
    
    conn.commit()
    conn.close()


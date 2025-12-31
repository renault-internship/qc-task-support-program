"""
SAP 테이블에 협력사 데이터 대량 추가 스크립트
- 중복된 협력사는 제외
- 룰 테이블이 이미 있으면 건너뛰기
사용법: python insert_sap_data_bulk.py
"""
import sqlite3
from pathlib import Path
from src.database import init_database, get_company_info, create_rule_table, DB_PATH

# 데이터베이스 초기화
init_database()

# 추가할 데이터
suppliers = [
    {"sap_name": "Hanrim Intech", "sap_code": "I806", "renault_code": "247744"},
    {"sap_name": "SMC Co.", "sap_code": "C508", "renault_code": "250034"},
    {"sap_name": "Dk Austech", "sap_code": "C202", "renault_code": "247751"},
    {"sap_name": "HBPO", "sap_code": "B933", "renault_code": "260864"},
    {"sap_name": "Dongwon Tech.", "sap_code": "I201", "renault_code": "247750"},
    {"sap_name": "AMS", "sap_code": "B907", "renault_code": "247736"},
    {"sap_name": "Kongsberg Automotive Wuxi", "sap_code": "Z239", "renault_code": "25524"},
    {"sap_name": "Autoliv KOREA", "sap_code": "I505", "renault_code": "247734"},
    {"sap_name": "LS Mtron", "sap_code": "R805", "renault_code": "253320"},
    {"sap_name": "HanJoo metal", "sap_code": "G802", "renault_code": "247748"},
    {"sap_name": "DAEWON KANG UP", "sap_code": "C702", "renault_code": "270538"},
    {"sap_name": "Cheil Elect.", "sap_code": "I602", "renault_code": "247845"},
    {"sap_name": "DY AUTO(Dongyang Mechatronics)", "sap_code": "B923", "renault_code": "253276"},
    {"sap_name": "Nexen Tech", "sap_code": "E703", "renault_code": "247836"},
    {"sap_name": "Nobel", "sap_code": "C910", "renault_code": "274799"},
    {"sap_name": "Daehan Calsonic", "sap_code": "A201", "renault_code": "253303"},
    {"sap_name": "Kunhwa", "sap_code": "C103", "renault_code": "247742"},
    {"sap_name": "S.P.L", "sap_code": "I202", "renault_code": "247743"},
    {"sap_name": "Dgenx (Dongwon Tech_Haman)", "sap_code": "C514", "renault_code": "262210"},
    {"sap_name": "GMB Korea", "sap_code": "G307", "renault_code": "261852"},
    {"sap_name": "ZF Sachs Korea Co.", "sap_code": "C906", "renault_code": "247746"},
    {"sap_name": "Continental Automotive (Cheongwon)", "sap_code": "C810", "renault_code": "253290"},
    {"sap_name": "Central CMS", "sap_code": "C404", "renault_code": "250030"},
    {"sap_name": "COAVIS", "sap_code": "C924", "renault_code": "288041"},
    {"sap_name": "Shinsung technics", "sap_code": "I908", "renault_code": "247747"},
    {"sap_name": "Valeo Automotive Korea", "sap_code": "I910", "renault_code": "257848"},
    {"sap_name": "Valeo Shenzhen", "sap_code": "Z204", "renault_code": "265287"},
    {"sap_name": "Webasto Dong hee", "sap_code": "B946", "renault_code": "247840"},
    {"sap_name": "Hankuk Sekurit", "sap_code": "I803", "renault_code": "247830"},
    {"sap_name": "ADI", "sap_code": "I217", "renault_code": "247733"},
    {"sap_name": "Kumho Tire", "sap_code": "C102", "renault_code": "253234"},
    {"sap_name": "Kwang Jin", "sap_code": "B910", "renault_code": "283779"},
    {"sap_name": "KBWS", "sap_code": "B928", "renault_code": "253289"},
    {"sap_name": "Jinyoung Elect.", "sap_code": "E601", "renault_code": "247838"},
    {"sap_name": "PHA", "sap_code": "B508", "renault_code": "266626"},
    {"sap_name": "Magna Power Korea (한온시스템)", "sap_code": "G804", "renault_code": "253301"},
    {"sap_name": "Hwa Seuug T&C", "sap_code": "B932", "renault_code": "269448"},
    {"sap_name": "Robert Bosch Malaysia", "sap_code": "Z117", "renault_code": "20963"},
    {"sap_name": "SJ&S (성주음향)", "sap_code": "E920", "renault_code": "253251"},
    {"sap_name": "SKF Aut.", "sap_code": "C222", "renault_code": "247745"},
    {"sap_name": "Valeo Elect. Korea", "sap_code": "G302", "renault_code": "253316"},
    {"sap_name": "Visteon Electronics korea", "sap_code": "E915", "renault_code": "29179"},
    {"sap_name": "Winner COM", "sap_code": "E604", "renault_code": "271083"},
    {"sap_name": "Wookyung M.I.T", "sap_code": "I100", "renault_code": "253334"},
    {"sap_name": "Yooil rubber", "sap_code": "B904", "renault_code": "275286"},
    {"sap_name": "Yujin-Reydel(Visteon korea)", "sap_code": "I904", "renault_code": "248587"},
    {"sap_name": "다우정밀공업", "sap_code": "C203", "renault_code": "250026"},
    {"sap_name": "ILJIN GLOBAL", "sap_code": "C940", "renault_code": "287639"},
    {"sap_name": "유수로지스틱스 (GKN Driveline Korea)", "sap_code": "C802", "renault_code": "253346"},
    {"sap_name": "Hutchinsonkorea", "sap_code": "C933", "renault_code": "289009"},
    {"sap_name": "Daesung Eltec", "sap_code": "E910", "renault_code": "253255"},
    {"sap_name": "DTR", "sap_code": "R201", "renault_code": "251697"},
    {"sap_name": "BROSE KOREA LTD.", "sap_code": "X081", "renault_code": "267145"},
    {"sap_name": "Erae Auto(Korea Delphi)", "sap_code": "C814", "renault_code": "247731"},
    {"sap_name": "HUMAX AUTOMOTIVE", "sap_code": "E909", "renault_code": "269779"},
    {"sap_name": "Alps Korea", "sap_code": "E801", "renault_code": "250025"},
    {"sap_name": "Corea Elect.", "sap_code": "E101", "renault_code": "253296"},
    {"sap_name": "Samick Kiriu", "sap_code": "G405", "renault_code": "247739"},
    {"sap_name": "Samsong(Daechang)", "sap_code": "I405", "renault_code": "253262"},
    {"sap_name": "Shinsung Automotive", "sap_code": "I922", "renault_code": "247747"},
    {"sap_name": "Tung Thih Electronic", "sap_code": "Z286", "renault_code": "269258"},
    {"sap_name": "CK Malaysia", "sap_code": "Z235", "renault_code": "270028"},
    {"sap_name": "TRW PRC (Anting plant in Shanghai", "sap_code": "Z480", "renault_code": "294491"},
    {"sap_name": "TRW PRC (Wuhan)", "sap_code": "Z479", "renault_code": "293666"},
    {"sap_name": "SMR China", "sap_code": "Z383", "renault_code": "291989"},
    {"sap_name": "TRW Steering", "sap_code": "C905", "renault_code": "253270"},
    {"sap_name": "Taizhou Valeo Wenling Automotive", "sap_code": "Z111", "renault_code": "248581"},
    {"sap_name": "CAP", "sap_code": "C401", "renault_code": "253310"},
    {"sap_name": "영천배기시스템 (포레시아배기)", "sap_code": "C939", "renault_code": "258930"},
    {"sap_name": "서한워너", "sap_code": "G930", "renault_code": "293698"},
    {"sap_name": "APTIV Philiphine", "sap_code": "Z432", "renault_code": "290153"},
    {"sap_name": "sumitomo (Mitsubishi Corporation Technos)", "sap_code": "Z543", "renault_code": "286338"},
    {"sap_name": "Autoliv Rumania", "sap_code": "E00314", "renault_code": "612056"},
    {"sap_name": "Yanfeng Viesteon", "sap_code": "Z389", "renault_code": "281077"},
    {"sap_name": "Valeo Compressor Europe s.r.o.", "sap_code": "Z246", "renault_code": "119894"},
    {"sap_name": "Bando Korea Co. Ltd", "sap_code": "G912", "renault_code": "262780"},
    {"sap_name": "Bosch(Kamco)", "sap_code": "E702", "renault_code": "253289"},
    {"sap_name": "Continental Automotive Sys. China (Shanghai)", "sap_code": "Z369", "renault_code": "282803"},
    {"sap_name": "Calsonickansei Korea I", "sap_code": "C902", "renault_code": "248588"},
    {"sap_name": "Cooper Standard", "sap_code": "R603", "renault_code": "247832"},
    {"sap_name": "SJK(Sejin Elect)", "sap_code": "E902", "renault_code": "253249"},
    {"sap_name": "Kiekert CS", "sap_code": "Z308", "renault_code": "274770"},
    {"sap_name": "LG에너지 솔루션", "sap_code": "E928", "renault_code": "273107"},
    {"sap_name": "Kwang shin", "sap_code": "G103", "renault_code": "253364"},
    {"sap_name": "Donghyun Mahle", "sap_code": "G210", "renault_code": "253352"},
    {"sap_name": "Eberspacher", "sap_code": "C704", "renault_code": "271370"},
    {"sap_name": "Guangzhou Fuyao Glass CO.,Ltd", "sap_code": "Z472", "renault_code": "405092"},
    {"sap_name": "Kongsberg Automotive", "sap_code": "G812", "renault_code": "276896"},
    {"sap_name": "Kiekert AG", "sap_code": "Z262", "renault_code": "28701"},
    {"sap_name": "Kiekert", "sap_code": "Z260", "renault_code": "28701"},
    {"sap_name": "HYUNDAI MOBIS", "sap_code": "E603", "renault_code": "270985"},
    {"sap_name": "Hankook Datwyler", "sap_code": "R801", "renault_code": "247835"},
    {"sap_name": "JTEKT", "sap_code": "Z157", "renault_code": "262990"},
    {"sap_name": "Inalfa Korea", "sap_code": "B205", "renault_code": "253279"},
    {"sap_name": "LS Cable", "sap_code": "E911", "renault_code": "286363"},
    {"sap_name": "Continental Automotive(icheon)", "sap_code": "G417", "renault_code": "269411"},
    {"sap_name": "Plastic Omnium(Inergy)", "sap_code": "C204", "renault_code": "247740"},
    {"sap_name": "ILHEUNG", "sap_code": "B908", "renault_code": "284795"},
    {"sap_name": "DAEHEUNG RUBBER AND TECH", "sap_code": "C106", "renault_code": "250038"},
    {"sap_name": "Delphi Korea(패커드코리아)", "sap_code": "E205", "renault_code": "258829"},
    {"sap_name": "DELKO", "sap_code": "E914", "renault_code": "290231"},
    {"sap_name": "Michang Cable(KMCS)", "sap_code": "C915", "renault_code": "250021"},
    {"sap_name": "MANDO CORPORATION (원주)", "sap_code": "C930", "renault_code": "265566"},
    {"sap_name": "Lear Automotive Electronics(Z456) (by Lear Wofe)", "sap_code": "Z456", "renault_code": "417900"},
    {"sap_name": "Korea Fuel-Tech Corporation", "sap_code": "C917", "renault_code": "253253"},
]

# 디폴트값
default_warranty_mileage = 60000
default_warranty_period = 3 * 365  # 3년을 일 단위로 변환 (1095일)

print("SAP 테이블에 협력사 데이터 대량 추가 중...")
print("=" * 70)
print(f"총 {len(suppliers)}개 협력사 처리 예정")
print("=" * 70)

added_count = 0
skipped_count = 0
error_count = 0

conn = sqlite3.connect(str(DB_PATH))
cursor = conn.cursor()

for supplier in suppliers:
    sap_code = supplier["sap_code"]
    sap_name = supplier["sap_name"]
    renault_code = supplier["renault_code"]
    
    # rule_table_name 자동 생성
    rule_table_name = f"rule_{sap_code}"
    
    try:
        # 기존 데이터 확인 (중복 체크)
        existing = get_company_info(sap_code)
        
        if existing:
            # 이미 존재하는 협력사는 건너뛰기
            print(f"⊘ 건너뜀: {sap_name} ({sap_code}) - 이미 존재함")
            skipped_count += 1
            continue
        
        # 새 협력사 추가
        cursor.execute("""
            INSERT INTO sap (sap_code, sap_name, warranty_mileage, warranty_period, rule_table_name, renault_code)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (sap_code, sap_name, default_warranty_mileage, default_warranty_period, rule_table_name, renault_code))
        
        # 룰 테이블 생성 (이미 있으면 건너뛰기)
        try:
            create_rule_table(rule_table_name, cursor)
            print(f"✓ 추가: {sap_name} ({sap_code}) - RENAULT CODE: {renault_code} [룰 테이블 생성됨]")
        except Exception as e:
            # 룰 테이블이 이미 있거나 생성 실패해도 협력사는 저장됨
            error_msg = str(e)
            if "이미 존재" in error_msg or "already exists" in error_msg.lower():
                print(f"✓ 추가: {sap_name} ({sap_code}) - RENAULT CODE: {renault_code} [룰 테이블 이미 존재]")
            else:
                print(f"⚠ 추가: {sap_name} ({sap_code}) - RENAULT CODE: {renault_code} [룰 테이블 생성 실패: {error_msg}]")
        
        added_count += 1
        
    except sqlite3.IntegrityError as e:
        # 중복 키 에러 (다른 방식으로 중복이 감지된 경우)
        print(f"⊘ 건너뜀: {sap_name} ({sap_code}) - 중복 키 에러")
        skipped_count += 1
    except Exception as e:
        print(f"✗ 오류: {sap_name} ({sap_code}) - {str(e)}")
        error_count += 1

# 모든 변경사항 커밋
conn.commit()
conn.close()

print("=" * 70)
print(f"처리 완료!")
print(f"  - 추가됨: {added_count}개")
print(f"  - 건너뜀: {skipped_count}개 (중복)")
print(f"  - 오류: {error_count}개")
print("=" * 70)


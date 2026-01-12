# sap_remarks.py
# sap_code -> remark(멀티라인 가능)

SAP_REMARKS = {
    "Z456": "업체동의완료, 2021.7.8",

    "Z389": "\n".join([
        "35 % (apply liability as the same rule with R-SAS / 2021.3.23 협의완료 / No retroactivity",
    ]),

    "Z386": "\n".join([
        "- COMEX ratio (Claim period 2020.3~2021.2) :",
        "    - Continental Changchun 50% :",
        "    - RSM 50%",
        "- COMEX ratio (Claim period 2021.3 ~ ) :",
        "    - Continental Changchun 11.54% :",
        "    - RSM 88.46%",
        "    -> BCM reprogramming : 업체 귀책아님",
        "    - 285902039R (READER-BADGE) : 생산 공장 틀림 (확인 중. 22.10.20)",
    ]),

    "Z369": "\n".join([
        "sop~2021.2 lumpsum 지급협의완료 (16 MW) -- 2021.3월",
    ]),

    "Z235": "송부 후 업체에서 중복여부 회신 -> 우리가 다시 WCB 요청",
    "Z111": "해외 클레임은 58%",

    "X081": "\n".join([
        "807215098R 80% / others 25% (인티바 모터 10%? 추가 추진 중)",
    ]),

    "R908": "initial comex",

    "R805": "\n".join([
        "All items(High press hose_Leak)",
        "- 차계구분없이 현상코드중에 누유 건에서 High가 있으면 100, 아니면 50",
        "All items(Low press hose_Leak)",
        "- 현상코드중에 소음(갸르릉,바람,노이즈 등등) 있으면 0",
        "All items(All hose_Noise)",
        "- 이음도 0",
        "Others (All hose_Other claims)",
        "- 그외 기타는 10",
        "All Items (All hose)",
        "- initial comex",
    ]),

    "I928": "initial comex",

    "I908": "\n".join([
        "LFD_ Tier2 유로스타일(offline, sop~ 2017.2 -- 50% --> 35%조정, 향후 LFD 35%) / HZG는 변동없이 50%임.",
    ]),

    "I906": "라디에이터만 51%",
    "I801": "“0155800451” (닛산 납품)는 제외하고 청구함",
    "I601": "initial comex",
    "I401": "initial comex",

    "I302": "\n".join([
        "Comex 협의 완료",
        "HZG outside mirror작동불량 (folding, unfolding 불량), auto reverse 불량 : 청구제외",
    ]),

    "I100": "Initial COMEX",
    "G933": "촉매 불량 이의제가가 오면 반영",
    "G930": "보증기한 엔진 5년 10만km",

    "G924": "\n".join([
        "2019.01 ~ (2014.10~2018.10 Claim 비용 493 MW * 10% 일시불 정산 by 2018.12)",
    ]),

    "G812": "해외 디테일 없으면 청구안함",
    "G804": "Initial COMEX",
    "G215": "initial comex",
    "G212": "initial comex",
    "E923": "initial comex",
    "E915": "L47 - INSTRUMENT PANEL",

    "E914": "Initial COMEX --> 30% (SOP~ 소급 적용)",

    "E913": "initial comex",

    "E910": "\n".join([
        "Initial COMEX",
        "LJL은 모두 60%",
    ]),

    "E902": "L38 선루프 스위치는 구상X",
    "E703": "Harness와 brkt간 간섭 원인인 클레임",

    "E604": "\n".join([
        "2018.03.20 합의 / (2013.2 ~2018.2) 소급  30% 적용 정산 / 2018.03 ~ 40%적용",
        "Initial 50% 진행",
    ]),

    "E601": "Initial COMEX",
    "E505": "initial comex",
    "C934": "initial comex",
    "C933": "Initial COMEX",

    "C931": "\n".join([
        "410110522R / 410018397R 부번 E차계로 표시될 떄 F차계로 수정 후 송부",
    ]),

    "C930": "차계구분없음",
    "C922": "Claim Approval Date(처리년월일): 201907~",
    "C917": "보증기한  5년 8만km",

    "C912": "\n".join([
        "LJL차종 진동소음클레임 건 협의안(23.01.13)  60%  40%",
        "2.제외 불량 항목 : NVH (진동, 소음), 누유    -> crack, 파손의 경우는 현 구상율 40% 적용함( 2024.6.5) :2024.03~04 청구 부터~ 적용",
    ]),

    "C911": "Initial COMEX",

    "C910": "\n".join([
        "LMx - 2018.3.20 합의 완료 2013.01~2018.02 소급분 35% 5회 분납 /2018.03~ 35% 적용",
        "LFD - 2018.3.20 합의완료",
        "HZG - 2019.11~12 처음 발생",
    ]),

    "C902": "\n".join([
        "L43 / H45 - DPF 센서 감지불량, 엔진 MISFIRE 클레임은 구상 제외",
        "LFD / HZG - W37 to be rechecked / DPF 센서 감지불량, 엔진 MISFIRE 클레임은 구상 제외",
    ]),

    "B508": "Initial COMEX",

    "B904": "\n".join([
        "LJL 마찰이음은 청구하지 않음",
        "21.10.21 이후 생산차량에 대해서는 모두 청구 실시함 (23.5.26)",
    ]),

    "B908": "Initial COMEX",
    "B923": "Initial COMEX",
    "B928": "Initial COMEX",
    "B932": "Initial COMEX",

    "B933": "\n".join([
        "Koleos 전체 구상금액중 21%의 금액을 구상청구에서 제외(별도 업체 이의제기과정 없이), 사급품 5%",
        "MOTOR ASSY, SHROUD ASSY-W/MOTOR FAN 부품은 Motor 부품으로 국내는 5.00 해외는 구상안함",
        "ENG는 사급품, 구상안함",
    ]),

    "B946": "\n".join([
        "단, 해외의 공임 및 외주공임 비용합은 max. 40만원(LFD,HZG는 미협의)",
        "Initial COMEX",
    ]),

    "C102": "LJL 40%",
    "C802": "내수 해외 모두 공임비 최대 200,000원",

    "C810": "\n".join([
        "RSM 품번 / 품명 : 2346015900 / Knock Sensor M4R 제품 구상 제외",
        "Initial COMEX",
    ]),

    "B204": "initial comex",

    "B401": "\n".join([
        "initial comex,HZG MBR CROSS ROOF의 경우는 구상율 ZERO",
        "2022년 년말 구상율 재협의 예정",
        "HZG MBR CROSS ROOF의 경우는 구상율 ZERO",
    ]),

    "B901": "\n".join([
        "LJL CCB (67870 7828R, 67870 4868R, 67870 8486R)의 조향 이음 현상(조타시 이음,크로스멤버 고정볼트 마찰 이음 등)에 대하여 22년 7월 이전 생산분에 대해 WCB 내역에서 제외한다.",
        "(내수,수출 동일) - 2023년 1~2월 청구부터~",
    ]),

    "B913": "\n".join([
        "- LJL Frame Assy-FR door mobile, RR의 glass 이탈 : 구상율 0% 적용",
        "     (2021.12.1 이후 차량에서 동일 불량 발생시 구상율 60% 적용)",
        "- 나머지 차량 및 불량 현상에 대해서는 LJL 60%, 기타 차량 50% 적용함",
        "2022년 년말 구상율 재협의 예정",
    ]),

    "B919": "\n".join([
        "Initial COMEX",
        "18년 6월 구상 협의 예정",
    ]),

    "B921": "\n".join([
        "initial comex",
        "2022년 년말 구상율 재협의 예정",
    ]),

    "B925": "initial comex",
    "B935": "원인이 불분명한 비용에 대해서는 RSM과 Gruopo Antolin-KOR가 50:50",
    "B945": "initial comex",
    "C104": "C934 : 부산공장,  FUEL TUBE - LFD /HZG",
    "C406": "HZG back door striker noise 중 2020.2.7 이전 차량은 제외함",
    "C509": "LJL 30%(2024.1월 부터 적용)",

    "Z551": "LJL은 60%",
    "Z472": "initial comex LJL은 60%",
    "Z395": "Export는 LFD청구X",

    "Z388": "\n".join([
        "STALK-COMBINATION : 56% (w/noise)(2ABC등)",
        "50% (w/out noise - other items)",
        "-> LJL 구상협의 미완료 (60% 적용)",
    ]),

    "Z401": "\n".join([
        "Initial 25%",
        'G차계 (LFD)는 HeadLamp랑 266058183R&266007584R / H차계 (HZG)는 HeadLamp만',
        "나머지는 삭제",
        "(Wuhu Valeo 클레임 포함해서 발송/대표업체설정완료)",
    ]),

    "Z374": "\n".join([
        "Initial 50% --> 최종 31.25% 동의완료 (르노 Corp.합의)",
        "KD 파트라 르노 센트럴과 논의 필요/ initial comex 50% 이나, 업체 답변없음",
        "284B69159R 제외하고 산출함",
    ]),

    "Z298": "Rabat(71.88) = 296053434R / 296055209R / 296059672R & Valls(30) = 296098929R / 296092131R (2020.11 ~, 16.67% by Renault / before : 30%)",

    "Z157": "공임대 Max 30만원",

    "I904": "\n".join([
        "해외 공임 및 외주공임 비용합은 max. 171,000원으로 한다(LFD/HZG 제외)",
        "Initial COMEX",
        "모든차종 PANEL-INST COMPL는 구상율을 40%조정(2025.01~ )",
    ]),

    "I806": "\n".join([
        "내수 부품만 19 % 임",
        "단, 수출 클레임의 공임 및 외주공임 비용합은 max. 75,000원((LFD/HZG 제외)",
        "Initial COMEX",
    ]),

    "I803": "Initial COMEX",

    "I602": "\n".join([
        "단, 해외 공임 및 외주공임 비용합은 max. 8만원으로 한다",
        "차종에 관계없이 모든 item에 대해서 40% 적용",
        "-> 구상율 변경 적용 : 25.5~6 청구분부터 적용, AR1:50%",
        "확정 COMEX,  AR1:50%",
    ]),

    "I505": "\n".join([
        "Symptom : Only Noise",
        "Initial COMEX",
        "HZG 경고등 중 생산일 20190827~20191029 만 해당",
    ]),

    "I217": "\n".join([
        "해외 공임 및 외주공임 비용합은 max. 25만원으로 한다(LFD/HZG 제외)",
        "Initial COMEX(해외 Duoai LFD는 청구X, 그리고 삭제할것!)",
    ]),

    "I202": "\n".join([
        "단, 해외 공임 및 외주공임 비용합은 max. 7만원(LFD HZG 제외), 구상제외: Nozzle 막힘과 분사 각도 불량건",
        "구상제외: Nozzle 막힘, 분사 각도 불량",
        "Initial COMEX",
        "2018년 상반기 구상 협의",
    ]),

    "I201": "\n".join([
        "해외 클레임의 공임 및 외주공임 비용합은 max. 35,000원(LFD/HZG 제외)",
        "26년 1월 청구분부터",
        "2021.07~08부터 적용",
    ]),

    "G917": "\n".join([
        "initial comex(5년 10만)",
        "2022년 년말 구상율 재협의 예정",
    ]),

    "G607": "\n".join([
        "Same with Corp COMEX",
        "Bosch KR 생산 (대전공장)",
        "INJECTOR-GAS",
        "PUMP-PETROL, HIGH PRESS",
    ]),

    "G417": "\n".join([
        "MODULE-TURN INDICATOR도 30",
        "앞으로 C810으로 청구(대표업체로 해놓을예정)",
    ]),

    "G307": "\n".join([
        "2016년 1월 21일 ~ 2016년 3월 19일 생산일자 차량(풀리파손 건)",
        "누유, 누수는 50",
        "H5H,210100078R : 25%",
        "차량 조립 기준 2020년11월30일 이전 생산차량은 청구에서 삭제",
        "ALL: 보증기간 36개월 적용(2025.11~)",
        "M9R (diesel) : 제외",
    ]),

    "G302": "\n".join([
        "단, 해외의 공임 및 외주공임 비용합은 max. 50만원(LFD/HZG 제외)",
        "Initial COMEX 동의",
    ]),

    "G103": "나머지 차계 일단 청구X",

    "E929": "\n".join([
        "LJL 차종의 경우",
        "1. 21년 5월 ∼ 10월 청구분 : 당사 40% 적용.",
        "2. 21년 11월 ∼ 청구분 : 당사 30% 적용.",
        "3. 당사 25% 분담율 재협의 22년 2월 중 (문신동 과장, 정동훈 수석. 21.5월)",
        "4. 22.7~8월분부터 구상율 25% 적용 (재체결 건은 이의제기 받아주기로 함.",
        "   Lot성 불량 발생시 구상율 협의 재실시. 22.4.7 회의)",
        '"W" code (2years, 40,000km) -> 일반고객도 동일기준',
        "일반 공임 Max 15,400원 / 외주 공임 Max 24,000원",
    ]),

    "E920": "\n".join([
        "E404 (광원텍)이 구 업체",
        "공임 Max 20,000원(차계 구분 없이 국내, 해외 모두)",
    ]),

    "E916": "\n".join([
        "수출건 LFD는 청구안함(삭제)(2020.11~12 클레임부터 적용)",
        "메일 참조(2025년 9월청구분 적용)",
    ]),

    "E909": "해외청구 안함",
    "E702": "단, 해외 공임 및 외주공임 비용합은 max. 10만원",

    "E503": "국내, 해외 모두 공임 및 외주공임 비용합은 max. 12만원(LFD/HZG 제외)",

    "C940": "\n".join([
        "Initial COMEX",
        "field claim 항목 중 무교환 수리 청구 제외 : 22.7~",
        "Other (HZG FRT/RR, LFD FRT) : 40%, LJL RR : 50% (2023.2.7 회의록 조정)",
        "LFD 구상완료, 50% (OFFLINE 2017.7.06~, 국산화 시점부터 구상진행, KD업체 NTN)",
    ]),

    "C929": "LJL 8%",

    "C915": "단, 해외 공임 및 외주공임 비용합은 max. 10만원(LFD/HZG 제외)",

    "C907": "\n".join([
        "LJL은 60%",
        "- 40300-1305R : 22.3.4 이후 생산차량만 해당 업체 제품임.",
        "- 40300-9516R : 22.7.11 이후 생산차량만 해당 업체 제품임.",
        "- 40300-9932R : 22.7.1 이후 생산차량만 해당 업체 제품임.",
    ]),

    "C905": "\n".join([
        "구상제외: Wheel alignment (현상코드 5S, 8E, 8F), 49001_0059R(Manual type의 누유는 구상 제외)",
        "21.6.24 이전 생산 청구삭제",
    ]),

    "C814": "단, 공임 및 외주공임 비용합은 max. 10만원으로 한다",

    "C222": "\n".join([
        '"""틱""" noise 구상 제외(보류)',
        "단, 해외공임 및 외주공임 비용합은 max. 15만원",
    ]),

    "C514": "\n".join([
        "(1) 을이 갑에게 납품 후 96개월 / (2) 신차 등록후 84개월 (3) 주행거리 120,000km",
        "배출가스 관련 주요 부품   - ECU 및 정화용 촉매",
        "그외 일반 배출가스 부품은 60개월 80,000km",
    ]),
}

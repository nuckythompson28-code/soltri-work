from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ─── 폰트 등록 ───
pdfmetrics.registerFont(TTFont('Malgun', 'C:/Windows/Fonts/malgun.ttf'))
pdfmetrics.registerFont(TTFont('MalgunBd', 'C:/Windows/Fonts/malgunbd.ttf'))
pdfmetrics.registerFont(TTFont('Nirmala', 'C:/Windows/Fonts/Nirmala.ttc', subfontIndex=0))
pdfmetrics.registerFont(TTFont('NirmalaBd', 'C:/Windows/Fonts/Nirmala.ttc', subfontIndex=1))

# ─── 색상 ───
NAVY      = HexColor('#1B2A4A')
DARK_BLUE = HexColor('#2C3E6B')
ACCENT    = HexColor('#3B82F6')
TEAL      = HexColor('#0D9488')
RED       = HexColor('#EF4444')
ORANGE    = HexColor('#F97316')
GREEN     = HexColor('#10B981')
GRAY      = HexColor('#64748B')
LIGHT_BG  = HexColor('#F0F4F8')
BLUE_BG   = HexColor('#E3F2FD')
RED_BG    = HexColor('#FFEBEE')
GREEN_BG  = HexColor('#E8F5E9')
YELLOW_BG = HexColor('#FFFACD')


# ═══════════════════════════════════════════════════════
# 번역 데이터
# ═══════════════════════════════════════════════════════

LANGS = {
    "en": {
        "font": "Malgun", "fontBd": "MalgunBd",
        "filename": "O0852_Training_EN.pdf",
        "cover_title": "O0852 Program Complete Analysis",
        "cover_sub": "CNC Macro Program Training Material",
        "cover_desc": "Ring-shaped Part Auto Continuous Machining System  |  FANUC Macro",
        "cover_features": [
            ["Macro Automation", "Set material/dimension/conditions by variable input only"],
            ["Continuous Machining", "Face→Step→Chamfer→Cut cycle auto repeat"],
            ["Auto Link", "Auto material pull when material runs short"],
            ["Safety Validation", "30+ alarms to prevent input errors"],
        ],
        "toc_title": "Table of Contents",
        "toc_items": [
            ("01", "Overview", "What the program does at a glance"),
            ("02", "Program Structure", "4 sub-program relationship"),
            ("03", "Input Parameters", "Variables modified by operator (#101~#123)"),
            ("04", "Feed Rate by Material", "CN / RS / CM auto setting"),
            ("05", "Advanced Settings & Safety", "System variables and validation"),
            ("06", "Machining Cycle Detail", "Face → Step → Chamfer → Cut"),
            ("07", "Auto Link & Remaining", "Auto material pull system"),
            ("08", "Alarm Code Table", "Error codes and responses"),
            ("09", "Machine M-Code Map", "9 machine configuration"),
            ("10", "Operator Precautions", "Safety rules and checklist"),
        ],
        "s1_title": "1. Overview",
        "s1_desc": "O0852 is a FANUC macro program that automatically and continuously machines<br/>pipe material on a CNC lathe to mass-produce ring-shaped parts (bushings/bearings).",
        "s1_summary_title": "Current Settings Summary",
        "s1_summary": [
            ["Item", "Material", "Finished Part", "Production"],
            ["Specs", "CN, OD70 × ID56\nLength 543mm", "OD64.9 × ID58.3\nLength 7.823mm", "9.843mm per piece\n~6 per pull cycle"],
        ],
        "s1_feat_title": "Key Features",
        "s1_features": [
            ["Macro Automation", "Set material, dimensions, cutting conditions by variable input. Multi-product support without code changes"],
            ["Continuous Machining", "Face → Step cut → Chamfer → Cut-off cycle auto repeats"],
            ["Auto Link", "When material runs short, auto-pulls material and continues machining"],
            ["Safety Validation", "30+ alarm checks prevent input errors, dimension conflicts, safety issues"],
        ],
        "s2_title": "2. Program Structure",
        "s2_desc": "4 sub-programs are called sequentially.",
        "s2_table": [
            ["Program", "Name", "Role", "Call Method"],
            ["O0852", "Main Setup", "Parameter input, validation, machine config", "Direct run"],
            ["O9001", "Main Logic", "Length/count calc, alarm check, start machining", "M98 P9001"],
            ["O9002", "Main Machining", "Face→Step→Chamfer→Cut repeat, Auto Link", "M98 P9002"],
            ["O9003", "Remaining", "Remaining material pull, re-setup, final cycle", "M98 P9003"],
        ],
        "s2_flow_title": "Call Flow",
        "s2_flow": "O0852 (Setup) → O9001 (Calc/Validate) → O9002 (Machining Cycle) → O9003 (Remaining)",
        "s2_tool_title": "Tools Used",
        "s2_tools": [
            ["Tool", "Purpose", "Main Process", "Note"],
            ["T01", "Boring bar (ID machining)", "Rough, Finish, Chamfer", "Main tool"],
            ["T02", "Cut-off insert", "Facing, Cut-off", "Face + Parting"],
            ["T03", "Auto Link", "Material pull", "Clamp/Unclamp"],
        ],
        "s3_title": "3. Input Parameters",
        "s3_info": "Operator modification area (Line 4~29, modify #101~#123 only)",
        "s3_1_title": "3-1. Basic Settings (#101 ~ #108)",
        "s3_1_data": [
            ["Variable", "Description", "Value", "Input Rule"],
            ["#101", "Material length ones digit", "43", "Range 0~99"],
            ["#102", "Material length hundreds", "500", "Only 0,100,200,300,400,500"],
            ["#103", "Chuck length (grip)", "75", "mm unit"],
            ["#104", "Initial face cut", "1", "Max 40"],
            ["#105", "Material type", "1", "1=CN, 2=RS, 3=CM"],
            ["#106", "Process type", "3", "3=Chamfer, 4=Chamfer(ext)"],
            ["#107", "Rough ON/OFF", "0", "0=OFF, 1=ON"],
            ["#108", "Single mode", "1", "1=One at a time"],
        ],
        "s3_1_note": "Material length: #101 + #102 = Total → Ex: 500 + 43 = 543mm",
        "s3_2_title": "3-2. Dimension Data (#109 ~ #118)",
        "s3_2_data": [
            ["Variable", "Description", "Value", "Unit"],
            ["#109", "Raw OD (Outside Diameter)", "70", "mm"],
            ["#110", "Raw ID (Inside Diameter)", "56", "mm"],
            ["#111", "Finished OD", "64.90", "mm"],
            ["#112", "Finished ID", "58.30", "mm"],
            ["#113", "Finished Length", "7.823", "mm"],
            ["#114", "Tip Width (Cut-off)", "2.02", "mm"],
            ["#115", "ID Radius", "0.34", "mm"],
            ["#116", "OD Radius", "0.36", "mm"],
            ["#117", "ID Chamfer", "0.37", "mm"],
            ["#118", "OD Chamfer", "0.60", "mm"],
            ["#119", "Length>40mm flag", "0", "Set 1 if >40mm"],
        ],
        "s3_3_title": "3-3. Cutting Conditions (#120 ~ #123)",
        "s3_3_data": [
            ["Variable", "Description", "Value", "Note"],
            ["#120", "T01 RPM", "1700", "Boring bar speed"],
            ["#121", "T02 RPM", "1700", "Cut-off insert speed"],
            ["#122", "Pull Distance", "1", "Auto Link pull distance"],
            ["#123", "Margin Value", "0.2", "Range 0~3"],
        ],
        "s4_title": "4. Feed Rate by Material Type",
        "s4_desc": "Feed coefficients are auto-determined by material type (#105).",
        "s4_data": [
            ["Material", "#124 (T01 Rough)", "#125 (T01 Step)", "#126 (T01 Chamfer)", "#127 (T02 Cut)"],
            ["CN (#105=1)", "0.15", "0.12", "0.15", "0.09"],
            ["RS (#105=2)", "0.13", "0.08", "0.05", "0.08"],
            ["CM (#105=3)", "0.13", "0.09", "0.09", "0.09"],
            ["Default", "0.15", "0.10", "0.10", "0.10"],
        ],
        "s4_formula_title": "Feed Rate Formula",
        "s4_formula": "F value = RPM × Feed coefficient",
        "s4_example_title": "Calculation Examples (RPM = 1700)",
        "s4_examples": [
            ["Process", "Material", "Calculation", "Result"],
            ["T01 Rough", "CN", "1700 × 0.15", "F255 mm/min"],
            ["T01 Step", "CN", "1700 × 0.12", "F204 mm/min"],
            ["T02 Cut-off", "CN", "1700 × 0.09", "F153 mm/min"],
            ["T01 Rough", "RS", "1700 × 0.13", "F221 mm/min"],
        ],
        "s4_note": "RS material has lower feed rates overall (higher hardness requires conservative cutting).",
        "s5_title": "5. Advanced Settings & Safety",
        "s5_warn": "DANGER: STOP EDITING — This section is auto-configured. DO NOT modify!",
        "s5_1_title": "5-1. System Variables (#500 series)",
        "s5_1_data": [
            ["Variable", "Description", "Value", "Role"],
            ["#530", "Safety length", "12", "Safety distance from chuck end"],
            ["#531", "Min remaining length", "15", "Min 15mm must stay in jaw after machining\n→ Short remaining = workpiece ejection risk"],
            ["#532", "Remaining-safety offset", "1.03", "Calibration value for accurate positioning\nduring last Auto Link pull"],
            ["#506", "Z rough max depth", "55", "Max Z-direction length per pass"],
            ["#507", "Z rough clearance", "1", "Z-direction clearance"],
            ["#508", "T01 nose R", "0.2", "Tool tip radius"],
            ["#509", "Rough interval", "0.4", "X-direction interval"],
            ["#510", "Chamfer interval", "1", "Chamfer cut interval"],
        ],
        "s5_2_title": "5-2. RS40 / Small Diameter Safety",
        "s5_2_info": "RS40 material (#105=2) or OD < 30mm (#109<30) → #129=15 auto-set<br/>Adds 15mm residual offset to prevent workpiece ejection due to low rigidity.",
        "s6_title": "6. Machining Cycle Detail",
        "s6_1_title": "6-1. Overall Flow",
        "s6_1_data": [
            ["Step", "Process", "Tool", "Main Action"],
            ["(1)", "Coordinate set", "-", "G10 L2 P0 Z[#103] — Set G54 work Z origin"],
            ["(2)", "Facing", "T02", "Face end surface (OD→ID direction)"],
            ["(3)", "Step cut", "T01", "Rough(option) + Finish — Z-direction cut"],
            ["(4)", "Chamfer", "T01", "OD-R, ID-R chamfer"],
            ["(5)", "Cut-off", "T02", "Part separation + M12 count"],
            ["", "Repeat", "", "Repeat (3)~(5) for count (#517)"],
        ],
        "s6_calc_title": "Key Calculation Variables",
        "s6_calc_data": [
            ["Formula", "Meaning", "Current Value"],
            ["#505 = #113 + #114", "Unit per piece (length + cut width)", "9.843mm"],
            ["#500 = #103 - #530", "Effective machining length", "63mm"],
            ["#517 = FIX[#500/#505]", "Pieces per pull cycle", "6 pcs"],
            ["#518 = #517 × #505", "Total machining length per cycle", "59.058mm"],
            ["#140 = #101+#102-#129", "Total material length (with offset)", "543mm"],
        ],
        "s7_title": "7. Auto Link & Remaining Material",
        "s7_1_title": "7-1. Auto Link Sequence (N210)",
        "s7_1_desc": "Core function that automatically pulls material when it runs short.",
        "s7_1_data": [
            ["Step", "Code", "Action", "Description"],
            ["(1)", "M05", "Spindle stop", "Stop rotation for safety"],
            ["(2)", "G00 T03, X0", "Tool select", "Auto Link tool, center move"],
            ["(3)", "Z[-#522+#122]", "Approach pull pos", "Move to pull position at F200"],
            ["(4)", "M[#132]", "Auto Loader CLOSE", "Loader grips material"],
            ["(5)", "M69 + G04 P2500", "Chuck UNCLAMP", "Chuck releases (2.5s wait)"],
            ["(6)", "W[length+margin]", "Pull material", "Pull material forward"],
            ["(7)", "M68 + G04 P2500", "Chuck CLAMP", "Chuck re-grips (2.5s wait)"],
            ["(8)", "M[#131]", "Auto Loader OPEN", "Loader releases"],
            ["(9)", "G00 Z100→GOTO100", "Return", "Safe pos → restart from facing"],
        ],
        "s7_warn": "SAFETY WARNING: Chuck opens and closes during Auto Link.<br/>NEVER touch material or tools during this sequence!<br/>Do NOT reduce G04 wait times.",
        "s8_title": "8. Alarm Code Table",
        "s8_desc": "When an error occurs, check the alarm number (#3000) and refer to the table below.",
        "s8_1_title": "8-1. Machining Logic Alarms",
        "s8_1_data": [
            ["Alarm", "Condition", "Meaning", "Action"],
            ["1", "#140 > 580", "Material length > 580mm", "Check #101, #102"],
            ["2", "#109 < #111", "Raw OD < Finished OD", "Check #109, #111"],
            ["3", "#110 > #112", "Raw ID > Finished ID", "Check #110, #112"],
            ["4", "#125 >= 0.14", "T01 step feed too high", "Check material type (#105)"],
            ["5", "#102 unit error", "Not in 100s unit", "Use 0/100/200/300/400/500"],
            ["6", "#101 > 99", "Ones digit out of range", "Change to 0~99"],
            ["7", "#140 <= #103", "Material too short", "Check material length"],
            ["10", "#104 > 40", "Init face cut too large", "Set #104 to 40 or less"],
        ],
        "s8_2_title": "8-2. Safety Validation Alarms",
        "s8_2_data": [
            ["Alarm", "Condition", "Meaning", "Action"],
            ["203", "OD <= ID", "Raw OD <= Raw ID", "Check #109, #110"],
            ["204", "FIN OD <= FIN ID", "Finished OD <= ID", "Check #111, #112"],
            ["205", "#109 <= 0", "OD is 0 or negative", "Check #109"],
            ["206", "#120 <= 0", "T01 RPM is 0 or less", "Check #120"],
            ["207", "#121 <= 0", "T02 RPM is 0 or less", "Check #121"],
            ["208", "#114 >= #113", "Tip width >= Fin length", "Check #114, #113"],
        ],
        "s8_3_title": "8-3. Missing Variable Alarms (101~130, 506~532)",
        "s8_3_info": "Alarm 101~130 or 506~532: The corresponding variable (#number) is not entered.<br/>Ex) Alarm 109 → #109 (Raw OD) not entered → Enter the value.",
        "s9_title": "9. Machine M-Code Mapping",
        "s9_desc": "M-codes for Auto Loader/Boring bar are auto-mapped by machine number (#130).",
        "s9_data": [
            ["Machine\nNo.", "Name", "AL OPEN\n(#131)", "AL CLOSE\n(#132)", "BR UP\n(#133)", "BR DOWN\n(#134)"],
            ["1", "AL-1", "M56", "M55", "M54", "M53"],
            ["4", "HA-4", "M52", "M51", "M54", "M53"],
            ["5", "AL-5", "M171", "M170", "M53", "M54"],
            ["7", "AL-7", "M64", "M63", "M53", "M54"],
            ["8", "AL-8", "M64", "M63", "M54", "M53"],
            ["9", "AL-9", "M63", "M64", "M53", "M54"],
            ["10", "AL-10", "M64", "M63", "M54", "M53"],
            ["13", "HA-S3", "M56", "M55", "M54", "M53"],
            ["14", "HA-S4", "M56", "M55", "M54", "M53"],
        ],
        "s9_info": "AL = Auto Loader | BR = Boring bar<br/>To change machine: only modify #130. M-codes auto-map.",
        "s9_warn": "BR UP/DOWN differs by machine. AL-5, AL-7, AL-9 are reversed.<br/>Always modify #130 only. Never edit M-codes directly.",
        "s10_title": "10. Operator Precautions",
        "s10_1_title": "10-1. Required Checks",
        "s10_1_data": [
            ["No.", "Check Item", "Detail"],
            ["1", "Editable area", "#101~#123 only.\nNEVER edit after DANGER section"],
            ["2", "#119 check (long parts)", "If #113 > 40mm,\nmust set #119=1"],
            ["3", "#102 input format", "Only 0,100,200,300,400,500\nOther values → Alarm 5"],
            ["4", "#101 input range", "0~99 only\n100+ → Alarm 6"],
            ["5", "Machine change", "Only modify #130\nNever edit M-codes directly"],
        ],
        "s10_2_title": "10-2. Safety Rules",
        "s10_2_info": "(1) NO intervention during Auto Link — NEVER touch material/tools during chuck open/close!<br/>(2) RS40 material — 15mm safety offset auto-applied, do NOT disable<br/>(3) Small OD (less than 30mm) — Same safety offset applied<br/>(4) On alarm — Check alarm number → Refer to code table → Fix variable → Restart<br/>(5) G04 wait times — Do NOT reduce Auto Link wait times",
        "s10_3_title": "10-3. Pre-Machining Checklist",
        "s10_3_items": [
            "Material type (#105) matches actual material?",
            "Raw OD/ID (#109, #110) matches measured values?",
            "Finished dimensions (#111~#118) match drawing?",
            "Machine number (#130) matches actual machine?",
            "RPM (#120, #121) suitable for material/tool?",
            "First run in single mode (#108=1) for test?",
            "Auto Loader operation verified?",
            "Cut-off insert (T02) wear checked?",
        ],
        "summary_title": "O0852 Program Key Summary",
        "summary_items": [
            "Operators only modify variables #101~#123 (material, dimensions, conditions)",
            "Feed rates auto-set by material type (CN/RS/CM)",
            "Face → Step → Chamfer → Cut-off cycle auto-repeats",
            "Auto Link pulls material when it runs short",
            "30+ alarms prevent input errors and safety issues",
            "9 machine M-codes auto-mapped (change #130 only)",
        ],
        "summary_footer": "If you have questions, please ask anytime",
    },

    "bn": {
        "font": "Nirmala", "fontBd": "NirmalaBd",
        "filename": "O0852_Training_BN.pdf",
        "cover_title": "O0852 প্রোগ্রাম সম্পূর্ণ বিশ্লেষণ",
        "cover_sub": "CNC ম্যাক্রো প্রোগ্রাম প্রশিক্ষণ উপকরণ",
        "cover_desc": "রিং-আকৃতির যন্ত্রাংশ স্বয়ংক্রিয় ধারাবাহিক মেশিনিং | FANUC ম্যাক্রো",
        "cover_features": [
            ["ম্যাক্রো অটোমেশন", "শুধু ভেরিয়েবল ইনপুট দিয়ে উপকরণ/মাপ/শর্ত সেট করুন"],
            ["ধারাবাহিক মেশিনিং", "ফেসিং→স্টেপ→চ্যামফার→কাটিং সাইকেল স্বয়ংক্রিয় পুনরাবৃত্তি"],
            ["অটো লিঙ্ক", "উপকরণ কম হলে স্বয়ংক্রিয়ভাবে টানা হয়"],
            ["নিরাপত্তা যাচাই", "৩০+ অ্যালার্ম দিয়ে ইনপুট ত্রুটি প্রতিরোধ"],
        ],
        "toc_title": "সূচিপত্র",
        "toc_items": [
            ("০১", "সামগ্রিক সারসংক্ষেপ", "প্রোগ্রাম কী করে এক নজরে"),
            ("০২", "প্রোগ্রামের গঠন", "৪টি সাব-প্রোগ্রামের সম্পর্ক"),
            ("০৩", "ইনপুট প্যারামিটার", "অপারেটর পরিবর্তন করে (#101~#123)"),
            ("০৪", "উপকরণ অনুযায়ী ফিড রেট", "CN / RS / CM স্বয়ংক্রিয় সেটিং"),
            ("০৫", "উন্নত সেটিংস ও নিরাপত্তা", "সিস্টেম ভেরিয়েবল ও যাচাই"),
            ("০৬", "মেশিনিং সাইকেল বিস্তারিত", "ফেসিং → স্টেপ → চ্যামফার → কাটিং"),
            ("০৭", "অটো লিঙ্ক ও অবশিষ্ট মেশিনিং", "স্বয়ংক্রিয় উপকরণ টানার সিস্টেম"),
            ("০৮", "অ্যালার্ম কোড তালিকা", "ত্রুটি নম্বর ও সমাধান"),
            ("০৯", "মেশিন M-কোড ম্যাপিং", "৯টি মেশিনের কনফিগারেশন"),
            ("১০", "অপারেটর সতর্কতা", "নিরাপত্তা নিয়ম ও চেকলিস্ট"),
        ],
        "s1_title": "১. সামগ্রিক সারসংক্ষেপ",
        "s1_desc": "O0852 হলো একটি FANUC ম্যাক্রো প্রোগ্রাম যা CNC লেদে পাইপ উপকরণ থেকে<br/>স্বয়ংক্রিয়ভাবে রিং-আকৃতির যন্ত্রাংশ (বুশিং/বিয়ারিং) গণ-উৎপাদন করে।",
        "s1_summary_title": "বর্তমান সেটিংস সারসংক্ষেপ",
        "s1_summary": [
            ["বিষয়", "উপকরণ", "সমাপ্ত যন্ত্রাংশ", "উৎপাদন"],
            ["স্পেক", "CN, OD70 × ID56\nদৈর্ঘ্য 543mm", "OD64.9 × ID58.3\nদৈর্ঘ্য 7.823mm", "প্রতিটিতে 9.843mm\nপ্রতি পুলে ~৬টি"],
        ],
        "s1_feat_title": "মূল বৈশিষ্ট্য",
        "s1_features": [
            ["ম্যাক্রো অটোমেশন", "ভেরিয়েবল ইনপুট দিয়ে উপকরণ, মাপ, কাটিং শর্ত সেট করুন"],
            ["ধারাবাহিক মেশিনিং", "ফেসিং → স্টেপ কাট → চ্যামফার → কাটিং সাইকেল স্বয়ংক্রিয় পুনরাবৃত্তি"],
            ["অটো লিঙ্ক", "উপকরণ কম হলে স্বয়ংক্রিয়ভাবে টানা হয়"],
            ["নিরাপত্তা যাচাই", "৩০+ অ্যালার্ম চেক দিয়ে ইনপুট ত্রুটি প্রতিরোধ"],
        ],
        "s2_title": "২. প্রোগ্রামের গঠন",
        "s2_desc": "৪টি সাব-প্রোগ্রাম ধারাবাহিকভাবে কল করা হয়।",
        "s2_table": [
            ["প্রোগ্রাম", "নাম", "ভূমিকা", "কল পদ্ধতি"],
            ["O0852", "মেইন সেটআপ", "প্যারামিটার ইনপুট, যাচাই, মেশিন কনফিগ", "সরাসরি চালান"],
            ["O9001", "মেইন লজিক", "দৈর্ঘ্য/সংখ্যা গণনা, অ্যালার্ম চেক", "M98 P9001"],
            ["O9002", "মেইন মেশিনিং", "ফেসিং→স্টেপ→চ্যামফার→কাটিং, অটো লিঙ্ক", "M98 P9002"],
            ["O9003", "অবশিষ্ট মেশিনিং", "অবশিষ্ট উপকরণ টানা, পুনরায় সেটআপ", "M98 P9003"],
        ],
        "s2_flow_title": "কল ফ্লো",
        "s2_flow": "O0852 (সেটআপ) → O9001 (গণনা/যাচাই) → O9002 (মেশিনিং) → O9003 (অবশিষ্ট)",
        "s2_tool_title": "ব্যবহৃত টুল",
        "s2_tools": [
            ["টুল", "উদ্দেশ্য", "প্রধান প্রক্রিয়া", "মন্তব্য"],
            ["T01", "বোরিং বার (ID)", "রাফিং, ফিনিশিং, চ্যামফার", "প্রধান টুল"],
            ["T02", "কাটিং ইনসার্ট", "ফেসিং, কাটিং", "ফেস + পার্টিং"],
            ["T03", "অটো লিঙ্ক", "উপকরণ টানা", "ক্ল্যাম্প/আনক্ল্যাম্প"],
        ],
        "s3_title": "৩. ইনপুট প্যারামিটার",
        "s3_info": "অপারেটর পরিবর্তনযোগ্য এলাকা (লাইন 4~29, শুধু #101~#123 পরিবর্তন করুন)",
        "s3_1_title": "৩-১. মৌলিক সেটিংস (#101 ~ #108)",
        "s3_1_data": [
            ["ভেরিয়েবল", "বিবরণ", "মান", "ইনপুট নিয়ম"],
            ["#101", "উপকরণ দৈর্ঘ্য একক", "43", "পরিসর 0~99"],
            ["#102", "উপকরণ দৈর্ঘ্য শতক", "500", "শুধু 0,100,200,300,400,500"],
            ["#103", "চাক দৈর্ঘ্য", "75", "mm একক"],
            ["#104", "প্রাথমিক ফেস কাট", "1", "সর্বোচ্চ 40"],
            ["#105", "উপকরণের ধরন", "1", "1=CN, 2=RS, 3=CM"],
            ["#106", "প্রক্রিয়ার ধরন", "3", "3=চ্যামফার, 4=চ্যামফার(বর্ধিত)"],
            ["#107", "রাফিং ON/OFF", "0", "0=OFF, 1=ON"],
            ["#108", "একক মোড", "1", "1=একটি করে"],
        ],
        "s3_1_note": "উপকরণ দৈর্ঘ্য: #101 + #102 = মোট → উদাহরণ: 500 + 43 = 543mm",
        "s3_2_title": "৩-২. মাপ ডেটা (#109 ~ #118)",
        "s3_2_data": [
            ["ভেরিয়েবল", "বিবরণ", "মান", "একক"],
            ["#109", "কাঁচা OD (বাইরের ব্যাস)", "70", "mm"],
            ["#110", "কাঁচা ID (ভিতরের ব্যাস)", "56", "mm"],
            ["#111", "সমাপ্ত OD", "64.90", "mm"],
            ["#112", "সমাপ্ত ID", "58.30", "mm"],
            ["#113", "সমাপ্ত দৈর্ঘ্য", "7.823", "mm"],
            ["#114", "টিপ প্রস্থ (কাটিং)", "2.02", "mm"],
            ["#115", "ID রেডিয়াস", "0.34", "mm"],
            ["#116", "OD রেডিয়াস", "0.36", "mm"],
            ["#117", "ID চ্যামফার", "0.37", "mm"],
            ["#118", "OD চ্যামফার", "0.60", "mm"],
            ["#119", "দৈর্ঘ্য>40mm ফ্ল্যাগ", "0", "40 ছাড়ালে 1 সেট করুন"],
        ],
        "s3_3_title": "৩-৩. কাটিং শর্ত (#120 ~ #123)",
        "s3_3_data": [
            ["ভেরিয়েবল", "বিবরণ", "মান", "মন্তব্য"],
            ["#120", "T01 RPM", "1700", "বোরিং বার গতি"],
            ["#121", "T02 RPM", "1700", "কাটিং ইনসার্ট গতি"],
            ["#122", "পুল দূরত্ব", "1", "অটো লিঙ্ক পুল দূরত্ব"],
            ["#123", "মার্জিন মান", "0.2", "পরিসর 0~3"],
        ],
        "s4_title": "৪. উপকরণ অনুযায়ী ফিড রেট",
        "s4_desc": "উপকরণের ধরন (#105) অনুযায়ী ফিড গুণাঙ্ক স্বয়ংক্রিয়ভাবে নির্ধারিত হয়।",
        "s4_data": [
            ["উপকরণ", "#124 (T01 রাফ)", "#125 (T01 স্টেপ)", "#126 (T01 চ্যামফার)", "#127 (T02 কাট)"],
            ["CN (#105=1)", "0.15", "0.12", "0.15", "0.09"],
            ["RS (#105=2)", "0.13", "0.08", "0.05", "0.08"],
            ["CM (#105=3)", "0.13", "0.09", "0.09", "0.09"],
            ["ডিফল্ট", "0.15", "0.10", "0.10", "0.10"],
        ],
        "s4_formula_title": "ফিড রেট সূত্র",
        "s4_formula": "F মান = RPM × ফিড গুণাঙ্ক",
        "s4_example_title": "গণনার উদাহরণ (RPM = 1700)",
        "s4_examples": [
            ["প্রক্রিয়া", "উপকরণ", "গণনা", "ফলাফল"],
            ["T01 রাফিং", "CN", "1700 × 0.15", "F255 mm/min"],
            ["T01 স্টেপ", "CN", "1700 × 0.12", "F204 mm/min"],
            ["T02 কাটিং", "CN", "1700 × 0.09", "F153 mm/min"],
        ],
        "s4_note": "RS উপকরণ বেশি শক্ত, তাই সামগ্রিকভাবে ফিড কম সেট করা হয়।",
        "s5_title": "৫. উন্নত সেটিংস ও নিরাপত্তা",
        "s5_warn": "বিপদ: সম্পাদনা বন্ধ — এই অংশ সিস্টেম স্বয়ংক্রিয় সেটিং। পরিবর্তন করবেন না!",
        "s5_1_title": "৫-১. সিস্টেম ভেরিয়েবল (#500 সিরিজ)",
        "s5_1_data": [
            ["ভেরিয়েবল", "বিবরণ", "মান", "ভূমিকা"],
            ["#530", "নিরাপত্তা দৈর্ঘ্য", "12", "চাক প্রান্ত থেকে নিরাপদ দূরত্ব"],
            ["#531", "ন্যূনতম অবশিষ্ট দৈর্ঘ্য", "15", "মেশিনিং-এর পর ন্যূনতম 15mm জ-তে ধরা থাকতে হবে"],
            ["#532", "অবশিষ্ট-নিরাপত্তা অফসেট", "1.03", "শেষ অটো লিঙ্ক পুলে সঠিক অবস্থান গণনার জন্য"],
            ["#506", "Z রাফ সর্বোচ্চ গভীরতা", "55", "প্রতি পাসে সর্বোচ্চ Z দৈর্ঘ্য"],
            ["#508", "T01 নোজ R", "0.2", "টুলের টিপ ব্যাসার্ধ"],
            ["#509", "রাফিং ব্যবধান", "0.4", "X দিকে ব্যবধান"],
            ["#510", "চ্যামফার ব্যবধান", "1", "চ্যামফার কাটের ব্যবধান"],
        ],
        "s5_2_title": "৫-২. RS40 / ছোট ব্যাসের নিরাপত্তা",
        "s5_2_info": "RS40 উপকরণ (#105=2) বা OD 30mm-এর কম → #129=15 স্বয়ংক্রিয়<br/>কম দৃঢ়তার কারণে উপকরণ ছিটকে যাওয়া রোধে 15mm অফসেট যোগ হয়।",
        "s6_title": "৬. মেশিনিং সাইকেল বিস্তারিত",
        "s6_1_title": "৬-১. সামগ্রিক প্রবাহ",
        "s6_1_data": [
            ["ধাপ", "প্রক্রিয়া", "টুল", "প্রধান কাজ"],
            ["(১)", "কোঅর্ডিনেট সেট", "-", "G10 L2 P0 Z[#103] — G54 Z অরিজিন সেট"],
            ["(২)", "ফেসিং", "T02", "প্রান্ত তল সমতল করা (OD→ID দিক)"],
            ["(৩)", "স্টেপ কাট", "T01", "রাফিং(ঐচ্ছিক) + ফিনিশিং — Z দিকে কাটা"],
            ["(৪)", "চ্যামফার", "T01", "OD-R, ID-R চ্যামফার"],
            ["(৫)", "কাটিং", "T02", "যন্ত্রাংশ আলাদা করা + M12 গণনা"],
            ["", "পুনরাবৃত্তি", "", "(৩)~(৫) সংখ্যা (#517) পর্যন্ত পুনরাবৃত্তি"],
        ],
        "s6_calc_title": "মূল গণনা ভেরিয়েবল",
        "s6_calc_data": [
            ["সূত্র", "অর্থ", "বর্তমান মান"],
            ["#505 = #113 + #114", "প্রতি পিসে একক (দৈর্ঘ্য + কাটিং প্রস্থ)", "9.843mm"],
            ["#500 = #103 - #530", "কার্যকর মেশিনিং দৈর্ঘ্য", "63mm"],
            ["#517 = FIX[#500/#505]", "প্রতি পুল সাইকেলে সংখ্যা", "৬টি"],
            ["#518 = #517 × #505", "প্রতি সাইকেলে মোট মেশিনিং দৈর্ঘ্য", "59.058mm"],
        ],
        "s7_title": "৭. অটো লিঙ্ক ও অবশিষ্ট মেশিনিং",
        "s7_1_title": "৭-১. অটো লিঙ্ক ক্রম (N210)",
        "s7_1_desc": "উপকরণ কম হলে স্বয়ংক্রিয়ভাবে টানার মূল ফাংশন।",
        "s7_1_data": [
            ["ধাপ", "কোড", "কাজ", "বিবরণ"],
            ["(১)", "M05", "স্পিন্ডল বন্ধ", "নিরাপত্তার জন্য ঘূর্ণন বন্ধ"],
            ["(২)", "G00 T03, X0", "টুল নির্বাচন", "অটো লিঙ্ক টুল, কেন্দ্রে সরান"],
            ["(৩)", "Z[-#522+#122]", "পুল অবস্থানে যান", "F200 এ অবস্থানে সরান"],
            ["(৪)", "M[#132]", "অটো লোডার বন্ধ", "লোডার উপকরণ ধরে"],
            ["(৫)", "M69+G04 P2500", "চাক আনক্ল্যাম্প", "চাক ছাড়ে (2.5সে অপেক্ষা)"],
            ["(৬)", "W[দৈর্ঘ্য+মার্জিন]", "উপকরণ টানুন", "উপকরণ সামনে টানা হয়"],
            ["(৭)", "M68+G04 P2500", "চাক ক্ল্যাম্প", "চাক আবার ধরে (2.5সে অপেক্ষা)"],
            ["(৮)", "M[#131]", "অটো লোডার খোলা", "লোডার ছেড়ে দেয়"],
            ["(৯)", "G00 Z100→GOTO100", "ফিরে যান", "নিরাপদ অবস্থান → ফেসিং থেকে শুরু"],
        ],
        "s7_warn": "নিরাপত্তা সতর্কতা: অটো লিঙ্কের সময় চাক খোলে ও বন্ধ হয়।<br/>এই সময় কখনো উপকরণ বা টুল স্পর্শ করবেন না!<br/>G04 অপেক্ষার সময় কমাবেন না।",
        "s8_title": "৮. অ্যালার্ম কোড তালিকা",
        "s8_desc": "ত্রুটি হলে অ্যালার্ম নম্বর (#3000) দেখুন এবং নিচের তালিকা অনুসরণ করুন।",
        "s8_1_title": "৮-১. মেশিনিং লজিক অ্যালার্ম",
        "s8_1_data": [
            ["অ্যালার্ম", "শর্ত", "অর্থ", "সমাধান"],
            ["1", "#140 > 580", "উপকরণ দৈর্ঘ্য > 580mm", "#101, #102 দেখুন"],
            ["2", "#109 < #111", "কাঁচা OD < সমাপ্ত OD", "#109, #111 দেখুন"],
            ["3", "#110 > #112", "কাঁচা ID > সমাপ্ত ID", "#110, #112 দেখুন"],
            ["7", "#140 <= #103", "উপকরণ খুব ছোট", "উপকরণ দৈর্ঘ্য দেখুন"],
        ],
        "s8_2_title": "৮-২. নিরাপত্তা অ্যালার্ম",
        "s8_2_data": [
            ["অ্যালার্ম", "শর্ত", "অর্থ", "সমাধান"],
            ["203", "OD <= ID", "কাঁচা OD <= কাঁচা ID", "#109, #110 দেখুন"],
            ["204", "FIN OD <= ID", "সমাপ্ত OD <= ID", "#111, #112 দেখুন"],
            ["208", "#114 >= #113", "টিপ প্রস্থ >= সমাপ্ত দৈর্ঘ্য", "#114, #113 দেখুন"],
        ],
        "s8_3_title": "৮-৩. ভেরিয়েবল অনুপস্থিত অ্যালার্ম",
        "s8_3_info": "অ্যালার্ম 101~130 বা 506~532: সংশ্লিষ্ট ভেরিয়েবল (#নম্বর) ইনপুট হয়নি।<br/>উদাহরণ: অ্যালার্ম 109 → #109 (কাঁচা OD) ইনপুট করুন।",
        "s9_title": "৯. মেশিন M-কোড ম্যাপিং",
        "s9_desc": "#130 মেশিন নম্বর অনুযায়ী অটো লোডার/বোরিং বার M-কোড স্বয়ংক্রিয়ভাবে সেট হয়।",
        "s9_data": [
            ["মেশিন\nনম্বর", "নাম", "AL খোলা\n(#131)", "AL বন্ধ\n(#132)", "BR উপরে\n(#133)", "BR নিচে\n(#134)"],
            ["1", "AL-1", "M56", "M55", "M54", "M53"],
            ["4", "HA-4", "M52", "M51", "M54", "M53"],
            ["5", "AL-5", "M171", "M170", "M53", "M54"],
            ["7", "AL-7", "M64", "M63", "M53", "M54"],
            ["8", "AL-8", "M64", "M63", "M54", "M53"],
            ["9", "AL-9", "M63", "M64", "M53", "M54"],
            ["10", "AL-10", "M64", "M63", "M54", "M53"],
            ["13", "HA-S3", "M56", "M55", "M54", "M53"],
            ["14", "HA-S4", "M56", "M55", "M54", "M53"],
        ],
        "s9_info": "AL = অটো লোডার | BR = বোরিং বার<br/>মেশিন পরিবর্তনে শুধু #130 পরিবর্তন করুন।",
        "s9_warn": "মেশিন পরিবর্তনে শুধু #130 পরিবর্তন করুন, M-কোড সরাসরি সম্পাদনা করবেন না।",
        "s10_title": "১০. অপারেটর সতর্কতা",
        "s10_1_title": "১০-১. অবশ্যই যাচাই করুন",
        "s10_1_data": [
            ["নং", "বিষয়", "বিস্তারিত"],
            ["1", "সম্পাদনযোগ্য এলাকা", "শুধু #101~#123\nDANGER-এর পরে সম্পাদনা করবেন না"],
            ["2", "#119 যাচাই", "#113 > 40mm হলে\n#119=1 সেট করতে হবে"],
            ["3", "#102 ফরম্যাট", "শুধু 0,100,200,300,400,500"],
            ["4", "#101 পরিসর", "শুধু 0~99"],
            ["5", "মেশিন পরিবর্তন", "শুধু #130 পরিবর্তন করুন"],
        ],
        "s10_2_title": "১০-২. নিরাপত্তা নিয়ম",
        "s10_2_info": "(১) অটো লিঙ্কের সময় হস্তক্ষেপ করবেন না!<br/>(২) RS40 উপকরণ — 15mm নিরাপত্তা অফসেট স্বয়ংক্রিয়<br/>(৩) ছোট OD (30mm-এর কম) — একই নিরাপত্তা অফসেট<br/>(৪) অ্যালার্ম হলে — নম্বর দেখুন → তালিকা দেখুন → ভেরিয়েবল ঠিক করুন<br/>(৫) G04 অপেক্ষার সময় কমাবেন না",
        "s10_3_title": "১০-৩. মেশিনিং-পূর্ব চেকলিস্ট",
        "s10_3_items": [
            "উপকরণের ধরন (#105) আসল উপকরণের সাথে মিলছে?",
            "কাঁচা OD/ID (#109, #110) পরিমাপের সাথে মিলছে?",
            "সমাপ্ত মাপ (#111~#118) ড্রয়িং-এর সাথে মিলছে?",
            "মেশিন নম্বর (#130) আসল মেশিনের সাথে মিলছে?",
            "RPM (#120, #121) উপকরণ/টুলের জন্য উপযুক্ত?",
            "প্রথম চালনায় একক মোড (#108=1) পরীক্ষা?",
            "অটো লোডার কাজ যাচাই করেছেন?",
            "কাটিং ইনসার্ট (T02) ক্ষয় দেখেছেন?",
        ],
        "summary_title": "O0852 প্রোগ্রাম মূল সারসংক্ষেপ",
        "summary_items": [
            "অপারেটর শুধু #101~#123 ভেরিয়েবল পরিবর্তন করেন",
            "উপকরণের ধরন (CN/RS/CM) অনুযায়ী ফিড রেট স্বয়ংক্রিয়",
            "ফেসিং → স্টেপ → চ্যামফার → কাটিং সাইকেল স্বয়ংক্রিয় পুনরাবৃত্তি",
            "উপকরণ কম হলে অটো লিঙ্কে স্বয়ংক্রিয় টানা হয়",
            "৩০+ অ্যালার্ম দিয়ে ত্রুটি ও নিরাপত্তা সমস্যা প্রতিরোধ",
            "৯টি মেশিনের M-কোড স্বয়ংক্রিয় ম্যাপিং (#130 পরিবর্তনে)",
        ],
        "summary_footer": "প্রশ্ন থাকলে যেকোনো সময় জিজ্ঞাসা করুন",
    },

    "ne": {
        "font": "Nirmala", "fontBd": "NirmalaBd",
        "filename": "O0852_Training_NE.pdf",
        "cover_title": "O0852 कार्यक्रम पूर्ण विश्लेषण",
        "cover_sub": "CNC म्याक्रो कार्यक्रम तालिम सामग्री",
        "cover_desc": "रिङ आकारको पार्ट्स स्वचालित निरन्तर मेसिनिङ | FANUC म्याक्रो",
        "cover_features": [
            ["म्याक्रो स्वचालन", "भेरिएबल इनपुट मात्रले सामग्री/माप/सर्त सेट गर्नुहोस्"],
            ["निरन्तर मेसिनिङ", "फेसिङ→स्टेप→च्याम्फर→काटिङ साइकल स्वचालित दोहोरिन्छ"],
            ["अटो लिङ्क", "सामग्री कम हुँदा स्वचालित रूपमा तानिन्छ"],
            ["सुरक्षा जाँच", "३०+ अलार्मले इनपुट त्रुटि रोक्छ"],
        ],
        "toc_title": "विषयसूची",
        "toc_items": [
            ("०१", "समग्र सारांश", "कार्यक्रमले के गर्छ एक झलकमा"),
            ("०२", "कार्यक्रमको संरचना", "४ वटा सब-प्रोग्रामको सम्बन्ध"),
            ("०३", "इनपुट प्यारामिटर", "अपरेटरले परिवर्तन गर्ने (#101~#123)"),
            ("०४", "सामग्री अनुसार फिड रेट", "CN / RS / CM स्वचालित सेटिङ"),
            ("०५", "उन्नत सेटिङ र सुरक्षा", "प्रणाली भेरिएबल र जाँच"),
            ("०६", "मेसिनिङ साइकल विस्तृत", "फेसिङ → स्टेप → च्याम्फर → काटिङ"),
            ("०७", "अटो लिङ्क र बाँकी मेसिनिङ", "स्वचालित सामग्री तान्ने प्रणाली"),
            ("०८", "अलार्म कोड तालिका", "त्रुटि नम्बर र समाधान"),
            ("०९", "मेसिन M-कोड म्यापिङ", "९ वटा मेसिनको कन्फिगरेसन"),
            ("१०", "अपरेटर सावधानी", "सुरक्षा नियम र चेकलिस्ट"),
        ],
        "s1_title": "१. समग्र सारांश",
        "s1_desc": "O0852 एक FANUC म्याक्रो कार्यक्रम हो जसले CNC लेथमा पाइप सामग्रीबाट<br/>स्वचालित रूपमा रिङ आकारका पार्ट्स (बुशिङ/बेयरिङ) ठूलो मात्रामा उत्पादन गर्छ।",
        "s1_summary_title": "हालको सेटिङ सारांश",
        "s1_summary": [
            ["विषय", "सामग्री", "तयार पार्ट", "उत्पादन"],
            ["विवरण", "CN, OD70 × ID56\nलम्बाइ 543mm", "OD64.9 × ID58.3\nलम्बाइ 7.823mm", "प्रति पिसमा 9.843mm\nप्रति पुलमा ~६ वटा"],
        ],
        "s1_feat_title": "मुख्य विशेषताहरू",
        "s1_features": [
            ["म्याक्रो स्वचालन", "भेरिएबल इनपुटले सामग्री, माप, काटिङ सर्त सेट गर्छ"],
            ["निरन्तर मेसिनिङ", "फेसिङ → स्टेप कट → च्याम्फर → काटिङ साइकल स्वचालित दोहोरिन्छ"],
            ["अटो लिङ्क", "सामग्री कम हुँदा स्वचालित रूपमा तानेर मेसिनिङ जारी राख्छ"],
            ["सुरक्षा जाँच", "३०+ अलार्म चेकले इनपुट त्रुटि, माप विरोधाभास रोक्छ"],
        ],
        "s2_title": "२. कार्यक्रमको संरचना",
        "s2_desc": "४ वटा सब-प्रोग्राम क्रमशः कल गरिन्छ।",
        "s2_table": [
            ["कार्यक्रम", "नाम", "भूमिका", "कल विधि"],
            ["O0852", "मुख्य सेटअप", "प्यारामिटर इनपुट, जाँच, मेसिन कन्फिग", "सिधै चलाउनुहोस्"],
            ["O9001", "मुख्य लजिक", "लम्बाइ/संख्या गणना, अलार्म जाँच", "M98 P9001"],
            ["O9002", "मुख्य मेसिनिङ", "फेसिङ→स्टेप→च्याम्फर→काटिङ, अटो लिङ्क", "M98 P9002"],
            ["O9003", "बाँकी मेसिनिङ", "बाँकी सामग्री तानिन्छ, पुन: सेटअप", "M98 P9003"],
        ],
        "s2_flow_title": "कल फ्लो",
        "s2_flow": "O0852 (सेटअप) → O9001 (गणना/जाँच) → O9002 (मेसिनिङ) → O9003 (बाँकी)",
        "s2_tool_title": "प्रयोग गरिने उपकरण",
        "s2_tools": [
            ["उपकरण", "उद्देश्य", "मुख्य प्रक्रिया", "टिप्पणी"],
            ["T01", "बोरिङ बार (ID)", "रफिङ, फिनिसिङ, च्याम्फर", "मुख्य उपकरण"],
            ["T02", "काटिङ इन्सर्ट", "फेसिङ, काटिङ", "फेस + पार्टिङ"],
            ["T03", "अटो लिङ्क", "सामग्री तान्ने", "क्ल्याम्प/अनक्ल्याम्प"],
        ],
        "s3_title": "३. इनपुट प्यारामिटर",
        "s3_info": "अपरेटरले परिवर्तन गर्ने क्षेत्र (लाइन 4~29, #101~#123 मात्र)",
        "s3_1_title": "३-१. आधारभूत सेटिङ (#101 ~ #108)",
        "s3_1_data": [
            ["भेरिएबल", "विवरण", "मान", "इनपुट नियम"],
            ["#101", "सामग्री लम्बाइ एकाइ", "43", "दायरा 0~99"],
            ["#102", "सामग्री लम्बाइ सयौं", "500", "0,100,200,300,400,500 मात्र"],
            ["#103", "चक लम्बाइ", "75", "mm एकाइ"],
            ["#104", "प्रारम्भिक फेस कट", "1", "अधिकतम 40"],
            ["#105", "सामग्रीको प्रकार", "1", "1=CN, 2=RS, 3=CM"],
            ["#106", "प्रक्रियाको प्रकार", "3", "3=च्याम्फर, 4=च्याम्फर(विस्तारित)"],
            ["#107", "रफिङ ON/OFF", "0", "0=OFF, 1=ON"],
            ["#108", "एकल मोड", "1", "1=एक एक गरेर"],
        ],
        "s3_1_note": "सामग्री लम्बाइ: #101 + #102 = जम्मा → उदाहरण: 500 + 43 = 543mm",
        "s3_2_title": "३-२. माप डाटा (#109 ~ #118)",
        "s3_2_data": [
            ["भेरिएबल", "विवरण", "मान", "एकाइ"],
            ["#109", "कच्चा OD (बाहिरी व्यास)", "70", "mm"],
            ["#110", "कच्चा ID (भित्री व्यास)", "56", "mm"],
            ["#111", "तयार OD", "64.90", "mm"],
            ["#112", "तयार ID", "58.30", "mm"],
            ["#113", "तयार लम्बाइ", "7.823", "mm"],
            ["#114", "टिप चौडाइ (काटिङ)", "2.02", "mm"],
            ["#115", "ID रेडियस", "0.34", "mm"],
            ["#116", "OD रेडियस", "0.36", "mm"],
            ["#117", "ID च्याम्फर", "0.37", "mm"],
            ["#118", "OD च्याम्फर", "0.60", "mm"],
            ["#119", "लम्बाइ>40mm फ्ल्याग", "0", "40 नाघे 1 सेट गर्नुहोस्"],
        ],
        "s3_3_title": "३-३. काटिङ सर्त (#120 ~ #123)",
        "s3_3_data": [
            ["भेरिएबल", "विवरण", "मान", "टिप्पणी"],
            ["#120", "T01 RPM", "1700", "बोरिङ बार गति"],
            ["#121", "T02 RPM", "1700", "काटिङ इन्सर्ट गति"],
            ["#122", "पुल दूरी", "1", "अटो लिङ्क पुल दूरी"],
            ["#123", "मार्जिन मान", "0.2", "दायरा 0~3"],
        ],
        "s4_title": "४. सामग्री अनुसार फिड रेट",
        "s4_desc": "सामग्रीको प्रकार (#105) अनुसार फिड गुणाङ्क स्वचालित रूपमा निर्धारण हुन्छ।",
        "s4_data": [
            ["सामग्री", "#124 (T01 रफ)", "#125 (T01 स्टेप)", "#126 (T01 च्याम्फर)", "#127 (T02 कट)"],
            ["CN (#105=1)", "0.15", "0.12", "0.15", "0.09"],
            ["RS (#105=2)", "0.13", "0.08", "0.05", "0.08"],
            ["CM (#105=3)", "0.13", "0.09", "0.09", "0.09"],
            ["पूर्वनिर्धारित", "0.15", "0.10", "0.10", "0.10"],
        ],
        "s4_formula_title": "फिड रेट सूत्र",
        "s4_formula": "F मान = RPM × फिड गुणाङ्क",
        "s4_example_title": "गणना उदाहरण (RPM = 1700)",
        "s4_examples": [
            ["प्रक्रिया", "सामग्री", "गणना", "परिणाम"],
            ["T01 रफिङ", "CN", "1700 × 0.15", "F255 mm/min"],
            ["T01 स्टेप", "CN", "1700 × 0.12", "F204 mm/min"],
            ["T02 काटिङ", "CN", "1700 × 0.09", "F153 mm/min"],
        ],
        "s4_note": "RS सामग्री बढी कडा हुन्छ, त्यसैले फिड कम सेट गरिन्छ।",
        "s5_title": "५. उन्नत सेटिङ र सुरक्षा",
        "s5_warn": "खतरा: सम्पादन रोक्नुहोस् — यो भाग प्रणाली स्वचालित सेटिङ हो। परिवर्तन नगर्नुहोस्!",
        "s5_1_title": "५-१. प्रणाली भेरिएबल (#500 शृंखला)",
        "s5_1_data": [
            ["भेरिएबल", "विवरण", "मान", "भूमिका"],
            ["#530", "सुरक्षा लम्बाइ", "12", "चक छेउबाट सुरक्षित दूरी"],
            ["#531", "न्यूनतम बाँकी लम्बाइ", "15", "मेसिनिङ पछि कम्तीमा 15mm जब(jaw)मा समातिएको हुनुपर्छ"],
            ["#532", "बाँकी-सुरक्षा अफसेट", "1.03", "अन्तिम अटो लिङ्क पुलमा सही स्थिति गणनाको लागि"],
            ["#506", "Z रफ अधिकतम गहिराइ", "55", "प्रति पासमा अधिकतम Z लम्बाइ"],
            ["#508", "T01 नोज R", "0.2", "उपकरण टिप त्रिज्या"],
            ["#509", "रफिङ अन्तराल", "0.4", "X दिशामा अन्तराल"],
            ["#510", "च्याम्फर अन्तराल", "1", "च्याम्फर कटको अन्तराल"],
        ],
        "s5_2_title": "५-२. RS40 / सानो व्यासको सुरक्षा",
        "s5_2_info": "RS40 सामग्री (#105=2) वा OD 30mm भन्दा कम → #129=15 स्वचालित<br/>कम कठोरताका कारण सामग्री छिटकिन नदिन 15mm अफसेट थपिन्छ।",
        "s6_title": "६. मेसिनिङ साइकल विस्तृत",
        "s6_1_title": "६-१. समग्र प्रवाह",
        "s6_1_data": [
            ["चरण", "प्रक्रिया", "उपकरण", "मुख्य कार्य"],
            ["(१)", "कोअर्डिनेट सेट", "-", "G10 L2 P0 Z[#103] — G54 Z मूल बिन्दु सेट"],
            ["(२)", "फेसिङ", "T02", "छेउको सतह समतल पार्ने (OD→ID दिशा)"],
            ["(३)", "स्टेप कट", "T01", "रफिङ(वैकल्पिक) + फिनिसिङ — Z दिशामा काट्ने"],
            ["(४)", "च्याम्फर", "T01", "OD-R, ID-R च्याम्फर"],
            ["(५)", "काटिङ", "T02", "पार्ट अलग गर्ने + M12 गणना"],
            ["", "दोहोरिने", "", "(३)~(५) संख्या (#517) सम्म दोहोरिन्छ"],
        ],
        "s6_calc_title": "मुख्य गणना भेरिएबल",
        "s6_calc_data": [
            ["सूत्र", "अर्थ", "हालको मान"],
            ["#505 = #113 + #114", "प्रति पिसको एकाइ (लम्बाइ + काटिङ चौडाइ)", "9.843mm"],
            ["#500 = #103 - #530", "प्रभावकारी मेसिनिङ लम्बाइ", "63mm"],
            ["#517 = FIX[#500/#505]", "प्रति पुल साइकलमा संख्या", "६ वटा"],
            ["#518 = #517 × #505", "प्रति साइकलमा जम्मा मेसिनिङ लम्बाइ", "59.058mm"],
        ],
        "s7_title": "७. अटो लिङ्क र बाँकी मेसिनिङ",
        "s7_1_title": "७-१. अटो लिङ्क क्रम (N210)",
        "s7_1_desc": "सामग्री कम हुँदा स्वचालित रूपमा तान्ने मुख्य कार्य।",
        "s7_1_data": [
            ["चरण", "कोड", "कार्य", "विवरण"],
            ["(१)", "M05", "स्पिन्डल बन्द", "सुरक्षाको लागि घुमाइ बन्द"],
            ["(२)", "G00 T03, X0", "उपकरण छनोट", "अटो लिङ्क उपकरण, केन्द्रमा सार्ने"],
            ["(३)", "Z[-#522+#122]", "पुल स्थानमा जाने", "F200 मा स्थानमा सार्ने"],
            ["(४)", "M[#132]", "अटो लोडर बन्द", "लोडरले सामग्री समात्छ"],
            ["(५)", "M69+G04 P2500", "चक अनक्ल्याम्प", "चकले छाड्छ (2.5से पर्खनु)"],
            ["(६)", "W[लम्बाइ+मार्जिन]", "सामग्री तान्ने", "सामग्री अगाडि तानिन्छ"],
            ["(७)", "M68+G04 P2500", "चक क्ल्याम्प", "चकले फेरि समात्छ (2.5से पर्खनु)"],
            ["(८)", "M[#131]", "अटो लोडर खोल्ने", "लोडर छाड्छ"],
            ["(९)", "G00 Z100→GOTO100", "फर्कने", "सुरक्षित स्थान → फेसिङबाट सुरु"],
        ],
        "s7_warn": "सुरक्षा चेतावनी: अटो लिङ्कको समयमा चक खुल्छ र बन्द हुन्छ।<br/>यो समयमा कहिल्यै सामग्री वा उपकरण नछुनुहोस्!<br/>G04 पर्खने समय नघटाउनुहोस्।",
        "s8_title": "८. अलार्म कोड तालिका",
        "s8_desc": "त्रुटि हुँदा अलार्म नम्बर (#3000) जाँच गर्नुहोस् र तलको तालिका हेर्नुहोस्।",
        "s8_1_title": "८-१. मेसिनिङ लजिक अलार्म",
        "s8_1_data": [
            ["अलार्म", "सर्त", "अर्थ", "समाधान"],
            ["1", "#140 > 580", "सामग्री लम्बाइ > 580mm", "#101, #102 जाँच"],
            ["2", "#109 < #111", "कच्चा OD < तयार OD", "#109, #111 जाँच"],
            ["3", "#110 > #112", "कच्चा ID > तयार ID", "#110, #112 जाँच"],
            ["7", "#140 <= #103", "सामग्री धेरै छोटो", "सामग्री लम्बाइ जाँच"],
        ],
        "s8_2_title": "८-२. सुरक्षा अलार्म",
        "s8_2_data": [
            ["अलार्म", "सर्त", "अर्थ", "समाधान"],
            ["203", "OD <= ID", "कच्चा OD <= कच्चा ID", "#109, #110 जाँच"],
            ["204", "FIN OD <= ID", "तयार OD <= ID", "#111, #112 जाँच"],
            ["208", "#114 >= #113", "टिप चौडाइ >= तयार लम्बाइ", "#114, #113 जाँच"],
        ],
        "s8_3_title": "८-३. भेरिएबल नभएको अलार्म",
        "s8_3_info": "अलार्म 101~130 वा 506~532: सम्बन्धित भेरिएबल (#नम्बर) इनपुट गरिएको छैन।<br/>उदाहरण: अलार्म 109 → #109 (कच्चा OD) इनपुट गर्नुहोस्।",
        "s9_title": "९. मेसिन M-कोड म्यापिङ",
        "s9_desc": "#130 मेसिन नम्बर अनुसार अटो लोडर/बोरिङ बार M-कोड स्वचालित सेट हुन्छ।",
        "s9_data": [
            ["मेसिन\nनं.", "नाम", "AL खोल्ने\n(#131)", "AL बन्द\n(#132)", "BR माथि\n(#133)", "BR तल\n(#134)"],
            ["1", "AL-1", "M56", "M55", "M54", "M53"],
            ["4", "HA-4", "M52", "M51", "M54", "M53"],
            ["5", "AL-5", "M171", "M170", "M53", "M54"],
            ["7", "AL-7", "M64", "M63", "M53", "M54"],
            ["8", "AL-8", "M64", "M63", "M54", "M53"],
            ["9", "AL-9", "M63", "M64", "M53", "M54"],
            ["10", "AL-10", "M64", "M63", "M54", "M53"],
            ["13", "HA-S3", "M56", "M55", "M54", "M53"],
            ["14", "HA-S4", "M56", "M55", "M54", "M53"],
        ],
        "s9_info": "AL = अटो लोडर | BR = बोरिङ बार<br/>मेसिन परिवर्तनमा #130 मात्र परिवर्तन गर्नुहोस्।",
        "s9_warn": "मेसिन परिवर्तनमा #130 मात्र परिवर्तन गर्नुहोस्, M-कोड सिधै सम्पादन नगर्नुहोस्।",
        "s10_title": "१०. अपरेटर सावधानी",
        "s10_1_title": "१०-१. अनिवार्य जाँच",
        "s10_1_data": [
            ["नं.", "विषय", "विस्तृत"],
            ["1", "सम्पादनयोग्य क्षेत्र", "#101~#123 मात्र\nDANGER पछि सम्पादन नगर्नुहोस्"],
            ["2", "#119 जाँच", "#113 > 40mm भए\n#119=1 सेट गर्नुपर्छ"],
            ["3", "#102 ढाँचा", "0,100,200,300,400,500 मात्र"],
            ["4", "#101 दायरा", "0~99 मात्र"],
            ["5", "मेसिन परिवर्तन", "#130 मात्र परिवर्तन गर्नुहोस्"],
        ],
        "s10_2_title": "१०-२. सुरक्षा नियम",
        "s10_2_info": "(१) अटो लिङ्कको समयमा हस्तक्षेप नगर्नुहोस्!<br/>(२) RS40 सामग्री — 15mm सुरक्षा अफसेट स्वचालित<br/>(३) सानो OD (30mm भन्दा कम) — उही सुरक्षा अफसेट<br/>(४) अलार्म आएमा — नम्बर जाँच → तालिका हेर्नुहोस् → भेरिएबल ठीक गर्नुहोस्<br/>(५) G04 पर्खने समय नघटाउनुहोस्",
        "s10_3_title": "१०-३. मेसिनिङ-पूर्व चेकलिस्ट",
        "s10_3_items": [
            "सामग्रीको प्रकार (#105) वास्तविक सामग्रीसँग मिल्छ?",
            "कच्चा OD/ID (#109, #110) मापनसँग मिल्छ?",
            "तयार माप (#111~#118) नक्सासँग मिल्छ?",
            "मेसिन नम्बर (#130) वास्तविक मेसिनसँग मिल्छ?",
            "RPM (#120, #121) सामग्री/उपकरणको लागि उपयुक्त?",
            "पहिलो चालनमा एकल मोड (#108=1) परीक्षण?",
            "अटो लोडर सञ्चालन जाँच गरिसक्नुभयो?",
            "काटिङ इन्सर्ट (T02) घिसावट जाँच?",
        ],
        "summary_title": "O0852 कार्यक्रम मुख्य सारांश",
        "summary_items": [
            "अपरेटरले #101~#123 भेरिएबल मात्र परिवर्तन गर्छन्",
            "सामग्रीको प्रकार (CN/RS/CM) अनुसार फिड रेट स्वचालित",
            "फेसिङ → स्टेप → च्याम्फर → काटिङ साइकल स्वचालित दोहोरिन्छ",
            "सामग्री कम हुँदा अटो लिङ्कले स्वचालित तान्छ",
            "३०+ अलार्मले त्रुटि र सुरक्षा समस्या रोक्छ",
            "९ वटा मेसिनको M-कोड स्वचालित म्यापिङ (#130 परिवर्तनमा)",
        ],
        "summary_footer": "प्रश्न भएमा जुनसुकै बेला सोध्नुहोस्",
    },

    "fil": {
        "font": "Malgun", "fontBd": "MalgunBd",
        "filename": "O0852_Training_FIL.pdf",
        "cover_title": "O0852 Program Kumpletong Pagsusuri",
        "cover_sub": "CNC Macro Program Training Material",
        "cover_desc": "Ring-shaped na Parte Awtomatikong Tuloy-tuloy na Machining  |  FANUC Macro",
        "cover_features": [
            ["Macro Automation", "Itakda ang materyal/sukat/kundisyon sa pamamagitan ng variable input lamang"],
            ["Tuloy-tuloy na Machining", "Face→Step→Chamfer→Cut cycle na awtomatikong umuulit"],
            ["Auto Link", "Awtomatikong hinahatak ang materyal kapag maikli na"],
            ["Safety Validation", "30+ alarm para maiwasan ang input error"],
        ],
        "toc_title": "Talaan ng Nilalaman",
        "toc_items": [
            ("01", "Pangkalahatang Buod", "Ano ang ginagawa ng program sa isang tingin"),
            ("02", "Istruktura ng Program", "Relasyon ng 4 na sub-program"),
            ("03", "Input Parameter", "Mga variable na binabago ng operator (#101~#123)"),
            ("04", "Feed Rate ayon sa Materyal", "CN / RS / CM awtomatikong setting"),
            ("05", "Advanced Settings at Safety", "System variable at validation"),
            ("06", "Detalye ng Machining Cycle", "Face → Step → Chamfer → Cut"),
            ("07", "Auto Link at Natitirang Materyal", "Awtomatikong material pull system"),
            ("08", "Alarm Code Table", "Error code at solusyon"),
            ("09", "Machine M-Code Mapping", "9 na machine configuration"),
            ("10", "Paalala sa Operator", "Safety rules at checklist"),
        ],
        "s1_title": "1. Pangkalahatang Buod",
        "s1_desc": "Ang O0852 ay isang FANUC macro program na awtomatiko at tuloy-tuloy na<br/>nagma-machine ng pipe material sa CNC lathe para mass-produce ng ring-shaped na parte (bushing/bearing).",
        "s1_summary_title": "Buod ng Kasalukuyang Setting",
        "s1_summary": [
            ["Item", "Materyal", "Tapos na Parte", "Produksyon"],
            ["Specs", "CN, OD70 × ID56\nHaba 543mm", "OD64.9 × ID58.3\nHaba 7.823mm", "9.843mm bawat piraso\n~6 bawat pull cycle"],
        ],
        "s1_feat_title": "Pangunahing Tampok",
        "s1_features": [
            ["Macro Automation", "Itakda ang materyal, sukat, cutting condition sa pamamagitan ng variable input. Multi-product support nang walang code change"],
            ["Tuloy-tuloy na Machining", "Face → Step cut → Chamfer → Cut-off cycle na awtomatikong umuulit"],
            ["Auto Link", "Kapag maikli na ang materyal, awtomatikong hinahatak at nagpapatuloy ng machining"],
            ["Safety Validation", "30+ alarm check para maiwasan ang input error, dimension conflict, safety issue"],
        ],
        "s2_title": "2. Istruktura ng Program",
        "s2_desc": "4 na sub-program ang sunud-sunod na tinatawag.",
        "s2_table": [
            ["Program", "Pangalan", "Tungkulin", "Paraan ng Pagtawag"],
            ["O0852", "Main Setup", "Parameter input, validation, machine config", "Direktang patakbuhin"],
            ["O9001", "Main Logic", "Haba/bilang kalkulasyon, alarm check", "M98 P9001"],
            ["O9002", "Main Machining", "Face→Step→Chamfer→Cut ulit, Auto Link", "M98 P9002"],
            ["O9003", "Natitira", "Natitirang materyal pull, re-setup, final cycle", "M98 P9003"],
        ],
        "s2_flow_title": "Call Flow",
        "s2_flow": "O0852 (Setup) → O9001 (Kalkulasyon/Validate) → O9002 (Machining Cycle) → O9003 (Natitira)",
        "s2_tool_title": "Mga Ginagamit na Tool",
        "s2_tools": [
            ["Tool", "Layunin", "Pangunahing Proseso", "Tala"],
            ["T01", "Boring bar (ID machining)", "Rough, Finish, Chamfer", "Pangunahing tool"],
            ["T02", "Cut-off insert", "Facing, Cut-off", "Face + Parting"],
            ["T03", "Auto Link", "Material pull", "Clamp/Unclamp"],
        ],
        "s3_title": "3. Input Parameter",
        "s3_info": "Lugar na maaaring baguhin ng operator (Line 4~29, #101~#123 lamang ang baguhin)",
        "s3_1_title": "3-1. Pangunahing Setting (#101 ~ #108)",
        "s3_1_data": [
            ["Variable", "Paglalarawan", "Halaga", "Panuntunan sa Input"],
            ["#101", "Haba ng materyal ones digit", "43", "Saklaw 0~99"],
            ["#102", "Haba ng materyal hundreds", "500", "0,100,200,300,400,500 lamang"],
            ["#103", "Haba ng chuck (grip)", "75", "mm unit"],
            ["#104", "Unang face cut", "1", "Max 40"],
            ["#105", "Uri ng materyal", "1", "1=CN, 2=RS, 3=CM"],
            ["#106", "Uri ng proseso", "3", "3=Chamfer, 4=Chamfer(ext)"],
            ["#107", "Rough ON/OFF", "0", "0=OFF, 1=ON"],
            ["#108", "Single mode", "1", "1=Isa-isa"],
        ],
        "s3_1_note": "Haba ng materyal: #101 + #102 = Kabuuan → Hal: 500 + 43 = 543mm",
        "s3_2_title": "3-2. Dimension Data (#109 ~ #118)",
        "s3_2_data": [
            ["Variable", "Paglalarawan", "Halaga", "Unit"],
            ["#109", "Raw OD (Labas na Diameter)", "70", "mm"],
            ["#110", "Raw ID (Loob na Diameter)", "56", "mm"],
            ["#111", "Finished OD", "64.90", "mm"],
            ["#112", "Finished ID", "58.30", "mm"],
            ["#113", "Finished na Haba", "7.823", "mm"],
            ["#114", "Lapad ng Tip (Cut-off)", "2.02", "mm"],
            ["#115", "ID Radius", "0.34", "mm"],
            ["#116", "OD Radius", "0.36", "mm"],
            ["#117", "ID Chamfer", "0.37", "mm"],
            ["#118", "OD Chamfer", "0.60", "mm"],
            ["#119", "Haba>40mm flag", "0", "Itakda 1 kung >40mm"],
        ],
        "s3_3_title": "3-3. Cutting Condition (#120 ~ #123)",
        "s3_3_data": [
            ["Variable", "Paglalarawan", "Halaga", "Tala"],
            ["#120", "T01 RPM", "1700", "Bilis ng boring bar"],
            ["#121", "T02 RPM", "1700", "Bilis ng cut-off insert"],
            ["#122", "Pull Distance", "1", "Auto Link pull distance"],
            ["#123", "Margin Value", "0.2", "Saklaw 0~3"],
        ],
        "s4_title": "4. Feed Rate ayon sa Uri ng Materyal",
        "s4_desc": "Ang feed coefficient ay awtomatikong natutukoy ayon sa uri ng materyal (#105).",
        "s4_data": [
            ["Materyal", "#124 (T01 Rough)", "#125 (T01 Step)", "#126 (T01 Chamfer)", "#127 (T02 Cut)"],
            ["CN (#105=1)", "0.15", "0.12", "0.15", "0.09"],
            ["RS (#105=2)", "0.13", "0.08", "0.05", "0.08"],
            ["CM (#105=3)", "0.13", "0.09", "0.09", "0.09"],
            ["Default", "0.15", "0.10", "0.10", "0.10"],
        ],
        "s4_formula_title": "Feed Rate Formula",
        "s4_formula": "F value = RPM × Feed coefficient",
        "s4_example_title": "Halimbawa ng Kalkulasyon (RPM = 1700)",
        "s4_examples": [
            ["Proseso", "Materyal", "Kalkulasyon", "Resulta"],
            ["T01 Rough", "CN", "1700 × 0.15", "F255 mm/min"],
            ["T01 Step", "CN", "1700 × 0.12", "F204 mm/min"],
            ["T02 Cut-off", "CN", "1700 × 0.09", "F153 mm/min"],
            ["T01 Rough", "RS", "1700 × 0.13", "F221 mm/min"],
        ],
        "s4_note": "Ang RS material ay mas mababa ang feed rate dahil mas matigas (kailangan ng maingat na pagputol).",
        "s5_title": "5. Advanced Settings at Safety",
        "s5_warn": "PANGANIB: HUWAG BAGUHIN — Ang bahaging ito ay awtomatikong naka-configure. HUWAG baguhin!",
        "s5_1_title": "5-1. System Variable (#500 series)",
        "s5_1_data": [
            ["Variable", "Paglalarawan", "Halaga", "Tungkulin"],
            ["#530", "Safety length", "12", "Ligtas na distansya mula sa dulo ng chuck"],
            ["#531", "Min remaining length", "15", "Min 15mm ang dapat nakakapit sa jaw pagkatapos ng machining\n→ Maikli = panganib na matanggal ang workpiece"],
            ["#532", "Remaining-safety offset", "1.03", "Calibration value para sa tumpak na positioning\nsa huling Auto Link pull"],
            ["#506", "Z rough max depth", "55", "Max Z-direction na haba bawat pass"],
            ["#508", "T01 nose R", "0.2", "Radius ng tool tip"],
            ["#509", "Rough interval", "0.4", "X-direction interval"],
            ["#510", "Chamfer interval", "1", "Chamfer cut interval"],
        ],
        "s5_2_title": "5-2. RS40 / Maliit na Diameter Safety",
        "s5_2_info": "RS40 material (#105=2) o OD < 30mm (#109<30) → #129=15 awtomatiko<br/>Nagdadagdag ng 15mm residual offset para maiwasan ang pagtanggal ng workpiece dahil sa mababang rigidity.",
        "s6_title": "6. Detalye ng Machining Cycle",
        "s6_1_title": "6-1. Pangkalahatang Daloy",
        "s6_1_data": [
            ["Hakbang", "Proseso", "Tool", "Pangunahing Aksyon"],
            ["(1)", "Coordinate set", "-", "G10 L2 P0 Z[#103] — Itakda ang G54 work Z origin"],
            ["(2)", "Facing", "T02", "Patag ang dulo (OD→ID direksyon)"],
            ["(3)", "Step cut", "T01", "Rough(opsyonal) + Finish — Z-direction cut"],
            ["(4)", "Chamfer", "T01", "OD-R, ID-R chamfer"],
            ["(5)", "Cut-off", "T02", "Paghiwalay ng parte + M12 count"],
            ["", "Ulit", "", "Ulitin (3)~(5) ayon sa bilang (#517)"],
        ],
        "s6_calc_title": "Mahahalagang Kalkulasyon",
        "s6_calc_data": [
            ["Formula", "Ibig Sabihin", "Kasalukuyang Halaga"],
            ["#505 = #113 + #114", "Unit bawat piraso (haba + cut width)", "9.843mm"],
            ["#500 = #103 - #530", "Epektibong machining length", "63mm"],
            ["#517 = FIX[#500/#505]", "Piraso bawat pull cycle", "6 pcs"],
            ["#518 = #517 × #505", "Kabuuang machining length bawat cycle", "59.058mm"],
            ["#140 = #101+#102-#129", "Kabuuang haba ng materyal (may offset)", "543mm"],
        ],
        "s7_title": "7. Auto Link at Natitirang Materyal",
        "s7_1_title": "7-1. Auto Link Sequence (N210)",
        "s7_1_desc": "Pangunahing function na awtomatikong humihila ng materyal kapag maikli na.",
        "s7_1_data": [
            ["Hakbang", "Code", "Aksyon", "Paglalarawan"],
            ["(1)", "M05", "Spindle stop", "Itigil ang pag-ikot para sa kaligtasan"],
            ["(2)", "G00 T03, X0", "Tool select", "Auto Link tool, ilipat sa gitna"],
            ["(3)", "Z[-#522+#122]", "Lapitan ang pull pos", "Ilipat sa pull position sa F200"],
            ["(4)", "M[#132]", "Auto Loader CLOSE", "Hawakan ng loader ang materyal"],
            ["(5)", "M69+G04 P2500", "Chuck UNCLAMP", "Binitawan ng chuck (2.5s hintay)"],
            ["(6)", "W[haba+margin]", "Hilahin ang materyal", "Hilahin ang materyal pasulong"],
            ["(7)", "M68+G04 P2500", "Chuck CLAMP", "Hawakan ulit ng chuck (2.5s hintay)"],
            ["(8)", "M[#131]", "Auto Loader OPEN", "Binitawan ng loader"],
            ["(9)", "G00 Z100→GOTO100", "Bumalik", "Ligtas na posisyon → magsimula mula sa facing"],
        ],
        "s7_warn": "BABALA SA KALIGTASAN: Nagbubukas at nagsasara ang chuck sa panahon ng Auto Link.<br/>HUWAG hawakan ang materyal o tool sa panahong ito!<br/>HUWAG bawasan ang G04 wait time.",
        "s8_title": "8. Alarm Code Table",
        "s8_desc": "Kapag may error, tingnan ang alarm number (#3000) at sumangguni sa talahanayan sa ibaba.",
        "s8_1_title": "8-1. Machining Logic Alarm",
        "s8_1_data": [
            ["Alarm", "Kundisyon", "Ibig Sabihin", "Aksyon"],
            ["1", "#140 > 580", "Haba ng materyal > 580mm", "Tingnan #101, #102"],
            ["2", "#109 < #111", "Raw OD < Finished OD", "Tingnan #109, #111"],
            ["3", "#110 > #112", "Raw ID > Finished ID", "Tingnan #110, #112"],
            ["4", "#125 >= 0.14", "T01 step feed masyadong mataas", "Tingnan ang uri ng materyal (#105)"],
            ["5", "#102 unit error", "Hindi nasa 100s unit", "Gamitin 0/100/200/300/400/500"],
            ["6", "#101 > 99", "Ones digit lampas sa saklaw", "Palitan ng 0~99"],
            ["7", "#140 <= #103", "Masyadong maikli ang materyal", "Tingnan ang haba ng materyal"],
            ["10", "#104 > 40", "Unang face cut masyadong malaki", "Itakda #104 sa 40 o mas mababa"],
        ],
        "s8_2_title": "8-2. Safety Validation Alarm",
        "s8_2_data": [
            ["Alarm", "Kundisyon", "Ibig Sabihin", "Aksyon"],
            ["203", "OD <= ID", "Raw OD <= Raw ID", "Tingnan #109, #110"],
            ["204", "FIN OD <= FIN ID", "Finished OD <= ID", "Tingnan #111, #112"],
            ["205", "#109 <= 0", "OD ay 0 o negatibo", "Tingnan #109"],
            ["206", "#120 <= 0", "T01 RPM ay 0 o mas mababa", "Tingnan #120"],
            ["207", "#121 <= 0", "T02 RPM ay 0 o mas mababa", "Tingnan #121"],
            ["208", "#114 >= #113", "Tip width >= Fin length", "Tingnan #114, #113"],
        ],
        "s8_3_title": "8-3. Nawawalang Variable Alarm (101~130, 506~532)",
        "s8_3_info": "Alarm 101~130 o 506~532: Hindi na-input ang katumbas na variable (#number).<br/>Hal) Alarm 109 → #109 (Raw OD) hindi na-input → Ilagay ang halaga.",
        "s9_title": "9. Machine M-Code Mapping",
        "s9_desc": "Ang M-code para sa Auto Loader/Boring bar ay awtomatikong naka-map ayon sa machine number (#130).",
        "s9_data": [
            ["Machine\nNo.", "Pangalan", "AL OPEN\n(#131)", "AL CLOSE\n(#132)", "BR UP\n(#133)", "BR DOWN\n(#134)"],
            ["1", "AL-1", "M56", "M55", "M54", "M53"],
            ["4", "HA-4", "M52", "M51", "M54", "M53"],
            ["5", "AL-5", "M171", "M170", "M53", "M54"],
            ["7", "AL-7", "M64", "M63", "M53", "M54"],
            ["8", "AL-8", "M64", "M63", "M54", "M53"],
            ["9", "AL-9", "M63", "M64", "M53", "M54"],
            ["10", "AL-10", "M64", "M63", "M54", "M53"],
            ["13", "HA-S3", "M56", "M55", "M54", "M53"],
            ["14", "HA-S4", "M56", "M55", "M54", "M53"],
        ],
        "s9_info": "AL = Auto Loader | BR = Boring bar<br/>Para magpalit ng machine: baguhin lang ang #130. Awtomatiko ang M-code mapping.",
        "s9_warn": "Iba-iba ang BR UP/DOWN sa bawat machine. Ang AL-5, AL-7, AL-9 ay baligtad.<br/>Laging #130 lang ang baguhin. Huwag direktang i-edit ang M-code.",
        "s10_title": "10. Paalala sa Operator",
        "s10_1_title": "10-1. Kinakailangang Suriin",
        "s10_1_data": [
            ["No.", "Suriin", "Detalye"],
            ["1", "Lugar na maaaring i-edit", "#101~#123 lamang.\nHUWAG i-edit pagkatapos ng DANGER section"],
            ["2", "#119 check (mahabang parte)", "Kung #113 > 40mm,\nitakda #119=1"],
            ["3", "#102 input format", "0,100,200,300,400,500 lamang\nIba → Alarm 5"],
            ["4", "#101 input range", "0~99 lamang\n100+ → Alarm 6"],
            ["5", "Pagpalit ng machine", "#130 lang ang baguhin\nHuwag direktang i-edit ang M-code"],
        ],
        "s10_2_title": "10-2. Mga Panuntunan sa Kaligtasan",
        "s10_2_info": "(1) HUWAG manghimasok sa panahon ng Auto Link — HUWAG hawakan ang materyal/tool habang bukas/sarado ang chuck!<br/>(2) RS40 material — 15mm safety offset awtomatikong naka-apply, HUWAG i-disable<br/>(3) Maliit na OD (mas mababa sa 30mm) — Parehong safety offset ang naka-apply<br/>(4) Sa alarm — Tingnan ang alarm number → Sumangguni sa code table → Ayusin ang variable → I-restart<br/>(5) G04 wait time — HUWAG bawasan ang Auto Link wait time",
        "s10_3_title": "10-3. Pre-Machining Checklist",
        "s10_3_items": [
            "Uri ng materyal (#105) tugma sa aktwal na materyal?",
            "Raw OD/ID (#109, #110) tugma sa sinukat na halaga?",
            "Finished dimension (#111~#118) tugma sa drawing?",
            "Machine number (#130) tugma sa aktwal na machine?",
            "RPM (#120, #121) angkop sa materyal/tool?",
            "Unang run sa single mode (#108=1) para sa test?",
            "Na-verify na ang Auto Loader operation?",
            "Na-check na ang cut-off insert (T02) wear?",
        ],
        "summary_title": "O0852 Program Pangunahing Buod",
        "summary_items": [
            "Ang operator ay nagbabago lamang ng variable #101~#123 (materyal, sukat, kundisyon)",
            "Awtomatikong naka-set ang feed rate ayon sa uri ng materyal (CN/RS/CM)",
            "Face → Step → Chamfer → Cut-off cycle na awtomatikong umuulit",
            "Awtomatikong hinahatak ng Auto Link ang materyal kapag maikli na",
            "30+ alarm ang pumipigil sa input error at safety issue",
            "9 na machine M-code awtomatikong naka-map (baguhin lang ang #130)",
        ],
        "summary_footer": "Kung may tanong, magtanong anumang oras",
    },
}


def build_pdf(lang_code, L):
    """Build a PDF for the given language"""
    FN = L["font"]
    FB = L["fontBd"]

    sty = {
        'title': ParagraphStyle('title', fontName=FB, fontSize=30, leading=40, textColor=NAVY, spaceAfter=5*mm),
        'h2': ParagraphStyle('h2', fontName=FB, fontSize=19, leading=26, textColor=DARK_BLUE, spaceBefore=7*mm, spaceAfter=4*mm),
        'h3': ParagraphStyle('h3', fontName=FB, fontSize=16, leading=22, textColor=ACCENT, spaceBefore=5*mm, spaceAfter=3*mm),
        'body': ParagraphStyle('body', fontName=FN, fontSize=14, leading=21, textColor=black, spaceAfter=3*mm),
        'body_bold': ParagraphStyle('body_bold', fontName=FB, fontSize=14, leading=21, textColor=black, spaceAfter=3*mm),
        'warn': ParagraphStyle('warn', fontName=FB, fontSize=14, leading=20, textColor=RED, spaceAfter=3*mm, leftIndent=5*mm),
        'small': ParagraphStyle('small', fontName=FN, fontSize=12, leading=17, textColor=GRAY, spaceAfter=2*mm),
        'center': ParagraphStyle('center', fontName=FB, fontSize=15, leading=22, textColor=DARK_BLUE, alignment=TA_CENTER, spaceAfter=4*mm),
        'toc': ParagraphStyle('toc', fontName=FN, fontSize=16, leading=26, textColor=black, leftIndent=8*mm, spaceAfter=2*mm),
    }

    def make_table(data, col_widths=None, header_color=NAVY):
        # Wrap all cells in Paragraph for auto word-wrap
        cell_sty_h = ParagraphStyle('tcH', fontName=FB, fontSize=12, leading=17,
                                     textColor=white, alignment=TA_CENTER)
        cell_sty_b = ParagraphStyle('tcB', fontName=FN, fontSize=12, leading=17,
                                     textColor=black, alignment=TA_CENTER)
        wrapped = []
        for ri, row in enumerate(data):
            new_row = []
            for cell in row:
                txt = str(cell).replace('\n', '<br/>')
                if ri == 0:
                    new_row.append(Paragraph(txt, cell_sty_h))
                else:
                    new_row.append(Paragraph(txt, cell_sty_b))
            wrapped.append(new_row)
        t = Table(wrapped, colWidths=col_widths, repeatRows=1)
        cmds = [
            ('BACKGROUND', (0,0), (-1,0), header_color),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.7, HexColor('#CBD5E1')),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ]
        for i in range(1, len(wrapped)):
            bg = LIGHT_BG if i % 2 == 0 else white
            cmds.append(('BACKGROUND', (0,i), (-1,i), bg))
        t.setStyle(TableStyle(cmds))
        return t

    def info_box(text, bg=BLUE_BG, tc=DARK_BLUE):
        p = Paragraph(text, ParagraphStyle('box', fontName=FB, fontSize=14, leading=21, textColor=tc))
        t = Table([[p]], colWidths=[170*mm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), bg),
            ('ROUNDEDCORNERS', [3,3,3,3]),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 12),
            ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ]))
        return t

    filepath = rf"C:\Users\admin\Desktop\work\CNC_CODES\{L['filename']}"
    doc = SimpleDocTemplate(filepath, pagesize=A4,
        topMargin=20*mm, bottomMargin=20*mm, leftMargin=18*mm, rightMargin=18*mm,
        title=f"O0852 CNC Training - {lang_code.upper()}")

    s = []  # story

    # ── Cover ──
    s.append(Spacer(1, 40*mm))
    s.append(Paragraph(L["cover_title"], ParagraphStyle('ct', fontName=FB, fontSize=38, leading=48, textColor=NAVY, alignment=TA_CENTER)))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["cover_sub"], ParagraphStyle('cs', fontName=FB, fontSize=22, leading=30, textColor=ACCENT, alignment=TA_CENTER)))
    s.append(Spacer(1, 3*mm))
    s.append(HRFlowable(width="60%", thickness=2, color=ACCENT, spaceAfter=8*mm, spaceBefore=5*mm))
    s.append(Paragraph(L["cover_desc"], ParagraphStyle('cd', fontName=FN, fontSize=16, leading=22, textColor=GRAY, alignment=TA_CENTER)))
    s.append(Spacer(1, 20*mm))
    for title, desc in L["cover_features"]:
        s.append(Paragraph(f'<b>{title}</b>  —  {desc}', ParagraphStyle('ci', fontName=FN, fontSize=15, leading=24, textColor=DARK_BLUE, alignment=TA_CENTER)))
    s.append(PageBreak())

    # ── TOC ──
    s.append(Paragraph(L["toc_title"], sty['title']))
    s.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=8*mm))
    for num, title, desc in L["toc_items"]:
        s.append(Paragraph(f'<font color="{ACCENT}">{num}</font>  <b>{title}</b>  <font color="{GRAY}">— {desc}</font>', sty['toc']))
    s.append(PageBreak())

    # Helper
    def section_header(title):
        s.append(Paragraph(title, sty['title']))
        s.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=5*mm))

    # ── S1 ──
    section_header(L["s1_title"])
    s.append(info_box(L["s1_desc"]))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s1_summary_title"], sty['h3']))
    s.append(make_table(L["s1_summary"], col_widths=[25*mm, 50*mm, 50*mm, 49*mm]))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s1_feat_title"], sty['h3']))
    for feat in L["s1_features"]:
        s.append(Paragraph(f'<font color="{ACCENT}"><b>{feat[0]}</b></font>: {feat[1]}', sty['body']))
    s.append(PageBreak())

    # ── S2 ──
    section_header(L["s2_title"])
    s.append(Paragraph(L["s2_desc"], sty['body']))
    s.append(make_table(L["s2_table"], col_widths=[22*mm, 30*mm, 70*mm, 32*mm]))
    s.append(Paragraph(L["s2_flow_title"], sty['h3']))
    s.append(Paragraph(L["s2_flow"], sty['center']))
    s.append(Paragraph(L["s2_tool_title"], sty['h3']))
    s.append(make_table(L["s2_tools"], col_widths=[20*mm, 45*mm, 50*mm, 45*mm]))
    s.append(PageBreak())

    # ── S3 ──
    section_header(L["s3_title"])
    s.append(info_box(L["s3_info"], GREEN_BG, GREEN))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s3_1_title"], sty['h2']))
    s.append(make_table(L["s3_1_data"], col_widths=[22*mm, 42*mm, 22*mm, 70*mm]))
    s.append(Spacer(1, 3*mm))
    s.append(info_box(L["s3_1_note"], YELLOW_BG, black))
    s.append(Paragraph(L["s3_2_title"], sty['h2']))
    s.append(make_table(L["s3_2_data"], col_widths=[22*mm, 55*mm, 22*mm, 55*mm]))
    s.append(Paragraph(L["s3_3_title"], sty['h2']))
    s.append(make_table(L["s3_3_data"], col_widths=[22*mm, 45*mm, 22*mm, 65*mm]))
    # S4 flows after S3-3
    section_header(L["s4_title"])
    s.append(Paragraph(L["s4_desc"], sty['body']))
    s.append(make_table(L["s4_data"], col_widths=[34*mm, 35*mm, 35*mm, 35*mm, 35*mm]))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s4_formula_title"], sty['h3']))
    s.append(info_box(L["s4_formula"], BLUE_BG, ACCENT))
    s.append(Paragraph(L["s4_example_title"], sty['h3']))
    s.append(make_table(L["s4_examples"], col_widths=[35*mm, 30*mm, 50*mm, 40*mm], header_color=TEAL))
    s.append(Paragraph(L["s4_note"], sty['small']))
    s.append(PageBreak())

    # ── S5 ──
    section_header(L["s5_title"])
    s.append(info_box(L["s5_warn"], RED_BG, RED))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s5_1_title"], sty['h2']))
    s.append(make_table(L["s5_1_data"], col_widths=[22*mm, 40*mm, 20*mm, 72*mm]))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s5_2_title"], sty['h2']))
    s.append(info_box(L["s5_2_info"], YELLOW_BG, black))
    s.append(PageBreak())

    # ── S6 ──
    section_header(L["s6_title"])
    s.append(Paragraph(L["s6_1_title"], sty['h2']))
    s.append(make_table(L["s6_1_data"], col_widths=[15*mm, 30*mm, 17*mm, 92*mm], header_color=DARK_BLUE))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s6_calc_title"], sty['h3']))
    s.append(make_table(L["s6_calc_data"], col_widths=[50*mm, 70*mm, 34*mm], header_color=TEAL))
    s.append(PageBreak())

    # ── S7 ──
    section_header(L["s7_title"])
    s.append(Paragraph(L["s7_1_title"], sty['h2']))
    s.append(Paragraph(L["s7_1_desc"], sty['body']))
    s.append(make_table(L["s7_1_data"], col_widths=[15*mm, 35*mm, 33*mm, 71*mm], header_color=ORANGE))
    s.append(Spacer(1, 5*mm))
    s.append(info_box(L["s7_warn"], RED_BG, RED))
    s.append(PageBreak())

    # ── S8 ──
    section_header(L["s8_title"])
    s.append(Paragraph(L["s8_desc"], sty['body']))
    s.append(Paragraph(L["s8_1_title"], sty['h2']))
    s.append(make_table(L["s8_1_data"], col_widths=[17*mm, 35*mm, 48*mm, 54*mm], header_color=RED))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s8_2_title"], sty['h2']))
    s.append(make_table(L["s8_2_data"], col_widths=[17*mm, 35*mm, 48*mm, 54*mm], header_color=ORANGE))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s8_3_title"], sty['h2']))
    s.append(info_box(L["s8_3_info"], YELLOW_BG, black))
    s.append(PageBreak())

    # ── S9 + S10 (merged) ──
    section_header(L["s9_title"])
    s.append(Paragraph(L["s9_desc"], sty['body']))
    s.append(make_table(L["s9_data"], col_widths=[18*mm, 22*mm, 28*mm, 28*mm, 28*mm, 30*mm]))
    s.append(Spacer(1, 5*mm))
    s.append(info_box(L["s9_info"], BLUE_BG, DARK_BLUE))
    s.append(Paragraph(L["s9_warn"], sty['warn']))

    section_header(L["s10_title"])
    s.append(Paragraph(L["s10_1_title"], sty['h2']))
    s.append(make_table(L["s10_1_data"], col_widths=[15*mm, 40*mm, 99*mm]))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s10_2_title"], sty['h2']))
    s.append(info_box(L["s10_2_info"], RED_BG, RED))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(L["s10_3_title"], sty['h2']))

    cl_data = [["", L["s10_3_items"][0].split("?")[0] + "?" if "?" in L["s10_3_items"][0] else L["s10_3_items"][0]]]
    cl_data = [["", "Check"]]  # header placeholder
    cl_data = []
    header_text = L["s10_3_items"][0].split("(")[0].strip() if "(" in L["s10_3_items"][0] else "Check"
    cl_data.append(["", header_text[:4] == "" and "Check" or "Check"])

    # Simpler: just build checklist with Paragraph wrapping
    cl_h_sty = ParagraphStyle('clH', fontName=FB, fontSize=13, leading=20, textColor=white, alignment=TA_CENTER)
    cl_c_sty = ParagraphStyle('clC', fontName=FN, fontSize=13, leading=20, textColor=black, alignment=TA_CENTER)
    cl_i_sty = ParagraphStyle('clI', fontName=FN, fontSize=13, leading=20, textColor=black, alignment=TA_LEFT)
    checklist = [[Paragraph("", cl_h_sty), Paragraph("Check Item", cl_h_sty)]]
    for item in L["s10_3_items"]:
        checklist.append([Paragraph("□", cl_c_sty), Paragraph(item, cl_i_sty)])
    t = Table(checklist, colWidths=[14*mm, 140*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), GREEN),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [white, GREEN_BG]),
    ]))
    s.append(t)

    # ── Summary ──
    s.append(Spacer(1, 10*mm))
    s.append(Paragraph(L["summary_title"], sty['title']))
    s.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=8*mm))
    for i, item in enumerate(L["summary_items"]):
        s.append(Paragraph(
            f'<font color="{ACCENT}"><b>{i+1}.</b></font>  {item}',
            ParagraphStyle('sum', fontName=FN, fontSize=16, leading=28, textColor=NAVY, leftIndent=5*mm, spaceAfter=4*mm)
        ))
    s.append(Spacer(1, 15*mm))
    s.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceAfter=5*mm, spaceBefore=5*mm))
    s.append(Paragraph(L["summary_footer"], ParagraphStyle('ft', fontName=FN, fontSize=15, textColor=GRAY, alignment=TA_CENTER)))

    doc.build(s)
    print(f"  [{lang_code.upper()}] Saved: {filepath}")


# ═══════════════════════════════════════════
# 생성
# ═══════════════════════════════════════════
print("Generating multilingual PDFs...")
for code, data in LANGS.items():
    build_pdf(code, data)
print("All done!")

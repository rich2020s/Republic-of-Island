from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUT = Path(r"C:\Users\richc\Documents\Codex\2026-10-02\4\outputs\島嶼共和_原型規則_v0.1.docx")
OUT.parent.mkdir(parents=True, exist_ok=True)

FONT = "Microsoft JhengHei"
NAVY = "24445F"
PALE_BLUE = "EDF4F8"
LIGHT_GRAY = "F4F5F6"
BORDER = "D9D9D9"
BLACK = RGBColor(0, 0, 0)


def set_run_font(run, size=None, bold=None, color=BLACK):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    run.font.color.rgb = color


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=110, start=125, bottom=110, end=125):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "4")
        tag.set(qn("w:color"), BORDER)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_row_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    set_run_font(run, 9, color=RGBColor(90, 90, 90))
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


doc = Document()
section = doc.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
section.top_margin = Inches(0.72)
section.bottom_margin = Inches(0.68)
section.left_margin = Inches(0.82)
section.right_margin = Inches(0.82)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = FONT
normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
normal.font.size = Pt(11.25)
normal.font.color.rgb = BLACK
normal.paragraph_format.space_after = Pt(5)
normal.paragraph_format.line_spacing = 1.35

for style_name, size, before, after in (
    ("Title", 26, 0, 14),
    ("Heading 1", 17, 16, 7),
    ("Heading 2", 13.5, 11, 5),
    ("Heading 3", 11.5, 8, 3),
):
    st = styles[style_name]
    st.font.name = FONT
    st._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = BLACK
    st.paragraph_format.space_before = Pt(before)
    st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True

# Word's built-in Title style can retain a theme border. Remove it explicitly.
title_ppr = styles["Title"]._element.get_or_add_pPr()
title_border = title_ppr.find(qn("w:pBdr"))
if title_border is not None:
    title_ppr.remove(title_border)

for list_style in ("List Bullet", "List Number"):
    st = styles[list_style]
    st.font.name = FONT
    st._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(11.25)
    st.paragraph_format.space_after = Pt(2.5)
    st.paragraph_format.line_spacing = 1.25


def add_heading(text, level=1):
    p = doc.add_heading(text, level=level)
    keep_with_next(p)
    return p


def add_para(text="", bold_lead=None, align=None, keep=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        set_run_font(r1, bold=True)
        r2 = p.add_run(text[len(bold_lead):])
        set_run_font(r2)
    else:
        r = p.add_run(text)
        set_run_font(r)
    if keep:
        keep_with_next(p)
    return p


def add_bullets(items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(item)
        set_run_font(r)


def add_numbered(items):
    for index, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.28)
        p.paragraph_format.first_line_indent = Inches(-0.22)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(f"{index}. {item}")
        set_run_font(r)


def add_table(headers, rows, widths=None, font_size=10.4):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    header = table.rows[0]
    set_row_cant_split(header)
    set_repeat_table_header(header)
    for i, text in enumerate(headers):
        cell = header.cells[i]
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths:
            cell.width = Inches(widths[i])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        set_run_font(r, font_size, True, RGBColor(255, 255, 255))
    for r_index, values in enumerate(rows):
        row = table.add_row()
        set_row_cant_split(row)
        for i, text in enumerate(values):
            cell = row.cells[i]
            if r_index % 2 == 1:
                set_cell_shading(cell, PALE_BLUE)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths:
                cell.width = Inches(widths[i])
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.18
            if len(str(text)) < 12 and i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(text))
            set_run_font(r, font_size)
    for row_index, row in enumerate(table.rows):
        keep = row_index < len(table.rows) - 1
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = keep
    after = doc.add_paragraph()
    after.paragraph_format.space_after = Pt(1)
    return table


# Cover
p = doc.add_paragraph(style="Title")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(108)
r = p.add_run("島嶼共和原型規則")
set_run_font(r, 26, True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(30)
r = p.add_run("現代民主政治與利益分配桌上遊戲規則草案")
set_run_font(r, 13, False, RGBColor(65, 65, 65))

add_table(
    ["項目", "設定"],
    [
        ["版本", "原型 v0.1"],
        ["人數", "3 至 5 人，建議 4 人"],
        ["時間", "150 至 240 分鐘"],
        ["長度", "6 回合，每 2 回合舉行選舉"],
        ["類型", "半合作談判 政治經濟 分配治理"],
    ],
    widths=[1.35, 4.7],
    font_size=11,
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(32)
r = p.add_run("設計主軸")
set_run_font(r, 12, True)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.left_indent = Inches(0.55)
p.paragraph_format.right_indent = Inches(0.55)
r = p.add_run("共同維持國家運作，同時為自己的支持聯盟爭取利益與功勞")
set_run_font(r, 12)

doc.add_page_break()

# Introduction and quick flow
add_heading("規則書用途", 1)
add_para("本規則書提供一套可直接製作紙上原型並進行第一次測試的完整流程。遊戲以提案 談判 修正 表決 執行和選舉為核心，重點不是精確預測現實，而是讓玩家清楚看見每項政策由誰付錢 誰受益 誰取得功勞，以及分配結果如何改變下一輪政治力量。")
add_para("遊戲借用共和政治中共同治理與派系競爭並存的結構，但不逐條重現任何現實政黨或法律制度。所有數值都是原型參數，應在測試後調整。")

add_heading("一回合速覽", 1)
add_table(
    ["階段", "主要行動", "核心結果"],
    [
        ["1 事件", "揭露國內外事件", "改變成本 難度或政治壓力"],
        ["2 收入", "取得稅收並支付利息與基本支出", "更新國庫與財政危機"],
        ["3 議程", "政府決定總預算與政策順序", "建立本回合談判範圍"],
        ["4 總預算", "分配國庫並進行修正與表決", "決定公共資源流向"],
        ["5 政策立法", "提出至多兩項法案", "決定受益者 負擔者和功勞"],
        ["6 執行", "政治人物進行政策檢定", "政策成功 部分成功或失敗"],
        ["7 監督", "調查交易與處理醜聞", "更新曝光 聲望與職位"],
        ["8 民意", "結算集團利益與支持", "更新信任和分配壓力"],
        ["9 選舉", "偶數回合重新分配席次", "形成新政府"],
        ["10 檢查", "檢查危機 共同失敗和終局", "進入下一回合或結束"],
    ],
    widths=[1.0, 2.85, 2.85],
    font_size=9.8,
)

add_heading("第一部分 遊戲目標與組成", 1)
add_heading("1 遊戲目標", 2)
add_para("玩家扮演政治網絡而不是固定政黨。每個網絡控制政治人物 立法席次 社會支持 政治資金及政策目標。玩家要共同維持國家運作，同時取得政策功勞 安排職位 擴張支持聯盟並避免承擔過多政治責任。")
add_para("若國家崩潰，所有玩家共同失敗。若國家維持到第六回合，政治聲望最高的玩家獲勝。")

add_heading("2 建議遊戲配件", 2)
add_bullets([
    "一張國家狀態板，包含國庫 生產能力 行政能力 社會信任 分配壓力與國家債務。",
    "六張社會集團板，每張可記錄所有玩家的支持度。",
    "五張政治網絡板及至少十五張政治人物牌。",
    "政策牌四十張，事件牌三十張，派系目標牌二十張。",
    "利益 負擔 功勞 派系利益 曝光 危機及職位籌碼。",
    "五種玩家顏色的支持標記 席次標記及投票標記。",
    "六面骰一顆，以及公開協議與秘密交易紀錄紙。",
])

add_heading("3 遊戲設置", 2)
add_numbered([
    "將六條國家軌道設定為國庫 10 生產能力 5 行政能力 5 社會信任 6 分配壓力 4 國家債務 2。",
    "洗勻事件牌 政策牌 政治人物牌和派系目標牌，分別放置。",
    "每位玩家選擇一個政治網絡，取得三名政治人物 五點影響力 五點政治資金和一張秘密派系目標。",
    "依政治網絡板設定六個社會集團的初始支持。每位玩家的初始支持總和建議為 12。",
    "依初始支持進行一次席次分配。每位玩家另取得一個基本席次。",
    "席次最多者取得首次組閣權。完成政府職位分配後開始第一回合。",
])

add_heading("第二部分 國家與社會", 1)
add_heading("4 國家狀態", 2)
add_table(
    ["軌道", "範圍與初值", "遊戲意義"],
    [
        ["國庫", "0 至 20 初值 10", "支付政策 預算 基本行政與危機成本。降到 0 不會立即失敗，但可能無法履行義務。"],
        ["生產能力", "0 至 10 初值 5", "決定每回合稅收，代表產業 技術 勞動力與基礎建設。"],
        ["行政能力", "0 至 10 初值 5", "提供政策執行加成，反映政府是否能把預算轉成結果。"],
        ["社會信任", "0 至 10 初值 6", "代表制度正當性。預算失敗 醜聞與承諾落空會使其下降。"],
        ["分配壓力", "0 至 10 初值 4", "代表貧富 世代 城鄉 住房及產業衝突。數值越高，危機越容易發生。"],
        ["國家債務", "0 至 10 初值 2", "舉債能提供短期資金，但會增加利息與共同失敗風險。"],
    ],
    widths=[1.05, 1.35, 4.3],
    font_size=9.7,
)

add_heading("5 社會集團", 2)
add_para("遊戲使用六個互相重疊的利益領域。它們不是完全互斥的人口分類，而是政治動員與政策分配的主要對象。每位玩家在每個集團都有 0 至 5 的支持度。")
add_table(
    ["社會集團", "主要關切"],
    [
        ["科技與出口產業", "國際市場 能源 研發 人才 匯率及產業補貼"],
        ["中小企業及自營業者", "稅負 融資 租金 工資 法規及內需"],
        ["受薪階級及公共部門", "薪資 工時 年金 公共服務及就業穩定"],
        ["房產 土地及屋主", "房價 稅制 都市計畫 開發與財產權"],
        ["青年 租屋者及照顧家庭", "住房 教育 育兒 長照 薪資及世代公平"],
        ["地方 農漁及基層網絡", "地方建設 補助 交通 農漁政策與人際服務"],
    ],
    widths=[2.0, 4.7],
)

add_heading("第三部分 玩家與政治人物", 1)
add_heading("6 玩家資源", 2)
add_table(
    ["資源", "用途"],
    [
        ["政治影響力", "提出修正案 爭取議程 發動調查 強化執行或影響程序。"],
        ["政治資金", "選舉活動 建立組織 處理醜聞及履行私人交易。"],
        ["曝光籌碼", "代表可疑交易與利益衝突。曝光越多，調查越容易成功。"],
        ["政治聲望", "主要勝利分數，來自政策 職位 選舉與危機處理。"],
        ["立法席次", "每席在正式表決時等於一票。"],
    ],
    widths=[1.45, 5.25],
)

add_heading("7 政治人物", 2)
add_para("每位玩家初始控制三名政治人物。每名政治人物具有能力 人脈 清廉與專業標籤。能力用於執行政策，人脈用於取得資金與地方利益，清廉用於抵抗調查。若政治人物執行符合專業標籤的政策，檢定獲得加一。")
add_table(
    ["屬性", "範圍", "作用"],
    [
        ["能力", "0 至 3", "政策執行與危機處理"],
        ["人脈", "0 至 3", "派系資源 募款與地方動員"],
        ["清廉", "0 至 3", "降低調查與醜聞成功率"],
        ["專業", "一至兩項", "符合政策領域時提供加成"],
    ],
    widths=[1.25, 1.2, 4.25],
)

add_heading("第四部分 核心分配機制", 1)
add_heading("8 利益籌碼", 2)
add_table(
    ["籌碼", "放置位置", "效果"],
    [
        ["公共投資", "國家軌道或政策牌", "改善生產 行政 韌性或長期財政，通常不立即增加派系支持。"],
        ["集團利益", "社會集團板", "在民意階段與負擔相抵，並透過政策功勞轉為支持。"],
        ["負擔", "社會集團板", "代表稅負 管制 漲價或福利削減，可能使贊成法案的派系失去支持。"],
        ["政策功勞", "玩家與社會集團之間", "政策成功時讓指定玩家取得支持，失敗時可能轉成責任。"],
        ["派系利益", "玩家板", "可轉換為兩點政治資金，或一點影響力加一點資金，同時產生曝光。"],
    ],
    widths=[1.25, 1.55, 3.9],
    font_size=9.8,
)

add_heading("9 一項法案的六個部分", 2)
add_numbered([
    "政策主體。選擇政策牌，確定基本成本 執行難度 生效時間與允許的分配方向。",
    "財源。使用國庫 加稅 舉債或刪減其他預算支付全部成本。",
    "受益與負擔。把集團利益 公共投資及負擔籌碼放在法案板上。",
    "執行者。指定一名具有相關職位或專業的政治人物。",
    "政治功勞。決定成功後哪些玩家會被哪些社會集團視為政策推動者。",
    "派系利益。若政策允許，加入尋租或人事條款，換取私人資源但增加曝光與執行難度。",
])

add_heading("10 財源規則", 2)
add_table(
    ["財源", "立即效果", "限制與代價"],
    [
        ["國庫支出", "國庫減少一點，支付一點成本", "不能使國庫低於零"],
        ["加稅或收費", "指定集團承擔一枚負擔，國庫增加一點", "同一集團在一項法案最多承擔兩枚"],
        ["舉債", "債務增加一點，國庫增加兩點", "每項法案最多使用一次"],
        ["刪減預算", "移除另一領域一枚預算，支付一點成本", "受影響集團承擔一枚負擔"],
    ],
    widths=[1.35, 2.65, 2.7],
)

add_heading("11 法案提案流程", 2)
add_numbered([
    "提出政策。擁有議程權的玩家選擇政策牌，宣告希望處理的問題。",
    "填入財源。提案者決定國庫 加稅 舉債及刪減支出，資金不足的法案不能進入表決。",
    "分配利益。依政策牌規定放置公共投資 集團利益與負擔。不能把政策效果分配到不相關領域。",
    "指定執行者。被指定政治人物的玩家可以拒絕，拒絕後提案者必須改派他人。",
    "分配功勞。把政策牌提供的功勞籌碼配置給參與聯盟。功勞的對象和社會集團必須明確。",
    "加入派系利益。使用尋租格換取私人資源，每格令受益玩家取得一枚曝光，並使執行檢定減一。",
    "公開談判。玩家可以要求調整利益 負擔 執行者 功勞 財源或派系條款。建議限時五分鐘。",
    "提出修正。玩家花一點影響力提出正式修正；提案者可接受，否則進行修正案表決。每案最多接受三項修正。",
    "最終表決。贊成票超過全部席次的一半即通過。棄權不算贊成，所有投票公開記錄。",
])

add_heading("12 修正案規則", 2)
add_para("一項修正案只能改變一個明確項目，例如增加或移動一枚利益 改變一枚負擔 更換執行者 改變一枚功勞，或增加與移除一項派系利益。修正後若增加成本，提案者必須同時提出財源。")
add_para("提案者若拒絕修正，提出者可以要求表決。修正案獲得過半席次支持即自動加入法案，即使原提案者反對。原提案者可以在最終表決前撤回整項法案，但失去一點影響力。")

add_heading("13 公開協議與秘密協議", 2)
add_para("公開協議寫在協議卡上並由參與玩家簽署，可約定未來表決 職位 政策功勞 政治資金或調查行動。違反者失去兩點聲望 兩點影響力，並在自己支持最高的社會集團失去一點支持。")
add_para("秘密協議可以包含任何合法行動，但規則不強制履行。背叛沒有直接處罰。若交易涉及派系利益，調查成功時可能轉成額外曝光。")

add_heading("第五部分 政策執行與民意", 1)
add_heading("14 政策執行檢定", 2)
add_para("法案通過不代表政策一定成功。執行時擲一顆六面骰，使用下式判定：")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(5)
p.paragraph_format.space_after = Pt(7)
r = p.add_run("骰子 + 執行者能力 + 專業加成 + 行政能力加成 - 尋租懲罰")
set_run_font(r, 12, True)
add_para("將結果與政策牌的執行難度比較。行政能力 0 至 3 不加成，4 至 6 加一，7 至 8 加二，9 至 10 加三。符合專業標籤再加一。每個尋租格減一。")
add_table(
    ["結果", "條件", "效果"],
    [
        ["重大成功", "高於難度至少 2", "所有效果生效，額外產生一枚功勞，執行者獲得一點聲望。"],
        ["成功", "達到難度", "所有效果正常生效。"],
        ["部分成功", "低於難度 1 至 2", "執行一半利益，或追加一點國庫以完成全部利益。"],
        ["失敗", "低於難度至少 3", "成本與負擔保留，利益取消，負責者失去一點聲望，社會信任減一。"],
    ],
    widths=[1.2, 1.65, 3.85],
    font_size=9.8,
)

add_heading("15 利益轉換成支持", 2)
add_para("每項政策完成後，逐一計算六個社會集團的淨利益。")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("淨利益 = 集團利益籌碼 - 負擔籌碼")
set_run_font(r, 12, True)
add_bullets([
    "淨利益為正：每一點淨利益最多啟動一枚對應功勞。持有功勞的玩家在該集團增加一點支持。",
    "利益多於功勞：多出的利益不屬於特定派系。若全案至少有兩點未被派系占有的公共成果，社會信任增加一點。",
    "淨利益為零：支持度不變。",
    "淨利益為負：提案者在該集團失去一點支持。若淨利益為負二或更低，所有投贊成票的玩家各失去一點支持。",
    "反對派動員：投反對票者可以花一點影響力，在受損集團增加一點支持。",
    "每位玩家因同一法案在同一集團最多增加或失去兩點支持。",
])

add_heading("16 分配壓力", 2)
add_para("回合結束時檢查利益分布。每符合一項條件，分配壓力調整一點；每回合最多增加或降低兩點。")
add_table(
    ["方向", "條件"],
    [
        ["增加", "某集團取得至少四點利益，同時另有集團完全沒有利益。"],
        ["增加", "至少三個集團承擔負擔，但只有一個集團得到利益。"],
        ["增加", "同一集團連續兩回合沒有得到任何利益。"],
        ["增加", "本回合產生至少三枚派系利益。"],
        ["降低", "至少五個集團本回合得到利益。"],
        ["降低", "一項具有再分配標籤的政策成功。"],
        ["降低", "玩家把兩枚可分配功勞主動轉為公共功勞。"],
    ],
    widths=[1.2, 5.5],
)

add_heading("第六部分 每回合流程", 1)
add_heading("17 事件階段", 2)
add_para("翻開一張事件牌並立即處理。事件可以改變政策成本 執行難度 社會反應或危機條件。事件類型包含景氣 能源 天災 疫情 房市 國安 社會運動 事故與外交壓力。未解決事件可能留下危機標記。")

add_heading("18 收入與債務階段", 2)
add_para("政府取得本回合收入：2 加上生產能力的一半，向下取整。收入加入國庫後，先支付兩點基本行政支出，再支付債務利息。")
add_table(
    ["債務", "每回合利息"],
    [["0 至 2", "0"], ["3 至 5", "1"], ["6 至 8", "2"], ["9 至 10", "3"]],
    widths=[2.3, 2.3],
)
add_para("若政府無法支付基本支出或利息，社會信任與行政能力各下降一點，並增加一枚財政危機標記。")

add_heading("19 政府與議程階段", 2)
add_para("政府安排本回合一項強制總預算案 最多兩項政策法案，以及最多一項人事 調查或不信任提案。政府首腦決定第一項政策議程，議會主持者決定第二項。若職位由同一玩家控制，席次第二多的反對派決定第二項。")

add_heading("20 總預算階段", 2)
add_para("總預算案每回合必須提出並表決。預算至少投入三枚國庫，涵蓋至少三個政策領域，並支付事件要求的緊急支出。五個領域為社會福利 產業與經濟 地方建設 國防與安全，以及教育與行政。")
add_para("預算案同樣經過財源 分配 談判 修正與表決。第一次失敗後可以重新提出一次。第二次仍未通過時啟動看守預算：國庫支付兩點 本回合不能產生新的集團利益 社會信任減一，所有政府政策執行檢定減一。")

add_heading("21 政策立法與執行階段", 2)
add_para("依議程順序處理最多兩項政策法案。所有法案完成表決後，再依通過順序進行執行檢定。延遲生效的政策把牌留在桌上，於牌面指定的未來回合結算。")

add_heading("22 監督階段", 2)
add_para("每位玩家每回合最多發動一次調查。發動者支付兩點影響力，或使用議會監督職位與事件效果。調查目標可以是玩家 政治人物或特定法案。")
add_para("調查時擲一顆六面骰，將目標曝光籌碼加入結果，再減去涉案政治人物的清廉。結果達到 6，醜聞成立。")
add_table(
    ["調查結果", "效果"],
    [
        ["達到 6", "目標失去兩點資金與一點聲望，在兩個相關集團各失去一點支持，涉案人物辭去職位。"],
        ["達到 9", "除一般效果外，涉案政治人物直接離場。"],
        ["未達 6", "發動者額外失去一點影響力。"],
        ["自然骰為 1", "視為政治操作失敗，發動者再失去一點聲望。"],
    ],
    widths=[1.5, 5.2],
    font_size=9.8,
)

add_heading("23 民意階段", 2)
add_numbered([
    "依每項已執行政策計算六個集團的淨利益。",
    "啟動政策功勞，調整玩家支持度。",
    "處理負擔與反對派動員。",
    "檢查公共成果，調整社會信任。",
    "依整體利益分布調整分配壓力。",
    "清除本回合利益與負擔籌碼，保留支持 曝光 聲望及長期政策。",
])

add_heading("24 回合結束", 2)
add_para("在第二 第四與第六回合進行選舉。其他回合直接進行危機與共同失敗檢查。完成後將回合標記前進一格。")

add_heading("第七部分 選舉與政府", 1)
add_heading("25 選舉席次", 2)
add_para("每個社會集團提供三個民意席次。支持最高的玩家取得兩席，第二高者取得一席。每位玩家另保有一個基本席次，以確保仍能參與正式表決。")
add_para("平手時依序比較本屆在該集團取得的政策功勞 投入的競選資金及政治人物能力。仍然平手則各擲一顆骰決定。完成六個集團後，重新計算議會總席次與過半門檻。")

add_heading("26 組成政府", 2)
add_para("席次最多的玩家首先取得組閣權，提出完整政府名單並進行信任表決。過半席次支持即成立。第一次失敗後，組閣權交給席次第二高者。兩次都失敗時成立看守政府，社會信任減一，下一回合只能提出總預算與一項政策法案。")
add_table(
    ["職位", "權力與責任"],
    [
        ["政府首腦", "決定第一項政策議程，每回合取得一點影響力，重大政策失敗時承擔責任。"],
        ["財政首長", "編列總預算，每回合一次使政策成本減少一點，財政危機時失去聲望。"],
        ["經濟首長", "執行產業 能源與經濟政策時加一，經濟危機時承擔責任。"],
        ["社會及內政首長", "執行福利 住宅與地方政策時加一，每回合可調整一枚地方利益。"],
        ["國安首長", "處理國安與外部危機時加一，成功時取得一點聲望。"],
        ["議會主持者", "決定第二項議程，修正案平票時決定結果，每回合免費提出一次修正。"],
    ],
    widths=[1.55, 5.15],
    font_size=9.8,
)

add_heading("第八部分 危機與勝利", 1)
add_heading("27 共同失敗", 2)
add_para("符合下列任一條件，所有玩家共同失敗：")
add_bullets([
    "社會信任降到 0。",
    "行政能力降到 0。",
    "國家債務達到 10 且政府無法支付利息。",
    "累積三枚未解決危機。",
    "連續兩回合沒有通過總預算。",
    "重大國安事件處理失敗，且事件牌指定共同失敗。",
])

add_heading("28 個人勝利", 2)
add_para("第六回合選舉與危機檢查完成後計算政治聲望。若國家沒有共同失敗，總分最高者獲勝。", keep=True)
add_table(
    ["項目", "分數"],
    [
        ["立法席次", "每席一分"],
        ["派系目標", "依牌面分數"],
        ["主導重大成功政策", "每項兩分"],
        ["成功處理重大危機", "每次兩分"],
        ["完整擔任一屆政府首腦", "兩分"],
        ["剩餘政治資金", "每三點一分"],
        ["終局沒有曝光", "兩分"],
        ["未處理曝光", "每枚扣一分"],
        ["違反公開協議", "每次扣兩分"],
        ["政治人物因醜聞離場", "每人扣兩分"],
    ],
    widths=[3.6, 2.1],
)
add_para("平手時依序比較社會集團支持總和 成功政策數量及未使用政治資金。仍然平手則共同勝利。")

add_heading("第九部分 完整法案範例", 1)
add_heading("29 住宅公平與都市更新法", 2)
add_para("政策牌資料：基本成本三點 執行難度六 政策領域為住宅 地方與社會福利，可分配四枚利益，產生兩枚政策功勞，並具有兩個尋租格。")

add_heading("原始提案", 3)
add_bullets([
    "財源：國庫支付兩點，房產及屋主承擔一枚稅負。",
    "利益：青年及租屋者兩枚，受薪家庭一枚，地方網絡一枚。",
    "執行者：提案者控制的社會及內政首長。",
    "功勞：提案者在青年與受薪階級各取得一枚。",
])

add_heading("談判要求", 3)
add_bullets([
    "地方派系要求一枚地方利益 執行權與一枚地方政策功勞。",
    "產業派系要求增加一枚建築產業利益，並願意支持追加國庫支出。",
    "青年派系要求保留兩枚青年利益，否則投反對票。",
])

add_heading("最終方案", 3)
add_table(
    ["項目", "內容"],
    [
        ["財源", "國庫支付三點，房產集團承擔一枚負擔。"],
        ["青年利益", "青年及租屋者取得兩枚。"],
        ["其他利益", "受薪階級 地方網絡 房產及建築產業各取得一枚。"],
        ["政治功勞", "提案者取得青年功勞，地方派系取得地方功勞。"],
        ["執行者", "地方派系控制的政治人物。"],
        ["派系利益", "地方派系取得一枚承包利益和一枚曝光，執行檢定減一。"],
    ],
    widths=[1.45, 5.25],
)
add_para("此案同時分配物質利益 執行權 私人資源 政策功勞與失敗風險，呈現遊戲的核心決策。")

doc.add_page_break()
add_heading("附錄 原型測試重點", 1)
add_para("第一次測試不必追求完全平衡。觀察玩家行為是否產生下列現象：")
add_bullets([
    "玩家是否真的需要交換利益，而不是只按照固定立場投票。",
    "是否存在兩種以上可行的執政聯盟與政策路線。",
    "公共投資 集團利益 政策功勞及派系利益是否都具有吸引力。",
    "執政者是否因控制議程而受益，同時承擔更高的失敗責任。",
    "反對派是否有合理的得分方法，而不是只能阻撓。",
    "遊戲結束時，玩家是否能清楚說明國家為何繁榮 停滯或崩潰。",
])
add_para("若玩家無法說明政策結果，應先減少籌碼與例外規則。若談判缺乏張力，應增加財政限制 功勞競爭或集團之間的衝突，而不是先增加更多經濟指標。")

# Footer and metadata
for sec in doc.sections:
    add_page_number(sec.footer.paragraphs[0])

doc.core_properties.title = "島嶼共和原型規則"
doc.core_properties.subject = "現代民主政治與利益分配桌上遊戲規則草案"
doc.core_properties.author = ""
doc.core_properties.keywords = "桌上遊戲 政治 經濟 利益分配 規則"

doc.save(OUT)
print("rulebook docx created")

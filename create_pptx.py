"""
MNP×激安不動産投資JVコンサル セミナースライド生成スクリプト
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import os

# ─── カラー定義 ───────────────────────────────────────
BG_COLOR    = RGBColor(0x0f, 0x1b, 0x2d)   # ダークネイビー
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_BLUE = RGBColor(0x4f, 0xc3, 0xf7)   # 水色（タイトル・強調）
ACCENT_GOLD = RGBColor(0xc9, 0xa8, 0x4c)   # ゴールド（金額・数値）
LIGHT_GRAY  = RGBColor(0xCC, 0xCC, 0xCC)
DARK_GRAY   = RGBColor(0x88, 0x88, 0x88)
RED_WARN    = RGBColor(0xFF, 0x55, 0x55)

# ─── フォント ─────────────────────────────────────────
FONT_PRIMARY   = "Noto Sans JP"
FONT_FALLBACK  = "Yu Gothic"
FONT_FALLBACK2 = "MS Gothic"

# ─── スライドサイズ (16:9) ────────────────────────────
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ─── ヘルパー関数 ─────────────────────────────────────

def set_slide_bg(slide, color: RGBColor):
    """スライド背景色を設定"""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                font_size=24, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, font_name=FONT_PRIMARY,
                italic=False, wrap=True):
    """テキストボックスを追加"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    """塗りつぶし矩形を追加"""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_divider(slide, top, color=ACCENT_BLUE):
    """水平区切り線を追加"""
    left   = Inches(0.8)
    width  = Inches(11.73)
    height = Pt(2)
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_slide(prs) -> object:
    """ブランクスライドを追加し背景色を設定して返す"""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide, BG_COLOR)
    return slide


def add_multiline_textbox(slide, lines, left, top, width, height,
                          font_size=20, bold=False, color=WHITE,
                          align=PP_ALIGN.LEFT, font_name=FONT_PRIMARY,
                          line_spacing=Pt(8)):
    """複数行・段落テキストボックスを追加 (lines: list of (text, color, bold, size))"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, col, b, sz = item, color, bold, font_size
        else:
            text = item.get("text", "")
            col  = item.get("color", color)
            b    = item.get("bold", bold)
            sz   = item.get("size", font_size)

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.alignment = align
        p.space_after = line_spacing
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(sz)
        run.font.bold = b
        run.font.color.rgb = col
    return txBox


# ─── スライド生成関数 ──────────────────────────────────

def slide_01_title(prs):
    """スライド1: タイトル"""
    slide = add_slide(prs)

    # 背景アクセント矩形（左端縦ライン）
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    # サブタイトル帯
    add_rect(slide, Inches(0.15), Inches(1.4), SLIDE_W, Inches(0.06), ACCENT_BLUE)

    # 上部ラベル
    add_textbox(slide, "MNP × 激安不動産投資 JVコンサル",
                Inches(0.5), Inches(0.25), Inches(12), Inches(0.7),
                font_size=20, color=ACCENT_BLUE, bold=True)

    # メインタイトル
    add_textbox(slide, "MNP × 激安不動産投資",
                Inches(0.5), Inches(1.55), Inches(12), Inches(1.5),
                font_size=54, bold=True, color=WHITE)

    # サブキャッチ
    add_textbox(slide, "自己資金ゼロから",
                Inches(0.5), Inches(3.1), Inches(12), Inches(0.8),
                font_size=34, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "年利100%超えの不動産オーナーへ",
                Inches(0.5), Inches(3.75), Inches(12), Inches(0.8),
                font_size=34, bold=True, color=ACCENT_GOLD)

    # 登壇者
    add_rect(slide, Inches(0.5), Inches(5.0), Inches(5.5), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide, "富岡信成（税理士）× 池田宜史（MNP歴12年・指導500名超）",
                Inches(0.5), Inches(5.1), Inches(12), Inches(0.6),
                font_size=18, color=LIGHT_GRAY)

    # 対象ライン
    add_textbox(slide, "対象: 不動産投資未経験のサラリーマン・副業で資産形成したい方",
                Inches(0.5), Inches(5.8), Inches(12), Inches(0.6),
                font_size=15, color=DARK_GRAY)
    return slide


def slide_02_agenda(prs):
    """スライド2: 本日の内容"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "本日の内容", Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=36, bold=True, color=ACCENT_BLUE)
    add_divider(slide, Inches(1.1))

    parts = [
        ("PART 1", "MNPで 0 → 1 の資金を作る（池田パート）",
         "携帯乗り換えで数十万円の資金を作る再現性の高い手法"),
        ("PART 2", "激安不動産投資の全貌（富岡パート）",
         "0円〜30万円の激安物件で年利100%超えを狙う方法"),
        ("PART 3", "メンバー実績 & オファー",
         "リアルな成果事例 + 限定コンサルのご案内"),
    ]

    for i, (part, title, desc) in enumerate(parts):
        top = Inches(1.35 + i * 1.7)
        add_rect(slide, Inches(0.5), top, Inches(12.5), Inches(1.5),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, Inches(0.5), top, Inches(0.08), Inches(1.5), ACCENT_GOLD)
        add_textbox(slide, part, Inches(0.7), top + Pt(12), Inches(2), Inches(0.5),
                    font_size=14, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, title, Inches(0.7), top + Inches(0.35), Inches(11.5), Inches(0.55),
                    font_size=22, bold=True, color=WHITE)
        add_textbox(slide, desc, Inches(0.7), top + Inches(0.9), Inches(11.5), Inches(0.5),
                    font_size=14, color=LIGHT_GRAY)
    return slide


def slide_03_mnp_intro(prs):
    """スライド3: PART1 MNP概要"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "PART 1", Inches(0.5), Inches(0.25), Inches(12), Inches(0.5),
                font_size=18, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "MNPで「0→1」の資金を作る",
                Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
                font_size=38, bold=True, color=WHITE)
    add_divider(slide, Inches(1.65))

    add_textbox(slide, "池田宜史 監修",
                Inches(0.5), Inches(1.8), Inches(12), Inches(0.5),
                font_size=16, color=ACCENT_GOLD, bold=True)

    bullets = [
        "MNP（携帯番号ポータビリティ）= キャリア乗り換えで発生するキャッシュバックを活用",
        "12年間・500名超の指導実績がある再現性の高い手法",
        "貯金ゼロでも短期間に数十万円の現金を生み出せる",
        "不動産投資の「最初の元手」を用意するための最速ルート",
    ]
    for i, b in enumerate(bullets):
        top = Inches(2.4 + i * 0.95)
        add_rect(slide, Inches(0.5), top + Inches(0.15), Inches(0.4), Inches(0.4), ACCENT_GOLD)
        add_textbox(slide, "✓", Inches(0.52), top + Inches(0.1), Inches(0.4), Inches(0.45),
                    font_size=18, bold=True, color=BG_COLOR, align=PP_ALIGN.CENTER)
        add_textbox(slide, b, Inches(1.05), top, Inches(11.5), Inches(0.7),
                    font_size=20, color=WHITE)
    return slide


def slide_04_mnp_steps(prs):
    """スライド4: MNP実践ステップ"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "MNP 実践ステップ",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    steps = [
        ("STEP 1", "現在のキャリアを確認", "docomo / au / SoftBank → 乗り換え先候補を選定"),
        ("STEP 2", "キャッシュバック条件を精査", "端末・プラン・オプションを組み合わせ最大化"),
        ("STEP 3", "手続き実行（最短1日）", "店頭またはオンラインで手続き完了"),
        ("STEP 4", "キャッシュバック受取", "数万円〜最大30万円超の現金・ポイントを獲得"),
    ]

    for i, (step, title, desc) in enumerate(steps):
        col = 0 if i < 2 else 1
        row = i % 2
        left = Inches(0.5 + col * 6.5)
        top  = Inches(1.35 + row * 2.8)
        add_rect(slide, left, top, Inches(6.0), Inches(2.4),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(6.0), Inches(0.08), ACCENT_GOLD)
        add_textbox(slide, step, left + Inches(0.2), top + Inches(0.15),
                    Inches(2), Inches(0.4), font_size=13, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, title, left + Inches(0.2), top + Inches(0.55),
                    Inches(5.5), Inches(0.6), font_size=22, bold=True, color=WHITE)
        add_textbox(slide, desc, left + Inches(0.2), top + Inches(1.2),
                    Inches(5.5), Inches(1.0), font_size=16, color=LIGHT_GRAY)
    return slide


def slide_05_mnp_result(prs):
    """スライド5: MNP実績・資金作り"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "MNPで作れる資金の目安",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    # 強調数値
    add_rect(slide, Inches(1.0), Inches(1.4), Inches(11.0), Inches(2.2),
             RGBColor(0x0a, 0x14, 0x24))
    add_textbox(slide, "数十万円〜", Inches(1.5), Inches(1.55),
                Inches(10), Inches(1.0), font_size=64, bold=True, color=ACCENT_GOLD,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, "短期間で現金化できる",
                Inches(1.5), Inches(2.55), Inches(10), Inches(0.5),
                font_size=20, color=WHITE, align=PP_ALIGN.CENTER)

    points = [
        "貯金ゼロでもスタートできる",
        "本業・副業と並行して実施可能",
        "これが不動産投資の「最初の種銭」になる",
        "池田が12年間・500名超に指導した再現性の高い方法",
    ]
    for i, pt in enumerate(points):
        top = Inches(3.85 + i * 0.75)
        add_rect(slide, Inches(0.6), top + Inches(0.17), Pt(10), Pt(10), ACCENT_GOLD)
        add_textbox(slide, f"• {pt}", Inches(0.85), top, Inches(11.5), Inches(0.65),
                    font_size=20, color=WHITE)
    return slide


def slide_06_part2_intro(prs):
    """スライド6: PART2 激安不動産投資の全貌"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "PART 2", Inches(0.5), Inches(0.25), Inches(12), Inches(0.5),
                font_size=18, bold=True, color=ACCENT_BLUE)
    add_textbox(slide, "激安不動産投資の全貌",
                Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
                font_size=40, bold=True, color=WHITE)
    add_divider(slide, Inches(1.65))

    add_textbox(slide, "富岡信成（税理士）監修",
                Inches(0.5), Inches(1.8), Inches(12), Inches(0.5),
                font_size=16, color=ACCENT_BLUE, bold=True)

    # 3本柱
    pillars = [
        ("ゼロ円", "〜30万円以下の", "激安取得"),
        ("利回り", "100%〜436%の", "超高回転"),
        ("税理士", "サポートによる", "キャッシュ最大化"),
    ]
    for i, (num, sub1, sub2) in enumerate(pillars):
        left = Inches(0.5 + i * 4.2)
        top  = Inches(2.5)
        add_rect(slide, left, top, Inches(3.9), Inches(4.3),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(3.9), Inches(0.08), ACCENT_BLUE)
        add_textbox(slide, num, left + Inches(0.2), top + Inches(0.3),
                    Inches(3.5), Inches(0.9), font_size=40, bold=True, color=ACCENT_BLUE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, sub1, left + Inches(0.2), top + Inches(1.3),
                    Inches(3.5), Inches(0.5), font_size=16, color=LIGHT_GRAY,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, sub2, left + Inches(0.2), top + Inches(1.8),
                    Inches(3.5), Inches(0.5), font_size=18, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
    return slide


def slide_07_zero_yen_case(prs):
    """スライド7: 0円物件の実例（青森県5DK）"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "0円物件の実例",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    # 物件詳細カード
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(12.0), Inches(4.8),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(12.0), Inches(0.08), ACCENT_BLUE)

    # 物件情報
    details = [
        ("所在地", "青森県"),
        ("築年数", "35年"),
        ("間取り", "5DK"),
        ("土地", "350平米"),
        ("建物", "110平米"),
    ]
    for i, (k, v) in enumerate(details):
        left = Inches(0.8 + (i % 3) * 4.0)
        top  = Inches(1.55 + (i // 3) * 1.2)
        add_textbox(slide, k, left, top, Inches(1.2), Inches(0.4),
                    font_size=13, color=DARK_GRAY)
        add_textbox(slide, v, left, top + Inches(0.35), Inches(3.5), Inches(0.55),
                    font_size=26, bold=True, color=WHITE)

    add_divider(slide, Inches(3.6))

    # 取得価格 → 売却
    add_textbox(slide, "取得価格", Inches(1.2), Inches(3.8), Inches(3), Inches(0.5),
                font_size=16, color=DARK_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, "0 円", Inches(1.2), Inches(4.2), Inches(3), Inches(0.8),
                font_size=52, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.CENTER)
    add_textbox(slide, "（手続き費用 約9万円のみ）",
                Inches(0.8), Inches(4.95), Inches(3.5), Inches(0.4),
                font_size=12, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    add_textbox(slide, "→", Inches(4.3), Inches(4.2), Inches(1.0), Inches(0.8),
                font_size=42, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)

    add_textbox(slide, "売却価格", Inches(5.3), Inches(3.8), Inches(3.5), Inches(0.5),
                font_size=16, color=DARK_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, "145万円", Inches(5.3), Inches(4.2), Inches(3.5), Inches(0.8),
                font_size=52, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    add_textbox(slide, "写真を撮っただけ", Inches(5.3), Inches(4.95), Inches(3.5), Inches(0.4),
                font_size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    add_textbox(slide, "利益 約130万円", Inches(9.0), Inches(4.1), Inches(4.0), Inches(0.9),
                font_size=28, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    return slide


def slide_08_profit_structure(prs):
    """スライド8: 利益構造"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "利益構造：なぜ初心者でも勝てるのか",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=30, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    # フロー図
    boxes = [
        (ACCENT_BLUE,  "0〜30万円", "激安取得"),
        (ACCENT_GOLD,  "家賃収入", "月3〜5万円"),
        (ACCENT_GOLD,  "売却益", "数十〜100万円以上"),
    ]
    for i, (col, num, label) in enumerate(boxes):
        left = Inches(0.6 + i * 4.2)
        top  = Inches(1.4)
        add_rect(slide, left, top, Inches(3.7), Inches(2.5),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(3.7), Inches(0.08), col)
        add_textbox(slide, num, left + Inches(0.15), top + Inches(0.25),
                    Inches(3.4), Inches(0.8), font_size=34, bold=True, color=col,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, label, left + Inches(0.15), top + Inches(1.1),
                    Inches(3.4), Inches(0.6), font_size=18, color=WHITE,
                    align=PP_ALIGN.CENTER)
        if i < 2:
            add_textbox(slide, "→", Inches(0.6 + i * 4.2 + 3.7), top + Inches(0.8),
                        Inches(0.5), Inches(0.8), font_size=32, bold=True, color=ACCENT_GOLD,
                        align=PP_ALIGN.CENTER)

    # 再投資ループ
    add_rect(slide, Inches(0.6), Inches(4.1), Inches(12.1), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide, "利益を再投資 → 物件を増やす → 資産が雪だるま式に拡大",
                Inches(0.6), Inches(4.25), Inches(12.1), Inches(0.7),
                font_size=20, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)

    # ポイント
    points = [
        "何千万ものローン不要 → リスクが桁違いに低い",
        "最悪でも「元値に少し上乗せ売却」で損切り可能",
        "税理士が税務コスト最小化 → 手取りを最大化",
    ]
    for i, pt in enumerate(points):
        add_textbox(slide, f"• {pt}", Inches(0.6), Inches(5.1 + i * 0.65),
                    Inches(12.1), Inches(0.6), font_size=18, color=WHITE)
    return slide


def slide_09_rental_merit(prs):
    """スライド9: 戸建て賃貸のメリット"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "戸建て賃貸が最強な理由",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    merits = [
        ("競合ゼロ",
         "相場の半値で貸せばすぐ埋まる\nアパート1室分の家賃で「戸建て丸ごと」に住める"),
        ("管理ラク",
         "入居者が自分で庭掃除 → 管理業者すら不要なケースも\n退去も少なく長期安定収入"),
        ("税務有利",
         "減価償却・経費計上で課税所得を圧縮\n税理士サポートで手取りを最大化"),
        ("出口自由",
         "賃貸でも売却でも選べる\n利回り投資家に売れば数十万〜百万超の売却益"),
    ]

    for i, (title, desc) in enumerate(merits):
        col = i % 2
        row = i // 2
        left = Inches(0.5 + col * 6.4)
        top  = Inches(1.4 + row * 2.75)
        add_rect(slide, left, top, Inches(6.0), Inches(2.5),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(0.08), Inches(2.5), ACCENT_BLUE)
        add_textbox(slide, title, left + Inches(0.2), top + Inches(0.15),
                    Inches(5.5), Inches(0.55), font_size=22, bold=True, color=ACCENT_BLUE)
        add_textbox(slide, desc, left + Inches(0.2), top + Inches(0.8),
                    Inches(5.5), Inches(1.6), font_size=15, color=WHITE)
    return slide


def slide_10_nationwide(prs):
    """スライド10: 全国のチャンス（田舎の放置物件）"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "全国にチャンスは眠っている",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    # 大見出し
    add_textbox(slide, "日本全国の「放置・空き家」が狙い目",
                Inches(0.5), Inches(1.3), Inches(12), Inches(0.7),
                font_size=24, color=ACCENT_GOLD, bold=True)

    stats = [
        ("全国の空き家数", "900万戸以上"),
        ("空き家率",       "過去最高水準"),
        ("対象エリア",     "東京都心・大阪・名古屋 中心部以外の全国"),
    ]
    for i, (k, v) in enumerate(stats):
        top = Inches(2.2 + i * 1.35)
        add_rect(slide, Inches(0.5), top, Inches(12.0), Inches(1.2),
                 RGBColor(0x0a, 0x14, 0x24))
        add_textbox(slide, k, Inches(0.8), top + Inches(0.1), Inches(4), Inches(0.4),
                    font_size=14, color=DARK_GRAY)
        add_textbox(slide, v, Inches(0.8), top + Inches(0.45), Inches(11), Inches(0.65),
                    font_size=22, bold=True, color=WHITE)

    add_textbox(slide, "「田舎は借り手がいない」は思い込み",
                Inches(0.5), Inches(6.15), Inches(12), Inches(0.5),
                font_size=20, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.CENTER)
    add_textbox(slide, "→ 相場の半値で出せば、すぐに入居者が決まる",
                Inches(0.5), Inches(6.7), Inches(12), Inches(0.5),
                font_size=18, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    return slide


def slide_11_part3_intro(prs):
    """スライド11: PART3 メンバー実績"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "PART 3", Inches(0.5), Inches(0.25), Inches(12), Inches(0.5),
                font_size=18, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "メンバー実績",
                Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
                font_size=44, bold=True, color=WHITE)
    add_divider(slide, Inches(1.65))

    # 大数値
    add_textbox(slide, "年間利益",
                Inches(0.5), Inches(2.2), Inches(6), Inches(0.5),
                font_size=18, color=DARK_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, "100万〜300万円",
                Inches(0.5), Inches(2.65), Inches(6), Inches(1.1),
                font_size=42, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)

    add_textbox(slide, "利回り",
                Inches(6.8), Inches(2.2), Inches(6), Inches(0.5),
                font_size=18, color=DARK_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, "100%〜436%",
                Inches(6.8), Inches(2.65), Inches(6), Inches(1.1),
                font_size=42, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.CENTER)

    add_divider(slide, Inches(3.9))

    add_textbox(slide, "全員が不動産投資「初心者」 ／ 全員が30万円以下でスタート",
                Inches(0.5), Inches(4.1), Inches(12), Inches(0.7),
                font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    return slide


def slide_12_case_a_b(prs):
    """スライド12: 事例A/B"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "事例紹介 ①",
                Inches(0.5), Inches(0.25), Inches(12), Inches(0.6),
                font_size=28, bold=True, color=WHITE)
    add_divider(slide, Inches(0.95))

    cases = [
        {
            "label": "事例A",
            "person": "Tさん（40代・会社員・東京）",
            "start": "初期投資 約50万円",
            "result": "1年間で利益 300万円",
            "comment": "「富岡さんが『この物件はOK』『これは見送り』と判断してくれるから安心して進められました」",
        },
        {
            "label": "事例B",
            "person": "Mさん（30代・主婦・子育て中）",
            "start": "18万円で戸建て取得",
            "result": "家賃3.3万円 / 利回り220%",
            "comment": "「18万円の投資で年間39.6万円の収入。パートの月収を超えちゃいました」",
        },
    ]

    for i, c in enumerate(cases):
        left = Inches(0.5 + i * 6.4)
        top  = Inches(1.15)
        add_rect(slide, left, top, Inches(6.0), Inches(5.9),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(6.0), Inches(0.08), ACCENT_GOLD)
        add_textbox(slide, c["label"], left + Inches(0.2), top + Inches(0.15),
                    Inches(5.5), Inches(0.5), font_size=18, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, c["person"], left + Inches(0.2), top + Inches(0.65),
                    Inches(5.5), Inches(0.5), font_size=16, color=LIGHT_GRAY)
        add_textbox(slide, c["start"], left + Inches(0.2), top + Inches(1.25),
                    Inches(5.5), Inches(0.5), font_size=17, color=WHITE)
        add_rect(slide, left + Inches(0.2), top + Inches(1.85), Inches(5.5), Inches(0.06),
                 ACCENT_GOLD)
        add_textbox(slide, c["result"], left + Inches(0.2), top + Inches(2.0),
                    Inches(5.5), Inches(0.8), font_size=28, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, c["comment"], left + Inches(0.2), top + Inches(3.0),
                    Inches(5.5), Inches(2.6), font_size=14, color=LIGHT_GRAY, italic=True)
    return slide


def slide_13_case_c(prs):
    """スライド13: 事例C"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "事例紹介 ②",
                Inches(0.5), Inches(0.25), Inches(12), Inches(0.6),
                font_size=28, bold=True, color=WHITE)
    add_divider(slide, Inches(0.95))

    # Case C メイン
    add_rect(slide, Inches(0.5), Inches(1.15), Inches(12.0), Inches(4.5),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(0.5), Inches(1.15), Inches(12.0), Inches(0.08), ACCENT_GOLD)

    add_textbox(slide, "事例C", Inches(0.8), Inches(1.3), Inches(3), Inches(0.5),
                font_size=18, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "Kさん（60代・定年退職後）",
                Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
                font_size=18, color=LIGHT_GRAY)
    add_textbox(slide, "取得価格 0円",
                Inches(0.8), Inches(2.4), Inches(5), Inches(0.7),
                font_size=30, bold=True, color=ACCENT_BLUE)
    add_textbox(slide, "→",
                Inches(5.8), Inches(2.4), Inches(0.8), Inches(0.7),
                font_size=30, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "利回り 300%超",
                Inches(6.7), Inches(2.4), Inches(5), Inches(0.7),
                font_size=30, bold=True, color=ACCENT_GOLD)
    add_textbox(slide,
                "「0円で家が手に入るなんて信じられなかった。でも実際にできてしまった。\n"
                "税理士の富岡さんが税金面もアドバイスしてくれるから、手取りを最大化できています」",
                Inches(0.8), Inches(3.3), Inches(11.5), Inches(2.1),
                font_size=17, color=WHITE, italic=True)

    # D/E/F小カード
    small = [
        ("事例D", "10万円取得", "年利 124%"),
        ("事例E", "1万円取得", "年利 436%（!）"),
        ("事例F", "30万円取得", "リフォームなし売却益"),
    ]
    for i, (lbl, inv, ret) in enumerate(small):
        left = Inches(0.5 + i * 4.3)
        top  = Inches(6.0)
        add_rect(slide, left, top, Inches(4.0), Inches(1.2),
                 RGBColor(0x05, 0x0e, 0x1a))
        add_rect(slide, left, top, Inches(4.0), Inches(0.06), ACCENT_GOLD)
        add_textbox(slide, lbl, left + Inches(0.15), top + Inches(0.08),
                    Inches(1.5), Inches(0.35), font_size=12, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, inv, left + Inches(0.15), top + Inches(0.4),
                    Inches(3.5), Inches(0.35), font_size=14, color=LIGHT_GRAY)
        add_textbox(slide, ret, left + Inches(0.15), top + Inches(0.75),
                    Inches(3.5), Inches(0.4), font_size=16, bold=True, color=ACCENT_GOLD)
    return slide


def slide_14_results_table(prs):
    """スライド14: 実績一覧テーブル"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "メンバー実績一覧",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.65),
                font_size=32, bold=True, color=WHITE)
    add_divider(slide, Inches(1.05))

    # テーブルヘッダー
    headers = ["取得価格", "月家賃 / 売却益", "年利回り", "状態"]
    col_w   = [Inches(2.8), Inches(3.5), Inches(2.8), Inches(3.3)]
    col_l   = [Inches(0.5), Inches(3.3), Inches(6.8), Inches(9.6)]

    for j, (h, l, w) in enumerate(zip(headers, col_l, col_w)):
        add_rect(slide, l, Inches(1.25), w, Inches(0.55), ACCENT_BLUE)
        add_textbox(slide, h, l + Inches(0.1), Inches(1.3), w, Inches(0.5),
                    font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    rows = [
        ("0円",   "家賃 3.3万円", "300%超", "入居付け成功"),
        ("1万円", "家賃 3.3万円", "436%",   "安定稼働中"),
        ("10万円","家賃収入あり", "124%",   "安定稼働中"),
        ("18万円","家賃 3.3万円", "220%",   "入居付け成功"),
        ("30万円","家賃収入あり", "150%超", "賃貸中"),
        ("50万円","年間利益 300万円","600%超","1年間の実績"),
    ]

    for i, (row) in enumerate(rows):
        bg = RGBColor(0x0a, 0x14, 0x24) if i % 2 == 0 else RGBColor(0x07, 0x10, 0x1c)
        top_r = Inches(1.8 + i * 0.78)
        for j, (val, l, w) in enumerate(zip(row, col_l, col_w)):
            add_rect(slide, l, top_r, w, Inches(0.75), bg)
            col = ACCENT_GOLD if j in (0, 2) else WHITE
            add_textbox(slide, val, l + Inches(0.1), top_r + Inches(0.12),
                        w - Inches(0.1), Inches(0.5), font_size=16, color=col,
                        bold=(j in (0, 2)), align=PP_ALIGN.CENTER)

    add_textbox(slide, "※ 全員が不動産投資初心者  ※ グループ開始2年未満の実績",
                Inches(0.5), Inches(7.1), Inches(12), Inches(0.35),
                font_size=12, color=DARK_GRAY)
    return slide


def slide_15_roadmap(prs):
    """スライド15: 自己資金ゼロからのロードマップ"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "自己資金ゼロから不動産オーナーへのロードマップ",
                Inches(0.5), Inches(0.3), Inches(12.5), Inches(0.7),
                font_size=28, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    steps = [
        ("STEP 1", "MNPで資金調達", "携帯乗り換えキャッシュバック\n→ 数十万円を短期で現金化"),
        ("STEP 2", "激安物件を探す", "全国の放置物件・0〜30万円\nオンラインで物件検討可能"),
        ("STEP 3", "専門家と物件判断", "富岡が「買い/見送り」を判定\n初心者でも失敗しにくい"),
        ("STEP 4", "貸す or 売る", "相場半値で即入居 / 売却益\n年利100〜436%を狙う"),
        ("STEP 5", "利益を再投資", "雪だるま式に物件を増やす\n不動産オーナーへ"),
    ]

    for i, (step, title, desc) in enumerate(steps):
        left = Inches(0.5 + i * 2.55)
        top  = Inches(1.4)
        add_rect(slide, left, top, Inches(2.3), Inches(4.8),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(2.3), Inches(0.08), ACCENT_BLUE)
        add_textbox(slide, step, left + Inches(0.1), top + Inches(0.15),
                    Inches(2.1), Inches(0.4), font_size=12, bold=True, color=ACCENT_BLUE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, title, left + Inches(0.1), top + Inches(0.6),
                    Inches(2.1), Inches(0.7), font_size=16, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, desc, left + Inches(0.1), top + Inches(1.4),
                    Inches(2.1), Inches(3.2), font_size=13, color=LIGHT_GRAY,
                    align=PP_ALIGN.CENTER)
        if i < 4:
            add_textbox(slide, "→", left + Inches(2.3), top + Inches(1.8),
                        Inches(0.25), Inches(0.6), font_size=18, bold=True, color=ACCENT_GOLD,
                        align=PP_ALIGN.CENTER)
    return slide


def slide_16_offer_overview(prs):
    """スライド16: オファー概要"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "【ご案内】MNP × 激安不動産投資 コンサルティング",
                Inches(0.5), Inches(0.2), Inches(12.5), Inches(0.75),
                font_size=26, bold=True, color=WHITE)
    add_divider(slide, Inches(1.05))

    add_textbox(slide, "含まれる内容 6項目",
                Inches(0.5), Inches(1.15), Inches(12), Inches(0.5),
                font_size=20, bold=True, color=ACCENT_GOLD)

    items = [
        ("01", "激安不動産投資の完全コンテンツ",
         "物件探し・購入方法（10種類以上）・入居付け・売却交渉 を網羅"),
        ("02", "月1回のZoom勉強会",
         "最新市場動向・成功事例・Q&Aをライブで"),
        ("03", "バンドグループでの質問サポート",
         "24時間以内に回答 → 一人で悩む時間ゼロ"),
        ("04", "物件購入相談",
         "資料を送れば税理士の目で「買い/見送り」を判断"),
        ("05", "税務アドバイス",
         "確定申告・経費計上・キャッシュフロー最大化を税理士がサポート"),
        ("06", "MNP初期資金構築プログラム（池田宜史監修）",
         "自己資金ゼロから始められる資金作りを完全サポート"),
    ]

    for i, (num, title, desc) in enumerate(items):
        top = Inches(1.75 + i * 0.88)
        add_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.82),
                 RGBColor(0x0a, 0x14, 0x24))
        add_textbox(slide, num, Inches(0.65), top + Inches(0.12),
                    Inches(0.6), Inches(0.55), font_size=14, bold=True, color=ACCENT_GOLD)
        add_textbox(slide, title, Inches(1.3), top + Inches(0.05),
                    Inches(5.5), Inches(0.4), font_size=16, bold=True, color=WHITE)
        add_textbox(slide, desc, Inches(1.3), top + Inches(0.43),
                    Inches(11.2), Inches(0.35), font_size=13, color=LIGHT_GRAY)
    return slide


def slide_17_anchoring(prs):
    """スライド17: アンカリング（300万以上の価値）"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "この価値、いくらに相当するか？",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=32, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    anchors = [
        ("一般的な不動産投資スクール",      "100万〜300万円", "「知識」だけで物件相談なし"),
        ("税理士顧問料（年間）",            "30万〜50万円",   "税務アドバイスのみ"),
        ("MNP指導（個別コーチング）",       "数十万円",       "資金作りのみ"),
    ]
    total_top = Inches(1.3)
    for i, (item, price, note) in enumerate(anchors):
        top = Inches(1.3 + i * 1.5)
        add_rect(slide, Inches(0.5), top, Inches(12.0), Inches(1.35),
                 RGBColor(0x0a, 0x14, 0x24))
        add_textbox(slide, item, Inches(0.8), top + Inches(0.12),
                    Inches(6.5), Inches(0.5), font_size=18, color=WHITE)
        add_textbox(slide, price, Inches(7.4), top + Inches(0.05),
                    Inches(4.5), Inches(0.6), font_size=26, bold=True, color=ACCENT_GOLD,
                    align=PP_ALIGN.RIGHT)
        add_textbox(slide, note, Inches(0.8), top + Inches(0.7),
                    Inches(11), Inches(0.45), font_size=13, color=DARK_GRAY)

    # 合計
    add_rect(slide, Inches(0.5), Inches(5.7), Inches(12.0), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide, "合計すれば", Inches(0.5), Inches(5.85), Inches(5), Inches(0.6),
                font_size=20, color=LIGHT_GRAY)
    add_textbox(slide, "300万円以上の価値", Inches(5.5), Inches(5.75), Inches(7.0), Inches(0.8),
                font_size=36, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.RIGHT)
    return slide


def slide_18_plan_compare(prs):
    """スライド18: プランA/B比較"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "ご案内プラン",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.65),
                font_size=34, bold=True, color=WHITE)
    add_divider(slide, Inches(1.05))

    # プランA
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(5.9), Inches(5.6),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(5.9), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide, "プラン A", Inches(0.7), Inches(1.45), Inches(5.5), Inches(0.55),
                font_size=20, bold=True, color=ACCENT_GOLD)
    add_textbox(slide, "98万円", Inches(0.7), Inches(2.0), Inches(5.5), Inches(1.1),
                font_size=58, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    plan_a = [
        "期間: 12ヶ月",
        "物件購入相談: 30回",
        "Zoom勉強会（月1回）",
        "バンドグループサポート",
        "税務アドバイス",
        "MNP資金構築プログラム",
    ]
    for i, pt in enumerate(plan_a):
        add_textbox(slide, f"✓ {pt}", Inches(0.8), Inches(3.25 + i * 0.52),
                    Inches(5.3), Inches(0.5), font_size=16, color=WHITE)

    # プランB
    add_rect(slide, Inches(6.8), Inches(1.3), Inches(5.9), Inches(5.6),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(6.8), Inches(1.3), Inches(5.9), Inches(0.08), ACCENT_BLUE)
    add_textbox(slide, "プラン B", Inches(7.0), Inches(1.45), Inches(5.5), Inches(0.55),
                font_size=20, bold=True, color=ACCENT_BLUE)
    add_textbox(slide, "59.8万円", Inches(7.0), Inches(2.0), Inches(5.5), Inches(1.1),
                font_size=52, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.CENTER)
    plan_b = [
        "期間: 6ヶ月",
        "物件購入相談: 10回",
        "Zoom勉強会（月1回）",
        "バンドグループサポート",
        "税務アドバイス",
        "MNP資金構築プログラム",
    ]
    for i, pt in enumerate(plan_b):
        add_textbox(slide, f"✓ {pt}", Inches(7.0), Inches(3.25 + i * 0.52),
                    Inches(5.3), Inches(0.5), font_size=16, color=WHITE)
    return slide


def slide_19_roi_calc(prs):
    """スライド19: 投資回収計算"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "98万円は「コスト」ではなく「投資」",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=30, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    add_textbox(slide, "ちょっと計算してみてください",
                Inches(0.5), Inches(1.3), Inches(12), Inches(0.5),
                font_size=20, color=LIGHT_GRAY)

    calcs = [
        ("最悪のケース", "1年以内に回収",
         "メンバー平均下限の\n年間利益 100万円 > 投資額 98万円"),
        ("標準ケース", "1.5〜2倍のリターン",
         "年間利益 150〜200万円\n余剰で次の物件に再投資"),
        ("最良のケース", "3倍以上のリターン",
         "年間利益 300万円\n2年で600万円の資産形成"),
    ]

    for i, (case, result, detail) in enumerate(calcs):
        left = Inches(0.5 + i * 4.3)
        top  = Inches(2.0)
        add_rect(slide, left, top, Inches(3.9), Inches(4.6),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, left, top, Inches(3.9), Inches(0.08), ACCENT_GOLD)
        add_textbox(slide, case, left + Inches(0.15), top + Inches(0.15),
                    Inches(3.6), Inches(0.45), font_size=14, color=DARK_GRAY,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, result, left + Inches(0.15), top + Inches(0.7),
                    Inches(3.6), Inches(0.8), font_size=24, bold=True, color=ACCENT_GOLD,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, detail, left + Inches(0.15), top + Inches(1.65),
                    Inches(3.6), Inches(2.7), font_size=16, color=WHITE,
                    align=PP_ALIGN.CENTER)

    add_textbox(slide, "しかも身についたスキルは「一生使える資産」",
                Inches(0.5), Inches(6.85), Inches(12), Inches(0.55),
                font_size=19, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    return slide


def slide_20_guarantee(prs):
    """スライド20: 限定・保証"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, RED_WARN)

    add_textbox(slide, "重要なお知らせ",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.65),
                font_size=30, bold=True, color=WHITE)
    add_divider(slide, Inches(1.05))

    # 限定枠
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(12.0), Inches(2.8),
             RGBColor(0x1a, 0x05, 0x05))
    add_rect(slide, Inches(0.5), Inches(1.3), Inches(12.0), Inches(0.08), RED_WARN)
    add_textbox(slide, "個別相談（ロードマップ作成会）",
                Inches(0.8), Inches(1.45), Inches(11.5), Inches(0.55),
                font_size=20, bold=True, color=WHITE)
    add_textbox(slide, "月 10名 限定",
                Inches(0.8), Inches(2.05), Inches(11.5), Inches(1.1),
                font_size=60, bold=True, color=RED_WARN, align=PP_ALIGN.CENTER)
    add_textbox(slide,
                "一人一人の物件相談にしっかり時間をかけるため\n大量受け入れは行いません。枠が埋まり次第、次回募集は未定。",
                Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.8),
                font_size=16, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    # Q&Aつぶし
    add_textbox(slide, "よくある疑問",
                Inches(0.5), Inches(4.2), Inches(12), Inches(0.45),
                font_size=18, bold=True, color=ACCENT_GOLD)

    qs = [
        ("初心者でも大丈夫？", "大丈夫。事例全員が初心者でした"),
        ("自己資金がない？",   "MNPで種銭を作るのがこのプログラムの核心"),
        ("忙しい会社員でも？", "スキマ時間+専門家サポートで本業と両立できます"),
    ]
    for i, (q, a) in enumerate(qs):
        top = Inches(4.75 + i * 0.72)
        add_textbox(slide, f"Q: {q}", Inches(0.6), top, Inches(5.5), Inches(0.6),
                    font_size=16, bold=True, color=WHITE)
        add_textbox(slide, f"→ {a}", Inches(6.3), top, Inches(6.4), Inches(0.6),
                    font_size=16, color=ACCENT_GOLD)
    return slide


def slide_21_cta_roadmap(prs):
    """スライド21: ロードマップ作成会（個別相談・無料）"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "まずは無料でお試しください",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.6),
                font_size=24, color=LIGHT_GRAY)
    add_textbox(slide, "ロードマップ作成会（個別相談）",
                Inches(0.5), Inches(0.85), Inches(12), Inches(0.85),
                font_size=38, bold=True, color=WHITE)
    add_textbox(slide, "完全無料",
                Inches(0.5), Inches(1.65), Inches(12), Inches(0.75),
                font_size=44, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    add_divider(slide, Inches(2.5))

    contents = [
        "あなたの現在の資産・収入状況をヒアリング",
        "MNP × 激安不動産投資でどう進めるかを個別設計",
        "セミナーでは話しきれなかった具体的な数字・プランをご提示",
        "「今すぐ始められるか」「まだ準備が必要か」を正直にお伝えします",
    ]
    for i, c in enumerate(contents):
        add_textbox(slide, f"• {c}", Inches(0.7), Inches(2.7 + i * 0.82),
                    Inches(12.0), Inches(0.75), font_size=19, color=WHITE)

    add_rect(slide, Inches(0.5), Inches(6.2), Inches(12.0), Inches(0.08), ACCENT_BLUE)
    add_textbox(slide, "富岡信成 × 池田宜史 が直接対応します",
                Inches(0.5), Inches(6.35), Inches(12.0), Inches(0.55),
                font_size=18, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.CENTER)
    return slide


def slide_22_limit(prs):
    """スライド22: 月10名限定"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, RED_WARN)

    add_textbox(slide, "枠は限られています",
                Inches(0.5), Inches(0.7), Inches(12), Inches(0.7),
                font_size=28, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
    add_textbox(slide, "月10名 限定",
                Inches(0.5), Inches(1.45), Inches(12), Inches(1.6),
                font_size=80, bold=True, color=RED_WARN, align=PP_ALIGN.CENTER)

    add_divider(slide, Inches(3.15))

    reasons = [
        "一人一人の物件を直接確認するため、大量受け入れは不可能",
        "税理士の富岡が個別に税務も見るため、クオリティを落とせない",
        "枠が埋まり次第、次回募集は「未定」",
    ]
    for i, r in enumerate(reasons):
        add_textbox(slide, f"• {r}", Inches(1.0), Inches(3.45 + i * 0.9),
                    Inches(11.0), Inches(0.8), font_size=20, color=WHITE,
                    align=PP_ALIGN.CENTER)

    add_textbox(slide, "「後で」は存在しないかもしれません",
                Inches(0.5), Inches(6.6), Inches(12), Inches(0.65),
                font_size=22, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    return slide


def slide_23_how_to_apply(prs):
    """スライド23: 申込方法"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "申込方法",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=36, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    steps = [
        ("STEP 1", "下のQRコード / URLにアクセス", "ロードマップ作成会の申込フォームへ"),
        ("STEP 2", "お名前・連絡先を入力", "24時間以内にご連絡します"),
        ("STEP 3", "日程調整（Zoom）", "ご都合の良い日時をご選択ください"),
        ("STEP 4", "当日、オンラインで面談", "約60分 / 完全無料"),
    ]

    for i, (step, title, desc) in enumerate(steps):
        top = Inches(1.4 + i * 1.35)
        add_rect(slide, Inches(0.5), top, Inches(12.3), Inches(1.2),
                 RGBColor(0x0a, 0x14, 0x24))
        add_rect(slide, Inches(0.5), top, Inches(0.08), Inches(1.2), ACCENT_BLUE)

        add_textbox(slide, step, Inches(0.8), top + Inches(0.1),
                    Inches(1.5), Inches(0.4), font_size=13, bold=True, color=ACCENT_BLUE)
        add_textbox(slide, title, Inches(0.8), top + Inches(0.5),
                    Inches(8.5), Inches(0.5), font_size=20, bold=True, color=WHITE)
        add_textbox(slide, desc, Inches(0.8), top + Inches(0.88),
                    Inches(11.5), Inches(0.4), font_size=14, color=LIGHT_GRAY)

    add_textbox(slide, "▼ 今すぐ申し込む ▼",
                Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.55),
                font_size=22, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    return slide


def slide_24_qa(prs):
    """スライド24: Q&A"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_BLUE)

    add_textbox(slide, "Q & A",
                Inches(0.5), Inches(0.3), Inches(12), Inches(0.7),
                font_size=40, bold=True, color=WHITE)
    add_divider(slide, Inches(1.1))

    qas = [
        ("Q. 不動産投資の知識がまったくありません",
         "A. 全く問題ありません。物件選びから購入・入居・売却まですべてサポートします。"),
        ("Q. 自己資金がほとんどゼロです",
         "A. まずMNPで資金を作るのがこのプログラムの出発点。ゼロからスタートできます。"),
        ("Q. 田舎の物件なんて借り手がいるの？",
         "A. います。相場の半値で貸せばすぐ決まります。事例で証明済み。"),
        ("Q. 怪しくないですか？",
         "A. 税理士（富岡）が実名・顔出し。国家資格保持者が詐欺をやれば資格剥奪です。"),
        ("Q. 損するリスクは？",
         "A. 激安取得だから最悪でも「少し上乗せ売却」で損切り可能。数千万投資とは次元が違います。"),
    ]

    for i, (q, a) in enumerate(qas):
        top = Inches(1.35 + i * 1.13)
        add_rect(slide, Inches(0.5), top, Inches(12.3), Inches(1.05),
                 RGBColor(0x0a, 0x14, 0x24))
        add_textbox(slide, q, Inches(0.8), top + Inches(0.08),
                    Inches(11.7), Inches(0.42), font_size=16, bold=True, color=ACCENT_BLUE)
        add_textbox(slide, a, Inches(0.8), top + Inches(0.52),
                    Inches(11.7), Inches(0.45), font_size=14, color=LIGHT_GRAY)
    return slide


def slide_25_final_message(prs):
    """スライド25: 最終メッセージ"""
    slide = add_slide(prs)
    add_rect(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_H, ACCENT_GOLD)

    add_textbox(slide, "1年後——あなたはどうなっていたいですか？",
                Inches(0.5), Inches(0.4), Inches(12), Inches(0.75),
                font_size=26, bold=True, color=WHITE)
    add_divider(slide, Inches(1.25))

    # Before
    add_rect(slide, Inches(0.5), Inches(1.45), Inches(5.7), Inches(3.4),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(0.5), Inches(1.45), Inches(5.7), Inches(0.08), RED_WARN)
    add_textbox(slide, "何もしなかった場合",
                Inches(0.7), Inches(1.6), Inches(5.3), Inches(0.5),
                font_size=16, color=RED_WARN, bold=True)
    bef = [
        "今と同じ給料・今と同じ不安",
        "「いつかやろう」と思い続ける",
        "AIに仕事を奪われるリスク増大",
        "老後・年金への不安が続く",
    ]
    for i, b in enumerate(bef):
        add_textbox(slide, f"• {b}", Inches(0.7), Inches(2.2 + i * 0.6),
                    Inches(5.3), Inches(0.55), font_size=16, color=LIGHT_GRAY)

    add_textbox(slide, "vs", Inches(6.3), Inches(2.7), Inches(0.8), Inches(0.8),
                font_size=28, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)

    # After
    add_rect(slide, Inches(7.2), Inches(1.45), Inches(5.7), Inches(3.4),
             RGBColor(0x0a, 0x14, 0x24))
    add_rect(slide, Inches(7.2), Inches(1.45), Inches(5.7), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide, "一歩踏み出した場合",
                Inches(7.4), Inches(1.6), Inches(5.3), Inches(0.5),
                font_size=16, color=ACCENT_GOLD, bold=True)
    aft = [
        "毎月3万円〜の家賃収入",
        "不動産という「実物資産」を保有",
        "AI時代も揺るがない資産基盤",
        "「動いてよかった」という実感",
    ]
    for i, a in enumerate(aft):
        add_textbox(slide, f"✓ {a}", Inches(7.4), Inches(2.2 + i * 0.6),
                    Inches(5.3), Inches(0.55), font_size=16, color=WHITE)

    # P.S.
    add_rect(slide, Inches(0.5), Inches(5.1), Inches(12.4), Inches(0.08), ACCENT_GOLD)
    add_textbox(slide,
                "P.S. 私（富岡）は一度すべてを失いました。それでも激安不動産投資で這い上がれた。\n"
                "正しい方法と正しいサポートがあれば、誰にでもチャンスがある——それを証明するために活動しています。",
                Inches(0.5), Inches(5.3), Inches(12.4), Inches(1.3),
                font_size=15, color=LIGHT_GRAY, italic=True)

    add_textbox(slide, "富岡信成 × 池田宜史",
                Inches(0.5), Inches(6.75), Inches(12.4), Inches(0.55),
                font_size=20, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.RIGHT)
    return slide


# ─── メイン ───────────────────────────────────────────

def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    builders = [
        slide_01_title,
        slide_02_agenda,
        slide_03_mnp_intro,
        slide_04_mnp_steps,
        slide_05_mnp_result,
        slide_06_part2_intro,
        slide_07_zero_yen_case,
        slide_08_profit_structure,
        slide_09_rental_merit,
        slide_10_nationwide,
        slide_11_part3_intro,
        slide_12_case_a_b,
        slide_13_case_c,
        slide_14_results_table,
        slide_15_roadmap,
        slide_16_offer_overview,
        slide_17_anchoring,
        slide_18_plan_compare,
        slide_19_roi_calc,
        slide_20_guarantee,
        slide_21_cta_roadmap,
        slide_22_limit,
        slide_23_how_to_apply,
        slide_24_qa,
        slide_25_final_message,
    ]

    for i, builder in enumerate(builders, 1):
        print(f"  スライド {i:02d} / {len(builders)} 生成中: {builder.__name__}")
        builder(prs)

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "セミナースライド.pptx"
    )
    prs.save(out_path)
    print(f"\n保存完了: {out_path}")
    return out_path


if __name__ == "__main__":
    main()

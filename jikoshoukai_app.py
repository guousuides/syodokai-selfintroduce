import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import os
import re
import unicodedata
from io import BytesIO
from PyPDF2 import PdfFileReader, PdfFileWriter

# ==========================================
# 定数・設定
# ==========================================

# ★ テンプレートPDFのパス ★
# このスクリプトと同じフォルダにある R8自己紹介.pdf を使います。
TEMPLATE_PDF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "R8自己紹介.pdf")

# ★ 太字設定 ★
USE_BOLD = True
BOLD_WIDTH_RATIO = 0.05

# --- フォント設定 ---
SPECIAL_CHARS_FONT_MAP = {
    '嵗': 'MSMincho',
    '俻': 'MSMincho',
}
DEFAULT_FONT = "MSMincho"


# ==========================================
# ★★★ 要素の追加・位置・サイズ設定 ★★★
# ==========================================
#
# 【新しい要素を追加する方法】
#   COORDINATES 辞書に新しいエントリを追加してください。
#   キー名 = CSVの列名（ヘッダー名）と一致させてください。
#
# 【設定できるパラメータ】
#
#   ● 必須パラメータ:
#     'x'          : 横位置 (左端が0、右端が595) ※A4の幅=595pt
#     'y'          : 縦位置 (下端が0、上端が842) ※A4の高さ=842pt
#     'font_size'  : フォントサイズ (例: 10, 14, 18, 24)
#
#   ● 折り返し設定 (長いテキスト用):
#     'wrap'        : True=折り返しあり, False=折り返しなし (デフォルト: False)
#     'max_chars'   : 1行の最大文字数 (wrapがTrueのとき必須)
#     'line_height' : 行高さ (px単位、デフォルト: font_size * 1.2)
#
#   ● 中央揃え:
#     'centered'   : True=中央揃え (xを中心位置として指定)
#
# 【追加例】
#
#   # 例1: シンプルな横書き（1行）
#   '趣味': {
#       'x': 100, 'y': 500, 'font_size': 14,
#   },
#
#   # 例2: 折り返しあり（長いテキスト）
#   '自己PR': {
#       'x': 50, 'y': 400, 'font_size': 12,
#       'wrap': True, 'max_chars': 30, 'line_height': 16,
#   },
#
#   # 例3: 中央揃え
#   'タイトル': {
#       'x': 297, 'y': 750, 'font_size': 20,
#       'wrap': True, 'max_chars': 20, 'line_height': 24,
#       'centered': True,
#   },
#
# 【画像の追加方法】
#   IMAGE_COORDINATES 辞書にエントリを追加してください。
#   キー名 = CSVの列名（ヘッダー名）と一致させてください。
#   CSVには画像ファイル名（例: photo.jpg）を記入してください。
#   画像ファイルはCSVと同じフォルダに置いてください。
#
#   設定パラメータ:
#     'x'      : 横位置 (左端が0)
#     'y'      : 縦位置 (下端が0) ※画像の左下が基準点
#     'width'  : 画像の幅 (pt)
#     'height' : 画像の高さ (pt)
#
#   # 例:
#   '写真': {
#       'x': 400, 'y': 650, 'width': 100, 'height': 130,
#   },
#

COORDINATES = {

    # --- 記入日 ---
    '記入日': {
        'x': 315, 'y': 730, 'font_size': 18,
    },
    # --- 記入月 ---
    '記入月': {
        'x': 270, 'y': 730, 'font_size': 18,
    },

    # --- 出身地 ---
    '出身地': {
        'x': 115, 'y': 637.5, 'font_size': 12,
    },

    # --- 出身高校 ---
    '出身高校': {
        'x': 115, 'y': 615, 'font_size': 18,
    },

    # --- 役職 ---
    '役職': {
        'x': 387.5, 'y': 615, 'font_size': 12,
    },
    # --- sns ---
    'sns': {
        'x': 387.5, 'y': 575, 'font_size': 12,
    },
  
    # --- 学部学年 ---
    '学部学科学年': {
        'x': 65, 'y': 575, 'font_size': 18,
    },

    # --- 名前 ---
    '氏名': {
        'x': 125, 'y': 685, 'font_size': 24,
    },

    # --- ふりがな ---
    'ふりがな': {
        'x': 125, 'y': 715, 'font_size': 12,
    },

    # --- コメント (折り返しあり) ---
    'コメント': {
        'x': 144, 'y': 127.5, 'font_size': 18,
        'wrap': True, 'max_chars': 19, 'line_height': 20,
    },

    # --- 年齢 ---
    '年齢': {
        'x': 287.5, 'y': 657.5, 'font_size': 18,
    },
    # --- 年 ---
    '年': {
        'x': 115, 'y': 657.5, 'font_size': 18,
    },
    # --- 年齢 ---
    '月': {
        'x': 155, 'y': 657.5, 'font_size': 18,
    },
    # --- 年齢 ---
    '日': {
        'x': 202.5, 'y': 657.5, 'font_size': 18,
    },

    # 上まで生年月日

    '履歴1年': {
        'x': 65, 'y': 516, 'font_size': 18,
    },

    '履歴1月': {
        'x': 113, 'y': 516, 'font_size': 18,
    },

    '履歴1内容': {
        'x': 144, 'y': 516, 'font_size': 18,
    },

    '履歴2年': {
        'x': 65, 'y': 491, 'font_size': 18,
    },

    '履歴2月': {
        'x': 113, 'y': 491, 'font_size': 18,
    },

    '履歴2内容': {
        'x': 144, 'y': 491, 'font_size': 18,
    },

    '履歴3年': {
        'x': 65, 'y': 467, 'font_size': 18,
    },

    '履歴3月': {
        'x': 113, 'y': 467, 'font_size': 18,
    },

    '履歴3内容': {
        'x': 144, 'y': 467, 'font_size': 18,
    },

    '履歴4年': {
        'x': 65, 'y': 443, 'font_size': 18,
    },

    '履歴4月': {
        'x': 113, 'y': 443, 'font_size': 18,
    },

    '履歴4内容': {
        'x': 144, 'y': 443, 'font_size': 18,
    },

    '履歴5年': {
        'x': 65, 'y': 419, 'font_size': 18,
    },

    '履歴5月': {
        'x': 113, 'y': 419, 'font_size': 18,
    },

    '履歴5内容': {
        'x': 144, 'y': 419, 'font_size': 18,
    },

    '入会理由': {
        'x': 144, 'y': 380, 'font_size': 18,
    },


    # ※ 気になる書体は SHOTAI_CHECKBOXES で別管理（下記参照）

    '得意書体': {
        'x': 144, 'y': 310, 'font_size': 18,
    },

    '中高の部活': {
        'x': 144, 'y': 285, 'font_size': 18,
    },

    '兼サー先': {
        'x': 265, 'y': 262, 'font_size': 18,
    },

    # 兼サー有無: x座標は draw_content_on_overlay 内で動的に変更
    # '有' → x=144, それ以外 → x=160
    '兼サー有無': {
        'x': 160, 'y': 265, 'font_size': 9,
    },

    '趣味特技': {
        'x': 144, 'y': 237, 'font_size': 18,
    },

    'アルバイト先': {
        'x': 265, 'y': 214, 'font_size': 18,
    },

    # アルバイト有無: x座標は draw_content_on_overlay 内で動的に変更
    # '有' → x=212.5, それ以外 → x=147.5
    'アルバイト有無': {
        'x': 160, 'y': 217, 'font_size': 9,
    },

    'おすすめワセ飯': {
        'x': 144, 'y': 189, 'font_size': 18,
    },

    'プチ野望': {
        'x': 144, 'y': 164, 'font_size': 18,
    },

    # ★★★ ここに新しい要素を追加してください ★★★
    # 上の【追加例】を参考に、CSVの列名をキーにして追加します。
}

# ==========================================
# ★★★ 気になる書体チェックボックス座標 ★★★
# ==========================================
# CSVの「気になる書体」列にカンマ区切りで書体名を記入してください。
# 例: 楷書,行書,草書
# 該当する書体の位置に✓が描画されます。
#
SHOTAI_CHECKBOXES = {
    '楷書':     {'x': 155,   'y': 358},
    '行書':     {'x': 190,   'y': 358},
    '草書':     {'x': 224,   'y': 358},
    '隷書':     {'x': 258,   'y': 358},
    '篆書':     {'x': 292,   'y': 358},
    'カリグラ':  {'x': 326,   'y': 358},
    '篆刻':     {'x': 377.5, 'y': 358},
    '刻字':     {'x': 410,   'y': 358},
    'ペン字':   {'x': 445,   'y': 358},
    '墨絵':     {'x': 155,   'y': 342.5},
    '写経':     {'x': 190,   'y': 342.5},
    '北魏':     {'x': 224,   'y': 342.5},
    'その他':   {'x': 290,   'y': 342},
}
SHOTAI_FONT_SIZE = 9

# ==========================================
# ★★★ 画像の位置・サイズ設定 ★★★
# ==========================================
# キー名 = CSVの列名。CSVには画像ファイル名を記入。
# 画像ファイルはCSVと同じフォルダに配置してください。
#
IMAGE_COORDINATES = {

    # --- 写真 ---
    '写真': {
        'x': 440, 'y': 717.5, 'width': 100, 'height': 130,
    },

    # ★★★ ここに新しい画像要素を追加してください ★★★
}


# ==========================================
# 描画関数群
# ==========================================

def draw_horizontal_text(c, text, x, y, font_name, font_size=12):
    """横書き（1行）"""
    c.setFont(font_name, font_size)
    if USE_BOLD:
        c.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        c.setStrokeColorRGB(0, 0, 0)
        c._code.append('2 Tr')
    c.drawString(x, y, str(text))
    if USE_BOLD:
        c._code.append('0 Tr')

def draw_horizontal_text_with_wrap(c, text, x, y, font_name, font_size=12, max_chars_per_line=40, line_height=14):
    """横書き（折り返しあり）"""
    c.setFont(font_name, font_size)
    if USE_BOLD:
        c.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        c.setStrokeColorRGB(0, 0, 0)
        c._code.append('2 Tr')
    lines = []
    text_str = str(text)
    start = 0
    while start < len(text_str):
        end = start + max_chars_per_line
        if end >= len(text_str):
            lines.append(text_str[start:])
            break
        newline_pos = text_str.find('\n', start, end)
        if newline_pos != -1:
            lines.append(text_str[start:newline_pos])
            start = newline_pos + 1
        else:
            lines.append(text_str[start:end])
            start = end
        continue
    for i, line in enumerate(lines):
        c.drawString(x, y - i * line_height, line)
    if USE_BOLD:
        c._code.append('0 Tr')

def draw_centered_horizontal_text_with_wrap(c, text, x, y, font_name, font_size=12, max_chars_per_line=40, line_height=14):
    """横書き・中央揃え（折り返しあり）"""
    c.setFont(font_name, font_size)
    if USE_BOLD:
        c.setLineWidth(font_size * BOLD_WIDTH_RATIO)
        c.setStrokeColorRGB(0, 0, 0)
        c._code.append('2 Tr')
    lines = []
    text_str = str(text)
    start = 0
    while start < len(text_str):
        end = start + max_chars_per_line
        if end >= len(text_str):
            lines.append(text_str[start:])
            break
        newline_pos = text_str.find('\n', start, end)
        if newline_pos != -1:
            lines.append(text_str[start:newline_pos])
            start = newline_pos + 1
        else:
            lines.append(text_str[start:end])
            start = end
        continue
    for i, line in enumerate(lines):
        line_width = pdfmetrics.stringWidth(line, font_name, font_size)
        draw_x = x - line_width / 2
        c.drawString(draw_x, y - i * line_height, line)
    if USE_BOLD:
        c._code.append('0 Tr')


# ==========================================
# データ処理関数
# ==========================================

def preprocess_data(data: pd.DataFrame) -> pd.DataFrame:
    """CSVデータの前処理。必要に応じてカスタマイズしてください。"""
    if "氏名" not in data.columns and "ユーザー名" in data.columns:
        data["氏名"] = data["ユーザー名"]

    data = data.fillna('')

    # if "学部" in data.columns and "学年" in data.columns:
    #     data["学部学年"] = data.apply(lambda r: f"{str(r.get('学部', '') or '')} {str(r.get('学年', '') or '')}", axis=1)
    #     data.drop(columns=["学部", "学年"], inplace=True)

    # if "ふりがな" in data.columns:
    #     data["ふりがな"] = data["ふりがな"].apply(lambda x: f"（{x}）" if x else "")

    if "学部" in data.columns and "学科" in data.columns and "学年" in data.columns:
        data["学部学科学年"] = data.apply(lambda r: f"{str(r.get('学部', '') or '')} {str(r.get('学科', '') or '')} {str(r.get('学年', '') or '')}", axis=1)
        data.drop(columns=["学部", "学科", "学年"], inplace=True)

    if "タイムスタンプ" in data.columns:
        data.drop(columns=["タイムスタンプ"], inplace=True)

    return data


# ==========================================
# テキスト描画（オーバーレイ作成）
# ==========================================

def draw_content_on_overlay(overlay_canvas, data_row, csv_dir=""):
    """COORDINATESに基づいてテキストと画像をオーバーレイキャンバスに描画"""
    for column, coord in COORDINATES.items():
        value = data_row.get(column, "")
        if not value or str(value).strip() == "":
            continue

        # 兼サー有無: 値に応じてx座標を動的に変更し、表示を✓にする
        if column == '兼サー有無':
            coord = dict(coord)  # 元の辞書を変更しないようにコピー
            if str(value).strip() == '有':
                coord['x'] = 212.5
            else:
                coord['x'] = 147.5
            value = '✓'

        # アルバイト有無: 値に応じてx座標を動的に変更し、表示を✓にする
        if column == 'アルバイト有無':
            coord = dict(coord)
            if str(value).strip() == '有':
                coord['x'] = 212.5
            else:
                coord['x'] = 147.5
            value = '✓'

        x = coord['x']
        y = coord['y']
        font_size = coord['font_size']
        line_height = coord.get('line_height', int(font_size * 1.2))

        if coord.get('centered', False):
            draw_centered_horizontal_text_with_wrap(
                overlay_canvas, value, x, y, DEFAULT_FONT,
                font_size, coord.get('max_chars', 40), line_height
            )
        elif coord.get('wrap', False):
            draw_horizontal_text_with_wrap(
                overlay_canvas, value, x, y, DEFAULT_FONT,
                font_size, coord.get('max_chars', 40), line_height
            )
        else:
            draw_horizontal_text(overlay_canvas, value, x, y, DEFAULT_FONT, font_size)

    # --- 気になる書体のチェック描画 ---
    shotai_value = data_row.get('気になる書体', '')
    if shotai_value and str(shotai_value).strip():
        selected = [s.strip() for s in str(shotai_value).split(';')]
        for shotai_name, pos in SHOTAI_CHECKBOXES.items():
            if shotai_name in selected:
                draw_horizontal_text(
                    overlay_canvas, '✓', pos['x'], pos['y'],
                    DEFAULT_FONT, SHOTAI_FONT_SIZE
                )

    # --- 画像の描画 ---
    for column, img_coord in IMAGE_COORDINATES.items():
        filename = data_row.get(column, "")
        if not filename or str(filename).strip() == "":
            continue
        img_path = os.path.join(csv_dir, str(filename).strip())
        if not os.path.exists(img_path):
            print(f"※ 画像が見つかりません: {img_path}（列: {column}）")
            continue
        try:
            overlay_canvas.drawImage(
                img_path,
                img_coord['x'] - img_coord['width'] / 2,
                img_coord['y'] - img_coord['height'] / 2,
                width=img_coord['width'], height=img_coord['height'],
                preserveAspectRatio=True, mask='auto'
            )
        except Exception as e:
            print(f"※ 画像の描画に失敗: {img_path} ({e})")


# ==========================================
# PDF生成（テンプレート + オーバーレイ合成）
# ==========================================

def create_overlay_for_row(data_row, csv_dir=""):
    """1行分のデータからオーバーレイPDFを作成"""
    packet = BytesIO()
    overlay_canvas = canvas.Canvas(packet, pagesize=portrait(A4))
    draw_content_on_overlay(overlay_canvas, data_row, csv_dir)
    overlay_canvas.save()
    packet.seek(0)
    return packet

def generate_combined_pdf(data, pdf_file_path, template_path, csv_dir=""):
    """全員を1つのPDFにまとめて生成（CSVの1行 = PDFの1ページ）"""
    writer = PdfFileWriter()

    for index, row in data.iterrows():
        template_reader = PdfFileReader(template_path)
        template_page = template_reader.getPage(0)

        overlay_packet = create_overlay_for_row(row, csv_dir)
        overlay_reader = PdfFileReader(overlay_packet)
        template_page.mergePage(overlay_reader.getPage(0))

        writer.addPage(template_page)

    with open(pdf_file_path, 'wb') as f:
        writer.write(f)
    messagebox.showinfo("完了", f"{len(data)}件のページを作成しました:\n{pdf_file_path}")

def generate_individual_pdfs(data, output_dir, template_path, csv_dir=""):
    """1人ずつ個別のPDFを生成"""
    for index, row in data.iterrows():
        student_name = row.get("氏名", f"no_name_{index}")
        safe_filename = re.sub(r'[\\/*?:"<>|]', "", str(student_name)) + ".pdf"
        pdf_path = os.path.join(output_dir, safe_filename)

        writer = PdfFileWriter()
        template_reader = PdfFileReader(template_path)
        template_page = template_reader.getPage(0)

        overlay_packet = create_overlay_for_row(row, csv_dir)
        overlay_reader = PdfFileReader(overlay_packet)
        template_page.mergePage(overlay_reader.getPage(0))

        writer.addPage(template_page)

        with open(pdf_path, 'wb') as f:
            writer.write(f)

    messagebox.showinfo("完了", f"{len(data)}件のPDFファイルを作成しました:\n{output_dir}")


# ==========================================
# メイン処理
# ==========================================

def main():
    root = tk.Tk()
    root.withdraw()

    if not os.path.exists(TEMPLATE_PDF):
        messagebox.showerror("エラー", f"テンプレートPDFが見つかりません:\n{TEMPLATE_PDF}\n\nスクリプトと同じフォルダに R8自己紹介.pdf を配置してください。")
        return

    csv_path = filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
    )
    if not csv_path:
        return

    try:
        data = pd.read_csv(csv_path, encoding='utf-8')
    except FileNotFoundError:
        messagebox.showerror("エラー", "CSVファイルが見つかりません。")
        return
    except UnicodeDecodeError:
        data = pd.read_csv(csv_path, encoding='shift-jis')
    except Exception as e:
        messagebox.showerror("エラー", f"CSV読み込みでエラー:\n{e}")
        return

    data = preprocess_data(data)

    try:
        pdfmetrics.registerFont(TTFont("HGRGE", "C:/Windows/Fonts/HGRGE.TTC"))
        pdfmetrics.registerFont(TTFont("HGRME", "C:/Windows/Fonts/HGRME.TTC"))
        pdfmetrics.registerFont(TTFont("MSMincho", "C:/Windows/Fonts/msmincho.ttc"))
    except Exception as e:
        messagebox.showerror("フォントエラー", f"フォントの読み込みに失敗しました。\n\nエラー詳細:\n{e}")
        return

    create_individual = messagebox.askyesno(
        "生成モードの選択",
        "はい(Yes): 1人ずつ個別のPDFを作成します。\nいいえ(No): 全員を1つのPDFにまとめて作成します。"
    )

    if create_individual:
        csv_dir = os.path.dirname(csv_path)
        output_dir = os.path.join(csv_dir, "自己紹介PDF")
        os.makedirs(output_dir, exist_ok=True)
        generate_individual_pdfs(data, output_dir, TEMPLATE_PDF, csv_dir)
    else:
        pdf_path = filedialog.asksaveasfilename(
            title="PDF保存先を選択してください",
            defaultextension=".pdf",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )
        if pdf_path:
            csv_dir = os.path.dirname(csv_path)
            generate_combined_pdf(data, pdf_path, TEMPLATE_PDF, csv_dir)

if __name__ == "__main__":
    main()

import sys
import os
import csv
import re
import urllib.request
import tempfile
from io import BytesIO
from datetime import datetime
from PIL import Image, ImageTk
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTableWidget, QTableWidgetItem,
                             QLabel, QComboBox, QLineEdit, QPushButton,
                             QHeaderView, QAbstractItemView, QFrame,
                             QSplitter, QScrollArea, QGridLayout, QFileDialog,
                             QMessageBox, QProgressDialog)
from PyQt5.QtCore import Qt, QUrl, QSize
from PyQt5.QtGui import QPixmap, QIcon, QFont
from PyQt5.QtPrintSupport import QPrinter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


class MemberManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("会員管理アプリケーション")
        self.setGeometry(100, 100, 1200, 800)

        # データ保存用
        self.data = []
        self.filtered_data = []

        # 並べ替え状態を保存
        self.sort_column = 2  # デフォルトは子供の名前
        self.sort_order = Qt.AscendingOrder  # 昇順

        # 詳細表示管理用の辞書
        self.expanded_rows = {}

        # 現在のCSVファイルパス
        self.current_csv_path = "ANS.csv"

        # 画像キャッシュ
        self.image_cache = {}

        # PDF用の日本語フォント設定
        self.initialize_pdf_fonts()

        # UI設定
        self.init_ui()

        # データ読み込み
        self.load_data(self.current_csv_path)

    def initialize_pdf_fonts(self):
        """PDF用のヒラギノ丸ゴシック体フォントを初期化"""
        try:
            # macOSの場合のヒラギノ丸ゴシック体のパス
            hiragino_path = "/Library/Fonts/ヒラギノ丸ゴ ProN W4.ttc"

            # Windowsの場合はフォントがインストールされているかを確認
            # (Windows環境にヒラギノ丸ゴシック体がインストールされている場合)
            windows_hiragino_path = "C:/Windows/Fonts/ヒラギノ丸ゴ ProN W4.ttc"

            # ヒラギノ丸ゴシック体のカスタムパス (ユーザーが指定する場合)
            custom_hiragino_path = os.path.join(os.path.expanduser("~"), "AppData/Local/Microsoft/Windows/Fonts/ヒラギノ丸ゴ ProN W4.ttc")

            # フォントファイルのパスを確認
            if os.path.exists(hiragino_path):
                font_path = hiragino_path
            elif os.path.exists(windows_hiragino_path):
                font_path = windows_hiragino_path
            elif os.path.exists(custom_hiragino_path):
                font_path = custom_hiragino_path
            else:
                # 代替のフォントパスを探索 (日本語フォント)
                meiryo_path = "C:/Windows/Fonts/meiryo.ttc"
                msgothic_path = "C:/Windows/Fonts/msgothic.ttc"
                msmincho_path = "C:/Windows/Fonts/msmincho.ttc"

                if os.path.exists(meiryo_path):
                    font_path = meiryo_path
                    print("ヒラギノ丸ゴシック体が見つからないため、メイリオを使用します")
                elif os.path.exists(msgothic_path):
                    font_path = msgothic_path
                    print("ヒラギノ丸ゴシック体が見つからないため、MS Gothicを使用します")
                elif os.path.exists(msmincho_path):
                    font_path = msmincho_path
                    print("ヒラギノ丸ゴシック体が見つからないため、MS Minchoを使用します")
                else:
                    print("日本語フォントが見つかりません。デフォルトフォントを使用します")
                    return

            # 見つかったフォントを登録
            print(f"フォントを登録します: {font_path}")
            if font_path.endswith('.ttc'):  # TrueTypeコレクション
                # ヒラギノ丸ゴシック体はTrueTypeコレクション(.ttc)ファイルのため、インデックスを指定
                pdfmetrics.registerFont(TTFont('HiraginoMaru', font_path, subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont('HiraginoMaru', font_path))

            # 登録に成功したメッセージ
            print("フォント登録成功")

        except Exception as e:
            print(f"フォント初期化エラー: {e}")
            # エラーが発生しても処理を続行

    def init_ui(self):
        # メインウィジェット
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # タイトル
        title_label = QLabel("会員管理システム")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)

        # フィルターエリア
        filter_frame = QFrame()
        filter_layout = QHBoxLayout(filter_frame)

        # CSV参照ボタン
        self.csv_button = QPushButton("CSVファイル選択")
        self.csv_button.clicked.connect(self.open_csv_file)
        filter_layout.addWidget(self.csv_button)

        # PDF出力ボタン
        self.pdf_button = QPushButton("PDF出力")
        self.pdf_button.clicked.connect(self.export_to_pdf)
        filter_layout.addWidget(self.pdf_button)

        # 区切り
        separator = QLabel("|")
        filter_layout.addWidget(separator)

        # 学年フィルタ
        grade_label = QLabel("学年フィルター:")
        self.grade_combo = QComboBox()
        self.grade_combo.addItem("すべて", "all")
        self.grade_combo.currentIndexChanged.connect(self.apply_filters)
        filter_layout.addWidget(grade_label)
        filter_layout.addWidget(self.grade_combo)

        # 検索フィルタ
        search_label = QLabel("検索:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("名前や四字熟語で検索...")
        self.search_input.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(search_label)
        filter_layout.addWidget(self.search_input)

        main_layout.addWidget(filter_frame)

        # テーブルウィジェット
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["詳細", "写真", "ご本人名", "お子様名", "学年", "次年度候補"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().sectionClicked.connect(self.sort_table)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)

        main_layout.addWidget(self.table)

        # 詳細表示エリア
        self.details_area = QScrollArea()
        self.details_area.setWidgetResizable(True)
        self.details_area.setVisible(False)
        self.details_widget = QWidget()
        self.details_layout = QGridLayout(self.details_widget)
        self.details_area.setWidget(self.details_widget)

        main_layout.addWidget(self.details_area)

        # ステータスバー
        self.statusBar().showMessage("データロード準備完了")

    def load_data(self, filename):
        try:
            # ファイルの存在チェック
            if not os.path.exists(filename):
                self.statusBar().showMessage(f"ファイルが見つかりません: {filename}")
                return

            with open(filename, 'r', encoding='utf-8') as f:
                csv_reader = csv.DictReader(f)
                self.data = []
                for row in csv_reader:
                    processed_row = {
                        'parent_name': row.get('ご本人のお名前', ''),
                        'child_name': row.get('お子様のお名前', ''),
                        'grade': row.get('お子様の学年', ''),
                        'child_phrase': row.get('お子様を表す四字熟語', ''),
                        'parent_phrase': row.get('ご自身を表す四字熟語', ''),
                        'next_year_candidate': row.get('次年度委員候補ですか？', ''),
                        'can_participate': row.get('委員会運営に参加可能ですか？', ''),
                        'reason': row.get('理由をお聞かせください', ''),
                        'impression': row.get('委員会への所感\nをお答えください', ''),
                        'photo_url': row.get('お写真', '')
                    }
                    self.data.append(processed_row)

                # 学年リストを更新
                grades = set(item['grade'] for item in self.data if item['grade'])
                self.grade_combo.clear()
                self.grade_combo.addItem("すべて", "all")
                for grade in sorted(grades):
                    self.grade_combo.addItem(grade, grade)

                # 初期表示
                self.filtered_data = self.data.copy()
                self.update_table()
                self.statusBar().showMessage(f"データ読み込み完了: {filename} ({len(self.data)}件)")

                # 現在のファイルパスを更新
                self.current_csv_path = filename
                self.setWindowTitle(f"会員管理アプリケーション - {os.path.basename(filename)}")
        except Exception as e:
            self.statusBar().showMessage(f"データ読み込みエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"CSVファイルの読み込み中にエラーが発生しました:\n{str(e)}")

    def open_csv_file(self):
        """CSVファイル選択ダイアログを開く"""
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(
            self,
            "CSVファイルを選択",
            os.path.dirname(self.current_csv_path),  # 前回のディレクトリを開く
            "CSVファイル (*.csv);;すべてのファイル (*)",
            options=options
        )

        if fileName:
            # 選択されたファイルを読み込む
            self.load_data(fileName)

    def apply_filters(self):
        # 学年フィルター
        grade_filter = self.grade_combo.currentData()

        # 検索フィルター
        search_term = self.search_input.text().lower()

        # フィルター適用
        self.filtered_data = []
        for item in self.data:
            # 学年フィルター
            if grade_filter != "all" and item['grade'] != grade_filter:
                continue

            # 検索フィルター
            if search_term and not (
                    search_term in item['parent_name'].lower() or
                    search_term in item['child_name'].lower() or
                    search_term in item['child_phrase'].lower() or
                    search_term in item['parent_phrase'].lower()
            ):
                continue

            self.filtered_data.append(item)

        # 現在のソート条件で並べ替え
        self.sort_data()

        # テーブル更新
        self.update_table()
        self.statusBar().showMessage(f"表示: {len(self.filtered_data)}/{len(self.data)}件")

    def sort_data(self):
        column_map = {
            2: 'parent_name',
            3: 'child_name',
            4: 'grade',
            5: 'next_year_candidate'
        }

        if self.sort_column in column_map:
            key = column_map[self.sort_column]
            reverse = (self.sort_order == Qt.DescendingOrder)
            self.filtered_data.sort(key=lambda x: x[key], reverse=reverse)

    def sort_table(self, column_index):
        if column_index in [2, 3, 4, 5]:  # ソート可能な列
            if self.sort_column == column_index:
                # 同じ列をクリックした場合は昇順/降順を切り替え
                self.sort_order = Qt.DescendingOrder if self.sort_order == Qt.AscendingOrder else Qt.AscendingOrder
            else:
                # 異なる列の場合は昇順に設定
                self.sort_column = column_index
                self.sort_order = Qt.AscendingOrder

            # データを並べ替え
            self.sort_data()

            # テーブル更新
            self.update_table()

    def convert_google_drive_url(self, url):
        if not url:
            return ''

        # 'open?id=' パターンの場合
        if 'open?id=' in url:
            file_id = url.split('open?id=')[1].split('&')[0]
            return f'https://drive.google.com/uc?export=view&id={file_id}'

        # 'file/d/' パターンの場合
        elif 'file/d/' in url:
            file_id = url.split('file/d/')[1].split('/')[0]
            return f'https://drive.google.com/uc?export=view&id={file_id}'

        # その他の場合はそのまま返す
        return url

    def load_image_from_url(self, url):
        # キャッシュにあれば使用
        if url in self.image_cache:
            return self.image_cache[url]

        try:
            converted_url = self.convert_google_drive_url(url)
            req = urllib.request.Request(
                converted_url,
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            with urllib.request.urlopen(req, timeout=5) as response:
                image_data = response.read()

            # PILで画像を開き、リサイズ
            image = Image.open(BytesIO(image_data))
            image = image.resize((64, 64), Image.LANCZOS)

            # PyQt用のQPixmapに変換
            pixmap = QPixmap()
            pixmap.loadFromData(image_data)
            pixmap = pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)

            # キャッシュに保存
            self.image_cache[url] = pixmap

            return pixmap
        except Exception as e:
            print(f"画像読み込みエラー: {str(e)}")
            return QPixmap()

    def update_table(self):
        self.table.setRowCount(0)  # テーブルをクリア

        for i, item in enumerate(self.filtered_data):
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)

            # 詳細ボタン
            detail_btn = QPushButton("詳細")
            detail_btn.clicked.connect(lambda checked, row=i: self.toggle_details(row))
            self.table.setCellWidget(row_position, 0, detail_btn)

            # 写真
            photo_label = QLabel()
            photo_label.setAlignment(Qt.AlignCenter)
            if item['photo_url']:
                # この部分は実際の環境で画像を読み込む処理に置き換え
                pixmap = self.load_image_from_url(item['photo_url'])
                if not pixmap.isNull():
                    photo_label.setPixmap(pixmap)
                else:
                    photo_label.setText("読み込みエラー")
            else:
                photo_label.setText("画像なし")
            photo_label.setFixedSize(70, 70)
            self.table.setCellWidget(row_position, 1, photo_label)

            # 親の名前
            parent_name_item = QTableWidgetItem(item['parent_name'])
            self.table.setItem(row_position, 2, parent_name_item)

            # 子供の名前
            child_name_item = QTableWidgetItem(item['child_name'])
            self.table.setItem(row_position, 3, child_name_item)

            # 学年
            grade_item = QTableWidgetItem(item['grade'])
            grade_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_position, 4, grade_item)

            # 次年度候補
            candidate_item = QTableWidgetItem(item['next_year_candidate'])
            candidate_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_position, 5, candidate_item)

        self.table.resizeRowsToContents()

    def toggle_details(self, row_index):
        # 詳細表示を切り替え
        if row_index in self.expanded_rows:
            self.expanded_rows.pop(row_index)
            self.details_area.setVisible(False)
        else:
            self.expanded_rows = {row_index: True}  # 他の行の詳細表示をクリア
            self.show_details(row_index)

    def show_details(self, row_index):
        item = self.filtered_data[row_index]

        # 詳細エリアをクリア
        for i in reversed(range(self.details_layout.count())):
            self.details_layout.itemAt(i).widget().setParent(None)

        # 詳細情報を追加
        self.details_layout.addWidget(QLabel("<b>基本情報</b>"), 0, 0)
        self.details_layout.addWidget(QLabel(f"<b>お子様の四字熟語:</b> {item['child_phrase']}"), 1, 0)
        self.details_layout.addWidget(QLabel(f"<b>ご自身の四字熟語:</b> {item['parent_phrase']}"), 2, 0)
        self.details_layout.addWidget(QLabel(f"<b>委員会運営参加:</b> {item['can_participate']}"), 3, 0)

        self.details_layout.addWidget(QLabel("<b>詳細情報</b>"), 0, 1)
        self.details_layout.addWidget(QLabel(f"<b>理由:</b> {item['reason'] or '-'}"), 1, 1)

        # 所感は複数行の可能性があるのでQTextEditで表示
        impression_label = QLabel(f"<b>委員会への所感:</b> {item['impression'] or '-'}")
        impression_label.setWordWrap(True)
        self.details_layout.addWidget(impression_label, 2, 1, 2, 1)

        self.details_area.setVisible(True)

    def export_to_pdf(self):
        """学年ごとにPDFを出力する（改良版画像処理）"""
        try:
            # 保存先を選択
            options = QFileDialog.Options()
            save_dir = QFileDialog.getExistingDirectory(
                self,
                "PDF保存先フォルダを選択",
                os.path.dirname(self.current_csv_path),
                options=options
            )

            if not save_dir:
                return

            # 学年ごとにデータをグループ化
            grade_groups = {}
            for item in self.data:
                grade = item['grade']
                if not grade:
                    grade = "未設定"

                if grade not in grade_groups:
                    grade_groups[grade] = []

                grade_groups[grade].append(item)

            # 進捗ダイアログ
            progress = QProgressDialog("PDFを生成中...", "キャンセル", 0, len(grade_groups), self)
            progress.setWindowTitle("PDF出力")
            progress.setWindowModality(Qt.WindowModal)
            progress.show()

            success_count = 0

            for i, (grade, items) in enumerate(grade_groups.items()):
                # キャンセルされた場合
                if progress.wasCanceled():
                    break

                progress.setValue(i)
                progress.setLabelText(f"{grade}のPDFを生成中...")

                # ファイル名を設定
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                safe_grade = re.sub(r'[\\/*?:"<>|]', '', grade)  # ファイル名に使えない文字を削除
                output_file = os.path.join(save_dir, f"プロフィール_{safe_grade}_{timestamp}.pdf")

                # PDFを生成
                try:
                    self.generate_profile_pdf(output_file, grade, items)
                    success_count += 1
                except Exception as e:
                    QMessageBox.warning(
                        self,
                        "警告",
                        f"{grade}のPDF生成中にエラーが発生しました: {str(e)}\n処理を続行します。"
                    )
                    print(f"PDF生成エラー ({grade}): {e}")
                    continue

            progress.setValue(len(grade_groups))

            if success_count > 0:
                QMessageBox.information(
                    self,
                    "完了",
                    f"{success_count}/{len(grade_groups)}学年のPDF出力が完了しました。\n保存先: {save_dir}"
                )
            else:
                QMessageBox.critical(
                    self,
                    "エラー",
                    "PDF出力に失敗しました。システム環境を確認してください。"
                )

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"PDF出力中にエラーが発生しました:\n{str(e)}")
            print(f"PDF出力全体エラー: {e}")

    def generate_profile_pdf(self, output_file, grade, items):
        """プロフィール形式のPDFを生成する（ヒラギノ丸ゴシック体使用）"""
        # PDF作成準備
        doc = SimpleDocTemplate(
            output_file,
            pagesize=A4,
            leftMargin=10 * mm,
            rightMargin=10 * mm,
            topMargin=10 * mm,
            bottomMargin=10 * mm
        )

        # スタイル定義
        styles = getSampleStyleSheet()

        # ヒラギノ丸ゴシック体を使用したスタイル
        # 登録されていれば使用、なければ代替フォント
        font_name = 'HiraginoMaru' if 'HiraginoMaru' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'

        japanese_style = ParagraphStyle(
            'JapaneseStyle',
            parent=styles['Normal'],
            fontName=font_name,
            fontSize=10,
            leading=12,
            wordWrap='CJK'
        )

        japanese_heading = ParagraphStyle(
            'JapaneseHeading',
            parent=styles['Heading1'],
            fontName=font_name,
            fontSize=14,
            leading=16,
            wordWrap='CJK'
        )

        # 内容作成
        elements = []

        # タイトル
        title = Paragraph(f"{grade} プロフィール一覧", japanese_heading)
        elements.append(title)
        elements.append(Spacer(1, 10 * mm))

        # 1ページに6名分(3列×2行)表示するためのレイアウト
        rows = []
        current_row = []

        # 一時ファイルのリスト（後で削除するため）
        temp_files = []

        # メンバーごとにプロフィールカードを作成
        for i, item in enumerate(items):
            try:
                # プロフィールカードを作成（画像あり）
                profile_card, tmp_file = self.create_improved_profile_card(item, japanese_style)
                if tmp_file:
                    temp_files.append(tmp_file)

                # 3列のレイアウト
                current_row.append(profile_card)
                if len(current_row) == 3:
                    rows.append(current_row)
                    current_row = []

                # 2行で新しいページ
                if len(rows) == 2 and current_row == []:
                    # テーブルでレイアウト
                    profile_table = Table(rows, colWidths=[6 * cm, 6 * cm, 6 * cm])
                    profile_table.setStyle(TableStyle([
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]))

                    elements.append(profile_table)
                    elements.append(Spacer(1, 5 * mm))

                    # ページを分ける
                    elements.append(Paragraph("", styles['Normal']))
                    elements.append(Spacer(1, 10 * mm))

                    rows = []
            except Exception as e:
                print(f"プロフィールカード作成エラー: {e}")
                # エラーの場合はその会員をスキップ
                continue

        # 残りのアイテムを処理
        if current_row:
            # 3列になるまで空のセルで埋める
            while len(current_row) < 3:
                current_row.append("")
            rows.append(current_row)

        if rows:
            profile_table = Table(rows, colWidths=[6 * cm, 6 * cm, 6 * cm])
            profile_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]))

            elements.append(profile_table)

        try:
            # PDFを保存
            doc.build(elements)
        finally:
            # 一時ファイルの削除
            for tmp_file in temp_files:
                try:
                    if os.path.exists(tmp_file):
                        os.unlink(tmp_file)
                except Exception as e:
                    print(f"一時ファイル削除エラー: {e}")

    def create_improved_profile_card(self, item, style):
        """一人分のプロフィールカードを作成（ヒラギノ丸ゴシック体使用）"""
        # プロフィール情報を整理
        texts = []
        texts.append(f"<b>{item['parent_name']}</b>")
        texts.append(f"お子様: {item['child_name']}")
        texts.append(f"学年: {item['grade']}")

        if item['child_phrase']:
            texts.append(f"お子様の四字熟語: {item['child_phrase']}")
        if item['parent_phrase']:
            texts.append(f"ご自身の四字熟語: {item['parent_phrase']}")

        texts.append(f"次年度候補: {item['next_year_candidate']}")
        if item['can_participate']:
            texts.append(f"委員会運営: {item['can_participate']}")

        # 全てのテキストを1つのパラグラフにまとめる
        content = Paragraph("<br/>".join(texts), style)

        # 画像の処理
        img = None
        temp_file_path = None

        if item['photo_url']:
            try:
                # URLの変換（Google Drive用）
                converted_url = self.convert_google_drive_url(item['photo_url'])

                # 画像をダウンロード
                req = urllib.request.Request(
                    converted_url,
                    headers={'User-Agent': 'Mozilla/5.0'}
                )

                with urllib.request.urlopen(req, timeout=10) as response:
                    image_data = response.read()

                # 一時ファイルを作成
                with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp_file:
                    temp_file.write(image_data)
                    temp_file_path = temp_file.name



                # ReportLabの画像オブジェクト作成
                img = ReportLabImage(temp_file_path, width=3 * cm, height=3 * cm)

            except Exception as e:
                print(f"画像ダウンロードエラー ({item['photo_url']}): {e}")
                # エラーが発生した場合は画像なしで続行
                img = None#????????
        # プロフィールカードの作成
        if img:
            # 画像とテキストを組み合わせる
            profile_table = Table([
                [img],
                [content]
            ], colWidths=[5.5 * cm])

            profile_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ('VALIGN', (0, 1), (0, 1), 'TOP'),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ]))

            return profile_table, temp_file_path
        else:
            # 画像なしのテキストのみ
            profile_table = Table([
                [content]
            ], colWidths=[5.5 * cm])

            profile_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ]))

            return profile_table, None
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MemberManagementApp()
    window.show()
    sys.exit(app.exec_())
import sys
import os

# .venv内のプラグインのパスに環境変数を設定
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = ".venv/lib/python3.12/site-packages/PyQt5/Qt5/plugins/platforms"

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
                             QMessageBox, QProgressDialog, QMenu)
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
        """PDF用の日本語フォントを初期化"""
        try:
            # フォント設定ファイルパス
            font_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "font_config.txt")

            # 保存されたフォント設定を読み込む
            font_path = None
            if os.path.exists(font_config_path):
                try:
                    with open(font_config_path, 'r', encoding='utf-8') as f:
                        saved_path = f.read().strip()
                        if os.path.exists(saved_path) and (saved_path.endswith('.ttf') or saved_path.endswith('.otf')):
                            font_path = saved_path
                            print(f"保存された設定からフォントを読み込みました: {font_path}")
                except Exception as e:
                    print(f"フォント設定の読み込みエラー: {e}")

            # 保存された設定がない場合は、デフォルトパスを探索
            if not font_path:
                default_paths = [
                    # ローカルフォント
                    "YuMincho.ttf",  # TrueType形式
                    "YuMincho.otf",  # OpenType形式

                    # システムフォントパス (macOS)
                    "/Library/Fonts/ヒラギノ明朝 ProN.ttc",
                    "/System/Library/Fonts/ヒラギノ明朝 ProN.ttc",

                    # システムフォントパス (Windows)
                    "C:/Windows/Fonts/msgothic.ttf",
                    "C:/Windows/Fonts/msmincho.ttf",
                    "C:/Windows/Fonts/meiryo.ttf",
                ]

                for path in default_paths:
                    if os.path.exists(path) and (path.endswith('.ttf') or path.endswith('.otf')):
                        font_path = path
                        print(f"デフォルトパスからフォントを見つけました: {font_path}")
                        break

            # フォントが見つからない場合は警告し、デフォルトフォントを使用
            if not font_path:
                print("日本語フォントが見つかりません。デフォルトフォントを使用します。")
                self.statusBar().showMessage("日本語フォントが見つかりません。PDF出力にはフォント選択が必要です。")
                return

            # 見つかったフォントを登録
            print(f"フォントを登録します: {font_path}")

            # otf/ttfファイルのみ登録
            if font_path.endswith('.ttf') or font_path.endswith('.otf'):
                pdfmetrics.registerFont(TTFont('JapaneseFont', font_path))
                print("フォント登録成功")

                # 設定ファイルに保存
                try:
                    with open(font_config_path, 'w', encoding='utf-8') as f:
                        f.write(font_path)
                    print(f"フォント設定を保存しました: {font_path}")
                except Exception as e:
                    print(f"フォント設定の保存エラー: {e}")
            else:
                print(f"未対応のフォント形式です: {font_path}")
                self.statusBar().showMessage("未対応のフォント形式です。.ttfまたは.otfファイルを選択してください。")

        except Exception as e:
            print(f"フォント初期化エラー: {e}")
            self.statusBar().showMessage(f"フォント初期化エラー: {e}")


    def check_font_before_pdf_export(self):
        """PDF出力前にフォント設定をチェックし、必要に応じてフォント選択ダイアログを表示"""
        # 登録済みのフォント名を確認
        registered_fonts = pdfmetrics.getRegisteredFontNames()

        if 'JapaneseFont' not in registered_fonts:
            # フォントが登録されていない場合
            reply = QMessageBox.question(
                self,
                "フォント設定",
                "PDF出力用の日本語フォントが設定されていません。\n今すぐフォントを設定しますか？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                # フォント選択ダイアログを表示
                self.select_font_file()

                # フォントが選択されたか再確認
                if 'JapaneseFont' not in pdfmetrics.getRegisteredFontNames():
                    QMessageBox.warning(
                        self,
                        "フォント未設定",
                        "日本語フォントが設定されていません。\nPDF出力時に日本語が正しく表示されない場合があります。"
                    )
                    return False
            else:
                # フォント設定をスキップする場合は警告
                QMessageBox.warning(
                    self,
                    "フォント未設定",
                    "日本語フォントが設定されていません。\nPDF出力時に日本語が正しく表示されない場合があります。"
                )
                return False

        return True

    def export_to_pdf(self):
        """学年ごとにPDFを出力する（フォントチェック付き）"""
        # フォント設定をチェック
        if not self.check_font_before_pdf_export():
            # ユーザーがキャンセルしたか、フォント設定に失敗した場合
            return

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
                    grade = ""

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

    def select_font_file(self):
        """フォントファイル選択ダイアログを表示"""
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(
            self,
            "PDF用日本語フォントを選択",
            "",
            "フォントファイル (*.ttf *.otf);;すべてのファイル (*)",
            options=options
        )

        if fileName:
            # フォントファイルの拡張子をチェック
            if fileName.lower().endswith(('.ttf', '.otf')):
                # フォント設定ファイルパス
                font_config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "font_config.txt")

                # 設定ファイルに保存
                try:
                    with open(font_config_path, 'w', encoding='utf-8') as f:
                        f.write(fileName)

                    # フォントを再登録
                    try:
                        pdfmetrics.registerFont(TTFont('JapaneseFont', fileName))
                        self.statusBar().showMessage(f"フォントを設定しました: {os.path.basename(fileName)}")
                        QMessageBox.information(
                            self,
                            "フォント設定完了",
                            f"フォントを設定しました:\n{os.path.basename(fileName)}"
                        )
                    except Exception as e:
                        print(f"フォント登録エラー: {e}")
                        self.statusBar().showMessage(f"フォント登録エラー: {e}")
                        QMessageBox.warning(
                            self,
                            "フォント登録エラー",
                            f"フォントの登録に失敗しました:\n{str(e)}"
                        )
                except Exception as e:
                    print(f"フォント設定の保存エラー: {e}")
                    self.statusBar().showMessage(f"フォント設定の保存エラー: {e}")
            else:
                self.statusBar().showMessage("未対応のフォント形式です。.ttfまたは.otfファイルを選択してください。")
                QMessageBox.warning(
                    self,
                    "未対応のフォント形式",
                    "選択されたファイルは対応していません。\n.ttfまたは.otfフォントファイルを選択してください。"
                )

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

        # PDF関連ボタン用のメニュー
        pdf_button = QPushButton("PDF")
        pdf_menu = QMenu(self)

        pdf_export_action = pdf_menu.addAction("PDF出力")
        pdf_export_action.triggered.connect(self.export_to_pdf)

        pdf_font_action = pdf_menu.addAction("フォント設定")
        pdf_font_action.triggered.connect(self.select_font_file)

        pdf_button.setMenu(pdf_menu)
        filter_layout.addWidget(pdf_button)

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
        self.table.setHorizontalHeaderLabels(["詳細", "写真", "ご本人名", "お子様名", "高等部二年生保護者の皆様へご挨拶","お住まいの地域"])
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
                        'parent_name': row.get('回答者のお名前', ''),
                        'child_name': row.get('お子様のお名前', ''),
                        'grade': row.get('お子様の学年', ''),
                        'child_phrase': row.get('高等部二年生保護者の皆様へご挨拶', ''),
                        'parent_phrase': row.get('お住まいの地域', ''),
                        'photo_url': row.get('お子様と回答者の写真', '')
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
            #
            # # 次年度候補
            # candidate_item = QTableWidgetItem(item['next_year_candidate'])
            # candidate_item.setTextAlignment(Qt.AlignCenter)
            # self.table.setItem(row_position, 5, candidate_item)

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
        self.details_layout.addWidget(QLabel(f"<b>お子様の名前:</b> {item['child_name']}"), 1, 0)
        self.details_layout.addWidget(QLabel(f"<b>高等部二年生保護者の皆様へご挨拶:</b> {item['child_phrase']}"), 2, 0)
        self.details_layout.addWidget(QLabel(f"<b>お住まいの地域:</b> {item['parent_phrase']}"), 3, 0)

        # 追加情報がある場合のみ表示（エラー防止）
        if 'can_participate' in item:
            self.details_layout.addWidget(QLabel(f"<b>委員会運営参加:</b> {item['can_participate']}"), 4, 0)

        self.details_layout.addWidget(QLabel("<b>詳細情報</b>"), 0, 1)

        # 追加情報がある場合のみ表示（エラー防止）
        if 'reason' in item:
            self.details_layout.addWidget(QLabel(f"<b>理由:</b> {item.get('reason') or '-'}"), 1, 1)

        if 'impression' in item:
            # 所感は複数行の可能性があるのでQTextEditで表示
            impression_label = QLabel(f"<b>委員会への所感:</b> {item.get('impression') or '-'}")
            impression_label.setWordWrap(True)
            self.details_layout.addWidget(impression_label, 2, 1, 2, 1)

        self.details_area.setVisible(True)

    def generate_profile_pdf(self, output_file, grade, items):
        """プロフィール形式のPDFを生成する（日本語フォント対応）"""
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

        # 日本語フォントを使用したスタイル
        # 'JapaneseFont'が登録されていれば使用、なければ代替フォント
        font_name = 'JapaneseFont' if 'JapaneseFont' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'

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

        # 以下、既存のコードと同じ...

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


    def convert_google_drive_url(self, url):
        """GoogleドライブのURLを直接アクセス可能なURLに変換する（最新版）"""
        if not url:
            return ''

        # ファイルIDを抽出
        file_id = None

        # 'open?id=' パターンの場合
        if 'open?id=' in url:
            file_id = url.split('open?id=')[1].split('&')[0] if '&' in url.split('open?id=')[1] else url.split('open?id=')[1]

        # 'file/d/' パターンの場合
        elif 'file/d/' in url:
            file_id = url.split('file/d/')[1].split('/')[0] if '/' in url.split('file/d/')[1] else url.split('file/d/')[1]

        # 'id=' パターンの場合
        elif 'id=' in url:
            file_id = url.split('id=')[1].split('&')[0] if '&' in url.split('id=')[1] else url.split('id=')[1]

        # その他のパターンでIDを探す
        else:
            # GoogleドライブのIDは通常25-33文字の英数字とハイフン
            import re
            matches = re.findall(r'[-\w]{25,}', url)
            if matches:
                file_id = matches[0]

        if not file_id:
            print(f"警告: GoogleドライブのファイルIDを抽出できませんでした: {url}")
            return url

        # 2023-2025年版の複数のアクセス方法を試す (優先順位あり)
        return f"https://lh3.googleusercontent.com/d/{file_id}"

    def fetch_image_with_retry(self, url, max_retries=3):
        """複数の方法を試して画像を取得する"""
        original_url = url

        # 取得を試みるURL形式のリスト
        file_id = None

        # ファイルIDを抽出
        if 'open?id=' in url:
            file_id = url.split('open?id=')[1].split('&')[0] if '&' in url.split('open?id=')[1] else url.split('open?id=')[1]
        elif 'file/d/' in url:
            file_id = url.split('file/d/')[1].split('/')[0] if '/' in url.split('file/d/')[1] else url.split('file/d/')[1]
        elif 'id=' in url:
            file_id = url.split('id=')[1].split('&')[0] if '&' in url.split('id=')[1] else url.split('id=')[1]
        else:
            import re
            matches = re.findall(r'[-\w]{25,}', url)
            if matches:
                file_id = matches[0]

        if not file_id:
            print(f"警告: GoogleドライブのファイルIDを抽出できませんでした: {url}")
            return None

        # 試すURLのパターン
        url_patterns = [
            f"https://lh3.googleusercontent.com/d/{file_id}",
            f"https://drive.usercontent.google.com/download?id={file_id}&export=view",
            f"https://drive.google.com/uc?export=view&id={file_id}",
            f"https://drive.google.com/uc?id={file_id}",
            f"https://drive.google.com/thumbnail?id={file_id}&sz=w2000"
        ]

        headers_list = [
            {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
                'Referer': 'https://drive.google.com/'
            },
            {
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.1 Safari/605.1.15',
                'Accept': 'image/webp,image/png,image/svg+xml,image/*;q=0.8',
                'Sec-Fetch-Site': 'cross-site',
                'Sec-Fetch-Mode': 'no-cors',
                'Sec-Fetch-Dest': 'image'
            }
        ]

        import urllib.request
        from urllib.error import HTTPError, URLError
        import time
        import random

        for retry in range(max_retries):
            # 各URLパターンを試す
            for url_pattern in url_patterns:
                # 各ヘッダーセットを試す
                for headers in headers_list:
                    try:
                        req = urllib.request.Request(url_pattern, headers=headers)
                        with urllib.request.urlopen(req, timeout=15) as response:
                            image_data = response.read()

                            if image_data and len(image_data) > 100:  # 最小サイズチェック
                                print(f"成功: {url_pattern} (試行 {retry+1}/{max_retries})")
                                return image_data
                            else:
                                print(f"画像データが不十分: {url_pattern}")
                    except (HTTPError, URLError) as e:
                        print(f"失敗 ({e}): {url_pattern}")
                    except Exception as e:
                        print(f"例外: {e} for {url_pattern}")

                    # アクセス制限回避のための短い待機
                    time.sleep(random.uniform(0.5, 1.5))

            # すべてのパターンが失敗した場合、次の再試行前に少し長く待機
            if retry < max_retries - 1:
                time.sleep(random.uniform(1.0, 3.0))

        print(f"すべての試行が失敗しました: {original_url}")
        return None

    def load_image_from_url(self, url):
        """URLから画像をロードする（複数の方法を試す改良版）"""
        # キャッシュにあれば使用
        if url in self.image_cache:
            return self.image_cache[url]

        try:
            if not url or url.strip() == '':
                print("空のURLが指定されました")
                return QPixmap()

            self.statusBar().showMessage(f"画像読み込み中: {url}")

            # 複数の方法で画像データの取得を試みる
            image_data = self.fetch_image_with_retry(url)

            if not image_data:
                print(f"画像を取得できませんでした: {url}")
                return QPixmap()

            # ----- QImageでロード -----
            from PyQt5.QtGui import QImage
            from PyQt5.QtCore import QByteArray

            # バイトデータをQByteArrayに変換
            byte_array = QByteArray(image_data)

            # QImageを作成
            image = QImage()
            loaded = image.loadFromData(byte_array)

            if loaded and not image.isNull():
                # 正常にロードできた場合
                image = image.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                pixmap = QPixmap.fromImage(image)
                self.image_cache[url] = pixmap
                self.statusBar().showMessage(f"画像読み込み完了: {url}")
                return pixmap

            # ----- 代替手段: 一時ファイル経由 -----
            import tempfile
            import os

            # 一時ファイルに保存して読み込み
            temp_file = tempfile.NamedTemporaryFile(delete=False)
            temp_path = temp_file.name
            temp_file.write(image_data)
            temp_file.close()

            # 拡張子の追加（メディアタイプを検出）
            from PIL import Image
            from io import BytesIO

            try:
                pil_image = Image.open(BytesIO(image_data))
                img_format = pil_image.format

                # 適切な拡張子を持つ新しい一時ファイル
                new_temp_path = temp_path + '.' + (img_format.lower() if img_format else 'jpg')
                os.rename(temp_path, new_temp_path)
                temp_path = new_temp_path

                pil_image.close()
            except Exception as pil_err:
                print(f"PILでの画像形式検出エラー: {pil_err}")

            # QPixmapで読み込み
            pixmap = QPixmap(temp_path)

            # 一時ファイルを削除
            try:
                os.unlink(temp_path)
            except Exception as e:
                print(f"一時ファイル削除エラー: {e}")

            if not pixmap.isNull():
                # リサイズ
                pixmap = pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)

                # キャッシュに保存
                self.image_cache[url] = pixmap
                self.statusBar().showMessage(f"画像読み込み完了（代替手段）: {url}")
                return pixmap

            # ----- 代替手段2: PILを使用した変換 -----
            try:
                from PIL import Image
                from io import BytesIO
                import numpy as np
                from PyQt5.QtGui import QImage, qRgb

                # PILで画像を開く
                pil_image = Image.open(BytesIO(image_data))

                # RGBAに変換
                pil_image = pil_image.convert("RGBA")

                # PILからnumpy配列に変換
                img_array = np.array(pil_image)

                # numpy配列からQImageに変換
                height, width, channels = img_array.shape
                bytes_per_line = channels * width

                # RGBA形式のQImageを作成
                qimg = QImage(img_array.data, width, height, bytes_per_line, QImage.Format_RGBA8888)

                # QImageからQPixmapに変換
                pixmap = QPixmap.fromImage(qimg)

                if not pixmap.isNull():
                    # リサイズ
                    pixmap = pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)

                    # キャッシュに保存
                    self.image_cache[url] = pixmap
                    self.statusBar().showMessage(f"画像読み込み完了（PIL変換）: {url}")
                    return pixmap

            except Exception as pil_err:
                print(f"PIL変換エラー: {pil_err}")

            # すべての方法が失敗した場合
            print(f"すべての画像読み込み方法が失敗: {url}")
            return QPixmap()

        except Exception as e:
            print(f"画像読み込み全体エラー: {str(e)} for URL: {url}")
            self.statusBar().showMessage(f"画像読み込みエラー: {url}")
            return QPixmap()  # 空の QPixmap を返す

    def create_improved_profile_card(self, item, style):
        """一人分のプロフィールカードを作成（複数のアクセス方法を試す）"""
        # プロフィール情報を整理
        texts = []
        texts.append(f"<b>{item['parent_name']}</b>")
        texts.append(f"お子様: {item['child_name']}")
        if item['child_phrase']:
            texts.append(f"高等部二年生保護者の皆様へご挨拶: {item['child_phrase']}")
        if item['parent_phrase']:
            texts.append(f"お住まいの地域: {item['parent_phrase']}")



        # 全てのテキストを1つのパラグラフにまとめる
        content = Paragraph("<br/>".join(texts), style)

        # 画像の処理
        img = None
        temp_file_path = None

        if item['photo_url'] and item['photo_url'].strip():
            try:
                # 複数の方法で画像データを取得
                image_data = self.fetch_image_with_retry(item['photo_url'])

                if not image_data:
                    raise ValueError(f"画像データを取得できませんでした: {item['photo_url']}")

                # 画像データの検証とファイル拡張子の決定
                import tempfile
                from PIL import Image
                from io import BytesIO

                try:
                    pil_image = Image.open(BytesIO(image_data))
                    img_format = pil_image.format
                    img_ext = img_format.lower() if img_format else 'jpg'
                    pil_image.close()
                except Exception as img_err:
                    print(f"画像検証エラー: {img_err}")
                    img_ext = 'jpg'  # デフォルト

                # 一時ファイルを作成
                with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{img_ext}') as temp_file:
                    temp_file.write(image_data)
                    temp_file_path = temp_file.name

                # レポートラボの画像オブジェクト作成を試みる
                try:
                    img = ReportLabImage(temp_file_path, width=3 * cm, height=3 * cm)
                except Exception as rl_err:
                    print(f"ReportLab画像読み込みエラー: {rl_err}")

                    # 代替手段: 画像形式を変換して再試行
                    try:
                        # PILを使ってJPEGに変換
                        pil_image = Image.open(temp_file_path)
                        pil_image = pil_image.convert('RGB')  # 透明度を排除

                        # 新しい一時ファイルに保存
                        new_temp_path = temp_file_path.rsplit('.', 1)[0] + '.jpg'
                        pil_image.save(new_temp_path, "JPEG", quality=90)
                        pil_image.close()

                        # 古い一時ファイルを削除
                        import os
                        try:
                            os.unlink(temp_file_path)
                        except Exception as del_err:
                            print(f"古い一時ファイル削除エラー: {del_err}")

                        # 新しいパスを設定
                        temp_file_path = new_temp_path

                        # 再度ReportLabの画像オブジェクト作成を試みる
                        img = ReportLabImage(temp_file_path, width=3 * cm, height=3 * cm)
                    except Exception as retry_err:
                        print(f"画像変換再試行エラー: {retry_err}")
                        img = None

                        # 一時ファイルのクリーンアップ
                        if temp_file_path and os.path.exists(temp_file_path):
                            try:
                                os.unlink(temp_file_path)
                                temp_file_path = None
                            except Exception as cleanup_err:
                                print(f"一時ファイル削除エラー: {cleanup_err}")

            except Exception as e:
                print(f"プロフィールカード画像処理エラー: {e}")
                img = None

                # 一時ファイルのクリーンアップ
                if temp_file_path and os.path.exists(temp_file_path):
                    try:
                        import os
                        os.unlink(temp_file_path)
                        temp_file_path = None
                    except Exception as cleanup_err:
                        print(f"エラー後のクリーンアップ失敗: {cleanup_err}")

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
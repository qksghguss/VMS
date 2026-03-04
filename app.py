import os
import sys
import traceback
import configparser

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QMessageBox, QTableWidget,
    QTableWidgetItem, QListWidget, QListWidgetItem, QAbstractItemView,
    QComboBox, QDateEdit, QSpinBox, QSplitter, QTextEdit, QCheckBox,
    QHeaderView
)

import pandas as pd
import matplotlib
matplotlib.use("QtAgg")
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib import font_manager, rcParams

from core import load_ini, AppConfig, SourceSpec, FIELDS, aggregate, export_excel

CONFIG_FILE = "config.ini"


def setup_matplotlib_font():
    preferred = ["Malgun Gothic", "AppleGothic", "NanumGothic", "DejaVu Sans"]
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in preferred:
        if name in installed:
            rcParams["font.family"] = name
            break
    rcParams["axes.unicode_minus"] = False


class WorkerThread(QThread):
    done = Signal(object, object, object, object, str)
    error = Signal(str)

    def __init__(self, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.cfg = cfg

    def run(self):
        try:
            expanded, daily, summary, detail, sort_col = aggregate(self.cfg)
            self.done.emit(expanded, daily, summary, detail, sort_col)
        except Exception:
            self.error.emit(traceback.format_exc())


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None):
        fig = Figure()
        super().__init__(fig)
        self.setParent(parent)
        self.ax = fig.add_subplot(111)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("작업일보 실적집계 (VMS)")
        self.resize(1920, 1080)

        self.cfg_path = os.path.join(os.getcwd(), CONFIG_FILE)
        self.cfg = None

        self.expanded = None
        self.daily = None
        self.summary = None
        self.detail = None
        self.sort_col = None

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self._build_settings_tab()
        self._build_dashboard_tab()
        self._apply_style()

        self.load_config()

    def _apply_style(self):
        self.setStyleSheet(
            """
            QWidget { font-size: 13px; }
            QTabBar::tab { min-width: 180px; min-height: 34px; padding: 6px 12px; }
            QPushButton { min-height: 30px; padding: 2px 10px; }
            QLineEdit, QComboBox, QDateEdit, QSpinBox { min-height: 28px; }
            QTableWidget { gridline-color: #d9d9d9; }
            QLabel#statCard {
                background: #f5f7fb;
                border: 1px solid #dce3ef;
                border-radius: 8px;
                padding: 12px;
                font-weight: 600;
            }
            """
        )

    def _build_settings_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        topbar = QHBoxLayout()
        self.btn_load = QPushButton("설정 불러오기")
        self.btn_save = QPushButton("설정 저장")
        self.btn_run = QPushButton("집계 실행")
        self.btn_export = QPushButton("결과 엑셀로 내보내기")
        self.btn_export.setEnabled(False)

        topbar.addWidget(QLabel("설정파일:"))
        self.lbl_cfg = QLineEdit()
        self.lbl_cfg.setReadOnly(True)
        topbar.addWidget(self.lbl_cfg, 1)
        topbar.addWidget(self.btn_load)
        topbar.addWidget(self.btn_save)
        topbar.addWidget(self.btn_run)
        topbar.addWidget(self.btn_export)
        layout.addLayout(topbar)

        splitter = QSplitter(Qt.Vertical)

        src_box = QWidget()
        src_layout = QVBoxLayout(src_box)
        src_header = QHBoxLayout()
        src_header.addWidget(QLabel("원본 작업일보 (Sources)"))
        self.btn_add_source = QPushButton("+ 추가")
        self.btn_del_source = QPushButton("- 삭제")
        src_header.addStretch(1)
        src_header.addWidget(self.btn_add_source)
        src_header.addWidget(self.btn_del_source)
        src_layout.addLayout(src_header)

        self.tbl_sources = QTableWidget(0, 6)
        self.tbl_sources.setHorizontalHeaderLabels(["사용", "작업명", "파일경로", "", "시트", "헤더행"])
        self.tbl_sources.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_sources.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self.tbl_sources.verticalHeader().setVisible(False)
        hh = self.tbl_sources.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Interactive)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.Interactive)
        hh.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tbl_sources.setColumnWidth(1, 180)
        self.tbl_sources.setColumnWidth(4, 180)
        src_layout.addWidget(self.tbl_sources)

        bottom = QWidget()
        bottom_layout = QHBoxLayout(bottom)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        map_box = QWidget()
        map_layout = QVBoxLayout(map_box)
        map_layout.addWidget(QLabel("열 매핑 (스크롤 없이 전체 표시)"))
        self.tbl_map = QTableWidget(len(FIELDS), 2)
        self.tbl_map.setHorizontalHeaderLabels(["필드", "원본컬럼"])
        self.tbl_map.verticalHeader().setVisible(False)
        mh = self.tbl_map.horizontalHeader()
        mh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        mh.setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl_map.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_map.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tbl_map.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        for i, f in enumerate(FIELDS):
            it = QTableWidgetItem(f)
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.tbl_map.setItem(i, 0, it)
            self.tbl_map.setItem(i, 1, QTableWidgetItem(f))
        self.tbl_map.resizeRowsToContents()
        self.tbl_map.setFixedHeight(self.tbl_map.horizontalHeader().height() + self.tbl_map.rowCount() * (self.tbl_map.verticalHeader().defaultSectionSize() + 1) + 4)
        map_layout.addWidget(self.tbl_map)

        ex_box = QWidget()
        ex_layout = QVBoxLayout(ex_box)
        ex_layout.addWidget(QLabel("제외자 (집계 제외, N빵 계산 포함)"))
        self.lst_ex = QListWidget()
        self.lst_ex.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        ex_layout.addWidget(self.lst_ex, 1)
        ex_btns = QHBoxLayout()
        self.ex_input = QLineEdit()
        self.ex_input.setPlaceholderText("이름 입력 후 Enter")
        self.btn_ex_add = QPushButton("추가")
        self.btn_ex_del = QPushButton("삭제")
        ex_btns.addWidget(self.ex_input, 1)
        ex_btns.addWidget(self.btn_ex_add)
        ex_btns.addWidget(self.btn_ex_del)
        ex_layout.addLayout(ex_btns)

        po_box = QWidget()
        po_layout = QVBoxLayout(po_box)
        po_layout.addWidget(QLabel("기간/출력/랭킹"))

        row0 = QHBoxLayout()
        self.chk_period = QCheckBox("기간필터 사용")
        row0.addWidget(self.chk_period)
        row0.addStretch(1)
        po_layout.addLayout(row0)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("시작일"))
        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        self.dt_start.setDisplayFormat("yyyy-MM-dd")
        row1.addWidget(self.dt_start)
        row1.addWidget(QLabel("종료일"))
        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)
        self.dt_end.setDisplayFormat("yyyy-MM-dd")
        row1.addWidget(self.dt_end)
        po_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("집계단위"))
        self.cmb_gran = QComboBox()
        self.cmb_gran.addItems(["daily", "weekly", "monthly"])
        row2.addWidget(self.cmb_gran)
        row2.addWidget(QLabel("랭킹기준"))
        self.cmb_rank = QComboBox()
        self.cmb_rank.addItems(["pages", "lot", "l1", "l2"])
        row2.addWidget(self.cmb_rank)
        po_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("반올림 자리수"))
        self.sp_dec = QSpinBox()
        self.sp_dec.setRange(0, 6)
        self.sp_dec.setValue(3)
        row3.addWidget(self.sp_dec)
        row3.addWidget(QLabel("결과파일명"))
        self.ed_out = QLineEdit("result.xlsx")
        row3.addWidget(self.ed_out, 1)
        po_layout.addLayout(row3)

        row4 = QHBoxLayout()
        self.btn_close = QPushButton("닫기")
        self.btn_exit = QPushButton("종료")
        row4.addStretch(1)
        row4.addWidget(self.btn_close)
        row4.addWidget(self.btn_exit)
        po_layout.addLayout(row4)
        po_layout.addStretch(1)

        bottom_layout.addWidget(map_box, 5)
        bottom_layout.addWidget(ex_box, 2)
        bottom_layout.addWidget(po_box, 3)

        splitter.addWidget(src_box)
        splitter.addWidget(bottom)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter, 1)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(130)
        layout.addWidget(QLabel("로그"))
        layout.addWidget(self.log)

        self.tabs.addTab(w, "설정/실행")

        self.btn_load.clicked.connect(self.load_config)
        self.btn_save.clicked.connect(lambda: self.save_config(notify=True))
        self.btn_run.clicked.connect(self.run_aggregate)
        self.btn_export.clicked.connect(self.export_result)
        self.btn_add_source.clicked.connect(self.add_source_row)
        self.btn_del_source.clicked.connect(self.del_source_row)
        self.btn_ex_add.clicked.connect(self.add_exclude)
        self.btn_ex_del.clicked.connect(self.del_exclude)
        self.ex_input.returnPressed.connect(self.add_exclude)
        self.btn_close.clicked.connect(self.close)
        self.btn_exit.clicked.connect(QApplication.instance().quit)

    def _build_dashboard_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)

        top = QHBoxLayout()
        self.lbl_status = QLabel("대기 중")
        self.lbl_status.setStyleSheet("font-weight:600;")
        top.addWidget(self.lbl_status)
        self.stat_total = QLabel("총 작업자: -")
        self.stat_total.setObjectName("statCard")
        self.stat_top = QLabel("TOP 작업자: -")
        self.stat_top.setObjectName("statCard")
        self.stat_pages = QLabel("총 매수: -")
        self.stat_pages.setObjectName("statCard")
        top.addWidget(self.stat_total)
        top.addWidget(self.stat_top)
        top.addWidget(self.stat_pages)
        top.addStretch(1)
        layout.addLayout(top)

        split = QSplitter(Qt.Horizontal)

        left = QWidget()
        lyt = QVBoxLayout(left)
        lyt.addWidget(QLabel("최종집계_요약 (작업자 x 작업일보, 상위 80)"))
        self.tbl_sum = QTableWidget(0, 7)
        self.tbl_sum.setHorizontalHeaderLabels(["순위", "작업자", "작업일보", "LOT", "매수", "Defect_L1", "Defect_L2"])
        self.tbl_sum.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_sum.setAlternatingRowColors(True)
        sh = self.tbl_sum.horizontalHeader()
        sh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        sh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        sh.setSectionResizeMode(2, QHeaderView.Stretch)
        for i in (3, 4, 5, 6):
            sh.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        lyt.addWidget(self.tbl_sum, 1)

        right = QWidget()
        rlyt = QVBoxLayout(right)
        rlyt.addWidget(QLabel("그래프"))
        self.canvas_top = MplCanvas()
        self.canvas_daily = MplCanvas()
        rlyt.addWidget(self.canvas_top, 1)
        rlyt.addWidget(self.canvas_daily, 1)

        split.addWidget(left)
        split.addWidget(right)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 4)

        layout.addWidget(split, 1)
        self.tabs.addTab(w, "대시보드")

    def _load_sheet_names(self, path: str):
        if not path or not os.path.exists(path):
            return ["Sheet1"]
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
            return xls.sheet_names or ["Sheet1"]
        except Exception:
            return ["Sheet1"]

    def _set_sheet_combo(self, row: int, path: str, selected_sheet: str = "Sheet1"):
        combo = QComboBox()
        names = self._load_sheet_names(path)
        combo.addItems(names)
        combo.setCurrentText(selected_sheet if selected_sheet in names else names[0])
        self.tbl_sources.setCellWidget(row, 4, combo)

    def _insert_source_row(self, src: SourceSpec):
        r = self.tbl_sources.rowCount()
        self.tbl_sources.insertRow(r)

        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        chk.setCheckState(Qt.Checked if src.use else Qt.Unchecked)
        self.tbl_sources.setItem(r, 0, chk)
        self.tbl_sources.setItem(r, 1, QTableWidgetItem(src.name))
        self.tbl_sources.setItem(r, 2, QTableWidgetItem(src.path))

        btn = QPushButton("...")
        btn.setMaximumWidth(36)
        btn.clicked.connect(lambda _, row=r: self.pick_source_file(row))
        self.tbl_sources.setCellWidget(r, 3, btn)

        self._set_sheet_combo(r, src.path, src.sheet)
        self.tbl_sources.setItem(r, 5, QTableWidgetItem(str(src.header_row)))

    def load_config(self):
        if not os.path.exists(self.cfg_path):
            QMessageBox.warning(self, "설정 없음", f"{self.cfg_path} 파일이 없습니다.")
            return
        self.lbl_cfg.setText(self.cfg_path)
        self.cfg = load_ini(self.cfg_path)
        self.log.append(f"설정 로드: {self.cfg_path}")

        self.tbl_sources.setRowCount(0)
        for s in self.cfg.sources:
            self._insert_source_row(s)

        for i, f in enumerate(FIELDS):
            self.tbl_map.item(i, 1).setText(self.cfg.mapping.get(f, f))

        self.lst_ex.clear()
        for n in sorted(self.cfg.exclude):
            self.lst_ex.addItem(QListWidgetItem(n))

        use_period = self.cfg.start_date is not None or self.cfg.end_date is not None
        self.chk_period.setChecked(use_period)
        if self.cfg.start_date:
            self.dt_start.setDate(self.cfg.start_date)
        if self.cfg.end_date:
            self.dt_end.setDate(self.cfg.end_date)
        self.cmb_gran.setCurrentText(self.cfg.granularity if self.cfg.granularity in ["daily", "weekly", "monthly"] else "monthly")
        self.cmb_rank.setCurrentText(self.cfg.rank_metric if self.cfg.rank_metric in ["pages", "lot", "l1", "l2"] else "pages")
        self.sp_dec.setValue(self.cfg.round_decimals)
        self.ed_out.setText(self.cfg.result_xlsx)

    def _gather_from_ui(self) -> AppConfig:
        sources = []
        for r in range(self.tbl_sources.rowCount()):
            use_item = self.tbl_sources.item(r, 0)
            use = use_item.checkState() == Qt.Checked if use_item else False
            name = self.tbl_sources.item(r, 1).text().strip() if self.tbl_sources.item(r, 1) else f"source_{r+1}"
            path = self.tbl_sources.item(r, 2).text().strip() if self.tbl_sources.item(r, 2) else ""
            sheet_combo = self.tbl_sources.cellWidget(r, 4)
            sheet = sheet_combo.currentText().strip() if isinstance(sheet_combo, QComboBox) else "Sheet1"
            header_item = self.tbl_sources.item(r, 5)
            header_row = int(header_item.text().strip() or "1") if header_item else 1
            if header_row <= 0:
                raise ValueError(f"헤더행은 1 이상이어야 합니다. (행: {r+1})")
            sources.append(SourceSpec(use=use, name=name, path=path, sheet=sheet, header_row=header_row))

        mapping = {}
        for i, f in enumerate(FIELDS):
            mapping[f] = self.tbl_map.item(i, 1).text().strip() if self.tbl_map.item(i, 1) else f

        exclude = {self.lst_ex.item(i).text().strip() for i in range(self.lst_ex.count())}

        start_date = self.dt_start.date().toPython() if self.chk_period.isChecked() else None
        end_date = self.dt_end.date().toPython() if self.chk_period.isChecked() else None
        granularity = self.cmb_gran.currentText()
        rank_metric = self.cmb_rank.currentText()
        round_decimals = int(self.sp_dec.value())
        result_xlsx = self.ed_out.text().strip() or "result.xlsx"

        return AppConfig(
            sources=sources,
            mapping=mapping,
            exclude={n.replace(" ", "") for n in exclude if n.strip()},
            start_date=start_date,
            end_date=end_date,
            granularity=granularity,
            rank_metric=rank_metric,
            round_decimals=round_decimals,
            result_xlsx=result_xlsx,
        )

    def save_config(self, notify=False):
        cfg = self._gather_from_ui()

        cp = configparser.ConfigParser()
        cp["app"] = {
            "window_mode": "fullscreen",
            "round_decimals": str(cfg.round_decimals),
            "rank_metric": cfg.rank_metric,
        }
        cp["period"] = {
            "start_date": cfg.start_date.isoformat() if cfg.start_date else "",
            "end_date": cfg.end_date.isoformat() if cfg.end_date else "",
            "granularity": cfg.granularity,
        }
        cp["output"] = {"result_xlsx": cfg.result_xlsx}

        for i, s in enumerate(cfg.sources, start=1):
            cp[f"source_{i}"] = {
                "use": "1" if s.use else "0",
                "name": s.name,
                "path": s.path,
                "sheet": s.sheet,
                "header_row": str(s.header_row),
            }

        cp["columns"] = {k: v for k, v in cfg.mapping.items()}
        cp["exclude"] = {"names": ",".join(sorted(cfg.exclude))}

        with open(self.cfg_path, "w", encoding="utf-8") as f:
            cp.write(f)

        self.log.append(f"설정 저장: {self.cfg_path}")
        if notify:
            QMessageBox.information(self, "저장 완료", "config.ini 저장 완료")

    def add_exclude(self):
        name = self.ex_input.text().strip()
        if not name:
            return
        norm = name.replace(" ", "")
        for i in range(self.lst_ex.count()):
            if self.lst_ex.item(i).text() == norm:
                self.ex_input.clear()
                return
        self.lst_ex.addItem(QListWidgetItem(norm))
        self.ex_input.clear()

    def del_exclude(self):
        for item in self.lst_ex.selectedItems():
            self.lst_ex.takeItem(self.lst_ex.row(item))

    def add_source_row(self):
        self._insert_source_row(SourceSpec(use=True, name=f"작업일보{self.tbl_sources.rowCount()+1}", path="", sheet="Sheet1", header_row=1))

    def pick_source_file(self, row=None):
        if row is None:
            rows = sorted({i.row() for i in self.tbl_sources.selectedIndexes()})
            if not rows:
                QMessageBox.information(self, "선택 필요", "파일 경로를 채울 소스 행을 먼저 선택하세요.")
                return
            row = rows[0]

        path, _ = QFileDialog.getOpenFileName(self, "작업일보 파일 선택", "", "Excel (*.xlsx *.xlsm *.xls)")
        if path:
            self.tbl_sources.setItem(row, 2, QTableWidgetItem(path))
            current_sheet = "Sheet1"
            combo = self.tbl_sources.cellWidget(row, 4)
            if isinstance(combo, QComboBox):
                current_sheet = combo.currentText()
            self._set_sheet_combo(row, path, current_sheet)

    def del_source_row(self):
        rows = sorted({i.row() for i in self.tbl_sources.selectedIndexes()}, reverse=True)
        for r in rows:
            self.tbl_sources.removeRow(r)

    def run_aggregate(self):
        try:
            cfg = self._gather_from_ui()
        except Exception as e:
            QMessageBox.warning(self, "설정 오류", str(e))
            return

        if not any(s.use for s in cfg.sources):
            QMessageBox.warning(self, "설정 오류", "사용할 소스를 1개 이상 체크하세요.")
            return

        self.save_config(notify=False)

        self.log.append("집계 시작...")
        self.lbl_status.setText("집계 중...")
        self.btn_export.setEnabled(False)

        self.worker = WorkerThread(cfg)
        self.worker.done.connect(self.on_done)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_done(self, expanded, daily, summary, detail, sort_col):
        self.expanded, self.daily, self.summary, self.detail, self.sort_col = expanded, daily, summary, detail, sort_col
        self.lbl_status.setText(f"완료 (랭킹 기준: {sort_col})")
        self.log.append(f"완료: 분해 {len(expanded)}행 / 집계대상 {len(expanded[expanded['제외여부']==False])}행")
        self.btn_export.setEnabled(True)
        self.update_dashboard()
        self.tabs.setCurrentIndex(1)

    def on_error(self, tb):
        self.lbl_status.setText("에러")
        self.log.append(tb)
        QMessageBox.critical(self, "에러", tb)

    def export_result(self):
        if self.daily is None:
            return
        default_name = self.ed_out.text().strip() or "result.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(self, "결과 엑셀 저장", default_name, "Excel (*.xlsx)")
        if not out_path:
            return
        export_excel(out_path, self.daily, self.summary, self.detail)
        self.log.append(f"엑셀 저장: {out_path}")
        QMessageBox.information(self, "저장 완료", "엑셀 저장 완료")

    def update_dashboard(self):
        df = self.summary.head(80).copy()
        self.tbl_sum.setRowCount(len(df))
        self.tbl_sum.setColumnCount(7)
        for r, row in enumerate(df.itertuples(index=False)):
            for c, val in enumerate(row[:7]):
                self.tbl_sum.setItem(r, c, QTableWidgetItem(str(val)))

        workers = self.summary["작업자"].nunique() if len(self.summary) else 0
        self.stat_total.setText(f"총 작업자: {workers}")
        if len(self.summary):
            top_row = self.summary.iloc[0]
            self.stat_top.setText(f"TOP 작업자: {top_row['작업자']} ({top_row['작업일보']})")
            self.stat_pages.setText(f"총 매수: {self.summary['매수'].sum():,.2f}")
        else:
            self.stat_top.setText("TOP 작업자: -")
            self.stat_pages.setText("총 매수: -")

        self.canvas_top.ax.clear()
        top10 = self.summary.head(10)
        if len(top10):
            x = [f"{r['작업자']}\n({r['작업일보']})" for _, r in top10.iterrows()]
            y = list(top10["매수"])
            self.canvas_top.ax.bar(x, y)
            self.canvas_top.ax.set_title("Top10 - 매수 (작업자 x 작업일보)")
            self.canvas_top.ax.tick_params(axis="x", rotation=20)
        self.canvas_top.draw()

        self.canvas_daily.ax.clear()
        daily_total = self.daily.groupby("기간", as_index=False)[["매수"]].sum().sort_values("기간")
        if len(daily_total):
            self.canvas_daily.ax.plot(daily_total["기간"], daily_total["매수"], marker="o")
            self.canvas_daily.ax.set_title(f"기간별 총 매수(단위: {self.cmb_gran.currentText()})")
            self.canvas_daily.ax.tick_params(axis="x", rotation=30)
        self.canvas_daily.draw()

    def closeEvent(self, event):
        try:
            self.save_config(notify=False)
        except Exception:
            pass
        super().closeEvent(event)


def main():
    setup_matplotlib_font()
    app = QApplication(sys.argv)
    win = MainWindow()
    win.showMaximized()
    win.showFullScreen()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

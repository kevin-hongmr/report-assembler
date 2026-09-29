# -*- coding: utf-8 -*-
"""主窗口：文件列表 + 封面字段 + 选项 + 一键汇编 + 预览。"""
import functools
import json
import logging
import os
import sys
import platform
import tempfile
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QPushButton, QLineEdit, QCheckBox, QSpinBox, QDoubleSpinBox,
    QLabel, QFileDialog, QScrollArea, QMessageBox, QGroupBox, QAbstractItemView,
    QSizePolicy, QComboBox)
from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal, QMimeData
from PyQt6.QtGui import QPixmap, QImage, QDrag, QIcon

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from app.core import pipeline, converter, fonts as fonts_mod
from app.core import assembler as assembler_mod
from app.core import logger as logger_mod

log = logging.getLogger(__name__)


def _guard_slot(fn):
    """槽函数兜底：把未捕获异常就地转成日志 + 提示，不让它拖垮进程。

    PyQt6（>=5.5 同源行为）在槽函数抛出未捕获异常时会调用 qFatal() 直接
    abort 进程：界面瞬间消失（即用户说的"闪退"），日志里除了 excepthook 记的
    那一行之外没有任何线索，工作内容也会丢。这里统一兜底。
    """
    @functools.wraps(fn)
    def wrapper(self, *args, **kwargs):
        try:
            return fn(self, *args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            log.exception("操作失败（已拦截，程序继续运行）：%s",
                          getattr(fn, "__name__", fn))
            try:
                QMessageBox.critical(
                    self, "操作失败",
                    "%s\n\n程序会继续运行。详细信息见日志：\n%s"
                    % (exc, logger_mod.LOG_FILE or "（无日志文件）"))
            except Exception:  # noqa: BLE001
                pass
            return None
    return wrapper

# 中文字号 -> 磅值（全档位；界面字号下拉框用，默认三号）
CN_FONT_SIZES = [
    ("初号", 42.0), ("小初", 36.0), ("一号", 26.0), ("小一", 24.0),
    ("二号", 22.0), ("小二", 18.0), ("三号", 16.0), ("小三", 15.0),
    ("四号", 14.0), ("小四", 12.0), ("五号", 10.5), ("小五", 9.0),
]
# 界面可调格式的四个对象（顺序即界面排列顺序）
FMT_KEYS = (("toc", "目录内容"), ("body", "正文"),
            ("h1", "一级标题"), ("h2", "二级标题"))


class DocTableWidget(QTableWidget):
    """文档列表表格：拖拽排序时完整保留单元格数据（含文件路径 UserRole）。

    覆盖默认的 InternalMove 模型移动（其在拖拽时仅编码文本，丢失自定义
    角色数据，并会触发 dataChanged() 无效索引）。改为在 dropEvent 中手动
    搬移整行 QTableWidgetItem 对象，保证文件路径等数据不丢失。
    """

    def __init__(self, parent=None):
        super().__init__(0, 3, parent)
        self.setHorizontalHeaderLabels(["文件", "单位", "标题"])
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragDropOverwriteMode(False)
        self.setDropIndicatorShown(True)
        self.horizontalHeader().setStretchLastSection(True)

    def startDrag(self, supportedActions):
        # 记录拖拽起始行，供 dropEvent 使用
        self._drag_source_rows = sorted({i.row() for i in self.selectedIndexes()})
        # 手动执行拖拽，绝不调用 super().startDrag()：
        # QAbstractItemView::startDrag 默认实现里，drag->exec() 返回 Qt::MoveAction
        # 后会调用 clearOrRemove() 删除当前选中行；而 dropEvent 已用 move_rows 把
        # 整行搬移到新位置并 selectRow 选中，随后 super 返回 MoveAction 会把新选中行
        # 误删（用户实测"拖动后文件消失"）。这里手动 exec 后不删除任何行，行数据已
        # 在 dropEvent 中完整搬移。
        drag = QDrag(self)
        mime = self.model().mimeData(self.selectedIndexes())
        if mime is None:
            mime = QMimeData()
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction, Qt.DropAction.MoveAction)

    def dropEvent(self, event):
        if event.source() is not self:
            event.ignore()
            return
        rows = sorted(set(getattr(self, "_drag_source_rows", None) or
                          {i.row() for i in self.selectedIndexes()}))
        if not rows:
            event.ignore()
            return

        pos = event.position().toPoint()
        idx = self.indexAt(pos)
        if idx.isValid():
            r = idx.row()
            rect = self.visualRect(idx)
            after = pos.y() >= rect.center().y()
            target = r + (1 if after else 0)
        else:
            target = self.rowCount()

        first = self.move_rows(rows, target)
        # 真正的防误删在 startDrag（手动 exec、不调 super 的 clearOrRemove），
        # 这里正常以 MoveAction 收尾即可，行数据已在 move_rows 中完整搬移。
        event.setDropAction(Qt.DropAction.MoveAction)
        event.accept()
        log.info("拖拽调整顺序：%s -> 位置 %d", rows, first)

    def move_rows(self, rows, target):
        """把若干行移动到 target 插入点（0..rowCount），保留单元格 UserRole 数据。

        返回移动后首行所在的新行号。供 dropEvent 与测试复用。
        """
        rows = sorted(set(rows))
        if not rows:
            return self.currentRow()
        target = max(0, min(target, self.rowCount()))

        # 取出整行数据（保留 QTableWidgetItem 对象及其 UserRole）
        payload = []
        for r in rows:
            payload.append([self.takeItem(r, c) for c in range(self.columnCount())])

        # 从大到小移除源行，避免索引前移影响
        for r in sorted(rows, reverse=True):
            self.removeRow(r)

        # 计算调整后的插入位置（源行中位于目标之前的会随删除前移）
        removed_before = sum(1 for r in rows if r < target)
        insert_at = max(0, min(target - removed_before, self.rowCount()))

        for row_items in payload:
            self.insertRow(insert_at)
            for c, it in enumerate(row_items):
                self.setItem(insert_at, c, it)
            insert_at += 1

        # 恢复选中到新位置
        self.clearSelection()
        first = max(0, min(target - removed_before, self.rowCount() - 1))
        for k in range(len(rows)):
            self.selectRow(first + k)
        self.setCurrentCell(first, 0)
        return first


def today_cn():
    d = datetime.now()
    return f"{d.year}年{d.month}月"


class Worker(QThread):
    """后台执行汇编/预览/文档识别，避免界面卡死。"""
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            self.finished.emit(self.fn(*self.args, **self.kwargs))
        except Exception as e:
            log.exception("后台任务执行失败：%s", e)
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文档汇编程序 · 政务版")
        self.resize(1100, 880)   # 高度加大，保证待汇编文档表格区与原来一样大
        self.work_dir = os.path.join(tempfile.gettempdir(), "doc_assembler_work")
        os.makedirs(self.work_dir, exist_ok=True)
        # 后台线程必须持有强引用直到真正结束（见 _run 注释），这里用列表统一管理
        self._workers = []
        self._env_box = None
        # 文件对话框复用单实例（懒创建，见 _file_dialog()）。原因是 fcitx 输入法插件
        # 在 Qt 6.7 + 反复创建/销毁 QFileDialog 时会触发 FcitxInputContextProxy 空指针
        # 连接（"Cannot connect (nullptr)::availabilityChanged..."）导致 SIGSEGV：
        # 实测「打开文档对话框」后再开「保存对话框」的第二次就崩溃（退出码 139）。
        # 只创建一个实例、动态切换 FileMode/AcceptMode 即可彻底规避运行期崩溃。
        self._file_dlg = None
        self._build_ui()
        self._load_fmt_config()
        # 环境自检延后到事件循环里执行（此时窗口已显示）：自检要跑一次真实的
        # LibreOffice 无头转换，最快也要约 1 秒，放到 __init__ 里会拖慢启动；
        # 更关键的是它以前会在 __init__ 中同步弹出模态对话框，主窗口必须等
        # 对话框被点掉才出现——无人值守/无头场景就此永久卡死（日志停在
        # "检测到转换引擎"之后，看起来像"一启动就卡死"）。
        QTimer.singleShot(0, self._start_env_check)
        log.info("主窗口初始化完成")

    # ---------------- UI ----------------
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)

        # 左侧：控制
        left = QVBoxLayout()
        layout.addLayout(left, 1)

        # 文件列表
        fbox = QGroupBox("待汇编文档（拖拽可排序；单位/标题可双击修改）")
        fv = QVBoxLayout(fbox)
        self.table = DocTableWidget()
        self.table.setMinimumHeight(220)   # 保证表格区域足够大，方便拖拽调整
        fv.addWidget(self.table)
        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("添加文件")
        self.btn_del = QPushButton("移除")
        self.btn_up = QPushButton("上移")
        self.btn_down = QPushButton("下移")
        btn_row.addWidget(self.btn_add); btn_row.addWidget(self.btn_del)
        btn_row.addWidget(self.btn_up); btn_row.addWidget(self.btn_down)
        fv.addLayout(btn_row)
        left.addWidget(fbox, 1)   # 表格区占左侧主要垂直空间（保持原有大小）

        # 封面字段
        cbox = QGroupBox("封面与目录")
        cv = QVBoxLayout(cbox)
        lab_title = QLabel("汇编总标题（会议名称）：")
        lab_title.setToolTip("封面会自动在总标题下另起一段显示“汇报材料”两行大字标题")
        cv.addWidget(lab_title)
        self.ed_title = QLineEdit("全市服务业大会")
        cv.addWidget(self.ed_title)
        cv.addWidget(QLabel("日期（如 2026年8月）："))
        self.ed_date = QLineEdit(today_cn())
        cv.addWidget(self.ed_date)
        hsz = QHBoxLayout()
        hsz.addWidget(QLabel("封面标题字号(磅)："))
        self.spin_size = QSpinBox(); self.spin_size.setRange(16, 56); self.spin_size.setValue(22)
        hsz.addWidget(self.spin_size)
        cv.addLayout(hsz)

        # 选项
        obox = QGroupBox("选项")
        ov = QVBoxLayout(obox)
        self.chk_signoff = QCheckBox("自动清除原文档落款（单位+日期）")
        self.chk_signoff.setChecked(True)
        self.chk_page = QCheckBox("每份材料另起一页")
        self.chk_page.setChecked(True)
        ov.addWidget(self.chk_signoff); ov.addWidget(self.chk_page)

        # 封面与目录、选项：并排显示（省垂直空间，保证表格区域大小）
        cover_opts_row = QHBoxLayout()
        cover_opts_row.addWidget(cbox, 3)
        cover_opts_row.addWidget(obox, 2)
        left.addLayout(cover_opts_row)

        # 格式设置（字号/行距：目录内容、正文、一级标题、二级标题）
        fmt_box = QGroupBox("格式设置（仅对汇编后：目录内容、正文、一级/二级标题生效）")
        fmtv = QVBoxLayout(fmt_box)
        tip = QLabel("字号、行距不影响正文中的表格、图片（表格图片保留原格式）。")
        tip.setStyleSheet("color:#808080;")
        fmtv.addWidget(tip)
        self._fmt_widgets = {}
        for key, label in FMT_KEYS:
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            cb_size = QComboBox()
            for name, pt in CN_FONT_SIZES:
                cb_size.addItem(name, pt)
            cb_size.setCurrentIndex(cb_size.findData(16.0))   # 默认三号
            cb_mode = QComboBox()
            cb_mode.addItem("固定值", "exact")
            cb_mode.addItem("行距倍数", "multiple")
            cb_mode.setCurrentIndex(0)                          # 默认固定值
            spin = QDoubleSpinBox()
            spin.setDecimals(2)
            spin.setRange(6.0, 100.0)
            spin.setSingleStep(0.5)
            spin.setValue(28.5)                                 # 默认固定 28.5 磅
            spin.setSuffix(" 磅")
            cb_mode.currentIndexChanged.connect(
                lambda _i, m=cb_mode, s=spin: self._fmt_mode_changed(m, s))
            row.addWidget(cb_size)
            row.addWidget(cb_mode)
            row.addWidget(spin)
            fmtv.addLayout(row)
            self._fmt_widgets[key] = {"size": cb_size, "line_mode": cb_mode, "line_val": spin}
        btn_fmt_reset = QPushButton("恢复默认（全部：三号 + 固定值 28.5 磅）")
        btn_fmt_reset.clicked.connect(self._fmt_reset_default)
        fmtv.addWidget(btn_fmt_reset)
        left.addWidget(fmt_box)

        # 操作
        act = QHBoxLayout()
        self.btn_assemble = QPushButton("一键汇编并保存")
        self.btn_assemble.setStyleSheet("font-weight:bold; padding:6px;")
        self.btn_preview = QPushButton("刷新预览")
        act.addWidget(self.btn_assemble); act.addWidget(self.btn_preview)
        left.addLayout(act)
        self.status = QLabel("就绪")
        # 状态栏是"窄提示条"而不是内容区：允许换行 + 横向尺寸策略设为 Ignored，
        # 防止长文本（如含完整路径的环境自检消息）把左栏最小宽度撑到接近整窗宽，
        # 把右侧预览挤成一条窄边（2026-09-29 用户截图实测）。
        self.status.setWordWrap(True)
        sp = self.status.sizePolicy()
        sp.setHorizontalPolicy(QSizePolicy.Policy.Ignored)
        self.status.setSizePolicy(sp)
        left.addWidget(self.status)

        # 右侧：预览
        right = QVBoxLayout()
        layout.addLayout(right, 1)
        right.addWidget(QLabel("预览（成品渲染）"))
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.preview_container = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.scroll.setWidget(self.preview_container)
        right.addWidget(self.scroll, 1)

        # 信号
        # 【务必注意】QPushButton.clicked 信号会附带一个 checked(bool) 实参。
        # 而下面的槽都被 _guard_slot 包成了变参 def wrapper(self, *args, **kwargs)，
        # PyQt 无法据此裁剪实参数量，会把那个 bool 原样传进来；若槽只声明 (self)，
        # 点击即抛 TypeError —— 弹「操作失败」框并打印
        # "add_files() takes 1 positional argument but 2 were given"（2026-09-29 用户实测）。
        # 因此**所有直连的槽都必须多接一个参数**（下面这些已统一为 _checked=False）。
        # 走 lambda 的连接（上移/下移/行距类型）自己吞掉了参数，不受影响。
        self.btn_add.clicked.connect(self.add_files)
        self.btn_del.clicked.connect(self.remove_row)
        self.btn_up.clicked.connect(lambda: self.move_row(-1))
        self.btn_down.clicked.connect(lambda: self.move_row(1))
        self.btn_assemble.clicked.connect(self.do_assemble)
        self.btn_preview.clicked.connect(self.do_preview)

    # ---------------- 文件列表操作 ----------------
    # 文件对话框统一走这里。关键：**复用同一个 QFileDialog 实例**，而不是每次调用
    # getOpenFileNames/getSaveFileName 新建一个临时对话框。
    # 原因（2026-09-29 实测定位）：fcitx 输入法插件在 Qt 6.7 + 反复创建/销毁
    # QFileDialog 时，会在第二个对话框创建输入上下文时触发空指针连接
    #（"QObject::connect: Cannot connect (nullptr)::availabilityChanged(bool) to
    #   FcitxInputContextProxy::availabilityChanged()"）→ 进程 SIGSEGV（闪退）。
    # 而用户的实际操作序列正是「先点添加文件(打开对话框) → 再点一键汇编并保存
    # (保存对话框)」，第二次创建对话框就崩。只创建一个实例、动态切换
    # FileMode / AcceptMode / 文件名过滤器，即可彻底规避（已用 Xvfb + 真实
    # fcitx 环境反复 open→save×N 验证：零崩溃）。
    # 同时显式用 Qt 自带（非原生）对话框，绕开 GTK3/xdg-desktop-portal 主题插件
    # 的叠加风险，也让各系统表现一致。
    def _file_dialog(self):
        """懒创建并返回唯一的文件对话框实例（进程内复用，规避 fcitx 崩溃）。"""
        if self._file_dlg is None:
            self._file_dlg = QFileDialog(self)
            self._file_dlg.setOption(QFileDialog.Option.DontUseNativeDialog, True)
        return self._file_dlg

    def _open_files(self):
        log.info("打开「选择文档」对话框…")
        dlg = self._file_dialog()
        dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
        dlg.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dlg.setWindowTitle("选择文档")
        dlg.setNameFilter("Word 文档 (*.doc *.docx *.wps)")
        dlg.selectFile("")
        if dlg.exec() != QFileDialog.DialogCode.Accepted:
            log.info("「选择文档」对话框返回：0 个文件（取消）")
            return []
        files = list(dlg.selectedFiles())
        log.info("「选择文档」对话框返回：%d 个文件", len(files))
        return files

    def _save_file(self):
        log.info("打开「保存汇编文档」对话框…")
        dlg = self._file_dialog()
        dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dlg.setFileMode(QFileDialog.FileMode.AnyFile)
        dlg.setWindowTitle("保存汇编文档")
        dlg.setNameFilter("Word (*.docx)")
        dlg.selectFile("汇编结果.docx")
        if dlg.exec() != QFileDialog.DialogCode.Accepted:
            log.info("「保存汇编文档」对话框返回：（取消）")
            return ""
        out = dlg.selectedFiles()[0] if dlg.selectedFiles() else ""
        log.info("「保存汇编文档」对话框返回：%s", out or "（取消）")
        return out

    @_guard_slot
    def add_files(self, _checked=False):
        files = self._open_files()
        if not files:
            return
        # 识别标题/单位要读文档内容，必要时还会调用 LibreOffice 转换 .doc/.wps，
        # 单份就可能耗几秒。放在后台线程里做：否则界面会长时间无响应
        # （实测用户添加 7 份文档时主线程被阻塞约 51 秒，期间窗口完全冻结）。
        self.status.setText("正在识别 %d 份文档的标题/单位…" % len(files))
        log.info("开始识别 %d 份文档的标题/单位", len(files))

        def work():
            rows = []
            for f in files:
                # 优先从文档内容自动识别标题/单位（需求：标题用原文，不用文件名）
                try:
                    info = pipeline.analyze_source(f, self.work_dir)
                    unit, title = info.get("unit", ""), info.get("title", "")
                except Exception:
                    log.exception("从内容识别标题/单位失败，回退文件名解析：%s", f)
                    unit, title = pipeline.parse_filename(f)
                rows.append((f, unit, title))
            return rows

        self._run(work, done=self._append_rows)

    @_guard_slot
    def _append_rows(self, rows):
        rows = rows or []
        for f, unit, title in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(os.path.basename(f)))
            self.table.setItem(r, 1, QTableWidgetItem(unit))
            self.table.setItem(r, 2, QTableWidgetItem(title))
            self.table.item(r, 0).setData(Qt.ItemDataRole.UserRole, f)  # 存完整路径
        self.status.setText("已添加 %d 份文档，共 %d 份"
                            % (len(rows), self.table.rowCount()))
        log.info("已添加 %d 份文档（列表共 %d 份）", len(rows), self.table.rowCount())

    def _row_paths(self):
        rows = []
        for r in range(self.table.rowCount()):
            it0 = self.table.item(r, 0)
            if not it0:
                continue
            path = it0.data(Qt.ItemDataRole.UserRole) or it0.text()
            unit = self.table.item(r, 1).text() if self.table.item(r, 1) else ""
            title = self.table.item(r, 2).text() if self.table.item(r, 2) else ""
            rows.append({"path": path, "unit": unit, "title": title})
        return rows

    @_guard_slot
    def remove_row(self, _checked=False):
        for i in sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(i)

    @_guard_slot
    def move_row(self, d):
        r = self.table.currentRow()
        if r < 0:
            return
        nr = r + d
        if not (0 <= nr < self.table.rowCount()):
            return
        # 直接交换相邻两行的全部单元格，保留 UserRole 文件路径
        for c in range(self.table.columnCount()):
            a = self.table.takeItem(r, c)
            b = self.table.takeItem(nr, c)
            self.table.setItem(nr, c, a)
            self.table.setItem(r, c, b)
        self.table.setCurrentCell(nr, 0)
        log.info("调整顺序：第 %d 行 -> 第 %d 行", r, nr)

    # ---------------- 汇编 / 预览 ----------------
    def _cfg_options(self):
        cfg = {
            "cover_title": self.ed_title.text().strip(),
            "date_text": self.ed_date.text().strip(),
            "cover_title_size": self.spin_size.value(),
        }
        options = {
            "remove_signoff": self.chk_signoff.isChecked(),
            "each_material_new_page": self.chk_page.isChecked(),
        }
        return cfg, options

    # ---------------- 格式设置（字号/行距） ----------------
    def _fmt_mode_changed(self, mode_combo, spin):
        """行距类型切换：数值框切换为该类型的默认值与单位。"""
        mode = mode_combo.currentData()
        if mode == "multiple":
            spin.setDecimals(2)
            spin.setRange(0.5, 10.0)
            spin.setSingleStep(0.25)
            spin.setSuffix(" 倍")
            spin.setValue(1.0)        # 行距倍数默认 1 倍
        else:
            spin.setDecimals(2)
            spin.setRange(6.0, 100.0)
            spin.setSingleStep(0.5)
            spin.setSuffix(" 磅")
            spin.setValue(28.5)       # 固定值默认 28.5 磅

    def _fmt_options(self):
        """读取界面当前格式设置 -> fmt dict（与 assembler 约定结构一致）。"""
        fmt = {}
        for key, _label in FMT_KEYS:
            w = self._fmt_widgets[key]
            fmt[key] = {
                "size": float(w["size"].currentData()),
                "line_mode": w["line_mode"].currentData(),
                "line_value": float(w["line_val"].value()),
            }
        return fmt

    def _apply_fmt_to_ui(self, key, d):
        """把一份格式 dict（含默认值兜底）应用到界面控件。"""
        w = self._fmt_widgets[key]
        idx = w["size"].findData(float(d.get("size", 16.0)))
        if idx >= 0:
            w["size"].setCurrentIndex(idx)
        mode = d.get("line_mode", "exact")
        idx_m = w["line_mode"].findData(mode)
        w["line_mode"].setCurrentIndex(max(idx_m, 0))
        # 行距类型切换可能触发默认值重置，这里再按配置值覆盖
        w["line_val"].setValue(float(d.get("line_value", 28.5)))

    @_guard_slot
    def _fmt_reset_default(self, _checked=False):
        """恢复默认格式：全部三号 + 固定值 28.5 磅。"""
        for key, _label in FMT_KEYS:
            self._apply_fmt_to_ui(key, assembler_mod.DEFAULT_FMT[key])
        self._save_fmt_config()
        self.status.setText("已恢复默认格式（三号 + 固定值 28.5 磅）")

    def _config_path(self):
        """config.json 首选位置：程序根目录（便携版场景）。"""
        root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        return os.path.join(root, "config.json")

    @staticmethod
    def _alt_config_path():
        """config.json 回退位置：用户主目录（程序目录不可写时）。"""
        return os.path.join(os.path.expanduser("~"), ".doc_assembler_config.json")

    def _load_fmt_config(self):
        """启动时从 config.json 载入格式设置（缺失/非法项回退默认）。"""
        data = {}
        for p in (self._config_path(), self._alt_config_path()):
            try:
                if os.path.isfile(p):
                    with open(p, "r", encoding="utf-8") as f:
                        data = json.load(f) or {}
                    if data:
                        break
            except Exception:
                log.exception("读取配置文件失败：%s", p)
        cfg = data.get("fmt") if isinstance(data, dict) else None
        for key, _label in FMT_KEYS:
            d = dict(assembler_mod.DEFAULT_FMT[key])
            u = (cfg or {}).get(key) or {}
            if isinstance(u, dict):
                if u.get("size") is not None:
                    try:
                        d["size"] = float(u["size"])
                    except (TypeError, ValueError):
                        pass
                if u.get("line_mode") in ("exact", "multiple"):
                    d["line_mode"] = u["line_mode"]
                if u.get("line_value") is not None:
                    try:
                        d["line_value"] = float(u["line_value"])
                    except (TypeError, ValueError):
                        pass
            self._apply_fmt_to_ui(key, d)

    def _save_fmt_config(self):
        """把当前格式设置写入 config.json（写入失败静默，不影响使用）。"""
        cfg = {"fmt": self._fmt_options()}
        for p in (self._config_path(), self._alt_config_path()):
            try:
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(cfg, f, ensure_ascii=False, indent=2)
                return
            except Exception:
                continue
        log.warning("保存配置文件失败（程序目录与用户主目录均不可写）")

    def closeEvent(self, event):
        """退出前保存格式设置，并等待仍在跑的后台线程收尾。

        不等就退出的话，Qt 会析构仍在运行的 QThread，进程以 abort 结束
        （表现为"关闭程序时闪退"）。
        """
        for w in list(self._workers):
            if w.isRunning():
                log.info("退出前等待后台任务结束…")
                if not w.wait(5000):
                    log.warning("后台任务未在 5 秒内结束，仍将退出")
        try:
            self._save_fmt_config()
        except Exception:
            log.exception("退出保存配置失败")
        super().closeEvent(event)

    @_guard_slot
    def do_assemble(self, _checked=False):
        sources = self._row_paths()
        if not sources:
            QMessageBox.warning(self, "提示", "请先添加待汇编文档。")
            return
        out = self._save_file()
        if not out:
            return
        cfg, options = self._cfg_options()
        fmt = self._fmt_options()
        self._save_fmt_config()
        self.status.setText("正在汇编…")
        log.info("开始汇编：%d 份文档 -> %s（格式：%s）", len(sources), out, fmt)
        self._run(pipeline.run_assemble, sources, out, cfg, options, self.work_dir, fmt,
                  done=lambda p: (self.status.setText(f"已保存：{p}"),
                                  log.info("汇编完成：%s", p),
                                  self.do_preview_after(p)))

    @_guard_slot
    def do_preview_after(self, docx_path):
        self._run(pipeline.run_preview, docx_path, self.work_dir,
                  done=lambda imgs: self.show_preview(imgs))

    @_guard_slot
    def do_preview(self, _checked=False):
        sources = self._row_paths()
        if not sources:
            QMessageBox.warning(self, "提示", "请先添加待汇编文档。")
            return
        cfg, options = self._cfg_options()
        fmt = self._fmt_options()
        tmp = os.path.join(self.work_dir, "preview_tmp.docx")
        self.status.setText("正在生成预览…")
        log.info("开始生成预览：%d 份文档", len(sources))
        def work():
            pipeline.run_assemble(sources, tmp, cfg, options, self.work_dir, fmt)
            return pipeline.run_preview(tmp, self.work_dir)
        self._run(work, done=lambda imgs: (self.status.setText("预览已更新"),
                                           log.info("预览生成完成：%d 页", len(imgs) if imgs else 0),
                                           self.show_preview(imgs)))

    def _reap_workers(self):
        """回收已经**真正结束**的后台线程。

        只回收 isFinished() 为真的：Qt 里销毁仍在运行的 QThread 会直接
        qFatal("QThread: Destroyed while thread is still running") 让进程 abort，
        且不会有 Python 回溯——这正是"闪退"最典型的成因之一。
        """
        self._workers = [w for w in self._workers if not w.isFinished()]

    def _run(self, fn, *args, done=None):
        """启动后台任务。

        线程对象一律放进 self._workers 列表持有强引用，**不再用单个 self._worker
        属性**：先前 do_assemble 的完成回调里会再调一次 _run（继续生成预览），
        旧写法在回调里直接覆盖了 self._worker，最后一个引用随之消失，QThread 可能
        在仍运行时被析构 → 进程 abort。
        """
        self._reap_workers()
        worker = Worker(fn, *args)
        self._workers.append(worker)
        worker.finished.connect(lambda r: done(r) if done else None)
        worker.error.connect(lambda e: QMessageBox.critical(self, "错误", e))
        worker.start()

    @_guard_slot
    def show_preview(self, images):
        # 清空旧预览
        for i in reversed(range(self.preview_layout.count())):
            w = self.preview_layout.itemAt(i).widget()
            if w:
                w.setParent(None)
        if not images:
            self.preview_layout.addWidget(QLabel("（无预览，可能缺少 WPS 或转换失败）"))
            return
        for p in images:
            lab = QLabel()
            px = QPixmap(p)
            lab.setPixmap(px.scaledToWidth(560, Qt.TransformationMode.SmoothTransformation))
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.preview_layout.addWidget(lab)

    # ---------------- 环境检查 ----------------
    def _start_env_check(self):
        """由事件循环触发：后台跑环境自检，完成后按需非模态提示。

        自检要真实调一次转换引擎（约 1 秒起），放后台线程可避免启动瞬间界面
        发顿；结果一律不阻塞界面——正常时只更新状态栏，异常时才弹**非模态**
        提示框（见 _on_env_checked）。
        """
        try:
            worker = Worker(self._check_env)
        except Exception:
            log.exception("环境自检启动失败（已忽略，不影响使用）")
            return
        self._workers.append(worker)
        worker.finished.connect(self._on_env_checked)
        worker.error.connect(lambda e: log.warning("环境自检失败（已忽略，不影响使用）：%s", e))
        worker.start()

    @_guard_slot
    def _on_env_checked(self, msgs):
        """呈现环境自检结果：正常→状态栏；有问题→非模态提示框。

        【为什么不再用 QMessageBox.information 模态弹窗】
        原来无论是否正常都会弹一个模态框，且弹在 __init__ 里：主窗口要等它被
        点掉才显示，无人值守/无头环境下永久卡死；正常时还平白打断用户。
        现在正常只轻提示，异常才提示，且用 show() 非模态——任何情况下都不会
        卡住启动或卡住自动化点击。
        """
        msgs = [m for m in (msgs or []) if m]
        if not msgs:
            return
        warns = [m for m in msgs if m.lstrip().startswith("⚠")]
        if not warns:
            # 一切正常：状态栏轻提示即可
            if self.status.text() == "就绪":
                self.status.setText(msgs[0])
            log.info("环境自检通过：%s", msgs[0])
            return
        log.warning("环境检测发现问题（已在界面提示）：%s", " | ".join(warns))
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("环境检测")
        box.setText("\n".join(warns))
        box.setModal(False)
        box.setWindowModality(Qt.WindowModality.NonModal)
        box.show()
        # 持有引用：否则对话框可能被 Python GC 提前回收而闪一下就没
        self._env_box = box

    def _check_env(self):
        """环境自检（纯逻辑，不弹窗、不阻塞）：返回需要提示的消息列表。

        带 "⚠" 前缀的是"确实有问题"，其余为正常信息。
        """
        try:
            return self._check_env_inner()
        except Exception:
            # 环境自检失败只记日志，绝不能拦住程序启动
            log.exception("环境自检异常（已忽略，不影响使用）")
            return []

    def _check_env_inner(self):
        msgs = []
        name, exe = converter.detect_converter()
        if not exe:
            log.warning("未检测到可用的转换引擎")
            if platform.system() == "Windows":
                msgs.append("⚠ 未检测到 WPS。本程序转换/预览依赖 WPS（专业版/政府版，"
                            "支持命令行转换）。\n请安装 WPS 后重试。\n"
                            "（WPS 个人版通常不支持命令行转换，无法用于本程序。）")
            else:
                msgs.append("⚠ 未检测到 LibreOffice，也未检测到 WPS。\n"
                            "国产系统（麒麟/统信）转换 .doc/.wps → .docx 依赖 "
                            "LibreOffice 无头转换，请先安装 LibreOffice 后重试。\n"
                            "（WPS for Linux 不支持命令行转换，无法用于本程序。）")
        else:
            log.info("检测到转换引擎：%s @ %s", name, exe)
            ok, why = converter.wps_can_convert()
            if ok:
                # 界面上只显示引擎名，不显示完整路径：路径对用户没有操作价值，
                # 且是撑宽状态栏/左栏的主要来源（完整路径已在上一行写入日志）。
                msgs.append("已检测到转换引擎（%s），可用于转换与预览。" % name)
            else:
                log.warning("%s 转换测试失败：%s", name, why)
                if name == "LibreOffice":
                    msgs.append("⚠ 已检测到 LibreOffice，但转换测试失败：\n" + why +
                                "\n请确认 LibreOffice 安装完整（soffice 可正常启动）；"
                                "否则 .doc/.wps 转换不可用，仅 .docx 汇编可正常进行。")
                else:
                    msgs.append("⚠ 已检测到 WPS，但命令行转换测试失败：\n" + why +
                                "\n很可能为 WPS 个人版（不支持无头命令行转换）。\n"
                                "请改用 WPS 专业版/政府版，否则 .doc/.wps 转换与预览不可用，"
                                "仅 .docx 汇编可正常进行。")
        from app.core import preview as preview_mod
        if not preview_mod.has_pymupdf():
            log.warning("未安装 PyMuPDF，预览不可用")
            if platform.system() == "Windows":
                how = "  app\\.venv\\Scripts\\python.exe -m pip install PyMuPDF"
            else:
                # 国产系统离线包首次运行会自动从 wheels/ 安装；手动装则走系统源
                how = ("  重新运行「启动汇编程序.sh」即可（会自动从 wheels/ 安装）；\n"
                       "  或执行：sudo apt install python3-pymupdf")
            msgs.append("⚠ 未安装 PyMuPDF，预览功能不可用。需预览请执行：\n" + how)
        missing, bundled = fonts_mod.check_fonts()
        if missing:
            log.warning("缺少字体：%s（随附字体：%s）", missing, bundled)
            if bundled:
                msgs.append("⚠ 系统可能缺少字体：" + "、".join(missing) +
                            "。可在程序 fonts_bundled/ 目录安装后重试。")
            else:
                msgs.append("⚠ 系统可能缺少字体：" + "、".join(missing) + "，请安装对应字体。")
        # 不再在此处弹模态框（见 _on_env_checked 说明），只把消息交回调用方
        return msgs


def _resolve_icon_path():
    """定位程序图标（项目根目录 图标.png），不存在则返回 None。"""
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    for rel in ("图标.png", os.path.join("app", "图标.png")):
        p = os.path.join(root, rel)
        if os.path.isfile(p):
            return p
    return None


def main():
    app = QApplication(sys.argv)

    # 程序身份与桌面集成：让任务栏显示自定义图标而不是 Python 图标。
    # 详见 app/ui/appicon.py 顶部说明（WM_CLASS / _NET_WM_ICON / .desktop 三者关系）。
    icon = None
    try:
        from app.ui import appicon
        icon = appicon.apply_app_identity(app)
        installed = appicon.install_desktop_entry()
        if installed:
            print("桌面集成：%s" % installed)
    except Exception as e:  # noqa: BLE001
        import logging
        logging.getLogger(__name__).warning("程序身份/桌面集成初始化失败（不影响使用）：%s", e)

    # 保底：即使 appicon 不可用，也尽量挂上图标
    if icon is None:
        icon_path = _resolve_icon_path()
        if icon_path:
            icon = QIcon(icon_path)
            app.setWindowIcon(icon)

    w = MainWindow()
    if icon is not None:
        # 同时在窗口上再设一次：部分 Qt 版本只在窗口自身图标非空时才写 _NET_WM_ICON
        w.setWindowIcon(icon)
    w.show()

    # 启动后记录一次真实的窗口身份（排障用）：任务栏图标不对时，先看这条日志。
    # WM_CLASS 两段都应是 wenhui-assembler，_NET_WM_ICON 应非空。
    try:
        from PyQt6.QtCore import QTimer
        from app.ui import appicon
        QTimer.singleShot(1500, lambda: appicon.log_window_identity(w))
    except Exception:  # noqa: BLE001
        pass

    rc = app.exec()
    # fcitx 输入法插件（libfcitxplatforminputcontextplugin.so）存在析构顺序 bug：
    # 只要进程里打开过 QFileDialog（创建过 FcitxInputContextProxy 输入上下文），
    # 正常退出走到 Qt 对象析构阶段就会空指针崩溃（SIGSEGV，退出码 139，且无 Python
    # 回溯）。此时窗口已关闭、closeEvent 已执行（格式配置已保存、后台线程已收尾），
    # 用 os._exit 直接结束进程、跳过 Qt 析构即可彻底规避。这也顺带覆盖了
    # "关闭窗口时闪退"的隐患。
    os._exit(rc if isinstance(rc, int) else 0)


if __name__ == "__main__":
    main()

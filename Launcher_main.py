from __future__ import annotations

import json
import os
import sys
import uuid
import subprocess
from dataclasses import dataclass, asdict
from typing import List, Optional

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStyle,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


# -----------------------------
# Storage and model
# -----------------------------


def _app_data_dir() -> str:
    home = os.path.expanduser("~")
    path = os.path.join(home, ".launcher")
    os.makedirs(path, exist_ok=True)
    return path


def _data_file() -> str:
    return os.path.join(_app_data_dir(), "launcher_data.json")


@dataclass
class LauncherEntry:
    id: str
    name: str
    path: str
    description: str = ""
    entry_type: str = "app"  # "app" or "separator"

    @staticmethod
    def from_file(filepath: str) -> "LauncherEntry":
        base = os.path.splitext(os.path.basename(filepath))[0]
        return LauncherEntry(
            id=str(uuid.uuid4()),
            name=base,
            path=os.path.abspath(filepath),
            description="",
            entry_type="app",
        )

    @staticmethod
    def create_separator(name: str) -> "LauncherEntry":
        return LauncherEntry(
            id=str(uuid.uuid4()),
            name=name,
            path="",
            description="",
            entry_type="separator",
        )


def load_entries() -> List[LauncherEntry]:
    fp = _data_file()
    if not os.path.exists(fp):
        return []
    try:
        with open(fp, "r", encoding="utf-8") as f:
            data = json.load(f)
        items = data.get("entries", [])
        result: List[LauncherEntry] = []
        for it in items:
            try:
                result.append(
                    LauncherEntry(
                        id=it.get("id", str(uuid.uuid4())),
                        name=it.get("name", ""),
                        path=it.get("path", ""),
                        description=it.get("description", ""),
                        entry_type=it.get("entry_type", "app"),
                    )
                )
            except Exception:
                continue
        return result
    except Exception:
        return []


def get_backup_files() -> List[str]:
    """Get list of available backup files"""
    fp = _data_file()
    backups = []
    for i in range(1, 11):
        backup_path = f"{fp}.bak{i}"
        if os.path.exists(backup_path):
            backups.append(backup_path)
    return backups


def restore_from_backup(backup_path: str) -> bool:
    """Restore data from backup file"""
    try:
        fp = _data_file()
        import shutil
        shutil.copy2(backup_path, fp)
        return True
    except Exception:
        return False


def _backup_data_file() -> None:
    """Create backup copies of data file, keeping up to 10 generations"""
    fp = _data_file()
    if not os.path.exists(fp):
        return

    # Rotate existing backups
    for i in range(9, 0, -1):
        old_backup = f"{fp}.bak{i}"
        new_backup = f"{fp}.bak{i+1}"
        if os.path.exists(old_backup):
            if os.path.exists(new_backup):
                os.remove(new_backup)
            os.rename(old_backup, new_backup)

    # Create new backup
    backup_path = f"{fp}.bak1"
    if os.path.exists(backup_path):
        os.remove(backup_path)
    import shutil
    shutil.copy2(fp, backup_path)


def save_entries(entries: List[LauncherEntry]) -> None:
    # Create backup before saving
    _backup_data_file()

    fp = _data_file()
    data = {"entries": [asdict(e) for e in entries]}
    with open(fp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -----------------------------
# UI widgets
# -----------------------------


class EntryWidget(QWidget):
    def __init__(self, entry: LauncherEntry, on_run, parent=None):
        super().__init__(parent)
        self.entry = entry
        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 4, 6, 4)  # 左マージンを18pxに増加して字下げ
        layout.setSpacing(6)

        text_box = QVBoxLayout()
        text_box.setContentsMargins(0, 0, 0, 0)
        text_box.setSpacing(2)

        self.name_label = QLabel(entry.name)
        font = QFont()
        font.setBold(True)
        self.name_label.setFont(font)
        self.name_label.setWordWrap(True)

        self.desc_label = QLabel(entry.description or " ")
        self.desc_label.setStyleSheet("color: gray;")
        self.desc_label.setWordWrap(True)

        text_box.addWidget(self.name_label)
        text_box.addWidget(self.desc_label)

        self.run_btn = QToolButton()
        self.run_btn.setText("Run")
        self.run_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.run_btn.setAutoRaise(False)
        self.run_btn.setFixedWidth(48)
        self.run_btn.setStyleSheet("""
            QToolButton {
                background-color: #404040;
                border: 1px solid #303030;
                border-radius: 3px;
                color: #ffffff;
                padding: 4px;
            }
            QToolButton:hover {
                background-color: #505050;
                border: 1px solid #404040;
            }
            QToolButton:pressed {
                background-color: #303030;
                border: 1px solid #202020;
            }
        """)
        self.run_btn.clicked.connect(lambda: on_run(self.entry))

        layout.addLayout(text_box)
        layout.addWidget(self.run_btn, 0, Qt.AlignRight | Qt.AlignVCenter)

        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def sizeHint(self) -> QSize:  # ensure compact height
        return QSize(180, 48)


class SeparatorWidget(QWidget):
    def __init__(self, entry: LauncherEntry, parent=None):
        super().__init__(parent)
        self.entry = entry
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 10, 6, 10)
        layout.setSpacing(2)

        self.name_label = QLabel(entry.name)
        font = QFont()
        font.setBold(True)
        self.name_label.setFont(font)
        self.name_label.setAlignment(Qt.AlignCenter)
        self.name_label.setStyleSheet("color: #666666; background-color: #f0f0f0; padding: 8px; border-radius: 3px; min-height: 16px;")

        layout.addWidget(self.name_label)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def sizeHint(self) -> QSize:
        return QSize(180, 46)


class LauncherListWidget(QListWidget):
    filesDropped = Signal(list)  # list[str]
    orderChanged = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QListWidget.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QListWidget.SingleSelection)
        self.setSpacing(2)

        # カスタムドラッグ&ドロップ用
        self._drag_start_row = None
        self._drag_item_id = None

        # Save order when internal rows moved
        try:
            self.model().rowsMoved.connect(lambda parent, start, end, dest, row: self._on_rows_moved(parent, start, end, dest, row))
        except Exception:
            pass

    def _on_rows_moved(self, parent, start, end, dest, row):
        self.orderChanged.emit()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            item = self.itemAt(event.pos())
            if item:
                self._drag_start_row = self.row(item)
                self._drag_item_id = item.data(Qt.UserRole)
        super().mousePressEvent(event)

    def dragEnterEvent(self, event):
        if event.source() is self:
            event.acceptProposedAction()
            return
        md = event.mimeData()
        if md.hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.source() is self:
            # 内部ドラッグでは必ずMoveActionを使用
            event.setDropAction(Qt.MoveAction)
            event.accept()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.source() is self:
            # カウント確認用
            count_before = self.count()

            # Qt標準処理を実行してドラッグ機能を有効にする
            super().dropEvent(event)

            count_after = self.count()

            # アイテム数が減少した場合はリストを強制リフレッシュ
            if count_after < count_before:
                # orderChangedシグナルを発火してリフレッシュを促す
                self.orderChanged.emit()
            else:
                self.orderChanged.emit()
            return

        md = event.mimeData()
        if md.hasUrls():
            paths = []
            for url in md.urls():
                local = url.toLocalFile()
                if local:
                    paths.append(local)
            if paths:
                self.filesDropped.emit(paths)
                event.acceptProposedAction()
                return
        super().dropEvent(event)


class EntryDialog(QDialog):
    def __init__(self, entry: Optional[LauncherEntry] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("登録 / 編集")
        self.setModal(True)

        self._is_edit = entry is not None
        self._entry = entry

        lay = QGridLayout(self)
        row = 0

        lay.addWidget(QLabel("タイプ"), row, 0)
        self.type_combo = QComboBox()
        self.type_combo.addItems(["アプリケーション", "カテゴリ"])
        if entry and entry.entry_type == "separator":
            self.type_combo.setCurrentIndex(1)
        else:
            self.type_combo.setCurrentIndex(0)
        self.type_combo.currentIndexChanged.connect(self._on_type_changed)
        lay.addWidget(self.type_combo, row, 1, 1, 2)
        row += 1

        lay.addWidget(QLabel("名前"), row, 0)
        self.name_edit = QLineEdit(entry.name if entry else "")
        lay.addWidget(self.name_edit, row, 1, 1, 2)
        row += 1

        self.path_label = QLabel("パス")
        lay.addWidget(self.path_label, row, 0)
        self.path_edit = QLineEdit(entry.path if entry else "")
        self.browse_btn = QPushButton("参照")
        self.browse_btn.clicked.connect(self._browse)
        lay.addWidget(self.path_edit, row, 1)
        lay.addWidget(self.browse_btn, row, 2)
        row += 1

        lay.addWidget(QLabel("概要"), row, 0)
        self.desc_edit = QLineEdit(entry.description if entry else "")
        lay.addWidget(self.desc_edit, row, 1, 1, 2)
        row += 1

        btns = QHBoxLayout()
        save_btn = QPushButton("保存")
        cancel_btn = QPushButton("キャンセル")
        save_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btns.addStretch(1)
        btns.addWidget(save_btn)
        btns.addWidget(cancel_btn)
        lay.addLayout(btns, row, 0, 1, 3)

        self.setFixedWidth(420)
        self._on_type_changed()

    def _on_type_changed(self):
        is_category = self.type_combo.currentIndex() == 1
        self.path_label.setVisible(not is_category)
        self.path_edit.setVisible(not is_category)
        self.browse_btn.setVisible(not is_category)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "ファイル選択", "", "実行ファイル (*.py *.pyw *.exe *.bat *.cmd);;すべてのファイル (*.*)")
        if path:
            self.path_edit.setText(path)
            if not self.name_edit.text().strip():
                self.name_edit.setText(os.path.splitext(os.path.basename(path))[0])

    def get_entry(self) -> Optional[LauncherEntry]:
        name = self.name_edit.text().strip()
        desc = self.desc_edit.text().strip()
        is_category = self.type_combo.currentIndex() == 1

        if not name:
            QMessageBox.warning(self, "未入力", "名前は必須です。")
            return None

        if is_category:
            # Category entry
            entry_type = "separator"
            path = ""
        else:
            # App entry
            entry_type = "app"
            path = self.path_edit.text().strip()
            if not path:
                QMessageBox.warning(self, "未入力", "アプリケーションにはパスが必須です。")
                return None
            if not os.path.exists(path):
                ret = QMessageBox.question(self, "ファイル未検出", "指定したパスが存在しません。保存しますか？")
                if ret != QMessageBox.Yes:
                    return None

        if self._is_edit and self._entry:
            self._entry.name = name
            self._entry.path = path
            self._entry.description = desc
            self._entry.entry_type = entry_type
            return self._entry
        return LauncherEntry(id=str(uuid.uuid4()), name=name, path=path, description=desc, entry_type=entry_type)


# -----------------------------
# Main window
# -----------------------------


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Application Launcher")
        # Default size: H=860, W=200
        self.resize(200, 860)
        self.setMinimumWidth(180)

        self.entries: List[LauncherEntry] = load_entries()

        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        top_bar = QHBoxLayout()
        add_btn = QToolButton()
        add_btn.setText("＋")
        add_btn.setToolTip("新規登録")
        add_btn.clicked.connect(self.add_entry_dialog)

        restore_btn = QToolButton()
        restore_btn.setText("↺")
        restore_btn.setToolTip("バックアップから復旧")
        restore_btn.clicked.connect(self.show_restore_dialog)

        top_bar.addStretch(1)
        top_bar.addWidget(add_btn)
        top_bar.addWidget(restore_btn)
        root.addLayout(top_bar)

        self.list = LauncherListWidget()
        self.list.itemDoubleClicked.connect(self.edit_selected)
        self.list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self._context_menu)
        self.list.filesDropped.connect(self._handle_files_dropped)
        self.list.orderChanged.connect(self._save_current_order)
        root.addWidget(self.list, 1)

        self._refresh_list()

    # ----- UI population
    def _refresh_list(self):
        self.list.clear()
        for e in self.entries:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, e.id)
            if e.entry_type == "separator":
                widget = SeparatorWidget(e)
            else:
                widget = EntryWidget(e, self._run_entry)
            # All items have the same drag/drop flags for simplicity
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled | Qt.ItemIsEnabled)
            item.setSizeHint(widget.sizeHint())
            self.list.addItem(item)
            self.list.setItemWidget(item, widget)

    def _find_entry_by_id(self, entry_id: str) -> Optional[LauncherEntry]:
        for e in self.entries:
            if e.id == entry_id:
                return e
        return None

    def _selected_entry(self) -> Optional[LauncherEntry]:
        item = self.list.currentItem()
        if not item:
            return None
        entry_id = item.data(Qt.UserRole)
        return self._find_entry_by_id(entry_id)

    # ----- Actions
    def add_entry_dialog(self, from_path: Optional[str] = None):
        if from_path:
            tmp = LauncherEntry.from_file(from_path)
            dlg = EntryDialog(tmp, self)
        else:
            dlg = EntryDialog(None, self)
        if dlg.exec() == QDialog.Accepted:
            entry = dlg.get_entry()
            if entry is None:
                return
            self.entries.append(entry)
            save_entries(self.entries)
            self._refresh_list()

    def edit_selected(self):
        entry = self._selected_entry()
        if not entry:
            return
        dlg = EntryDialog(entry, self)
        if dlg.exec() == QDialog.Accepted:
            entry2 = dlg.get_entry()
            if entry2 is None:
                return
            save_entries(self.entries)
            self._refresh_list()

    def _context_menu(self, pos):
        item = self.list.itemAt(pos)
        menu = QMenu(self)

        if not item:
            act_add = menu.addAction("新規追加")
            chosen = menu.exec(self.list.mapToGlobal(pos))
            if chosen == act_add:
                self.add_entry_dialog()
            return

        entry_id = item.data(Qt.UserRole)
        entry = self._find_entry_by_id(entry_id)
        if not entry:
            return

        if entry.entry_type == "separator":
            act_edit = menu.addAction("Edit")
            menu.addSeparator()
            act_delete = menu.addAction("Delete")
        else:
            act_run = menu.addAction("Run")
            act_edit = menu.addAction("Edit")
            menu.addSeparator()
            act_delete = menu.addAction("Delete")

        chosen = menu.exec(self.list.mapToGlobal(pos))
        if entry.entry_type != "separator" and chosen == act_run:
            self._run_entry(entry)
        elif chosen == act_edit:
            self.edit_selected()
        elif chosen == act_delete:
            self._delete_entry(entry)

    def _delete_entry(self, entry: LauncherEntry):
        ret = QMessageBox.question(self, "削除確認", f"『{entry.name}』を削除しますか？")
        if ret != QMessageBox.Yes:
            return
        self.entries = [e for e in self.entries if e.id != entry.id]
        save_entries(self.entries)
        self._refresh_list()

    def _handle_files_dropped(self, paths: List[str]):
        # Register each dropped file
        for p in paths:
            if not p:
                continue
            # Create dialog pre-filled; user can just type description and save
            self.add_entry_dialog(from_path=p)

    def _save_current_order(self):
        # Read items from QListWidget in order and reorder self.entries accordingly
        ordered_ids: List[str] = []
        for i in range(self.list.count()):
            item = self.list.item(i)
            if item:
                eid = item.data(Qt.UserRole)
                if eid:
                    ordered_ids.append(eid)

        # Check if we're missing any entries
        original_ids = {e.id for e in self.entries}
        ordered_ids_set = set(ordered_ids)
        missing_ids = original_ids - ordered_ids_set

        if missing_ids:
            # UIアイテムが欠損している場合、リストを再構築
            self._refresh_list()
            return

        if len(ordered_ids) != len(ordered_ids_set):
            return

        id_to_entry = {e.id: e for e in self.entries}
        new_entries = [id_to_entry[eid] for eid in ordered_ids if eid in id_to_entry]

        # Only save if we have the same number of entries
        if len(new_entries) == len(self.entries):
            self.entries = new_entries
            save_entries(self.entries)

    # ----- Run

    def show_restore_dialog(self):
        from PySide6.QtWidgets import QListWidget, QVBoxLayout, QPushButton, QHBoxLayout
        backups = get_backup_files()
        if not backups:
            QMessageBox.information(self, "復旧", "利用可能なバックアップファイルがありません。")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("バックアップから復旧")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)

        layout.addWidget(QLabel("復旧するバックアップを選択してください:"))

        backup_list = QListWidget()
        for i, backup_path in enumerate(backups):
            import time
            try:
                mtime = os.path.getmtime(backup_path)
                time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
                backup_list.addItem(f"バックアップ{i+1} ({time_str})")
            except:
                backup_list.addItem(f"バックアップ{i+1}")
        layout.addWidget(backup_list)

        buttons = QHBoxLayout()
        restore_btn = QPushButton("復旧")
        cancel_btn = QPushButton("キャンセル")

        def do_restore():
            current_row = backup_list.currentRow()
            if current_row >= 0:
                backup_path = backups[current_row]
                if restore_from_backup(backup_path):
                    QMessageBox.information(dialog, "復旧完了", "データが復旧されました。アプリケーションを再起動してください。")
                    dialog.accept()
                else:
                    QMessageBox.warning(dialog, "復旧失敗", "復旧に失敗しました。")

        restore_btn.clicked.connect(do_restore)
        cancel_btn.clicked.connect(dialog.reject)
        buttons.addStretch()
        buttons.addWidget(restore_btn)
        buttons.addWidget(cancel_btn)
        layout.addLayout(buttons)

        dialog.setFixedSize(400, 300)
        dialog.exec()

    def _run_entry(self, entry: LauncherEntry):
        if entry.entry_type == "separator":
            return

        path = entry.path
        if not os.path.exists(path):
            QMessageBox.warning(self, "起動失敗", "ファイルが見つかりません。編集で修正してください。")
            return
        ext = os.path.splitext(path)[1].lower()
        cwd = os.path.dirname(path) or None
        try:
            if ext in (".py", ".pyw"):
                # 新しいCMDウィンドウでPythonスクリプトを実行
                cmd = f'start "Python: {entry.name}" cmd /k "cd /d "{cwd}" && python "{path}" && pause"'
                subprocess.Popen(cmd, shell=True)
            elif ext in (".exe", ".bat", ".cmd"):
                # 新しいCMDウィンドウで実行ファイルを実行
                if cwd:
                    cmd = f'start "{entry.name}" cmd /k "cd /d "{cwd}" && "{path}" && pause"'
                else:
                    cmd = f'start "{entry.name}" cmd /k ""{path}" && pause"'
                subprocess.Popen(cmd, shell=True)
            else:
                # その他のファイルはシステムデフォルトで開く
                if sys.platform.startswith("win"):
                    os.startfile(path)  # type: ignore[attr-defined]
                else:
                    subprocess.Popen([path], cwd=cwd)
        except Exception as e:
            QMessageBox.critical(self, "起動エラー", f"起動に失敗しました:\n{e}")


def main():
    # Windows環境でコンソール文字化け対策
    if sys.platform.startswith("win"):
        import locale
        import codecs
        try:
            # コンソールのエンコーディングをUTF-8に設定
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
        except:
            pass

    # Windows環境でコンソールウィンドウを非表示にする
    if sys.platform.startswith("win"):
        import os
        if os.name == 'nt':
            import ctypes
            ctypes.windll.kernel32.FreeConsole()

    app = QApplication(sys.argv)
    w = MainWindow()
    # Keep narrow footprint; allow resizing vertically, width small
    w.setMinimumWidth(250)
    w.resize(250, 930)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()


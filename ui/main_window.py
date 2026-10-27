import os
import sys
import re
import json
import datetime
import webbrowser
from pathlib import Path
from typing import Optional, Dict, Any

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QTextCursor
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QComboBox,
    QCheckBox,
    QTextEdit,
    QProgressBar,
    QFileDialog,
    QMessageBox,
    QDialog,
    QSplitter,
    QTabWidget,
)

from core.google_client import GoogleCloudClient
from core.uploader_worker import (
    BatchUploadWorker,
    scan_available_files,
    generate_subfolder_names,
)
from core.sevenzip_helper import find_7z_executable
from ui.batch_delete_page import BatchDeletePage

# 确保打包为 .exe 后配置文件与 token 保存在 .exe 同级目录下
APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent.parent
)
CONFIG_FILE = str(APP_DIR / "user_config.json")
TOKEN_FILE = str(APP_DIR / "token.json")


class AuthTestThread(QThread):
    """后台测试 Google 授权连接的轻量线程，避免卡顿界面"""

    result_signal = pyqtSignal(bool, str)

    def __init__(self, auth_mode: str, cred_file: str, proxy: str, parent=None):
        super().__init__(parent)
        self.auth_mode = auth_mode
        self.cred_file = cred_file
        self.proxy = proxy

    def run(self):
        client = GoogleCloudClient(
            auth_mode=self.auth_mode,
            cred_file=self.cred_file,
            token_file=TOKEN_FILE,
            proxy=self.proxy,
        )
        ok, msg = client.authenticate()
        self.result_signal.emit(ok, msg)


class HelpGuideDialog(QDialog):
    """内置的 Google Drive & Sheets 授权配置图文式步骤教程弹窗"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📖 Google 云端硬盘与谷歌表格 授权凭证获取教程")
        self.resize(760, 620)

        layout = QVBoxLayout(self)
        browser = QTextEdit(self)
        browser.setReadOnly(True)
        browser.setHtml(
            """
            <h2 style="color:#1B4F72;">Google Drive & Google Sheets 授权配置指南</h2>
            <p>本工具同时支持 <b>方式一：OAuth 2.0 个人账号网页登录</b>（最推荐，直接传到您自己的谷歌云盘）和 <b>方式二：Service Account 服务账号密钥</b>。只需配置一次，后续使用无需重复配置。</p>
            <hr>
            <h3 style="color:#C0392B;">⚠️ 必看：如果提示 Google Drive API has not been used in project... (403 错误)</h3>
            <p>说明您的 Google Cloud 项目还没有点击开启 API 开关，只需点击下面两个链接，在页面上点击蓝色按钮 <b>「启用 (Enable)」</b>，等待 1 分钟即可正常上传：</p>
            <ul>
                <li><b>开启 Google Drive API：</b><a href="https://console.cloud.google.com/apis/library/drive.googleapis.com">https://console.cloud.google.com/apis/library/drive.googleapis.com</a></li>
                <li><b>开启 Google Sheets API：</b><a href="https://console.cloud.google.com/apis/library/sheets.googleapis.com">https://console.cloud.google.com/apis/library/sheets.googleapis.com</a></li>
            </ul>
            <hr>
            <h3 style="color:#117864;">✅ 方式一：OAuth 2.0 个人账号网页授权（推荐）</h3>
            <ol>
                <li>打开 <a href="https://console.cloud.google.com/">Google Cloud Console (谷歌云控制台)</a> 并登录您的谷歌账号。</li>
                <li>点击顶部的项目下拉框，点击 <b>「新建项目 (New Project)」</b>，随便起个名字（例如 <code>MyDriveUploader</code>）并创建。</li>
                <li>在左侧菜单选择 <b>「API 和服务」 -> 「库 (Library)」</b>，分别搜索并点击 <b>启用 (Enable)</b> 以下两个 API：
                    <ul>
                        <li><b>Google Drive API</b>（必须启用，用于创建文件夹、上传文件、开启下载权限与批量删除）</li>
                        <li><b>Google Sheets API</b>（必须启用，用于创建和写入谷歌表格）</li>
                    </ul>
                </li>
                <li>在左侧菜单选择 <b>「API 和服务」 -> 「OAuth 同意屏幕 (OAuth consent screen)」</b>：
                    <ul>
                        <li>User Type 选择 <b>外部 (External)</b>，填写应用名称和您的邮箱，一路保存。</li>
                        <li>在 <b>「测试用户 (Test users)」</b> 页面，点击 <b>「+ ADD USERS」</b>，把您自己的 Gmail 邮箱添加进去并保存。</li>
                    </ul>
                </li>
                <li>在左侧菜单选择 <b>「API 和服务」 -> 「凭据 (Credentials)」</b>：
                    <ul>
                        <li>点击顶部 <b>「+ 创建凭据」 -> 「OAuth 客户端 ID」</b>。</li>
                        <li>应用类型选择 <b>「桌面应用 (Desktop app)」</b>，点击创建。</li>
                        <li>在弹出的窗口中点击 <b>「下载 JSON (Download JSON)」</b>，将该文件保存到电脑上（例如命名为 <code>credentials.json</code>）。</li>
                    </ul>
                </li>
                <li>回到本工具，授权方式选择 <b>「OAuth 2.0 个人账号网页登录」</b>，选择刚才下载的 JSON 文件，点击 <b>「🔍 测试谷歌连接」</b>，在弹出的浏览器中登录您的谷歌账号（若提示“Google尚未验证此应用”，点击左下角“高级 -> 继续前往”即可），勾选全部权限完成授权！</li>
            </ol>
            <hr>
            <h3 style="color:#2E4053;">🔑 方式二：Service Account 服务账号密钥</h3>
            <ol>
                <li>同样在 Google Cloud 项目中启用 <b>Google Drive API</b> 和 <b>Google Sheets API</b>。</li>
                <li>进入 <b>「API 和服务」 -> 「凭据」 -> 「+ 创建凭据」 -> 「服务账号 (Service account)」</b>。</li>
                <li>创建完成后，点击进入该服务账号，切换到 <b>「密钥 (Keys)」</b> 标签页，点击 <b>「添加密钥 -> 创建新密钥 -> JSON」</b>，下载得到服务账号 JSON 文件。</li>
                <li><b>注意：</b>使用服务账号时，建议在您自己的 Google Drive 中新建一个总文件夹和谷歌表格，然后右键 <b>「共享 (Share)」</b>，把服务账号的邮箱（形如 <code>xxx@xxx.iam.gserviceaccount.com</code>）添加为 <b>编辑者 (Editor)</b>，再把该总文件夹链接和表格链接填到本工具中。</li>
            </ol>
            <hr>
            <h3 style="color:#B7950B;">🌐 关于网络代理 (HTTP/HTTPS Proxy)</h3>
            <p>如果您在国内网络环境使用代理软件（如 Clash / v2rayN），请在界面上的 <b>「代理地址」</b> 填入本地代理端口（例如 Clash 默认是 <code>http://127.0.0.1:7890</code>，v2rayN 默认是 <code>http://127.0.0.1:10809</code>）。如果在海外网络直连，留空即可。</p>
            """
        )
        layout.addWidget(browser)

        btn_close = QPushButton("我知道了，关闭窗口")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(
            "Google Drive 批量分类上传 & 7z自动加密 & 谷歌表格归档 & 批量删除工具"
        )
        self.resize(1180, 920)

        self.worker: Optional[BatchUploadWorker] = None
        self.auth_thread: Optional[AuthTestThread] = None
        self.current_available_count: int = 0

        self.last_excel_path: str = ""
        self.last_gsheet_url: str = ""
        self.last_master_folder_url: str = ""

        # 防抖定时器
        self.scan_timer = QTimer(self)
        self.scan_timer.setSingleShot(True)
        self.scan_timer.setInterval(250)
        self.scan_timer.timeout.connect(self._do_refresh_local_file_count)

        self._init_ui()
        self._load_config()
        self._do_refresh_local_file_count()
        self._update_naming_preview()
        self._on_7z_toggle_changed()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 8, 12, 8)
        main_layout.setSpacing(6)

        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #F4F6F9;
            }
            QGroupBox {
                font-family: "Microsoft YaHei";
                font-size: 13px;
                font-weight: bold;
                color: #1B4F72;
                border: 1px solid #D0D7DE;
                border-radius: 6px;
                margin-top: 7px;
                padding-top: 7px;
                background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
            }
            QLabel {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                color: #2C3E50;
            }
            QLineEdit, QComboBox, QSpinBox {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                padding: 4px 6px;
                border: 1px solid #BDC3C7;
                border-radius: 4px;
                background-color: #FAFAFA;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border: 1px solid #2980B9;
                background-color: #FFFFFF;
            }
            QPushButton {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                padding: 5px 12px;
                border-radius: 4px;
                border: 1px solid #BDC3C7;
                background-color: #ECF0F1;
                color: #2C3E50;
            }
            QPushButton:hover {
                background-color: #D6EAF8;
                border-color: #5DADE2;
            }
            QCheckBox {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                color: #2C3E50;
            }
            QTabWidget::pane {
                border: 1px solid #BDC3C7;
                border-radius: 6px;
                background-color: #F8F9F9;
            }
            QTabBar::tab {
                font-family: "Microsoft YaHei";
                font-size: 13px;
                font-weight: bold;
                padding: 8px 22px;
                margin-right: 4px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                background-color: #E5E8E8;
                color: #34495E;
            }
            QTabBar::tab:selected {
                background-color: #1F618D;
                color: #FFFFFF;
            }
            """
        )

        # 顶部标题与辅助工具栏
        top_bar = QHBoxLayout()
        title_lbl = QLabel(
            "☁️ Google Drive 批量分文件夹上传 & 7z自动加密 & 谷歌表格归档 & 批量清理工具"
        )
        title_lbl.setFont(QFont("Microsoft YaHei", 14, QFont.Weight.Bold))
        title_lbl.setStyleSheet("color: #154360;")
        top_bar.addWidget(title_lbl)
        top_bar.addStretch()

        self.btn_help = QPushButton("📖 谷歌授权凭证获取教程")
        self.btn_help.setStyleSheet(
            "background-color:#FEF9E7; color:#7D6608; border-color:#F7DC6F; font-weight:bold;"
        )
        self.btn_help.clicked.connect(self._on_show_help)
        top_bar.addWidget(self.btn_help)

        main_layout.addLayout(top_bar)

        # ==================== 顶部公共区：Google 云端授权与网络设置 ====================
        grp_auth = QGroupBox(
            "🔑 Google 云端账号授权与网络设置（上传页面与批量删除页面共用）"
        )
        grid_auth = QGridLayout(grp_auth)
        grid_auth.setSpacing(6)

        grid_auth.addWidget(QLabel("授权方式:"), 0, 0)
        self.combo_auth_mode = QComboBox()
        self.combo_auth_mode.addItem(
            "OAuth 2.0 个人账号网页授权 (credentials.json)", "oauth"
        )
        self.combo_auth_mode.addItem(
            "Service Account 服务账号密钥 (service_account.json)",
            "service_account",
        )
        grid_auth.addWidget(self.combo_auth_mode, 0, 1)

        grid_auth.addWidget(QLabel("JSON凭证文件:"), 0, 2)
        self.edit_cred_file = QLineEdit()
        self.edit_cred_file.setPlaceholderText(
            "请选择 credentials.json 或 service_account.json 路径..."
        )
        grid_auth.addWidget(self.edit_cred_file, 0, 3)

        btn_pick_cred = QPushButton("选择文件...")
        btn_pick_cred.clicked.connect(self._on_pick_cred_file)
        grid_auth.addWidget(btn_pick_cred, 0, 4)

        self.btn_test_auth = QPushButton("🔍 测试谷歌连接")
        self.btn_test_auth.setStyleSheet(
            "background-color:#D6EAF8; color:#1B4F72; font-weight:bold;"
        )
        self.btn_test_auth.clicked.connect(self._on_test_google_auth)
        grid_auth.addWidget(self.btn_test_auth, 0, 5)

        grid_auth.addWidget(QLabel("网络代理(选填):"), 1, 0)
        self.edit_proxy = QLineEdit()
        self.edit_proxy.setPlaceholderText(
            "国内使用代理请填写，如 http://127.0.0.1:7890 ；直连环境留空即可"
        )
        grid_auth.addWidget(self.edit_proxy, 1, 1, 1, 5)

        main_layout.addWidget(grp_auth)

        # ==================== 多页面切换标签栏 (QTabWidget) ====================
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs, stretch=1)

        # -------------------- 页面一：批量分类上传 & 7z自动加密 & 谷歌表格归档 --------------------
        self.page_upload = QWidget()
        upload_page_layout = QVBoxLayout(self.page_upload)
        upload_page_layout.setContentsMargins(8, 6, 8, 6)
        upload_page_layout.setSpacing(5)

        splitter = QSplitter(Qt.Orientation.Vertical)
        upload_page_layout.addWidget(splitter, stretch=1)

        config_container = QWidget()
        config_layout = QVBoxLayout(config_container)
        config_layout.setContentsMargins(0, 0, 0, 0)
        config_layout.setSpacing(5)

        # 1. 本地源文件夹与防重归档
        grp_local = QGroupBox("① 本地待上传文件目录 & 自动归档防重设置")
        grid_local = QGridLayout(grp_local)
        grid_local.setSpacing(6)

        grid_local.addWidget(QLabel("本地源文件夹:"), 0, 0)
        self.edit_source_dir = QLineEdit()
        self.edit_source_dir.setPlaceholderText(
            "选择存放待上传本地文件的文件夹..."
        )
        self.edit_source_dir.textChanged.connect(self._refresh_local_file_count)
        grid_local.addWidget(self.edit_source_dir, 0, 1, 1, 2)

        btn_pick_dir = QPushButton("📂 选择文件夹")
        btn_pick_dir.clicked.connect(self._on_pick_source_dir)
        grid_local.addWidget(btn_pick_dir, 0, 3)

        btn_refresh_dir = QPushButton("🔄 刷新扫描")
        btn_refresh_dir.clicked.connect(self._do_refresh_local_file_count)
        grid_local.addWidget(btn_refresh_dir, 0, 4)

        btn_open_src = QPushButton("打开源目录")
        btn_open_src.clicked.connect(self._on_open_source_dir)
        grid_local.addWidget(btn_open_src, 0, 5)

        grid_local.addWidget(QLabel("文件排序方式:"), 1, 0)
        self.combo_sort = QComboBox()
        self.combo_sort.addItem("按文件名自然排序 (1, 2, 10...)", "natural")
        self.combo_sort.addItem("按修改时间 (从早到晚)", "mtime_asc")
        self.combo_sort.addItem("按修改时间 (从新到旧)", "mtime_desc")
        self.combo_sort.addItem("按文件大小 (从小到大)", "size_asc")
        self.combo_sort.addItem("按文件大小 (从大到小)", "size_desc")
        self.combo_sort.currentIndexChanged.connect(self._refresh_local_file_count)
        grid_local.addWidget(self.combo_sort, 1, 1)

        grid_local.addWidget(QLabel("后缀过滤(选填):"), 1, 2)
        self.edit_ext_filter = QLineEdit()
        self.edit_ext_filter.setPlaceholderText("留空传所有类型；或填如 .mp4,.jpg,.zip")
        self.edit_ext_filter.textChanged.connect(self._refresh_local_file_count)
        grid_local.addWidget(self.edit_ext_filter, 1, 3)

        self.chk_recursive = QCheckBox("包含源目录内子文件夹中的文件")
        self.chk_recursive.stateChanged.connect(self._refresh_local_file_count)
        grid_local.addWidget(self.chk_recursive, 1, 4, 1, 2)

        grid_local.addWidget(QLabel("已上传归档夹名:"), 2, 0)
        self.edit_uploaded_folder = QLineEdit("已上传")
        self.edit_uploaded_folder.textChanged.connect(self._refresh_local_file_count)
        grid_local.addWidget(self.edit_uploaded_folder, 2, 1)

        self.chk_group_archive = QCheckBox(
            "在「已上传」文件夹内按对应的云端子文件夹名分组存放（清晰对照且防重）"
        )
        self.chk_group_archive.setChecked(True)
        grid_local.addWidget(self.chk_group_archive, 2, 2, 1, 3)

        btn_open_archive = QPushButton("📦 打开「已上传」目录")
        btn_open_archive.clicked.connect(self._on_open_archive_dir)
        grid_local.addWidget(btn_open_archive, 2, 5)

        config_layout.addWidget(grp_local)

        # 2. 云端文件夹层级、数量控制与命名规则
        grp_cloud = QGroupBox(
            "② 云端文件夹层级、数量配额控制（多的不传、少的不重复） & 命名规则"
        )
        grid_cloud = QGridLayout(grp_cloud)
        grid_cloud.setSpacing(6)

        grid_cloud.addWidget(QLabel("云端总文件夹:"), 0, 0)
        self.combo_master_mode = QComboBox()
        self.combo_master_mode.addItem(
            "自动新建云端总文件夹 (推荐)", "create_new"
        )
        self.combo_master_mode.addItem(
            "使用已有云端总文件夹 (填入文件夹链接或ID)", "use_existing"
        )
        self.combo_master_mode.currentIndexChanged.connect(
            self._on_master_mode_changed
        )
        grid_cloud.addWidget(self.combo_master_mode, 0, 1)

        self.lbl_master_input = QLabel("新建总文件夹名:")
        grid_cloud.addWidget(self.lbl_master_input, 0, 2)
        self.edit_master_input = QLineEdit()
        self.edit_master_input.setPlaceholderText(
            "留空则自动按「批量上传总目录_日期时间」命名"
        )
        grid_cloud.addWidget(self.edit_master_input, 0, 3)

        self.lbl_parent_drive = QLabel("上级目录ID(选填):")
        grid_cloud.addWidget(self.lbl_parent_drive, 0, 4)
        self.edit_parent_drive_id = QLineEdit()
        self.edit_parent_drive_id.setPlaceholderText(
            "留空则在云盘根目录创建；或粘贴上级文件夹链接/ID"
        )
        grid_cloud.addWidget(self.edit_parent_drive_id, 0, 5)

        grid_cloud.addWidget(QLabel("需创建子文件夹数(M):"), 1, 0)
        self.spin_folder_count = QSpinBox()
        self.spin_folder_count.setRange(1, 9999)
        self.spin_folder_count.setValue(3)
        self.spin_folder_count.setSuffix(" 个文件夹")
        self.spin_folder_count.valueChanged.connect(self._on_quota_changed)
        grid_cloud.addWidget(self.spin_folder_count, 1, 1)

        grid_cloud.addWidget(QLabel("每夹上传文件数(N):"), 1, 2)
        self.spin_files_per_folder = QSpinBox()
        self.spin_files_per_folder.setRange(1, 9999)
        self.spin_files_per_folder.setValue(4)
        self.spin_files_per_folder.setSuffix(" 个文件/夹")
        self.spin_files_per_folder.valueChanged.connect(self._on_quota_changed)
        grid_cloud.addWidget(self.spin_files_per_folder, 1, 3)

        self.lbl_quota_status = QLabel("正在计算配额...")
        self.lbl_quota_status.setFixedHeight(26)
        self.lbl_quota_status.setStyleSheet(
            "background-color:#EBF5FB; color:#1B4F72; padding:3px 8px; border-radius:4px; font-weight:bold;"
        )
        grid_cloud.addWidget(self.lbl_quota_status, 1, 4, 1, 2)

        grid_cloud.addWidget(QLabel("子文件夹命名规则:"), 2, 0)
        self.combo_naming_mode = QComboBox()
        self.combo_naming_mode.addItem(
            "纯数字补零序号 (如: 001, 002, 003)", "pure_num"
        )
        self.combo_naming_mode.addItem(
            "类别前缀 + 数字序号 (如: 分类_001, 分类_002)", "prefix_num"
        )
        self.combo_naming_mode.addItem(
            "日期 + 前缀 + 序号 (如: 20261027_分类_001)", "date_prefix_num"
        )
        self.combo_naming_mode.addItem(
            "自定义模板 (可用 {prefix} {num} {index} {date})",
            "custom_template",
        )
        self.combo_naming_mode.addItem(
            "自定义名称列表 (逗号分隔，如: 视频组,图片组,文档组)",
            "custom_list",
        )
        self.combo_naming_mode.setCurrentIndex(1)
        self.combo_naming_mode.currentIndexChanged.connect(
            self._update_naming_preview
        )
        grid_cloud.addWidget(self.combo_naming_mode, 2, 1)

        grid_cloud.addWidget(QLabel("类别前缀:"), 2, 2)
        self.edit_folder_prefix = QLineEdit("分类_")
        self.edit_folder_prefix.textChanged.connect(self._update_naming_preview)
        grid_cloud.addWidget(self.edit_folder_prefix, 2, 3)

        sub_num_layout = QHBoxLayout()
        sub_num_layout.addWidget(QLabel("起始序号:"))
        self.spin_start_index = QSpinBox()
        self.spin_start_index.setRange(0, 99999)
        self.spin_start_index.setValue(1)
        self.spin_start_index.valueChanged.connect(self._update_naming_preview)
        sub_num_layout.addWidget(self.spin_start_index)

        sub_num_layout.addWidget(QLabel("数字位数:"))
        self.spin_zero_pad = QSpinBox()
        self.spin_zero_pad.setRange(1, 6)
        self.spin_zero_pad.setValue(3)
        self.spin_zero_pad.setSuffix(" 位")
        self.spin_zero_pad.valueChanged.connect(self._update_naming_preview)
        sub_num_layout.addWidget(self.spin_zero_pad)
        grid_cloud.addLayout(sub_num_layout, 2, 4, 1, 2)

        grid_cloud.addWidget(QLabel("模板/自定义列表:"), 3, 0)
        self.edit_custom_naming = QLineEdit("{prefix}{num}")
        self.edit_custom_naming.setPlaceholderText(
            "自定义模板如 {date}_{prefix}{num} ；或自定义名称列表如 文件夹A, 文件夹B, 文件夹C"
        )
        self.edit_custom_naming.textChanged.connect(self._update_naming_preview)
        grid_cloud.addWidget(self.edit_custom_naming, 3, 1, 1, 2)

        self.lbl_naming_preview = QLabel("📁 子文件夹命名预览: ")
        self.lbl_naming_preview.setFixedHeight(26)
        self.lbl_naming_preview.setStyleSheet(
            "color:#0E6251; background-color:#E8F8F5; padding:3px 8px; border-radius:4px; font-weight:bold;"
        )
        grid_cloud.addWidget(self.lbl_naming_preview, 3, 3, 1, 3)

        config_layout.addWidget(grp_cloud)

        # 3. 上传前 7z 自动加密设置
        grp_7z = QGroupBox(
            "③ 上传前 7z 自动加密开关（开启后不论上传什么文件均自动转为 7z 加密压缩卷，且强制加密文件名 -mhe=on）"
        )
        grid_7z = QGridLayout(grp_7z)
        grid_7z.setSpacing(6)

        self.chk_enable_7z = QCheckBox(
            "🔒 开启上传前自动 7z 加密（强制加密文件名，云端仅存 .7z 加密压缩卷）"
        )
        self.chk_enable_7z.setStyleSheet(
            "color: #8E44AD; font-weight: bold; font-size: 13px;"
        )
        self.chk_enable_7z.stateChanged.connect(self._on_7z_toggle_changed)
        grid_7z.addWidget(self.chk_enable_7z, 0, 0, 1, 2)

        grid_7z.addWidget(QLabel("7z 加密密码(必填):"), 0, 2)
        pwd_layout = QHBoxLayout()
        self.edit_7z_password = QLineEdit()
        self.edit_7z_password.setPlaceholderText("请手动输入 7z 解压密码...")
        self.edit_7z_password.setEchoMode(QLineEdit.EchoMode.Normal)
        pwd_layout.addWidget(self.edit_7z_password)

        self.chk_show_pwd = QCheckBox("显示密码")
        self.chk_show_pwd.setChecked(True)
        self.chk_show_pwd.stateChanged.connect(self._on_toggle_pwd_echo)
        pwd_layout.addWidget(self.chk_show_pwd)
        grid_7z.addLayout(pwd_layout, 0, 3)

        grid_7z.addWidget(QLabel("压缩速度等级:"), 0, 4)
        self.combo_7z_level = QComboBox()
        self.combo_7z_level.addItem("极速加密 (-mx=0 仅打包加密，最快)", 0)
        self.combo_7z_level.addItem("快速压缩加密 (-mx=1 推荐)", 1)
        self.combo_7z_level.addItem("标准压缩加密 (-mx=5 常规)", 5)
        self.combo_7z_level.addItem("极限压缩加密 (-mx=9 最小体积)", 9)
        self.combo_7z_level.setCurrentIndex(1)
        grid_7z.addWidget(self.combo_7z_level, 0, 5)

        grid_7z.addWidget(QLabel("7z 打包粒度:"), 1, 0)
        self.combo_7z_mode = QComboBox()
        self.combo_7z_mode.addItem(
            "每个文件独立加密为单独的 .7z 压缩卷上传（每夹保持 N 个独立 .7z 文件及各自下载直链）",
            "single_file",
        )
        self.combo_7z_mode.addItem(
            "将每个子文件夹内的 N 个文件合并打包为 1 个 .7z 压缩卷上传",
            "folder_bundle",
        )
        grid_7z.addWidget(self.combo_7z_mode, 1, 1, 1, 2)

        self.chk_record_pwd_in_sheet = QCheckBox("在表格备注列中记录 7z 密码")
        self.chk_record_pwd_in_sheet.setChecked(True)
        grid_7z.addWidget(self.chk_record_pwd_in_sheet, 1, 3)

        self.chk_keep_7z_copy = QCheckBox(
            "在本地「已上传」内同时保留加密后的.7z副本"
        )
        self.chk_keep_7z_copy.setChecked(False)
        grid_7z.addWidget(self.chk_keep_7z_copy, 1, 4)

        detected_7z = find_7z_executable() or ""
        self.edit_7z_exe = QLineEdit(detected_7z)
        btn_pick_7z = QPushButton(
            "7z路径..." if not detected_7z else "✅ 7z引擎已就绪"
        )
        if detected_7z:
            btn_pick_7z.setStyleSheet(
                "color:#196F3D; background-color:#D4EFDF; font-weight:bold;"
            )
        btn_pick_7z.clicked.connect(self._on_pick_7z_exe)
        self.btn_pick_7z = btn_pick_7z
        grid_7z.addWidget(btn_pick_7z, 1, 5)

        config_layout.addWidget(grp_7z)

        # 4. 谷歌表格 (Google Sheets) 与本地报表设置
        grp_sheet = QGroupBox(
            "④ 谷歌表格 (Google Sheets) & 层级结构报表设置（包含总夹链接、子夹链接、单文件ID与直链）"
        )
        grid_sheet = QGridLayout(grp_sheet)
        grid_sheet.setSpacing(6)

        grid_sheet.addWidget(QLabel("谷歌表格链接/ID:"), 0, 0)
        self.edit_gsheet_input = QLineEdit()
        self.edit_gsheet_input.setPlaceholderText(
            "可粘贴已有 Google Sheets 表格链接或ID直接追加写入；留空则自动在云端新建谷歌表格"
        )
        grid_sheet.addWidget(self.edit_gsheet_input, 0, 1, 1, 2)

        self.chk_merge_cells = QCheckBox(
            "在表格「层级分栏总表」中纵向合并相同总文件夹、子文件夹单元格（同时自动生成「树状目录层级表」）"
        )
        self.chk_merge_cells.setChecked(True)
        grid_sheet.addWidget(self.chk_merge_cells, 0, 3, 1, 3)

        btn_row_sheet = QHBoxLayout()
        self.btn_open_excel = QPushButton("📊 打开最新本地 Excel 层级表")
        self.btn_open_excel.clicked.connect(self._on_open_last_excel)
        btn_row_sheet.addWidget(self.btn_open_excel)

        self.btn_open_gsheet = QPushButton("📗 在浏览器打开云端谷歌表格")
        self.btn_open_gsheet.clicked.connect(self._on_open_last_gsheet)
        btn_row_sheet.addWidget(self.btn_open_gsheet)

        self.btn_open_master_cloud = QPushButton("🌐 在浏览器打开云端总文件夹")
        self.btn_open_master_cloud.clicked.connect(
            self._on_open_last_master_folder
        )
        btn_row_sheet.addWidget(self.btn_open_master_cloud)
        btn_row_sheet.addStretch()

        grid_sheet.addLayout(btn_row_sheet, 1, 0, 1, 6)
        config_layout.addWidget(grp_sheet)

        splitter.addWidget(config_container)

        # 5. 底部执行按钮、进度条与彩色日志
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 4, 0, 0)
        bottom_layout.setSpacing(4)

        ctrl_bar = QHBoxLayout()
        self.btn_start = QPushButton(
            "🚀 开始批量分文件夹上传 & 写入层级表格"
        )
        self.btn_start.setFixedHeight(38)
        self.btn_start.setFont(QFont("Microsoft YaHei", 11, QFont.Weight.Bold))
        self.btn_start.setStyleSheet(
            """
            QPushButton {
                background-color: #1F618D;
                color: #FFFFFF;
                border: none;
                border-radius: 5px;
                padding: 6px 22px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
            QPushButton:disabled {
                background-color: #95A5A6;
            }
            """
        )
        self.btn_start.clicked.connect(self._on_start_upload)
        ctrl_bar.addWidget(self.btn_start)

        self.btn_stop = QPushButton("🛑 停止任务")
        self.btn_stop.setFixedHeight(38)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet(
            """
            QPushButton {
                background-color: #C0392B;
                color: #FFFFFF;
                border: none;
                border-radius: 5px;
                padding: 6px 16px;
                font-weight: bold;
            }
            QPushButton:disabled {
                background-color: #E6B0AA;
            }
            """
        )
        self.btn_stop.clicked.connect(self._on_stop_upload)
        ctrl_bar.addWidget(self.btn_stop)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(26)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ctrl_bar.addWidget(self.progress_bar, stretch=1)

        btn_export_log = QPushButton("💾 导出日志")
        btn_export_log.clicked.connect(self._on_export_log)
        ctrl_bar.addWidget(btn_export_log)

        btn_clear_log = QPushButton("🧹 清空日志")
        btn_clear_log.clicked.connect(lambda: self.txt_log.clear())
        ctrl_bar.addWidget(btn_clear_log)

        bottom_layout.addLayout(ctrl_bar)

        self.lbl_progress_detail = QLabel("就绪：等待任务开始...")
        self.lbl_progress_detail.setFixedHeight(22)
        self.lbl_progress_detail.setWordWrap(False)
        self.lbl_progress_detail.setStyleSheet(
            "color: #34495E; font-weight: bold;"
        )
        bottom_layout.addWidget(self.lbl_progress_detail)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFont(QFont("Consolas", 10))
        self.txt_log.setStyleSheet(
            "background-color: #1E272E; color: #D2DAE2; border-radius: 6px; padding: 6px;"
        )
        bottom_layout.addWidget(self.txt_log, stretch=1)

        splitter.addWidget(bottom_widget)
        splitter.setSizes([460, 290])

        self.tabs.addTab(
            self.page_upload,
            "📤 页面一：批量分类上传 & 7z自动加密 & 谷歌表格归档",
        )

        # -------------------- 页面二：云端链接批量删除清理 --------------------
        self.page_delete = BatchDeletePage(
            get_base_config_cb=self._collect_config,
            parent=self,
        )
        self.tabs.addTab(
            self.page_delete,
            "🗑️ 页面二：云端链接批量删除清理（服务结束后一键删链接）",
        )

        self._append_log(
            "欢迎使用 Google Drive 批量分类上传 & 7z自动加密 & 谷歌表格归档 & 批量删除工具！",
            "SUCCESS",
        )
        if detected_7z:
            self._append_log(
                f"🔒 已自动检测到 7-Zip 加密引擎: {detected_7z}（支持 AES-256 加密及 -mhe=on 强制加密文件名）",
                "INFO",
            )

    def _on_toggle_pwd_echo(self):
        if self.chk_show_pwd.isChecked():
            self.edit_7z_password.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.edit_7z_password.setEchoMode(QLineEdit.EchoMode.Password)

    def _on_7z_toggle_changed(self):
        enabled = self.chk_enable_7z.isChecked()
        self.edit_7z_password.setEnabled(enabled)
        self.chk_show_pwd.setEnabled(enabled)
        self.combo_7z_level.setEnabled(enabled)
        self.combo_7z_mode.setEnabled(enabled)
        self.chk_record_pwd_in_sheet.setEnabled(enabled)
        self.chk_keep_7z_copy.setEnabled(enabled)
        self.btn_pick_7z.setEnabled(enabled)

    def _on_pick_7z_exe(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 7-Zip (7z.exe) 程序文件",
            r"C:\Program Files\7-Zip",
            "Executable (7z.exe *.exe);;All Files (*)",
        )
        if path:
            self.edit_7z_exe.setText(path)
            self.btn_pick_7z.setText("✅ 7z引擎已就绪")
            self.btn_pick_7z.setStyleSheet(
                "color:#196F3D; background-color:#D4EFDF; font-weight:bold;"
            )
            self._save_config()

    def _set_status_single_line(self, text: str, is_error: bool = False):
        first_line = (text or "").splitlines()[0] if text else ""
        clean = " ".join(first_line.split())
        if len(clean) > 110:
            clean = clean[:110] + "..."
        self.lbl_progress_detail.setText(clean)
        if is_error:
            self.lbl_progress_detail.setStyleSheet(
                "color: #C0392B; font-weight: bold;"
            )
        else:
            self.lbl_progress_detail.setStyleSheet(
                "color: #34495E; font-weight: bold;"
            )

    def _append_log(self, msg: str, level: str = "INFO"):
        now_str = datetime.datetime.now().strftime("%H:%M:%S")
        color_map = {
            "INFO": "#D2DAE2",
            "SUCCESS": "#2ECC71",
            "WARN": "#F39C12",
            "ERROR": "#FF5E57",
        }
        color = color_map.get(level.upper(), "#D2DAE2")
        for line in (msg or "").splitlines():
            escaped = (
                line.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            html_line = f'<div style="color:{color}; margin-bottom:2px;">[{now_str}] [{level}] {escaped}</div>'
            self.txt_log.moveCursor(QTextCursor.MoveOperation.End)
            self.txt_log.insertHtml(html_line)
            self.txt_log.insertPlainText("\n")
        self.txt_log.moveCursor(QTextCursor.MoveOperation.End)

    def _prompt_enable_google_api_if_needed(self, err_msg: str):
        urls = re.findall(r"https://console\.[^\s]+", err_msg)
        if "尚未开启 Google Drive API" in err_msg and urls:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("需要在谷歌云后台开启 API 开关")
            box.setText(
                "您的谷歌账号已登录成功，但您的 Google Cloud 项目尚未开启【Google Drive API】或【Google Sheets API】开关！\n\n"
                "是否立即在浏览器中打开谷歌云开启页面？（在打开的网页中点击蓝色的【启用 / Enable】按钮，等 30 秒后再点开始即可）"
            )
            box.setDetailedText(err_msg)
            btn_open = box.addButton(
                "🌐 立即在浏览器打开 API 启用页面",
                QMessageBox.ButtonRole.AcceptRole,
            )
            box.addButton("稍后手动打开", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() == btn_open:
                for u in urls:
                    webbrowser.open(u)
            return True
        return False

    def _on_master_mode_changed(self):
        mode = self.combo_master_mode.currentData()
        if mode == "use_existing":
            self.lbl_master_input.setText("已有总文件夹链接/ID:")
            self.edit_master_input.setPlaceholderText(
                "请粘贴 Google Drive 已有总文件夹的分享链接或 Folder ID"
            )
            self.edit_parent_drive_id.setEnabled(False)
        else:
            self.lbl_master_input.setText("新建总文件夹名:")
            self.edit_master_input.setPlaceholderText(
                "留空则自动按「批量上传总目录_日期时间」命名"
            )
            self.edit_parent_drive_id.setEnabled(True)

    def _refresh_local_file_count(self):
        self.scan_timer.start()

    def _do_refresh_local_file_count(self):
        source_dir = self.edit_source_dir.text().strip()
        uploaded_name = self.edit_uploaded_folder.text().strip() or "已上传"
        sort_mode = self.combo_sort.currentData() or "natural"
        ext_filter = self.edit_ext_filter.text().strip()
        recursive = self.chk_recursive.isChecked()

        files = scan_available_files(
            source_dir=source_dir,
            uploaded_folder_name=uploaded_name,
            sort_mode=sort_mode,
            ext_filter=ext_filter,
            recursive=recursive,
        )
        self.current_available_count = len(files)
        self._update_quota_banner()

    def _on_quota_changed(self):
        self._update_quota_banner()
        self._update_naming_preview()

    def _update_quota_banner(self):
        m = self.spin_folder_count.value()
        n = self.spin_files_per_folder.value()
        planned_total = m * n
        avail = getattr(self, "current_available_count", 0)

        if not self.edit_source_dir.text().strip():
            self.lbl_quota_status.setText(
                f"计划总量: {m}夹 × {n}个 = {planned_total} 个文件 | 请先选择本地源文件夹"
            )
            self.lbl_quota_status.setStyleSheet(
                "background-color:#EBF5FB; color:#1B4F72; padding:3px 8px; border-radius:4px; font-weight:bold;"
            )
            return

        if avail == 0:
            self.lbl_quota_status.setText(
                f"⚠️ 本地可用: 0 个 | 计划: {m}夹×{n}个={planned_total}个（目录为空或已全部移入「已上传」）"
            )
            self.lbl_quota_status.setStyleSheet(
                "background-color:#FDEDEC; color:#922B21; padding:3px 8px; border-radius:4px; font-weight:bold;"
            )
        elif avail > planned_total:
            extra = avail - planned_total
            self.lbl_quota_status.setText(
                f"ℹ️ 本地可用: {avail} 个 > 计划: {planned_total} 个（只传前 {planned_total} 个，多余 {extra} 个不上传）"
            )
            self.lbl_quota_status.setStyleSheet(
                "background-color:#E8F8F5; color:#0E6251; padding:3px 8px; border-radius:4px; font-weight:bold;"
            )
        elif avail < planned_total:
            diff = planned_total - avail
            actual_folders = (avail + n - 1) // n
            self.lbl_quota_status.setText(
                f"⚠️ 本地可用: {avail} 个 < 计划: {planned_total} 个（差 {diff} 个，绝不重复上传！将只传 {avail} 个分入 {actual_folders} 个夹）"
            )
            self.lbl_quota_status.setStyleSheet(
                "background-color:#FEF9E7; color:#9A7D0A; padding:3px 8px; border-radius:4px; font-weight:bold;"
            )
        else:
            self.lbl_quota_status.setText(
                f"✅ 数量完美匹配: 本地可用 {avail} 个 == 计划 {m}夹×{n}个 = {planned_total} 个文件"
            )
            self.lbl_quota_status.setStyleSheet(
                "background-color:#D4EFDF; color:#196F3D; padding:3px 8px; border-radius:4px; font-weight:bold;"
            )

    def _update_naming_preview(self):
        mode = self.combo_naming_mode.currentData()
        count = min(self.spin_folder_count.value(), 5)
        prefix = self.edit_folder_prefix.text()
        start_idx = self.spin_start_index.value()
        zero_pad = self.spin_zero_pad.value()
        custom_text = self.edit_custom_naming.text()

        preview_names = generate_subfolder_names(
            count=count,
            naming_mode=mode,
            prefix=prefix,
            start_index=start_idx,
            zero_pad=zero_pad,
            custom_template=custom_text,
            custom_list_text=custom_text,
        )
        suffix = " ..." if self.spin_folder_count.value() > 5 else ""
        self.lbl_naming_preview.setText(
            f"📁 子文件夹命名预览: {', '.join(preview_names)}{suffix}"
        )

    def _on_pick_cred_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Google JSON 授权凭证文件",
            "",
            "JSON Files (*.json);;All Files (*)",
        )
        if path:
            self.edit_cred_file.setText(path)
            self._save_config()

    def _on_pick_source_dir(self):
        path = QFileDialog.getExistingDirectory(
            self, "选择本地待上传文件所在的源文件夹"
        )
        if path:
            self.edit_source_dir.setText(path)
            self._do_refresh_local_file_count()
            self._save_config()

    def _on_open_source_dir(self):
        src = self.edit_source_dir.text().strip()
        if src and os.path.isdir(src):
            os.startfile(src)
        else:
            QMessageBox.warning(self, "提示", "请先选择有效的本地源文件夹！")

    def _on_open_archive_dir(self):
        src = self.edit_source_dir.text().strip()
        if not src or not os.path.isdir(src):
            QMessageBox.warning(self, "提示", "请先选择有效的本地源文件夹！")
            return
        arch_name = self.edit_uploaded_folder.text().strip() or "已上传"
        arch_dir = Path(src) / arch_name
        arch_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(str(arch_dir))

    def _on_show_help(self):
        dlg = HelpGuideDialog(self)
        dlg.exec()

    def _on_test_google_auth(self):
        cred_file = self.edit_cred_file.text().strip()
        if not cred_file or not os.path.isfile(cred_file):
            QMessageBox.warning(
                self,
                "缺少凭证文件",
                "请先点击「选择文件...」选定您的 Google JSON 凭证文件！\n如果您还没有凭证，请点击右上角「📖 谷歌授权凭证获取教程」查看步骤。",
            )
            return

        self.btn_test_auth.setEnabled(False)
        self.btn_test_auth.setText("正在验证...")
        self._append_log("正在验证 Google Drive & Sheets 授权连接...", "INFO")

        self.auth_thread = AuthTestThread(
            auth_mode=self.combo_auth_mode.currentData(),
            cred_file=cred_file,
            proxy=self.edit_proxy.text().strip(),
            parent=self,
        )
        self.auth_thread.result_signal.connect(self._on_auth_test_finished)
        self.auth_thread.start()

    def _on_auth_test_finished(self, ok: bool, msg: str):
        self.btn_test_auth.setEnabled(True)
        self.btn_test_auth.setText("🔍 测试谷歌连接")
        if ok:
            self._append_log(f"✅ {msg}", "SUCCESS")
            self._set_status_single_line(f"✅ {msg}", is_error=False)
            QMessageBox.information(self, "连接成功", msg)
        else:
            self._append_log(f"❌ {msg}", "ERROR")
            self._set_status_single_line(f"❌ 连接未通过: {msg}", is_error=True)
            if not self._prompt_enable_google_api_if_needed(msg):
                QMessageBox.critical(self, "连接失败", msg)

    def _collect_config(self) -> Dict[str, Any]:
        return {
            "auth_mode": self.combo_auth_mode.currentData(),
            "cred_file": self.edit_cred_file.text().strip(),
            "token_file": TOKEN_FILE,
            "proxy": self.edit_proxy.text().strip(),
            "source_dir": self.edit_source_dir.text().strip(),
            "sort_mode": self.combo_sort.currentData(),
            "ext_filter": self.edit_ext_filter.text().strip(),
            "recursive_scan": self.chk_recursive.isChecked(),
            "uploaded_folder_name": self.edit_uploaded_folder.text().strip()
            or "已上传",
            "group_archive_by_subfolder": self.chk_group_archive.isChecked(),
            "master_folder_mode": self.combo_master_mode.currentData(),
            "master_folder_input": self.edit_master_input.text().strip(),
            "parent_drive_id": self.edit_parent_drive_id.text().strip(),
            "folder_count": self.spin_folder_count.value(),
            "files_per_folder": self.spin_files_per_folder.value(),
            "naming_mode": self.combo_naming_mode.currentData(),
            "folder_prefix": self.edit_folder_prefix.text(),
            "start_index": self.spin_start_index.value(),
            "zero_pad": self.spin_zero_pad.value(),
            "custom_template": self.edit_custom_naming.text().strip(),
            "custom_list_text": self.edit_custom_naming.text().strip(),
            "enable_7z_encrypt": self.chk_enable_7z.isChecked(),
            "sevenzip_password": self.edit_7z_password.text(),
            "sevenzip_level": self.combo_7z_level.currentData(),
            "sevenzip_mode": self.combo_7z_mode.currentData(),
            "record_pwd_in_sheet": self.chk_record_pwd_in_sheet.isChecked(),
            "keep_encrypted_7z_in_archive": self.chk_keep_7z_copy.isChecked(),
            "sevenzip_exe": self.edit_7z_exe.text().strip(),
            "google_sheet_input": self.edit_gsheet_input.text().strip(),
            "merge_hierarchy_cells": self.chk_merge_cells.isChecked(),
        }

    def _save_config(self):
        try:
            cfg = self._collect_config()
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_config(self):
        if not os.path.isfile(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            idx_auth = self.combo_auth_mode.findData(cfg.get("auth_mode", "oauth"))
            if idx_auth >= 0:
                self.combo_auth_mode.setCurrentIndex(idx_auth)

            self.edit_cred_file.setText(cfg.get("cred_file", ""))
            self.edit_proxy.setText(cfg.get("proxy", ""))
            self.edit_source_dir.setText(cfg.get("source_dir", ""))

            idx_sort = self.combo_sort.findData(cfg.get("sort_mode", "natural"))
            if idx_sort >= 0:
                self.combo_sort.setCurrentIndex(idx_sort)

            self.edit_ext_filter.setText(cfg.get("ext_filter", ""))
            self.chk_recursive.setChecked(bool(cfg.get("recursive_scan", False)))
            self.edit_uploaded_folder.setText(
                cfg.get("uploaded_folder_name", "已上传")
            )
            self.chk_group_archive.setChecked(
                bool(cfg.get("group_archive_by_subfolder", True))
            )

            idx_master = self.combo_master_mode.findData(
                cfg.get("master_folder_mode", "create_new")
            )
            if idx_master >= 0:
                self.combo_master_mode.setCurrentIndex(idx_master)
            self.edit_master_input.setText(cfg.get("master_folder_input", ""))
            self.edit_parent_drive_id.setText(cfg.get("parent_drive_id", ""))

            self.spin_folder_count.setValue(int(cfg.get("folder_count", 3)))
            self.spin_files_per_folder.setValue(int(cfg.get("files_per_folder", 4)))

            idx_naming = self.combo_naming_mode.findData(
                cfg.get("naming_mode", "prefix_num")
            )
            if idx_naming >= 0:
                self.combo_naming_mode.setCurrentIndex(idx_naming)

            self.edit_folder_prefix.setText(cfg.get("folder_prefix", "分类_"))
            self.spin_start_index.setValue(int(cfg.get("start_index", 1)))
            self.spin_zero_pad.setValue(int(cfg.get("zero_pad", 3)))
            self.edit_custom_naming.setText(
                cfg.get("custom_template", "{prefix}{num}")
            )

            self.chk_enable_7z.setChecked(bool(cfg.get("enable_7z_encrypt", False)))
            self.edit_7z_password.setText(cfg.get("sevenzip_password", ""))
            idx_7z_lvl = self.combo_7z_level.findData(cfg.get("sevenzip_level", 1))
            if idx_7z_lvl >= 0:
                self.combo_7z_level.setCurrentIndex(idx_7z_lvl)
            idx_7z_mode = self.combo_7z_mode.findData(
                cfg.get("sevenzip_mode", "single_file")
            )
            if idx_7z_mode >= 0:
                self.combo_7z_mode.setCurrentIndex(idx_7z_mode)
            self.chk_record_pwd_in_sheet.setChecked(
                bool(cfg.get("record_pwd_in_sheet", True))
            )
            self.chk_keep_7z_copy.setChecked(
                bool(cfg.get("keep_encrypted_7z_in_archive", False))
            )
            if cfg.get("sevenzip_exe"):
                self.edit_7z_exe.setText(cfg.get("sevenzip_exe", ""))

            self.edit_gsheet_input.setText(cfg.get("google_sheet_input", ""))
            self.chk_merge_cells.setChecked(
                bool(cfg.get("merge_hierarchy_cells", True))
            )
        except Exception:
            pass

    def _on_start_upload(self):
        cfg = self._collect_config()
        self._save_config()

        if not cfg["source_dir"] or not os.path.isdir(cfg["source_dir"]):
            QMessageBox.warning(
                self, "参数检查", "请先选择有效的「本地源文件夹」！"
            )
            return

        if cfg["enable_7z_encrypt"]:
            if not cfg["sevenzip_password"]:
                QMessageBox.warning(
                    self,
                    "缺少 7z 加密密码",
                    "您已勾选【🔒 开启上传前自动 7z 加密】，请手动输入「7z 加密密码」！",
                )
                self.edit_7z_password.setFocus()
                return
            if not find_7z_executable(cfg["sevenzip_exe"]):
                QMessageBox.warning(
                    self,
                    "未找到 7-Zip",
                    "未找到 7z.exe 加密程序！请点击「7z路径...」按钮选择本机 7z.exe 所在位置。",
                )
                return

        if not cfg["cred_file"] or not os.path.isfile(cfg["cred_file"]):
            QMessageBox.warning(
                self,
                "缺少 Google 授权凭证",
                "请先在顶部「🔑 Google 云端账号授权与网络设置」中选择有效的 Google JSON 凭证文件！",
            )
            return

        if (
            cfg["master_folder_mode"] == "use_existing"
            and not cfg["master_folder_input"]
        ):
            QMessageBox.warning(
                self,
                "参数检查",
                "您选择了「使用已有云端总文件夹」，请填入总文件夹的分享链接或 ID！",
            )
            return

        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.progress_bar.setValue(0)
        self._set_status_single_line("正在启动上传线程...", is_error=False)

        self.worker = BatchUploadWorker(cfg, parent=self)
        self.worker.log_signal.connect(self._append_log)
        self.worker.progress_signal.connect(self._on_worker_progress)
        self.worker.finished_signal.connect(self._on_worker_finished)
        self.worker.start()

    def _on_stop_upload(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self._append_log("正在请求安全停止任务，请稍候...", "WARN")
            self.btn_stop.setEnabled(False)

    def _on_worker_progress(self, current: int, total: int, status_text: str):
        pct = int(current * 100 / max(total, 1))
        self.progress_bar.setValue(pct)
        self._set_status_single_line(status_text, is_error=False)

    def _on_worker_finished(self, success: bool, result: dict):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self._do_refresh_local_file_count()

        if success:
            self.progress_bar.setValue(100)
            self.last_excel_path = result.get("excel_path", "")
            self.last_gsheet_url = result.get("gsheet_url", "")
            master_info = result.get("master_folder", {})
            self.last_master_folder_url = master_info.get("webViewLink", "")

            self._set_status_single_line(
                f"✅ 任务完成！共处理 {result.get('total_uploaded_files', 0)} 个文件，分配至 {len(result.get('subfolders', []))} 个子文件夹。",
                is_error=False,
            )
            msg_lines = [
                "✅ 批量分类上传与层级制表已完成！",
                "",
                f"• 云端总文件夹：{master_info.get('name', '')}",
                f"• 实际创建子夹：{len(result.get('subfolders', []))} 个",
                f"• 实际处理文件：{result.get('total_uploaded_files', 0)} 个",
                f"• 7z 加密状态：{'已开启 (含加密文件名 -mhe=on)' if result.get('enable_7z') else '未开启 (原文件直接上传)'}",
                f"• 数量说明：{result.get('summary_note', '')}",
                "",
                "• 本地层级 Excel 表格已保存至：",
                f"  {self.last_excel_path}",
            ]
            if self.last_gsheet_url:
                msg_lines.append("• 云端谷歌表格链接：")
                msg_lines.append(f"  {self.last_gsheet_url}")

            QMessageBox.information(self, "任务完成", "\n".join(msg_lines))
        else:
            err = result.get("error", "未知错误")
            self._set_status_single_line(f"❌ 任务未完成: {err}", is_error=True)
            if not self._prompt_enable_google_api_if_needed(err):
                QMessageBox.warning(self, "任务未完成", err)

    def _on_open_last_excel(self):
        if self.last_excel_path and os.path.isfile(self.last_excel_path):
            os.startfile(self.last_excel_path)
            return

        src = self.edit_source_dir.text().strip()
        if src:
            rep_dir = Path(src) / "上传报表_Reports"
            if rep_dir.is_dir():
                xlsx_files = sorted(
                    rep_dir.glob("*.xlsx"),
                    key=lambda p: p.stat().st_mtime,
                    reverse=True,
                )
                if xlsx_files:
                    os.startfile(str(xlsx_files[0]))
                    return

        QMessageBox.information(
            self, "提示", "暂未找到生成的本地 Excel 报表，请先执行一次上传任务！"
        )

    def _on_open_last_gsheet(self):
        url = self.last_gsheet_url or self.edit_gsheet_input.text().strip()
        if url and url.startswith("http"):
            webbrowser.open(url)
        else:
            QMessageBox.information(
                self,
                "提示",
                "当前还没有可打开的云端谷歌表格链接！",
            )

    def _on_open_last_master_folder(self):
        if self.last_master_folder_url and self.last_master_folder_url.startswith(
            "http"
        ):
            webbrowser.open(self.last_master_folder_url)
        else:
            QMessageBox.information(
                self, "提示", "请先执行一次上传任务以获取云端总文件夹链接！"
            )

    def _on_export_log(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出运行日志",
            f"上传日志_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "Text Files (*.txt)",
        )
        if path:
            Path(path).write_text(self.txt_log.toPlainText(), encoding="utf-8")
            self._append_log(f"运行日志已导出至: {path}", "SUCCESS")

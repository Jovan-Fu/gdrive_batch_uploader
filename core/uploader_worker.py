import os
import re
import shutil
import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

from PyQt6.QtCore import QThread, pyqtSignal

from core.google_client import GoogleCloudClient, format_size
from core.exporter import HierarchyReportExporter
from core.sevenzip_helper import find_7z_executable, create_encrypted_7z

IGNORED_FILENAMES = {
    "desktop.ini",
    "thumbs.db",
    ".ds_store",
    ".upload_history.json",
}


def natural_sort_key(path_obj: Path):
    """自然排序 Key，使 1.mp4, 2.mp4, 10.mp4 按数字顺序排列"""
    text = path_obj.name.lower()
    return [
        int(part) if part.isdigit() else part
        for part in re.split(r"(\d+)", text)
    ]


def scan_available_files(
    source_dir: str,
    uploaded_folder_name: str = "已上传",
    sort_mode: str = "natural",
    ext_filter: str = "",
    recursive: bool = False,
) -> List[Path]:
    """
    扫描本地源目录中的待上传文件：
    - 自动排除「已上传」归档目录、「上传报表_Reports」目录以及「.7z_staging_temp」临时加密目录
    - 自动排除隐藏文件、临时文件 (~$*)
    """
    if not source_dir or not os.path.isdir(source_dir):
        return []

    root_path = Path(source_dir).resolve()
    excluded_dir_names = {
        uploaded_folder_name.strip(),
        "已上传",
        "已上传_Uploaded",
        "上传报表_Reports",
        ".7z_staging_temp",
        "__pycache__",
        ".git",
    }

    allowed_exts = set()
    if ext_filter and ext_filter.strip() and ext_filter.strip() != "*":
        for part in re.split(r"[,;，；\s]+", ext_filter.strip()):
            p = part.strip().lower()
            if not p:
                continue
            if not p.startswith("."):
                p = "." + p
            allowed_exts.add(p)

    files: List[Path] = []

    if recursive:
        for dirpath, dirnames, filenames in os.walk(root_path):
            dirnames[:] = [
                d
                for d in dirnames
                if d not in excluded_dir_names and not d.startswith(".")
            ]
            for fn in filenames:
                if (
                    fn.lower() in IGNORED_FILENAMES
                    or fn.startswith("~$")
                    or fn.startswith(".")
                ):
                    continue
                f_path = Path(dirpath) / fn
                if allowed_exts and f_path.suffix.lower() not in allowed_exts:
                    continue
                if f_path.is_file():
                    files.append(f_path)
    else:
        for item in root_path.iterdir():
            if item.is_dir():
                continue
            fn = item.name
            if (
                fn.lower() in IGNORED_FILENAMES
                or fn.startswith("~$")
                or fn.startswith(".")
            ):
                continue
            if allowed_exts and item.suffix.lower() not in allowed_exts:
                continue
            if item.is_file():
                files.append(item)

    if sort_mode == "mtime_asc":
        files.sort(key=lambda p: (p.stat().st_mtime, natural_sort_key(p)))
    elif sort_mode == "mtime_desc":
        files.sort(key=lambda p: (-p.stat().st_mtime, natural_sort_key(p)))
    elif sort_mode == "size_asc":
        files.sort(key=lambda p: (p.stat().st_size, natural_sort_key(p)))
    elif sort_mode == "size_desc":
        files.sort(key=lambda p: (-p.stat().st_size, natural_sort_key(p)))
    else:
        files.sort(key=natural_sort_key)

    return files


def generate_subfolder_names(
    count: int,
    naming_mode: str = "prefix_num",
    prefix: str = "",
    start_index: int = 1,
    zero_pad: int = 3,
    custom_template: str = "{prefix}{num}",
    custom_list_text: str = "",
) -> List[str]:
    """根据用户配置的云端文件夹命名规则，生成 count 个子文件夹名称列表。"""
    now = datetime.datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")

    names: List[str] = []

    if naming_mode == "custom_list":
        raw_items = [
            x.strip()
            for x in re.split(r"[,，\n\r;；]+", custom_list_text or "")
            if x.strip()
        ]
        for i in range(count):
            idx_num = start_index + i
            padded = str(idx_num).zfill(max(1, zero_pad))
            if i < len(raw_items):
                names.append(raw_items[i])
            else:
                base = prefix.strip() or "分类_"
                names.append(f"{base}{padded}")
        return names

    for i in range(count):
        idx_num = start_index + i
        padded = str(idx_num).zfill(max(1, zero_pad))

        if naming_mode == "pure_num":
            names.append(padded)
        elif naming_mode == "prefix_num":
            names.append(f"{prefix}{padded}")
        elif naming_mode == "date_prefix_num":
            mid = f"_{prefix}" if prefix else "_"
            names.append(f"{date_str}{mid}{padded}")
        elif naming_mode == "custom_template":
            tpl = custom_template.strip() or "{prefix}{num}"
            try:
                folder_name = tpl.format(
                    num=padded,
                    index=idx_num,
                    prefix=prefix,
                    date=date_str,
                    time=time_str,
                )
            except Exception:
                folder_name = f"{prefix}{padded}"
            names.append(folder_name)
        else:
            names.append(f"{prefix}{padded}")

    return names


class BatchUploadWorker(QThread):
    """
    后台批量上传工作线程：
    - 严格按 M 个文件夹 × 每个文件夹 N 个文件 的总量规则执行
    - 多的不上传，少的绝不重复上传，并在日志中详细说明
    - 支持上传前自动 7z 加密开关（强制 -mhe=on 加密文件名 + 手动设置密码）
    - 每个文件上传后开启公开下载权限 (anyone -> reader)，获取文件 ID、下载直链、查看链接
    - 每上传成功一个本地文件，立即将其移动到本地「已上传」归档文件夹，彻底杜绝重复上传
    - 生成按【总文件夹 -> 子文件夹 -> 单个文件】层级排版的谷歌表格与本地 Excel 表格
    """

    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int, int, str)
    finished_signal = pyqtSignal(bool, dict)

    def __init__(self, config: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.config = config
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def _log(self, msg: str, level: str = "INFO"):
        self.log_signal.emit(msg, level)

    def _move_to_uploaded_archive(
        self,
        local_file: Path,
        source_dir: str,
        uploaded_folder_name: str,
        subfolder_name: str,
        group_by_subfolder: bool = True,
    ) -> Path:
        """将已上传成功的单个本地文件移动到「已上传」归档目录下，防止下次重复上传。"""
        archive_root = Path(source_dir) / (uploaded_folder_name.strip() or "已上传")
        if group_by_subfolder and subfolder_name:
            safe_sub = (
                re.sub(r'[<>:"/\\|?*]', "_", subfolder_name).strip() or "default"
            )
            target_dir = archive_root / safe_sub
        else:
            target_dir = archive_root

        target_dir.mkdir(parents=True, exist_ok=True)
        dest_path = target_dir / local_file.name

        if dest_path.exists():
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            dest_path = target_dir / f"{local_file.stem}_{ts}{local_file.suffix}"

        shutil.move(str(local_file), str(dest_path))
        return dest_path

    def run(self):
        start_dt = datetime.datetime.now()
        batch_time_str = start_dt.strftime("%Y-%m-%d %H:%M:%S")
        file_ts_str = start_dt.strftime("%Y%m%d_%H%M%S")
        staging_root: Optional[Path] = None

        try:
            source_dir = self.config.get("source_dir", "").strip()
            uploaded_folder_name = (
                self.config.get("uploaded_folder_name", "已上传").strip() or "已上传"
            )
            group_archive_by_subfolder = self.config.get(
                "group_archive_by_subfolder", True
            )
            sort_mode = self.config.get("sort_mode", "natural")
            ext_filter = self.config.get("ext_filter", "")
            recursive_scan = self.config.get("recursive_scan", False)

            planned_folders = int(self.config.get("folder_count", 1))
            files_per_folder = int(self.config.get("files_per_folder", 1))
            planned_total_files = planned_folders * files_per_folder

            enable_7z = bool(self.config.get("enable_7z_encrypt", False))
            sevenzip_password = self.config.get("sevenzip_password", "")
            sevenzip_mode = self.config.get("sevenzip_mode", "single_file")
            sevenzip_level = int(self.config.get("sevenzip_level", 1))
            sevenzip_exe_custom = self.config.get("sevenzip_exe", "").strip()
            keep_encrypted_copy = bool(
                self.config.get("keep_encrypted_7z_in_archive", False)
            )
            record_pwd_in_sheet = bool(
                self.config.get("record_pwd_in_sheet", False)
            )

            self._log("=" * 68, "INFO")
            self._log(f"🚀 开始执行批量分类上传任务 | 时间: {batch_time_str}", "INFO")
            self._log(f"📂 本地待上传源目录: {source_dir}", "INFO")

            sevenzip_exe = ""
            if enable_7z:
                if not sevenzip_password:
                    msg = "已开启【上传前自动 7z 加密】，但未设置 7z 加密密码！请在界面上输入加密密码后再启动。"
                    self._log(f"❌ {msg}", "ERROR")
                    self.finished_signal.emit(False, {"error": msg})
                    return

                sevenzip_exe = find_7z_executable(sevenzip_exe_custom) or ""
                if not sevenzip_exe:
                    msg = "未找到 7-Zip (7z.exe) 程序！请确认电脑已安装 7-Zip，或在界面指定 7z.exe 路径。"
                    self._log(f"❌ {msg}", "ERROR")
                    self.finished_signal.emit(False, {"error": msg})
                    return

                mode_desc = (
                    "每个文件独立加密为 .7z 压缩卷上传"
                    if sevenzip_mode == "single_file"
                    else "每个子文件夹内的文件合并加密为 1 个 .7z 压缩卷上传"
                )
                self._log(
                    f"🔒 【7z 自动加密已开启】打包方式: {mode_desc} | 强制加密文件名 (-mhe=on) | 7z引擎: {sevenzip_exe}",
                    "SUCCESS",
                )

            # 1. 扫描本地可用文件
            available_files = scan_available_files(
                source_dir=source_dir,
                uploaded_folder_name=uploaded_folder_name,
                sort_mode=sort_mode,
                ext_filter=ext_filter,
                recursive=recursive_scan,
            )
            available_count = len(available_files)

            self._log(
                f"📊 计划配置: 需要创建 【{planned_folders}】 个云端文件夹 × 每个文件夹上传 【{files_per_folder}】 个文件 = 计划上传总量 【{planned_total_files}】 个文件",
                "INFO",
            )
            self._log(
                f"🔍 本地扫描结果: 当前源目录共有 【{available_count}】 个未上传的可用文件（已自动排除「{uploaded_folder_name}」归档目录）",
                "INFO",
            )

            if available_count == 0:
                msg = f"本地目录中没有可上传的文件（0 个）！请检查源目录或确认文件是否已全部移入「{uploaded_folder_name}」文件夹。"
                self._log(f"⚠️ {msg}", "WARN")
                self.finished_signal.emit(False, {"error": msg})
                return

            # 2. 严格判断数量关系（多的不上传，少的绝不重复上传，日志详细说明）
            summary_note = ""
            if available_count > planned_total_files:
                extra_count = available_count - planned_total_files
                selected_files = available_files[:planned_total_files]
                summary_note = (
                    f"本地可用文件 {available_count} 个 > 计划总量 {planned_total_files} 个（{planned_folders}夹×{files_per_folder}个）。"
                    f"已严格按要求仅上传前 {planned_total_files} 个文件；多余的 {extra_count} 个文件未上传，保留在原目录。"
                )
                self._log(f"ℹ️ 【多的不上传说明】{summary_note}", "WARN")
            elif available_count < planned_total_files:
                shortfall = planned_total_files - available_count
                selected_files = available_files[:]
                full_f_count = available_count // files_per_folder
                rem_files = available_count % files_per_folder
                actual_f_count = full_f_count + (1 if rem_files > 0 else 0)
                uncreated_f_count = planned_folders - actual_f_count

                detail_parts = []
                if full_f_count > 0:
                    detail_parts.append(
                        f"前 {full_f_count} 个文件夹各装满 {files_per_folder} 个文件"
                    )
                if rem_files > 0:
                    detail_parts.append(
                        f"第 {actual_f_count} 个文件夹装入剩余 {rem_files} 个文件（不满 {files_per_folder} 个，绝不重复上传凑数）"
                    )
                if uncreated_f_count > 0:
                    detail_parts.append(
                        f"剩余计划的 {uncreated_f_count} 个文件夹因无更多本地文件不再创建"
                    )
                detail_str = "；".join(detail_parts)

                summary_note = (
                    f"本地可用文件仅 {available_count} 个 < 计划总量 {planned_total_files} 个（差 {shortfall} 个）。"
                    f"已严格遵守「少的也不重复上传」原则，仅将现有 {available_count} 个文件唯一上传：{detail_str}。"
                )
                self._log(f"⚠️ 【少的不重复上传说明】{summary_note}", "WARN")
            else:
                selected_files = available_files[:]
                summary_note = (
                    f"本地可用文件 {available_count} 个，与计划总量 {planned_total_files} 个（{planned_folders}夹×{files_per_folder}个）完全匹配，全部按序上传。"
                )
                self._log(f"✅ 【数量完全匹配】{summary_note}", "SUCCESS")

            file_groups: List[List[Path]] = []
            for i in range(0, len(selected_files), files_per_folder):
                file_groups.append(selected_files[i : i + files_per_folder])

            actual_folder_count = len(file_groups)

            subfolder_names = generate_subfolder_names(
                count=actual_folder_count,
                naming_mode=self.config.get("naming_mode", "prefix_num"),
                prefix=self.config.get("folder_prefix", ""),
                start_index=int(self.config.get("start_index", 1)),
                zero_pad=int(self.config.get("zero_pad", 3)),
                custom_template=self.config.get("custom_template", "{prefix}{num}"),
                custom_list_text=self.config.get("custom_list_text", ""),
            )

            # 3. 初始化并认证 Google 客户端
            client = GoogleCloudClient(
                auth_mode=self.config.get("auth_mode", "oauth"),
                cred_file=self.config.get("cred_file", ""),
                token_file=self.config.get("token_file", "token.json"),
                proxy=self.config.get("proxy", ""),
                log_callback=self._log,
            )

            auth_ok, auth_msg = client.authenticate()
            if not auth_ok:
                self.finished_signal.emit(False, {"error": auth_msg})
                return

            # 4. 准备云端总文件夹 (Master Folder)
            master_mode = self.config.get("master_folder_mode", "create_new")
            master_input = self.config.get("master_folder_input", "").strip()
            parent_drive_id = self.config.get("parent_drive_id", "").strip()

            if master_mode == "use_existing" and master_input:
                self._log(f"正在获取已有云端总文件夹信息: {master_input} ...", "INFO")
                master_info = client.get_folder_info(master_input)
                client.set_public_permission(master_info["id"])
                self._log(
                    f"📁 已绑定云端总文件夹: 「{master_info['name']}」 | ID: {master_info['id']} | 链接: {master_info['webViewLink']}",
                    "SUCCESS",
                )
            else:
                new_master_name = (
                    master_input
                    if master_input
                    else f"批量上传总目录_{start_dt.strftime('%Y%m%d_%H%M%S')}"
                )
                self._log(f"正在云端创建总文件夹: 「{new_master_name}」...", "INFO")
                master_info = client.create_folder(
                    folder_name=new_master_name,
                    parent_id=parent_drive_id if parent_drive_id else None,
                    make_public=True,
                )
                self._log(
                    f"📁 已创建云端总文件夹: 「{master_info['name']}」 | ID: {master_info['id']} | 链接: {master_info['webViewLink']}",
                    "SUCCESS",
                )

            if enable_7z:
                staging_root = Path(source_dir) / ".7z_staging_temp"
                staging_root.mkdir(parents=True, exist_ok=True)

            # 5. 按文件夹分组逐个创建子文件夹并（加密）上传文件
            total_to_upload = len(selected_files)
            uploaded_so_far = 0
            subfolders_result: List[Dict[str, Any]] = []

            for folder_idx, (sf_name, group_files) in enumerate(
                zip(subfolder_names, file_groups), start=1
            ):
                if self._stop_requested:
                    self._log("🛑 用户手动停止了上传任务！已上传的文件均已安全归档。", "WARN")
                    break

                self._log("-" * 56, "INFO")
                self._log(
                    f"📂 [{folder_idx}/{actual_folder_count}] 正在创建云端子文件夹: 「{sf_name}」 (本夹分配 {len(group_files)} 个本地文件)...",
                    "INFO",
                )
                sf_info = client.create_folder(
                    folder_name=sf_name,
                    parent_id=master_info["id"],
                    make_public=True,
                )
                self._log(
                    f"   ✓ 子文件夹创建成功: 「{sf_info['name']}」 | ID: {sf_info['id']} | 链接: {sf_info['webViewLink']}",
                    "SUCCESS",
                )

                sf_record = {
                    "index": folder_idx,
                    "id": sf_info["id"],
                    "name": sf_info["name"],
                    "webViewLink": sf_info["webViewLink"],
                    "files": [],
                }
                subfolders_result.append(sf_record)

                # 分支 A：开启了 7z 加密 且选择了「整夹合并为 1 个 .7z 压缩卷上传」
                if enable_7z and sevenzip_mode == "folder_bundle":
                    safe_sf_name = (
                        re.sub(r'[<>:"/\\|?*]', "_", sf_name).strip()
                        or f"folder_{folder_idx}"
                    )
                    bundle_7z_path = staging_root / f"{safe_sf_name}.7z"
                    self.progress_signal.emit(
                        uploaded_so_far,
                        total_to_upload,
                        f"正在 7z 加密打包子文件夹「{sf_name}」({len(group_files)}个文件)...",
                    )
                    self._log(
                        f"   🔒 正在将本夹 {len(group_files)} 个文件合并执行 7z 加密（含加密文件名 -mhe=on） -> {bundle_7z_path.name} ...",
                        "INFO",
                    )
                    ok_7z, msg_7z, out_7z = create_encrypted_7z(
                        input_files=group_files,
                        output_7z_path=bundle_7z_path,
                        password=sevenzip_password,
                        sevenzip_exe=sevenzip_exe,
                        compression_level=sevenzip_level,
                    )
                    if not ok_7z:
                        raise RuntimeError(msg_7z)

                    self._log(
                        f"   ⬆️ 正在上传 7z 加密压缩卷: {out_7z.name} ({format_size(out_7z.stat().st_size)})...",
                        "INFO",
                    )

                    def _bundle_cb(done_b: int, total_b: int):
                        pct = int(done_b * 100 / max(total_b, 1))
                        self.progress_signal.emit(
                            uploaded_so_far,
                            total_to_upload,
                            f"正在上传加密卷 {out_7z.name} ({pct}%)",
                        )

                    up_res = client.upload_file(
                        local_path=str(out_7z),
                        parent_folder_id=sf_info["id"],
                        make_public=True,
                        progress_callback=_bundle_cb,
                    )
                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    pwd_note = (
                        f" | 7z密码:{sevenzip_password}"
                        if record_pwd_in_sheet
                        else ""
                    )
                    perm_label = f"{up_res['permission_status']} | 7z加密含文件名(内含{len(group_files)}个文件){pwd_note}"

                    sf_record["files"].append(
                        {
                            "index_in_folder": 1,
                            "global_index": folder_idx,
                            "name": up_res["name"],
                            "id": up_res["id"],
                            "downloadLink": up_res["downloadLink"],
                            "webViewLink": up_res["webViewLink"],
                            "size_bytes": up_res["size_bytes"],
                            "size_str": up_res["size_str"],
                            "permission_status": perm_label,
                            "upload_time": now_str,
                        }
                    )
                    self._log(
                        f"      ✅ 7z压缩卷上传成功 & 权限已开 | 文件ID: {up_res['id']} | 下载直链: {up_res['downloadLink']}",
                        "SUCCESS",
                    )

                    for src_f in group_files:
                        moved_dest = self._move_to_uploaded_archive(
                            local_file=src_f,
                            source_dir=source_dir,
                            uploaded_folder_name=uploaded_folder_name,
                            subfolder_name=sf_name,
                            group_by_subfolder=group_archive_by_subfolder,
                        )
                        rel_moved = moved_dest.relative_to(Path(source_dir))
                        self._log(f"      📦 原文件已移入归档防重: {rel_moved}", "INFO")
                        uploaded_so_far += 1

                    if keep_encrypted_copy:
                        self._move_to_uploaded_archive(
                            local_file=out_7z,
                            source_dir=source_dir,
                            uploaded_folder_name=uploaded_folder_name,
                            subfolder_name=sf_name,
                            group_by_subfolder=group_archive_by_subfolder,
                        )
                    elif out_7z.exists():
                        out_7z.unlink()

                    self.progress_signal.emit(
                        uploaded_so_far,
                        total_to_upload,
                        f"已完成 {uploaded_so_far}/{total_to_upload} (子文件夹 {sf_name})",
                    )
                    continue

                # 分支 B（默认）：逐个文件处理
                for file_idx_in_folder, local_file_path in enumerate(
                    group_files, start=1
                ):
                    if self._stop_requested:
                        self._log("🛑 任务停止请求已响应，终止后续文件上传。", "WARN")
                        break

                    global_idx = uploaded_so_far + 1
                    upload_target_path = local_file_path
                    temp_7z_created: Optional[Path] = None

                    if enable_7z:
                        safe_sf = (
                            re.sub(r'[<>:"/\\|?*]', "_", sf_name).strip()
                            or f"sf_{folder_idx}"
                        )
                        archive_name = f"{local_file_path.stem}.7z"
                        temp_7z_path = (
                            staging_root
                            / safe_sf
                            / f"{file_idx_in_folder:03d}_{archive_name}"
                        )
                        self.progress_signal.emit(
                            uploaded_so_far,
                            total_to_upload,
                            f"正在 7z 加密 [{global_idx}/{total_to_upload}]: {local_file_path.name} -> {archive_name}",
                        )
                        self._log(
                            f"   🔒 [{file_idx_in_folder}/{len(group_files)}] 正在执行 7z 加密（含加密文件名 -mhe=on）: {local_file_path.name} -> {archive_name} ...",
                            "INFO",
                        )
                        ok_7z, msg_7z, out_7z = create_encrypted_7z(
                            input_files=[local_file_path],
                            output_7z_path=temp_7z_path,
                            password=sevenzip_password,
                            sevenzip_exe=sevenzip_exe,
                            compression_level=sevenzip_level,
                        )
                        if not ok_7z:
                            raise RuntimeError(msg_7z)

                        upload_target_path = out_7z
                        temp_7z_created = out_7z
                        cloud_display_name = archive_name
                    else:
                        cloud_display_name = local_file_path.name

                    status_label = (
                        f"正在上传 [{global_idx}/{total_to_upload}] -> 子文件夹「{sf_name}」({file_idx_in_folder}/{len(group_files)}): {cloud_display_name}"
                    )
                    self.progress_signal.emit(
                        uploaded_so_far, total_to_upload, status_label
                    )
                    self._log(
                        f"   ⬆️ [{file_idx_in_folder}/{len(group_files)}] 正在上传文件: {cloud_display_name} ({format_size(upload_target_path.stat().st_size)})...",
                        "INFO",
                    )

                    def _chunk_cb(done_b: int, total_b: int):
                        pct = int(done_b * 100 / max(total_b, 1))
                        self.progress_signal.emit(
                            uploaded_so_far,
                            total_to_upload,
                            f"{status_label} ({pct}%)",
                        )

                    up_res = client.upload_file(
                        local_path=str(upload_target_path),
                        parent_folder_id=sf_info["id"],
                        custom_name=cloud_display_name,
                        make_public=True,
                        progress_callback=_chunk_cb,
                    )

                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if enable_7z:
                        pwd_note = (
                            f" | 密码:{sevenzip_password}"
                            if record_pwd_in_sheet
                            else ""
                        )
                        perm_status_str = f"{up_res['permission_status']} | 7z加密含文件名(原文件:{local_file_path.name}){pwd_note}"
                    else:
                        perm_status_str = up_res["permission_status"]

                    file_record = {
                        "index_in_folder": file_idx_in_folder,
                        "global_index": global_idx,
                        "name": up_res["name"],
                        "id": up_res["id"],
                        "downloadLink": up_res["downloadLink"],
                        "webViewLink": up_res["webViewLink"],
                        "size_bytes": up_res["size_bytes"],
                        "size_str": up_res["size_str"],
                        "permission_status": perm_status_str,
                        "upload_time": now_str,
                    }
                    sf_record["files"].append(file_record)

                    self._log(
                        f"      ✅ 上传成功 & 权限已开 | 云端文件名: {up_res['name']} | 文件ID: {up_res['id']} | 下载直链: {up_res['downloadLink']}",
                        "SUCCESS",
                    )

                    moved_dest = self._move_to_uploaded_archive(
                        local_file=local_file_path,
                        source_dir=source_dir,
                        uploaded_folder_name=uploaded_folder_name,
                        subfolder_name=sf_name,
                        group_by_subfolder=group_archive_by_subfolder,
                    )
                    rel_moved = moved_dest.relative_to(Path(source_dir))
                    self._log(
                        f"      📦 本地原文件已移入归档目录防重: {rel_moved}",
                        "INFO",
                    )

                    if temp_7z_created and temp_7z_created.exists():
                        if keep_encrypted_copy:
                            clean_7z = temp_7z_created.parent / cloud_display_name
                            if clean_7z != temp_7z_created:
                                temp_7z_created.replace(clean_7z)
                            self._move_to_uploaded_archive(
                                local_file=clean_7z,
                                source_dir=source_dir,
                                uploaded_folder_name=uploaded_folder_name,
                                subfolder_name=sf_name,
                                group_by_subfolder=group_archive_by_subfolder,
                            )
                        else:
                            try:
                                temp_7z_created.unlink()
                            except Exception:
                                pass

                    uploaded_so_far += 1
                    self.progress_signal.emit(
                        uploaded_so_far,
                        total_to_upload,
                        f"已完成 {uploaded_so_far}/{total_to_upload}: {cloud_display_name}",
                    )

            if staging_root and staging_root.exists():
                shutil.rmtree(staging_root, ignore_errors=True)

            # 6. 汇总批次结果并导出本地 Excel 与 云端 Google Sheets
            batch_result = {
                "batch_time": batch_time_str,
                "master_folder": master_info,
                "planned_folders": planned_folders,
                "planned_files_per_folder": files_per_folder,
                "planned_total_files": planned_total_files,
                "available_local_files": available_count,
                "total_uploaded_files": uploaded_so_far,
                "summary_note": summary_note,
                "subfolders": subfolders_result,
                "excel_path": "",
                "gsheet_url": "",
                "enable_7z": enable_7z,
            }

            exporter = HierarchyReportExporter(log_callback=self._log)
            merge_cells = bool(self.config.get("merge_hierarchy_cells", True))

            report_dir = Path(source_dir) / "上传报表_Reports"
            report_dir.mkdir(parents=True, exist_ok=True)
            excel_filename = f"云端上传层级表_{master_info['name']}_{file_ts_str}.xlsx"
            safe_excel_filename = re.sub(r'[<>:"/\\|?*]', "_", excel_filename)
            excel_path = str(report_dir / safe_excel_filename)

            self._log("-" * 56, "INFO")
            self._log("正在生成按层级排版的本地 Excel 备份表格...", "INFO")
            exporter.export_local_excel(
                batch_result=batch_result,
                output_path=excel_path,
                merge_hierarchy_cells=merge_cells,
            )
            batch_result["excel_path"] = excel_path

            sheet_input = self.config.get("google_sheet_input", "").strip()
            new_sheet_title = f"云端上传层级表_{master_info['name']}"
            gs_ok, gs_url, gs_msg = exporter.write_to_google_sheet(
                google_client=client,
                batch_result=batch_result,
                sheet_id_or_url=sheet_input,
                new_sheet_title=new_sheet_title,
                merge_hierarchy_cells=merge_cells,
            )
            if gs_ok:
                batch_result["gsheet_url"] = gs_url
            else:
                self._log(
                    f"⚠️ 谷歌在线表格写入提醒: {gs_msg}（所有数据已完整保存在本地 Excel 表格中：{excel_path}）",
                    "WARN",
                )

            # 7. 最终总结日志
            self._log("=" * 68, "SUCCESS")
            self._log("🎉 批量上传与层级制表任务执行完毕！最终统计清单：", "SUCCESS")
            self._log(
                f"   • 云端总文件夹: 「{master_info['name']}」 | ID: {master_info['id']}",
                "SUCCESS",
            )
            self._log(f"   • 总文件夹链接: {master_info['webViewLink']}", "SUCCESS")
            if enable_7z:
                self._log(
                    "   • 7z 加密状态: 已启用自动 7z 加密（AES-256 + 强制加密文件名 -mhe=on），云端全部为加密 .7z 压缩卷",
                    "SUCCESS",
                )
            self._log(
                f"   • 计划配置总量: {planned_folders} 个文件夹 × 每夹 {files_per_folder} 个 = {planned_total_files} 个文件",
                "INFO",
            )
            self._log(
                f"   • 实际创建子夹: {len(subfolders_result)} 个子文件夹 | 实际处理本地文件: {uploaded_so_far} 个文件",
                "SUCCESS",
            )
            self._log(f"   • 数量控制说明: {summary_note}", "INFO")
            self._log(f"   • 本地备份表格: {excel_path}", "SUCCESS")
            if batch_result["gsheet_url"]:
                self._log(
                    f"   • 谷歌在线表格: {batch_result['gsheet_url']}", "SUCCESS"
                )
            self._log("=" * 68, "SUCCESS")

            self.finished_signal.emit(True, batch_result)

        except Exception as e:
            if staging_root and staging_root.exists():
                shutil.rmtree(staging_root, ignore_errors=True)
            err_text = f"任务执行过程中发生异常: {str(e)}"
            self._log(f"❌ {err_text}", "ERROR")
            self.finished_signal.emit(False, {"error": err_text})

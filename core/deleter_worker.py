import datetime
from typing import List, Dict, Any
from PyQt6.QtCore import QThread, pyqtSignal

from core.google_client import GoogleCloudClient


class BatchDeleteWorker(QThread):
    """
    后台批量删除云端链接/ID 工作线程：
    支持：
    - 彻底永久删除 (permanent)
    - 移入云盘回收站 (trash)
    - 仅关闭公开下载权限 (revoke_public)
    能够自动识别单文件直接下载链接、单文件预览链接、文件夹链接、纯ID，
    并在删除完毕后输出详细统计报告。
    """

    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int, int, str)
    item_result_signal = pyqtSignal(int, str, str, str, str)
    finished_signal = pyqtSignal(bool, dict)

    def __init__(
        self, config: Dict[str, Any], targets: List[Dict[str, str]], parent=None
    ):
        super().__init__(parent)
        self.config = config
        self.targets = targets
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def _log(self, msg: str, level: str = "INFO"):
        self.log_signal.emit(msg, level)

    def run(self):
        start_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        delete_mode = self.config.get("delete_mode", "permanent")
        fetch_name = bool(self.config.get("fetch_name_first", True))

        mode_names = {
            "permanent": "🔥 彻底永久删除（立即释放云盘空间）",
            "trash": "♻️ 移入云端回收站（30天内可恢复）",
            "revoke_public": "🔒 仅关闭公开下载权限（保留文件但外链失效）",
        }
        mode_desc = mode_names.get(delete_mode, delete_mode)

        try:
            total = len(self.targets)
            self._log("=" * 68, "INFO")
            self._log(
                f"🗑️ 开始执行云端链接批量清理任务 | 时间: {start_str}",
                "INFO",
            )
            self._log(
                f"📌 清理模式: {mode_desc} | 本次待处理去重目标数: 【{total}】 个",
                "INFO",
            )

            if total == 0:
                self.finished_signal.emit(
                    False, {"error": "没有解析到有效的链接或 ID"}
                )
                return

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

            success_count = 0
            not_found_count = 0
            error_count = 0
            succeeded_ids: List[str] = []

            for idx, item in enumerate(self.targets):
                if self._stop_requested:
                    self._log("🛑 用户手动停止了批量删除任务！", "WARN")
                    break

                fid = item["id"]
                type_hint = item.get("type_hint", "链接/ID")
                self.progress_signal.emit(
                    idx,
                    total,
                    f"正在处理 [{idx + 1}/{total}] ({type_hint}): {fid} ...",
                )

                status_code, item_name, detail_msg = client.delete_cloud_item(
                    file_or_folder_id=fid,
                    delete_mode=delete_mode,
                    fetch_name_first=fetch_name,
                )

                self.item_result_signal.emit(
                    idx, fid, item_name, status_code, detail_msg
                )

                if status_code == "SUCCESS":
                    success_count += 1
                    succeeded_ids.append(fid)
                    self._log(
                        f"   ✅ [{idx + 1}/{total}] ID: {fid} -> {detail_msg}",
                        "SUCCESS",
                    )
                elif status_code == "NOT_FOUND":
                    not_found_count += 1
                    succeeded_ids.append(fid)
                    self._log(
                        f"   ⚪ [{idx + 1}/{total}] ID: {fid} -> {detail_msg}",
                        "WARN",
                    )
                else:
                    error_count += 1
                    self._log(
                        f"   ❌ [{idx + 1}/{total}] ID: {fid} -> 失败: {detail_msg}",
                        "ERROR",
                    )

                self.progress_signal.emit(
                    idx + 1,
                    total,
                    f"已处理 {idx + 1}/{total}: 成功 {success_count} | 已不存在/跳过 {not_found_count} | 失败 {error_count}",
                )

            summary = {
                "total": total,
                "success_count": success_count,
                "not_found_count": not_found_count,
                "error_count": error_count,
                "succeeded_ids": succeeded_ids,
                "delete_mode": delete_mode,
                "mode_desc": mode_desc,
            }

            self._log("=" * 68, "SUCCESS")
            self._log(
                f"🎉 批量删除/清理任务执行结束！总计 {total} 个目标：成功处理 {success_count} 个，云端已不存在(自动跳过) {not_found_count} 个，失败 {error_count} 个。",
                "SUCCESS" if error_count == 0 else "WARN",
            )
            self._log("=" * 68, "SUCCESS")

            self.finished_signal.emit(True, summary)

        except Exception as e:
            err_msg = f"批量删除任务发生异常: {str(e)}"
            self._log(f"❌ {err_msg}", "ERROR")
            self.finished_signal.emit(False, {"error": err_msg})

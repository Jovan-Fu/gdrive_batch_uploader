import os
import re
import json
import time
import mimetypes
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List, Tuple

import requests
from google.oauth2.credentials import Credentials as OAuthCredentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request, AuthorizedSession
import gspread

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

DRIVE_API_V3 = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD_V3 = "https://www.googleapis.com/upload/drive/v3"


def extract_google_id(url_or_id: str) -> str:
    """
    从 Google Drive 文件夹链接、文件链接 或 Google Sheets 链接中提取 ID。
    如果传入的本身就是纯 ID，则直接返回去除首尾空格后的字符串。
    """
    text = (url_or_id or "").strip()
    if not text:
        return ""

    patterns = [
        r"/folders/([a-zA-Z0-9_-]{10,})",
        r"/spreadsheets/d/([a-zA-Z0-9_-]{10,})",
        r"/file/d/([a-zA-Z0-9_-]{10,})",
        r"[?&]id=([a-zA-Z0-9_-]{10,})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return m.group(1)

    if "/" not in text and len(text) >= 10:
        return text

    return text


def parse_drive_targets_from_text(raw_text: str) -> List[Dict[str, str]]:
    """
    从用户粘贴的多行文本、表格复制内容或链接列表中，智能提取所有 Google Drive 文件/文件夹链接与 ID，
    自动按 ID 去重并保持原有顺序。
    支持：
    - 单文件直接下载直链: https://drive.google.com/uc?export=download&id=xxx
    - 单文件在线预览链接: https://drive.google.com/file/d/xxx/view?usp=sharing
    - 总文件夹/子文件夹链接: https://drive.google.com/drive/folders/xxx?usp=sharing
    - 纯文件/文件夹 ID (单独成行或列)
    """
    if not raw_text:
        return []

    seen_ids = set()
    results: List[Dict[str, str]] = []

    url_patterns = [
        (
            "文件夹链接",
            re.compile(
                r"https?://drive\.google\.com/drive/(?:u/\d+/)?folders/([a-zA-Z0-9_-]{10,})[^\s\"\'\<\>]*"
            ),
        ),
        (
            "单文件下载直链",
            re.compile(
                r"https?://drive\.google\.com/uc\?[^\s\"\'\<\>]*?id=([a-zA-Z0-9_-]{10,})[^\s\"\'\<\>]*"
            ),
        ),
        (
            "单文件预览链接",
            re.compile(
                r"https?://drive\.google\.com/file/d/([a-zA-Z0-9_-]{10,})[^\s\"\'\<\>]*"
            ),
        ),
        (
            "云端链接",
            re.compile(
                r"https?://drive\.google\.com/open\?[^\s\"\'\<\>]*?id=([a-zA-Z0-9_-]{10,})[^\s\"\'\<\>]*"
            ),
        ),
    ]

    for line_idx, raw_line in enumerate(raw_text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        matched_in_line = False
        for kind_label, pat in url_patterns:
            for m in pat.finditer(line):
                fid = m.group(1)
                full_url = m.group(0)
                matched_in_line = True
                if fid not in seen_ids:
                    seen_ids.add(fid)
                    results.append(
                        {
                            "id": fid,
                            "type_hint": kind_label,
                            "raw_input": full_url,
                            "line_no": line_idx,
                        }
                    )

        if not matched_in_line:
            tokens = re.split(r"[\t,;，；\s]+", line)
            for tok in tokens:
                tok = tok.strip()
                if re.fullmatch(r"[a-zA-Z0-9_-]{15,60}", tok):
                    if tok.isdigit() or tok.isalpha():
                        continue
                    if tok not in seen_ids:
                        seen_ids.add(tok)
                        results.append(
                            {
                                "id": tok,
                                "type_hint": "纯云端ID",
                                "raw_input": tok,
                                "line_no": line_idx,
                            }
                        )

    return results


def format_size(size_bytes: int) -> str:
    """格式化字节大小为人类可读字符串"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def format_google_api_error(status_code: int, resp_text: str) -> str:
    """将 Google API 返回的冗长 JSON 错误解析为简明易懂的中文提示"""
    try:
        data = json.loads(resp_text)
        err_obj = data.get("error", {})
        msg = err_obj.get("message", "")

        if status_code == 403 and (
            "has not been used in project" in msg
            or "is disabled" in msg
            or "accessNotConfigured" in resp_text
            or "SERVICE_DISABLED" in resp_text
        ):
            proj_match = re.search(r"project\s+(\d+)", msg)
            proj_id = proj_match.group(1) if proj_match else ""
            drive_url = (
                f"https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project={proj_id}"
                if proj_id
                else "https://console.cloud.google.com/apis/library/drive.googleapis.com"
            )
            sheets_url = (
                f"https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project={proj_id}"
                if proj_id
                else "https://console.cloud.google.com/apis/library/sheets.googleapis.com"
            )
            return (
                f"您的谷歌云项目 ({proj_id or '当前项目'}) 尚未开启 Google Drive API 或 Google Sheets API！\n"
                f"请点击以下链接在浏览器中点击【启用 (Enable)】，等待1分钟后再试：\n"
                f"1. 开启 Google Drive API: {drive_url}\n"
                f"2. 开启 Google Sheets API: {sheets_url}"
            )

        if msg:
            return f"HTTP {status_code}: {msg}"
    except Exception:
        pass

    clean_one_line = " ".join((resp_text or "").split())
    return f"HTTP {status_code}: {clean_one_line[:240]}"


class GoogleCloudClient:
    """封装 Google Drive API v3 与 Google Sheets API v4 操作。"""

    def __init__(
        self,
        auth_mode: str = "oauth",
        cred_file: str = "",
        token_file: str = "token.json",
        proxy: str = "",
        log_callback: Optional[Callable[[str, str], None]] = None,
    ):
        self.auth_mode = auth_mode
        self.cred_file = cred_file
        self.token_file = token_file
        self.proxy = (proxy or "").strip()
        self.log_callback = log_callback

        self.creds = None
        self.session: Optional[AuthorizedSession] = None
        self.gspread_client: Optional[gspread.Client] = None

    def _log(self, msg: str, level: str = "INFO"):
        if self.log_callback:
            self.log_callback(msg, level)

    def _apply_proxy_env(self):
        if self.proxy:
            os.environ["HTTP_PROXY"] = self.proxy
            os.environ["HTTPS_PROXY"] = self.proxy
            os.environ["http_proxy"] = self.proxy
            os.environ["https_proxy"] = self.proxy

    def authenticate(self) -> Tuple[bool, str]:
        """执行认证并初始化 AuthorizedSession 和 gspread 客户端"""
        try:
            self._apply_proxy_env()

            if not self.cred_file or not os.path.isfile(self.cred_file):
                msg = f"找不到凭证文件：{self.cred_file or '未选择文件'}"
                self._log(msg, "ERROR")
                return False, msg

            if self.auth_mode == "service_account":
                self._log("正在使用 Service Account (服务账号) 密钥进行认证...", "INFO")
                self.creds = service_account.Credentials.from_service_account_file(
                    self.cred_file, scopes=SCOPES
                )
            else:
                self._log("正在检查 OAuth 2.0 授权状态...", "INFO")
                creds = None
                if self.token_file and os.path.exists(self.token_file):
                    try:
                        creds = OAuthCredentials.from_authorized_user_file(
                            self.token_file, SCOPES
                        )
                    except Exception as e:
                        self._log(f"读取本地缓存 token 失败，将重新授权: {e}", "WARN")
                        creds = None

                if not creds or not creds.valid:
                    if creds and creds.expired and creds.refresh_token:
                        self._log("正在刷新 OAuth 2.0 访问令牌...", "INFO")
                        req_session = requests.Session()
                        if self.proxy:
                            req_session.proxies = {
                                "http": self.proxy,
                                "https": self.proxy,
                            }
                        creds.refresh(Request(session=req_session))
                    else:
                        self._log(
                            "首次授权或令牌已失效，正在打开浏览器进行谷歌账号登录授权（限时120秒）...",
                            "INFO",
                        )
                        flow = InstalledAppFlow.from_client_secrets_file(
                            self.cred_file, SCOPES
                        )
                        creds = flow.run_local_server(
                            port=0, prompt="consent", timeout_seconds=120
                        )

                    if self.token_file:
                        with open(self.token_file, "w", encoding="utf-8") as f:
                            f.write(creds.to_json())
                        self._log(f"已保存授权令牌到 {self.token_file}", "SUCCESS")

                self.creds = creds

            self.session = AuthorizedSession(self.creds)
            if self.proxy:
                self.session.proxies = {"http": self.proxy, "https": self.proxy}

            self.gspread_client = gspread.Client(auth=self.creds, session=self.session)

            about_resp = self.session.get(
                f"{DRIVE_API_V3}/about",
                params={"fields": "user(displayName,emailAddress)"},
                timeout=20,
            )
            if about_resp.status_code == 200:
                user_info = about_resp.json().get("user", {})
                email = user_info.get("emailAddress", "未知邮箱")
                name = user_info.get("displayName", "")
                account_str = f"{name} <{email}>" if name else email
                self._log(f"谷歌云端授权成功！当前账号: {account_str}", "SUCCESS")
                return True, f"连接成功！当前账号: {account_str}"
            else:
                friendly_err = format_google_api_error(
                    about_resp.status_code, about_resp.text
                )
                for line in friendly_err.splitlines():
                    self._log(line, "ERROR")
                return False, friendly_err

        except Exception as e:
            err_msg = f"谷歌授权或网络请求失败: {str(e)}"
            self._log(err_msg, "ERROR")
            return False, err_msg

    def get_folder_info(self, folder_id_or_url: str) -> Dict[str, str]:
        """获取已有云端文件夹的名称、ID、链接"""
        folder_id = extract_google_id(folder_id_or_url)
        if not folder_id:
            raise ValueError("无效的云端文件夹 ID 或链接")

        resp = self.session.get(
            f"{DRIVE_API_V3}/files/{folder_id}",
            params={
                "fields": "id,name,mimeType,webViewLink",
                "supportsAllDrives": "true",
            },
            timeout=20,
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"无法获取云端文件夹 ({folder_id}) 信息: {format_google_api_error(resp.status_code, resp.text)}"
            )
        data = resp.json()
        fid = data["id"]
        web_link = data.get(
            "webViewLink", f"https://drive.google.com/drive/folders/{fid}?usp=sharing"
        )
        return {
            "id": fid,
            "name": data.get("name", fid),
            "webViewLink": web_link,
        }

    def create_folder(
        self,
        folder_name: str,
        parent_id: Optional[str] = None,
        make_public: bool = True,
    ) -> Dict[str, str]:
        """在 Google Drive 创建文件夹，可选开启公开查看权限，并返回 id、name、webViewLink。"""
        metadata: Dict[str, Any] = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
        }
        if parent_id:
            clean_parent = extract_google_id(parent_id)
            if clean_parent:
                metadata["parents"] = [clean_parent]

        resp = self.session.post(
            f"{DRIVE_API_V3}/files",
            params={
                "fields": "id,name,webViewLink",
                "supportsAllDrives": "true",
            },
            json=metadata,
            timeout=30,
        )
        if resp.status_code not in (200, 201):
            raise RuntimeError(
                f"创建云端文件夹 '{folder_name}' 失败: {format_google_api_error(resp.status_code, resp.text)}"
            )

        data = resp.json()
        folder_id = data["id"]
        web_link = data.get(
            "webViewLink",
            f"https://drive.google.com/drive/folders/{folder_id}?usp=sharing",
        )

        if make_public:
            self.set_public_permission(folder_id)

        return {
            "id": folder_id,
            "name": data.get("name", folder_name),
            "webViewLink": web_link,
        }

    def set_public_permission(self, file_or_folder_id: str) -> Tuple[bool, str]:
        """为文件或文件夹开启公开读取/下载权限 (anyone -> reader)"""
        for attempt in range(3):
            try:
                resp = self.session.post(
                    f"{DRIVE_API_V3}/files/{file_or_folder_id}/permissions",
                    params={"supportsAllDrives": "true"},
                    json={"role": "reader", "type": "anyone"},
                    timeout=25,
                )
                if resp.status_code in (200, 201):
                    return True, "已开启公开下载权限"
                elif (
                    resp.status_code == 403
                    and "cannotModifyInheritedPermission" in resp.text
                ):
                    return True, "继承上级公开权限"
                else:
                    if attempt == 2:
                        return (
                            False,
                            f"权限设置返回 {format_google_api_error(resp.status_code, resp.text)}",
                        )
            except Exception as e:
                if attempt == 2:
                    return False, f"权限设置异常: {e}"
                time.sleep(1.5)
        return False, "权限设置未知状态"

    def upload_file(
        self,
        local_path: str,
        parent_folder_id: str,
        custom_name: Optional[str] = None,
        make_public: bool = True,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> Dict[str, str]:
        """
        使用 Resumable Upload 上传单个本地文件到指定云端文件夹，开启下载权限，
        并返回文件的 ID、直接下载链接 (downloadLink) 与在线预览链接 (webViewLink)。
        """
        path_obj = Path(local_path)
        if not path_obj.is_file():
            raise FileNotFoundError(f"本地文件不存在: {local_path}")

        file_name = custom_name or path_obj.name
        file_size = path_obj.stat().st_size
        mime_type, _ = mimetypes.guess_type(local_path)
        if not mime_type:
            mime_type = "application/octet-stream"

        metadata = {
            "name": file_name,
            "parents": [extract_google_id(parent_folder_id)],
        }

        init_headers = {
            "X-Upload-Content-Type": mime_type,
            "X-Upload-Content-Length": str(file_size),
            "Content-Type": "application/json; charset=UTF-8",
        }
        init_resp = self.session.post(
            f"{DRIVE_UPLOAD_V3}/files",
            params={
                "uploadType": "resumable",
                "fields": "id,name,size,mimeType,webViewLink,webContentLink",
                "supportsAllDrives": "true",
            },
            headers=init_headers,
            data=json.dumps(metadata).encode("utf-8"),
            timeout=30,
        )
        if init_resp.status_code != 200 or "Location" not in init_resp.headers:
            raise RuntimeError(
                f"初始化文件上传失败 ({file_name}): {format_google_api_error(init_resp.status_code, init_resp.text)}"
            )

        upload_url = init_resp.headers["Location"]

        chunk_size = 8 * 1024 * 1024
        uploaded_bytes = 0
        final_data = None

        with open(local_path, "rb") as f:
            if file_size == 0:
                put_resp = self.session.put(
                    upload_url,
                    headers={"Content-Length": "0"},
                    data=b"",
                    timeout=60,
                )
                if put_resp.status_code not in (200, 201):
                    raise RuntimeError(
                        f"空文件上传失败 ({file_name}): {format_google_api_error(put_resp.status_code, put_resp.text)}"
                    )
                final_data = put_resp.json()
            else:
                while uploaded_bytes < file_size:
                    chunk = f.read(chunk_size)
                    if not chunk:
                        break
                    chunk_len = len(chunk)
                    start_byte = uploaded_bytes
                    end_byte = uploaded_bytes + chunk_len - 1
                    content_range = f"bytes {start_byte}-{end_byte}/{file_size}"

                    for attempt in range(3):
                        try:
                            put_resp = self.session.put(
                                upload_url,
                                headers={
                                    "Content-Length": str(chunk_len),
                                    "Content-Range": content_range,
                                },
                                data=chunk,
                                timeout=180,
                            )
                            if put_resp.status_code in (200, 201, 308):
                                break
                            elif attempt == 2:
                                raise RuntimeError(
                                    f"上传分块失败: {format_google_api_error(put_resp.status_code, put_resp.text)}"
                                )
                        except Exception as e:
                            if attempt == 2:
                                raise
                            time.sleep(2 * (attempt + 1))

                    uploaded_bytes += chunk_len
                    if progress_callback:
                        progress_callback(uploaded_bytes, file_size)

                    if put_resp.status_code in (200, 201):
                        final_data = put_resp.json()

        if not final_data or "id" not in final_data:
            raise RuntimeError(f"上传完成但未返回文件 ID: {file_name}")

        file_id = final_data["id"]

        perm_ok, perm_msg = (True, "未开启")
        if make_public:
            perm_ok, perm_msg = self.set_public_permission(file_id)

        download_link = final_data.get(
            "webContentLink",
            f"https://drive.google.com/uc?export=download&id={file_id}",
        )
        if not download_link:
            download_link = f"https://drive.google.com/uc?export=download&id={file_id}"

        view_link = final_data.get(
            "webViewLink",
            f"https://drive.google.com/file/d/{file_id}/view?usp=sharing",
        )

        return {
            "id": file_id,
            "name": final_data.get("name", file_name),
            "size_bytes": file_size,
            "size_str": format_size(file_size),
            "downloadLink": download_link,
            "webViewLink": view_link,
            "permission_ok": perm_ok,
            "permission_status": perm_msg,
        }

    def delete_cloud_item(
        self,
        file_or_folder_id: str,
        delete_mode: str = "permanent",
        fetch_name_first: bool = True,
    ) -> Tuple[str, str, str]:
        """根据云端文件/文件夹 ID 执行删除、移入回收站或关闭权限操作。"""
        clean_id = extract_google_id(file_or_folder_id)
        if not clean_id:
            return "ERROR", "未知", "无效的 ID 或链接"

        item_name = clean_id
        item_kind = "文件/文件夹"

        if fetch_name_first:
            try:
                meta_resp = self.session.get(
                    f"{DRIVE_API_V3}/files/{clean_id}",
                    params={
                        "fields": "id,name,mimeType,trashed",
                        "supportsAllDrives": "true",
                    },
                    timeout=15,
                )
                if meta_resp.status_code == 404:
                    return (
                        "NOT_FOUND",
                        clean_id,
                        "云端已不存在该项（可能已被删除或随上级文件夹一并删除）",
                    )
                elif meta_resp.status_code == 200:
                    mdata = meta_resp.json()
                    item_name = mdata.get("name", clean_id)
                    mime = mdata.get("mimeType", "")
                    item_kind = (
                        "文件夹"
                        if mime == "application/vnd.google-apps.folder"
                        else "文件"
                    )
            except Exception:
                pass

        try:
            if delete_mode == "permanent":
                resp = self.session.delete(
                    f"{DRIVE_API_V3}/files/{clean_id}",
                    params={"supportsAllDrives": "true"},
                    timeout=25,
                )
                if resp.status_code in (200, 204):
                    return "SUCCESS", item_name, f"{item_kind}「{item_name}」已彻底永久删除"
                elif resp.status_code == 404:
                    return (
                        "NOT_FOUND",
                        item_name,
                        f"{item_kind}「{item_name}」在云端已不存在（跳过）",
                    )
                else:
                    return (
                        "ERROR",
                        item_name,
                        format_google_api_error(resp.status_code, resp.text),
                    )

            elif delete_mode == "trash":
                resp = self.session.patch(
                    f"{DRIVE_API_V3}/files/{clean_id}",
                    params={"supportsAllDrives": "true"},
                    json={"trashed": True},
                    timeout=25,
                )
                if resp.status_code == 200:
                    return "SUCCESS", item_name, f"{item_kind}「{item_name}」已移入云端回收站"
                elif resp.status_code == 404:
                    return (
                        "NOT_FOUND",
                        item_name,
                        f"{item_kind}「{item_name}」在云端已不存在（跳过）",
                    )
                else:
                    return (
                        "ERROR",
                        item_name,
                        format_google_api_error(resp.status_code, resp.text),
                    )

            elif delete_mode == "revoke_public":
                resp = self.session.delete(
                    f"{DRIVE_API_V3}/files/{clean_id}/permissions/anyone",
                    params={"supportsAllDrives": "true"},
                    timeout=25,
                )
                if resp.status_code in (200, 204):
                    return (
                        "SUCCESS",
                        item_name,
                        f"{item_kind}「{item_name}」已关闭公开下载权限（外链已失效）",
                    )
                elif resp.status_code == 404:
                    return (
                        "NOT_FOUND",
                        item_name,
                        f"{item_kind}「{item_name}」原本无公开权限或文件已不存在",
                    )
                else:
                    return (
                        "ERROR",
                        item_name,
                        format_google_api_error(resp.status_code, resp.text),
                    )
            else:
                return "ERROR", item_name, f"未知的删除模式: {delete_mode}"

        except Exception as e:
            return "ERROR", item_name, f"请求异常: {str(e)}"

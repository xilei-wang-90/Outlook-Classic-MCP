"""Mail COM operations."""

from __future__ import annotations

import time
from typing import Any

import pythoncom
import win32com.client

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.errors import OutlookError
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_FOLDER_INBOX,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_IMPORTANCE_NORMAL,
    OL_MAIL_ITEM,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import from_iso, to_iso, truncate
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir
from outlook_mcp.utils.safety import safe_dasl


def _mail_summary(item: Any) -> dict[str, Any]:
    attachments = _safe_get(item, "Attachments")
    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "has_attachments": attachments.Count > 0 if attachments else False,
        "importance": _safe_get(item, "Importance"),
        "preview": truncate(_safe_get(item, "Body", ""), 200),
    }


def _mail_full(item: Any, include_body: bool = True) -> dict[str, Any]:
    attachments = []
    if _safe_get(item, "Attachments"):
        for i, att in enumerate(item.Attachments, start=1):
            attachments.append(
                {
                    "index": i,
                    "filename": att.FileName,
                    "size_bytes": _safe_get(att, "Size"),
                }
            )
    result = {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", ""),
        "attachments": attachments,
    }
    if include_body:
        result["body"] = _safe_get(item, "Body", "")
        result["html_body"] = _safe_get(item, "HTMLBody", "")
    return result


def list_mails(
    outlook: Any,
    namespace: Any,
    *,
    folder: str | None = "inbox",
    limit: int = 25,
    offset: int = 0,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    clauses: list[str] = []
    if unread_only:
        clauses.append("[UnRead] = True")
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    if since_dt:
        clauses.append(f"[ReceivedTime] >= '{since_dt.strftime('%m/%d/%Y %H:%M %p')}'")
    if until_dt:
        clauses.append(f"[ReceivedTime] <= '{until_dt.strftime('%m/%d/%Y %H:%M %p')}'")

    if clauses:
        items = items.Restrict(" AND ".join(clauses))

    from_lower = from_address.lower() if from_address else None
    results: list[dict[str, Any]] = []
    skipped = 0
    for item in items:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        if from_lower:
            sender = (_safe_get(item, "SenderEmailAddress") or "").lower()
            if from_lower not in sender:
                continue
        if skipped < offset:
            skipped += 1
            continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "folder": f.Name,
        "count": len(results),
        "offset": offset,
        "limit": limit,
        "items": results,
        "has_more": len(results) == limit,
        "next_offset": offset + len(results) if len(results) == limit else None,
    }


def search_mails(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    folder: str | None = "inbox",
    limit: int = 25,
    scope: str = "subject_body",
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    if scope == "dasl":
        # Caller is explicitly passing a raw DASL filter; don't mangle it.
        filtered = items.Restrict(query)
    elif scope == "subject":
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
        )
    elif scope == "from":
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
            f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%'"
        )
    else:  # subject_body
        # Restrict by subject only — body properties (textdescription) are not
        # locally indexed on Exchange Online and cause Restrict() to download
        # every message body from the server, hanging indefinitely. Instead,
        # pre-filter by subject via DASL, then match body in Python on the
        # smaller result set.
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
        )
        query_lower = query.lower()

    results: list[dict[str, Any]] = []
    for item in filtered:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        if scope == "subject_body":
            body = (_safe_get(item, "Body") or "").lower()
            subject = (_safe_get(item, "Subject") or "").lower()
            if query_lower not in subject and query_lower not in body:
                continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "query": query,
        "scope": scope,
        "folder": f.Name,
        "count": len(results),
        "items": results,
    }


_MAIL_FOLDER_TYPES = {0, 1}  # olMailItem, olMeetingItem default item types


def _walk_mail_folders(root: Any) -> list[Any]:
    """Recursively collect all mail-type folders under root."""
    folders: list[Any] = []
    for folder in root.Folders:
        if _safe_get(folder, "DefaultItemType", -1) in _MAIL_FOLDER_TYPES:
            folders.append(folder)
            folders.extend(_walk_mail_folders(folder))
    return folders


_WS_CONN_STR = "Provider=Search.CollatorDSO.1;Extended Properties='Application=Windows'"

# Windows Search property names for Outlook mail items
_WS_SQL = """
SELECT TOP {limit}
    System.Subject,
    System.Message.FromName,
    System.Message.DateReceived,
    System.ItemFolderNameDisplay,
    System.ItemPathDisplay,
    System.Message.ToName,
    System.Size
FROM SystemIndex
WHERE CONTAINS(*,{fts})
  AND System.Kind = 'email'
  {scope_clause}
ORDER BY System.Message.DateReceived DESC
"""


def _ws_quote(term: str) -> str:
    """Wrap a search term for Windows Search CONTAINS().

    Phrases need double-quotes inside the FTS expression; single words work
    without. We always phrase-quote to handle multi-word queries correctly.
    The term itself is wrapped in outer single quotes for the SQL string.
    """
    escaped = term.replace("'", "''").replace('"', '""')
    return f'\'"\\"{escaped}\\""\''.replace("\\\"", '"')


def remote_search_mails(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    scope: str = "all",
    limit: int = 50,
    search_subfolders: bool = True,
    timeout_sec: int = 30,
) -> dict[str, Any]:
    """Search mail via Windows Search — the same engine Outlook's search bar uses.

    Queries the SystemIndex OLE DB catalog with a CONTAINS() full-text expression,
    searching subject AND body in one pass. Results are then linked back to live
    MAPI items to populate EntryIDs and metadata.
    """
    # Build the FTS expression: '"phrase here"'
    esc = query.replace('"', '""')
    fts_expr = f'\'"{esc}"\''

    # Optionally scope to a specific folder subtree
    if scope != "all":
        folder = resolve_folder(namespace, scope)
        folder_path = _safe_get(folder, "FolderPath", "").replace("\\", "/")
        # ItemPathDisplay for Outlook items looks like /Mailbox/FolderName/...
        store_name = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent.Name
        ws_path = f"/{store_name}/{folder_path.lstrip('/')}".rstrip("/")
        scope_clause = f"AND System.ItemPathDisplay LIKE '{ws_path}%'"
    else:
        scope_clause = ""

    sql = _WS_SQL.format(limit=limit * 2, fts=fts_expr, scope_clause=scope_clause)

    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.Open(_WS_CONN_STR)
        rs = win32com.client.Dispatch("ADODB.Recordset")
        rs.Open(sql, conn)
    except Exception as exc:
        raise OutlookError(f"Windows Search query failed: {exc}") from exc

    # Collect raw hits from the index
    hits: list[dict[str, Any]] = []
    while not rs.EOF and len(hits) < limit * 2:
        hits.append({
            "subject":    rs.Fields.Item("SYSTEM.SUBJECT").Value,
            "from_name":  rs.Fields.Item("SYSTEM.MESSAGE.FROMNAME").Value,
            "date":       rs.Fields.Item("SYSTEM.MESSAGE.DATERECEIVED").Value,
            "folder":     rs.Fields.Item("SYSTEM.ITEMFOLDERNAMEDISPLAY").Value,
            "path":       rs.Fields.Item("SYSTEM.ITEMPATHDISPLAY").Value,
        })
        rs.MoveNext()
    rs.Close()
    conn.Close()

    # Resolve each hit back to a live MAPI item so we have the EntryID,
    # preview, attachments, etc. Match by folder path + subject + date.
    results: list[dict[str, Any]] = []
    for hit in hits:
        if len(results) >= limit:
            break
        try:
            folder_name = (hit["folder"] or "").strip()
            subject = (hit["subject"] or "").strip()
            if not folder_name or not subject:
                continue

            # Navigate to the folder via the path segments in ItemPathDisplay
            # e.g. /Lingkai.Shi@amd.com/Inbox/subject  -> store/Inbox
            path_parts = [p for p in (hit["path"] or "").split("/") if p]
            # path_parts[0]=store, path_parts[1]=folder, rest=subject segments
            folder_spec = "/".join(path_parts[1:-1]) if len(path_parts) > 2 else folder_name
            try:
                mapi_folder = resolve_folder(namespace, folder_spec)
            except Exception:
                mapi_folder = resolve_folder(namespace, folder_name)

            # Find the item by subject within that folder
            esc_subj = subject.replace("'", "''")
            item = mapi_folder.Items.Find(f"[Subject] = '{esc_subj}'")
            if item is None:
                # Fall back: use Restrict on subject DASL (handles special chars better)
                safe_subj = safe_dasl(subject)
                filtered = mapi_folder.Items.Restrict(
                    f"@SQL=\"urn:schemas:httpmail:subject\" = '{safe_subj}'"
                )
                item = next(iter(filtered), None)
            if item is None:
                continue

            cls = _safe_get(item, "Class")
            if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
                continue
            results.append({**_mail_summary(item), "folder": folder_name})
        except Exception:
            continue

    return {
        "query": query,
        "scope": scope,
        "search_subfolders": search_subfolders,
        "count": len(results),
        "items": results,
    }


def get_mail(outlook: Any, namespace: Any, *, entry_id: str, include_body: bool = True) -> dict[str, Any]:
    return _mail_full(get_item_by_id(namespace, entry_id), include_body=include_body)


def send_mail(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
) -> dict[str, Any]:
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
    mail.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)

    for raw_path in attachments or []:
        mail.Attachments.Add(validate_attachment_path(raw_path))

    if save_only:
        mail.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": mail.EntryID,
            "subject": mail.Subject,
        }

    mail.Send()
    return {
        "status": "sent",
        "to": to,
        "cc": cc or [],
        "bcc": bcc or [],
        "subject": subject,
    }


def reply_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        reply.HTMLBody = body + (reply.HTMLBody or "")
    else:
        reply.Body = body + "\n\n" + (reply.Body or "")
    for raw_path in attachments or []:
        reply.Attachments.Add(validate_attachment_path(raw_path))
    reply.Send()
    return {
        "status": "sent",
        "reply_all": reply_all,
        "in_reply_to": entry_id,
        "subject": reply.Subject,
    }


def forward_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    html: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if cc:
        fwd.CC = "; ".join(cc)
    if body:
        if html:
            fwd.BodyFormat = OL_FORMAT_HTML
            fwd.HTMLBody = body + (fwd.HTMLBody or "")
        else:
            fwd.Body = body + "\n\n" + (fwd.Body or "")
    fwd.Send()
    return {"status": "sent", "forwarded": entry_id, "to": to, "subject": fwd.Subject}


def move_mail(outlook: Any, namespace: Any, *, entry_id: str, target_folder: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    target = resolve_folder(namespace, target_folder)
    moved = item.Move(target)
    return {"status": "moved", "new_entry_id": moved.EntryID, "folder": target.Name}


def delete_mail(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def mark_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    out_dir = validate_output_dir(output_dir)
    saved: list[str] = []
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]
    import os

    for att in attachments:
        target = os.path.join(out_dir, att.FileName)
        att.SaveAsFile(target)
        saved.append(target)
    return {
        "status": "saved",
        "count": len(saved),
        "files": saved,
        "output_dir": out_dir,
    }

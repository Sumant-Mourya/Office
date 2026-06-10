"""Google Sheets sync – date × hour matrix layout + config sheet.

Layout:
  Row 1   : Headers → A1="DATE", B1="12:00AM-01:00AM", C1="01:00AM-02:00AM", ...
  Row 2+  : One row per date.  A=date, B..Y=hourly data cells.
  Each hourly cell is updated every minute with the latest stats for that hour.
  When the hour rolls over, the next column is used.
"""

import socket
import ipaddress
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import UI_PORT
from logger_setup import get_logger

log = get_logger("sheets.sync")

# 24 hour slots – column B through Y  (indices 1..24 in 0-based)
_HOUR_SLOTS = []
for _h in range(24):
    _start = datetime(2000, 1, 1, _h, 0)
    _end = datetime(2000, 1, 1, (_h + 1) % 24, 0) if _h < 23 else datetime(2000, 1, 2, 0, 0)
    _HOUR_SLOTS.append(f"{_start.strftime('%I:%M%p')}-{_end.strftime('%I:%M%p')}")

# Full header row: DATE + 24 hour slots
_HEADERS = ["DATE"] + _HOUR_SLOTS

# Column letters A-Y (25 columns)
def _col_letter(idx: int) -> str:
    """Convert 0-based column index to spreadsheet column letter(s)."""
    if idx < 26:
        return chr(65 + idx)
    # For safety, support beyond Z
    first = idx // 26 - 1
    second = idx % 26
    return chr(65 + first) + chr(65 + second)


def _get_local_ip() -> str:
    """Get a usable private IPv4 LAN address for this machine."""

    def _is_usable_private_ipv4(ip: str) -> bool:
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            return False
        return (
            addr.version == 4
            and addr.is_private
            and not addr.is_loopback
            and not addr.is_link_local
        )

    candidates: list[str] = []
    seen: set[str] = set()

    def _add_candidate(ip: str):
        if ip in seen:
            return
        seen.add(ip)
        if _is_usable_private_ipv4(ip):
            candidates.append(ip)

    # UDP probe usually reveals the active interface IP.
    for target in ("8.8.8.8", "1.1.1.1", "192.168.1.1", "10.255.255.255"):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect((target, 80))
            _add_candidate(s.getsockname()[0])
            s.close()
        except Exception:
            continue

    # Hostname resolution as a fallback source.
    try:
        hostname = socket.gethostname()
        for _, _, _, _, sockaddr in socket.getaddrinfo(
            hostname, None, socket.AF_INET, socket.SOCK_DGRAM
        ):
            _add_candidate(sockaddr[0])
    except Exception:
        pass

    def _rank(ip: str) -> int:
        if ip.startswith("192.168."):
            return 0
        if ip.startswith("10."):
            return 1
        if ip.startswith("172."):
            return 2
        return 3

    if candidates:
        candidates.sort(key=_rank)
        return candidates[0]

    return "127.0.0.1"


class SheetSync:
    """Date × Hour matrix sync + config tab."""

    def __init__(self, creds, spreadsheet_id: str, sheet_name: str):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.service = build("sheets", "v4", credentials=creds)
        self.sheets = self.service.spreadsheets()

    # ── Tab helpers ───────────────────────────────────────────────────

    def _tab_exists(self, name: str) -> bool:
        meta = self.sheets.get(spreadsheetId=self.spreadsheet_id).execute()
        return any(
            s["properties"]["title"] == name for s in meta.get("sheets", [])
        )

    def sheet_exists(self) -> bool:
        try:
            return self._tab_exists(self.sheet_name)
        except HttpError as exc:
            log.error("sheet_exists check failed: %s", exc)
            raise

    def create_sheet(self):
        """Create the PC tab if it doesn't exist, then write headers."""
        if not self._tab_exists(self.sheet_name):
            self.sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [
                    {"addSheet": {"properties": {"title": self.sheet_name}}}
                ]},
            ).execute()
            log.info("Created sheet tab '%s'", self.sheet_name)

        # Write full header row: DATE + 24 hour columns
        end_col = _col_letter(len(_HEADERS) - 1)
        self.sheets.values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self.sheet_name}'!A1:{end_col}1",
            valueInputOption="RAW",
            body={"values": [_HEADERS]},
        ).execute()

    def log_start_time(self):
        """Append a row with the program start timestamp."""
        now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        row = [[f"▶ Program started: {now_str}"] + [""] * 24]
        self.sheets.values().append(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self.sheet_name}'!A:Y",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": row},
        ).execute()
        log.info("Logged start time to sheet")

    # ── Config sheet ──────────────────────────────────────────────────

    def ensure_config_sheet(self, pc_name: str):
        """Create / update a 'config' tab with pc_name + local WiFi/LAN IPs."""
        tab = "config"
        if not self._tab_exists(tab):
            self.sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [
                    {"addSheet": {"properties": {"title": tab}}}
                ]},
            ).execute()
            self.sheets.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A1:D1",
                valueInputOption="RAW",
                body={"values": [["PC_NAME", "LOCAL_IP", "APP_URL", "LAST_SEEN"]]},
            ).execute()

        result = self.sheets.values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{tab}'!A2:D200",
        ).execute()
        rows = result.get("values", [])

        row_idx = None
        for i, r in enumerate(rows, start=2):
            if r and r[0] == pc_name:
                row_idx = i
                break

        local_ip = _get_local_ip()
        if local_ip == "127.0.0.1" and row_idx:
            existing_row = rows[row_idx - 2]
            existing_ip = existing_row[1] if len(existing_row) > 1 else ""
            if existing_ip and existing_ip != "127.0.0.1":
                local_ip = existing_ip
        if local_ip == "127.0.0.1":
            log.warning(
                "Could not detect private LAN IP for %s; storing localhost.",
                pc_name,
            )
        app_url = f"http://{local_ip}:{UI_PORT}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        values = [[pc_name, local_ip, app_url, now_str]]

        if row_idx:
            self.sheets.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A{row_idx}:D{row_idx}",
                valueInputOption="RAW",
                body={"values": values},
            ).execute()
        else:
            self.sheets.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A:D",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()
        log.info("Config sheet updated for %s (url=%s)", pc_name, app_url)

    def pc_name_taken(self, pc_name: str) -> bool:
        """Check if a PC name is already claimed in config.

        A tab can exist without a config entry if a previous start attempt failed
        midway. Those partial cases are treated as not-taken so the same machine
        can retry setup successfully.
        """
        if not self._tab_exists(pc_name):
            return False

        tab = "config"
        if not self._tab_exists(tab):
            return False

        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A2:A200",
            ).execute()
        except HttpError:
            # Be conservative when we cannot read ownership metadata.
            return True

        for row in result.get("values", []):
            if row and row[0] == pc_name:
                return True
        return False

    def read_config_pcs(self) -> list[dict]:
        """Return list of PCs from config tab."""
        tab = "config"
        if not self._tab_exists(tab):
            return []
        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{tab}'!A2:D100",
            ).execute()
            out = []
            for row in result.get("values", []):
                if row:
                    out.append({
                        "pc_name": row[0] if len(row) > 0 else "",
                        "local_ip": row[1] if len(row) > 1 else "",
                        "app_url": row[2] if len(row) > 2 else "",
                        "last_seen": row[3] if len(row) > 3 else "",
                    })
            return out
        except HttpError:
            return []

    # ── Formatting helpers ────────────────────────────────────────────

    @staticmethod
    def _fmt_usage(usage_dict: dict) -> str:
        """Format usage dict as multiline: 'name = Xh Ym Zs' per line."""
        items = sorted(usage_dict.items(), key=lambda x: -x[1])
        lines = []
        for name, sec in items:
            h, rem = divmod(int(sec), 3600)
            m, s = divmod(rem, 60)
            if h > 0:
                lines.append(f"{name} = {h}h {m}m {s}s")
            elif m > 0:
                lines.append(f"{name} = {m}m {s}s")
            else:
                lines.append(f"{name} = {s}s")
        return "\n".join(lines) if lines else ""

    @staticmethod
    def _fmt_seconds(sec: float) -> str:
        h, rem = divmod(int(sec), 3600)
        m, s = divmod(rem, 60)
        return f"{h}h {m}m {s}s"

    @staticmethod
    def _hour_slot_index(hour_slot: str) -> int | None:
        """Convert hour slot string like '10:00AM-11:00AM' to 0-based hour index."""
        try:
            start_str = hour_slot.split("-")[0].strip().upper()
            dt = datetime.strptime(start_str, "%I:%M%p")
            return dt.hour
        except Exception:
            return None

    # ── Sync – date × hour matrix ────────────────────────────────────

    def sync(self, data: dict):
        """Upsert hourly data into the date×hour matrix.

        Each date gets one row.  Each hour-slot maps to a column (B=00:00, C=01:00, …).
        The cell for the current hour is updated every call (every minute).
        """
        target_date = data["date"]
        hours_data = data.get("hours", {})

        # Build cell content for each active hour
        hour_cells: dict[int, str] = {}  # hour_index -> cell text
        for slot, hdata in hours_data.items():
            idx = self._hour_slot_index(slot)
            if idx is None:
                continue

            mouse = hdata.get("mouse_clicks", 0)
            keys = hdata.get("key_presses", 0)
            work = hdata.get("work_seconds", 0)
            idle = hdata.get("idle_seconds", 0)
            websites = hdata.get("websites", {})
            windows = hdata.get("windows", {})

            cell_lines = [
                f"Mouse: {mouse} | Keys: {keys}",
                f"Work: {self._fmt_seconds(work)} | Idle: {self._fmt_seconds(idle)}",
            ]

            # Top 5 apps
            if windows:
                top_apps = sorted(windows.items(), key=lambda x: -x[1])[:5]
                app_parts = [f"{a}: {self._fmt_seconds(s)}" for a, s in top_apps]
                cell_lines.append(f"Apps: {', '.join(app_parts)}")

            # Top 5 websites
            if websites:
                top_sites = sorted(websites.items(), key=lambda x: -x[1])[:5]
                site_parts = [f"{s}: {self._fmt_seconds(t)}" for s, t in top_sites]
                cell_lines.append(f"Sites: {', '.join(site_parts)}")

            hour_cells[idx] = "\n".join(cell_lines)

        if not hour_cells:
            log.debug("No hour data to sync for %s", target_date)
            return

        # Find existing date row (skip header row 1, skip start-time rows)
        try:
            result = self.sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{self.sheet_name}'!A:A",
            ).execute()
            existing = result.get("values", [])
        except HttpError as exc:
            log.error("Failed to read sheet: %s", exc)
            return

        row_idx = None
        for i, r in enumerate(existing):
            if r and r[0] == target_date:
                row_idx = i + 1  # 1-based
                break

        # If no row exists for this date, append a new row with the date in col A
        if not row_idx:
            try:
                # Append a row with just the date; we'll fill hour cells next
                self.sheets.values().append(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"'{self.sheet_name}'!A:A",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [[target_date]]},
                ).execute()
                # Re-read to find the actual row index
                result = self.sheets.values().get(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"'{self.sheet_name}'!A:A",
                ).execute()
                existing = result.get("values", [])
                for i, r in enumerate(existing):
                    if r and r[0] == target_date:
                        row_idx = i + 1
                        break
                if not row_idx:
                    log.error("Could not find newly appended row for %s", target_date)
                    return
            except HttpError as exc:
                log.error("Failed to append date row: %s", exc)
                return

        # Update each hour cell individually
        batch_data = []
        for hour_idx, cell_text in hour_cells.items():
            col = _col_letter(hour_idx + 1)  # +1 because col A is DATE
            cell_range = f"'{self.sheet_name}'!{col}{row_idx}"
            batch_data.append({
                "range": cell_range,
                "values": [[cell_text]],
            })

        if batch_data:
            try:
                self.sheets.values().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={
                        "valueInputOption": "RAW",
                        "data": batch_data,
                    },
                ).execute()
                log.debug(
                    "Synced %d hour cell(s) for %s (row %d)",
                    len(batch_data), target_date, row_idx,
                )
            except HttpError as exc:
                log.error("Sheet sync failed: %s", exc)
                raise

    # ── Access validation ─────────────────────────────────────────────

    def validate_access(self) -> bool:
        try:
            self.sheets.get(spreadsheetId=self.spreadsheet_id).execute()
            return True
        except HttpError:
            return False

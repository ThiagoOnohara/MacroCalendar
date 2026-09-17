"""Small Outlook Classic adapter with actionable error messages."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import List, Optional, Tuple


class OutlookError(RuntimeError):
    """An expected problem connecting to or using Outlook Classic."""


class OutlookClient:
    OL_FOLDER_CALENDAR = 9
    OL_APPOINTMENT_ITEM = 1
    OL_FREE = 0

    def __init__(self, calendar_name: Optional[str] = None):
        self.calendar_name = calendar_name
        self.application = None
        self.namespace = None
        self.calendar = None

    def connect(self) -> "OutlookClient":
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise OutlookError(
                "pywin32 não está instalado. Instale as dependências com "
                "'python -m pip install -e .' em Windows."
            ) from exc

        try:
            self.application = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.application.GetNamespace("MAPI")
            root = self.namespace.GetDefaultFolder(self.OL_FOLDER_CALENDAR)
        except Exception as exc:  # COM errors vary by Office version/profile.
            raise OutlookError(
                "Não foi possível abrir o Outlook via COM.\n"
                "Use o Outlook Classic para Windows, abra-o, faça login e tente novamente.\n"
                "O New Outlook não oferece suporte ao Outlook Object Model/COM."
            ) from exc

        if self.calendar_name:
            self.calendar = self._find_calendar(root, self.calendar_name)
            if self.calendar is None:
                available = ", ".join(self.calendar_names()) or "(nenhum encontrado)"
                raise OutlookError(
                    f"O calendário '{self.calendar_name}' não foi encontrado.\n"
                    f"Calendários disponíveis: {available}"
                )
        else:
            self.calendar = root
            self.calendar_name = str(getattr(root, "Name", "Calendário"))

        return self

    def _require_connection(self) -> None:
        if self.namespace is None:
            raise OutlookError("Outlook ainda não foi conectado. Execute connect() primeiro.")

    def _require_calendar(self) -> None:
        self._require_connection()
        if self.calendar is None:
            raise OutlookError("Nenhum calendário foi selecionado.")

    def _find_calendar(self, root, name: str):
        target = name.strip().casefold()

        def visit(folder, prefix: str = ""):
            folder_name = str(getattr(folder, "Name", "")).strip()
            path = f"{prefix}{folder_name}" if folder_name else prefix.rstrip(" / ")
            if folder_name.casefold() == target or path.casefold() == target:
                return folder
            for child in list(getattr(folder, "Folders", [])):
                found = visit(child, f"{path} / " if path else "")
                if found is not None:
                    return found
            return None

        return visit(root)

    def calendar_names(self) -> List[str]:
        self._require_connection()
        root = self.namespace.GetDefaultFolder(self.OL_FOLDER_CALENDAR)
        result: List[str] = []

        def visit(folder, prefix: str = "") -> None:
            name = str(getattr(folder, "Name", "")).strip()
            if name:
                result.append(f"{prefix}{name}")
            for child in list(getattr(folder, "Folders", [])):
                visit(child, f"{prefix}{name} / " if name else prefix)

        visit(root)
        return result

    def existing_appointments(self, subject: str, start: datetime) -> List[object]:
        self._require_calendar()
        matches: List[object] = []
        for item in list(self.calendar.Items):
            try:
                item_subject = str(item.Subject)
                item_start = item.Start
                same_minute = item_start.strftime("%Y-%m-%d %H:%M") == start.strftime("%Y-%m-%d %H:%M")
                if item_subject == subject and same_minute:
                    matches.append(item)
            except (AttributeError, TypeError, ValueError):
                # Calendar folders can contain non-appointment items.
                continue
        return matches

    def upsert_event(
        self,
        *,
        subject: str,
        start: datetime,
        duration_minutes: int,
        reminder_minutes: int,
        category: str,
        dry_run: bool = False,
    ) -> Tuple[str, bool]:
        """Create or update one event, returning (action, duplicate_exists)."""
        matches = self.existing_appointments(subject, start)
        action = "atualizar" if matches else "criar"
        if dry_run:
            return action, len(matches) > 1

        appointment = matches[0] if matches else self.calendar.Items.Add(self.OL_APPOINTMENT_ITEM)
        appointment.Subject = subject
        appointment.Start = start.strftime("%Y-%m-%d %H:%M")
        appointment.End = (start + timedelta(minutes=duration_minutes)).strftime(
            "%Y-%m-%d %H:%M"
        )
        appointment.BusyStatus = self.OL_FREE
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = reminder_minutes
        appointment.Categories = category
        # Do not add attendees: saving an appointment must not create invitations.
        # Existing attendees are intentionally preserved when updating an item.
        appointment.Save()
        return action, len(matches) > 1

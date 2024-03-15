# Copyright (C) 2024 - ForgeFlow S.L.
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import base64
from datetime import datetime
import icalendar
import re
from odoo.tools import plaintext2html

from odoo import _, fields, models
from odoo.exceptions import ValidationError


class CalendarImportIcs(models.TransientModel):
    """
    This wizard is used to import ics files to calendar
    """

    _name = "calendar.import.ics"
    _description = "Calendar Import Ics"

    import_ics_file = fields.Binary(required=True)
    import_ics_filename = fields.Char()
    import_start_date = fields.Date("Start Import Date")
    import_end_date = fields.Date("End Import Date")
    partner_id = fields.Many2one("res.partner", string="Partner")
    do_remove_old_event = fields.Boolean(
        string="Remove old events?",
        help="If checked, the previously imported events "
        "that are not in this import will be deleted",
        default=True,
    )

    def button_import(self):
        imported_uids = []
        self.ensure_one()
        assert self.import_ics_file
        extension = self.import_ics_filename.split(".")[1]
        if extension != "ics":
            raise ValidationError(_("Only ics files are supported"))
        if self.env.user and not self.partner_id:
            self.partner_id = self.env.user.partner_id.id
        file_decoded = base64.b64decode(self.import_ics_file)
        file_str = file_decoded.decode("utf-8")

        calendar = icalendar.Calendar.from_ical(file_str)
        for event in calendar.walk('VEVENT'):
            start_date = event.get("DTSTART") and event.decoded("DTSTART")
            end_date = event.get("DTEND") and event.decoded("DTEND")
            if not start_date or not end_date:
                continue
            if (not self.import_start_date or not self.import_end_date) or (
                self.import_start_date <= start_date.date()
                and self.import_end_date >= end_date.date()
            ):
                vals = self._prepare_event_vals(event)
                imported_uids.append(vals["event_identifier"])
                existing_event = self.env["calendar.event"].search(
                    [("event_identifier", "=", vals["event_identifier"])]
                )
                if existing_event:
                    self._update_event(existing_event, vals)
                else:
                    self._create_event(vals)

        if self.do_remove_old_event:
            self._delete_non_imported_events(imported_uids)

    def _prepare_event_vals(self, ical_event):
        vals = {
            "start": ical_event.decoded("DTSTART").strftime("%Y-%m-%d %H:%M:00"),
            "stop": ical_event.decoded("DTEND").strftime("%Y-%m-%d %H:%M:00"),
            "name": ical_event.decoded("SUMMARY").decode("UTF-8"),
            "event_identifier": ical_event.decoded("UID").decode("UTF-8"),
            "partner_ids": [(4, self.partner_id.id)],
        }
        if ical_event.get("DESCRIPTION"):
            desc = ical_event.decoded("DESCRIPTION").decode("UTF-8")
            vals['description'] = plaintext2html(desc)
            m = re.findall(r"^Address: (.+)$", desc, re.MULTILINE)
            if len(m) == 1:
                vals['location'] = m[0]
        return vals

    def _update_event(self, event, vals):
        update_vals = {}
        if event.start != vals["start"]:
            update_vals["start"] = vals["start"]
        if event.stop != vals["stop"]:
            update_vals["stop"] = vals["stop"]
        if event.name != vals["name"]:
            update_vals["name"] = vals["name"]
        if self.partner_id not in event.partner_ids:
            update_vals["partner_ids"] = [(4, self.partner_id.id, 0)]
        if event.location != vals.get("location"):
            update_vals["location"] = vals.get("location")
        if event.description != vals.get("description"):
            update_vals["description"] = vals.get("description")
        event.write(update_vals)

    def _create_event(self, vals):
        self.env["calendar.event"].create(vals)

    def _delete_non_imported_events(self, imported_events):
        domain = [
            ("event_identifier", "!=", False),
            ("event_identifier", "not in", imported_events),
            ("partner_ids", "in", self.partner_id.id),
        ]

        if self.import_start_date:
            domain.append(("start", ">=", self.import_start_date))

        if self.import_end_date:
            domain.append(("stop", "<=", self.import_end_date))

        non_imported_events = self.env["calendar.event"].search(domain)
        for non_imported_event in non_imported_events:
            non_imported_event.write({"partner_ids": [(3, self.partner_id.id)]})
        if not non_imported_events.partner_ids:
            non_imported_events.unlink()

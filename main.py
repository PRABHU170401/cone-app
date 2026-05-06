from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.widget import Widget
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.graphics import (Color, RoundedRectangle, Rectangle, Ellipse, Line)
from kivy.metrics import dp
from kivy.core.window import Window
from kivy.animation import Animation
from datetime import datetime
from kivy.clock import Clock
import os
import json

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# ── Palette ───────────────────────────────────────────────────────────────────
# Window.clearcolor moved to build()

C_BG        = (0.11, 0.09, 0.20, 1)
C_SURFACE   = (0.18, 0.14, 0.30, 1)
C_GLASS     = (0.22, 0.17, 0.38, 0.85)
C_PRIMARY   = (0.55, 0.40, 0.95, 1)
C_ACCENT    = (0.75, 0.55, 1.00, 1)
C_SUCCESS   = (0.30, 0.85, 0.65, 1)
C_DANGER    = (0.95, 0.35, 0.50, 1)
C_TEXT      = (0.92, 0.90, 1.00, 1)
C_MUTED     = (0.55, 0.50, 0.72, 1)
C_RESET     = (0.90, 0.40, 0.60, 1)
C_EDIT      = (0.35, 0.65, 0.95, 1)
C_WARN      = (0.95, 0.75, 0.20, 1)
C_RESULT    = (0.20, 0.75, 0.55, 1)
C_GOLD      = (1.00, 0.82, 0.20, 1)
C_PANEL_BG  = (0.08, 0.06, 0.16, 1)   # darker bg for side panel


# ── Canvas helpers ────────────────────────────────────────────────────────────
def _draw_glass_card(instance, value, radius=dp(18), color=None):
    c = color or C_BG
    instance.canvas.before.clear()
    with instance.canvas.before:
        Color(*c)
        RoundedRectangle(pos=instance.pos, size=instance.size, radius=[radius])
        Color(0.55, 0.40, 0.95, 0.18)
        Line(
            rounded_rectangle=(
                instance.pos[0] + 1, instance.pos[1] + 1,
                instance.size[0] - 2, instance.size[1] - 2,
                radius
            ),
            width=1.2,
        )


def make_card(padding=dp(16), radius=dp(18), color=None):
    layout = BoxLayout(orientation='vertical', padding=padding, spacing=dp(8),
                       size_hint_y=None)
    layout.bind(minimum_height=layout.setter('height'))

    def _draw(inst, val):
        _draw_glass_card(inst, val, radius=radius, color=color)

    layout.bind(pos=_draw, size=_draw)
    return layout


def section_label(text, color=None, font_size=None, bold=False):
    lbl = Label(
        text=text,
        color=color or C_TEXT,
        font_size=font_size or dp(15),
        bold=bold,
        halign='left',
        valign='middle',
        size_hint_y=None,
        height=dp(32),
    )
    lbl.bind(size=lambda inst, val: setattr(inst, 'text_size', (val[0], None)))
    return lbl


def make_input(hint='', next_input=None):
    ti = TextInput(
        hint_text=hint,
        multiline=False,
        input_filter='int',
        font_size=dp(18),
        size_hint_y=None,
        height=dp(52),
        padding=[dp(14), dp(15)],
        background_color=(*C_SURFACE[:3], 1),
        foreground_color=C_TEXT,
        cursor_color=C_PRIMARY,
        hint_text_color=(*C_MUTED[:3], 0.55),
    )
    if next_input is not None:
        def on_text_validate(instance, _nxt=next_input):
            Clock.schedule_once(lambda dt: setattr(_nxt, 'focus', True), 0.05)
        ti.bind(on_text_validate=on_text_validate)
    return ti


def make_button(text, bg_color, text_color=None, height=dp(56)):
    tc = text_color or C_TEXT
    btn = Button(
        text=text,
        font_size=dp(16),
        bold=True,
        size_hint_y=None,
        height=height,
        background_normal='',
        background_color=(0, 0, 0, 0),
        color=tc,
    )

    def _draw(inst, val):
        inst.canvas.before.clear()
        with inst.canvas.before:
            Color(*bg_color)
            RoundedRectangle(pos=inst.pos, size=inst.size, radius=[dp(28)])
            Color(1, 1, 1, 0.08)
            RoundedRectangle(
                pos=(inst.pos[0] + dp(4), inst.pos[1] + inst.size[1] * 0.55),
                size=(inst.size[0] - dp(8), inst.size[1] * 0.38),
                radius=[dp(28)],
            )

    btn.bind(pos=_draw, size=_draw)
    return btn


def result_label(text='', color=None):
    return Label(
        text=text,
        color=color or C_TEXT,
        font_size=dp(19),
        bold=True,
        halign='center',
        size_hint_y=None,
        height=dp(38),
    )


# ── Core calculation ──────────────────────────────────────────────────────────
def calculate(length, width, height, extra):
    try:
        length, width, height, extra = (
            int(length), int(width), int(height), int(extra)
        )
    except Exception:
        return None
    if height < 0:
        return None
    total = sum((length - i) * (width - i) for i in range(height))
    return total + extra


# ── Excel helpers ─────────────────────────────────────────────────────────────
def _get_excel_path(folder, date_str=None):
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(folder, f"{date_str}.xlsx")


def _get_all_excel_files(folder):
    if not os.path.exists(folder):
        return []
    files = [f[:-5] for f in os.listdir(folder)
             if f.endswith(".xlsx") and len(f) == 15]
    return sorted(files, reverse=True)


def _load_records(file_path):
    records = []
    if not os.path.exists(file_path):
        return records
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(v is not None for v in row):
                records.append(dict(zip(headers, row)))
        wb.close()
    except Exception:
        pass
    return records


def _save_records(file_path, records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Title", "Date",
               "Small Length", "Small Width", "Small Height", "Small Extra", "Small Result",
               "Big Length",   "Big Width",   "Big Height",   "Big Extra",   "Big Result",
               "Total"]

    header_fill  = PatternFill("solid", start_color="6B50C8")
    header_font  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="9980D4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill, cell.font  = header_fill, header_font
        cell.alignment, cell.border = header_align, border

    alt_fill    = PatternFill("solid", start_color="1E1535")
    normal_fill = PatternFill("solid", start_color="150F28")
    data_font   = Font(name="Calibri", size=10, color="D4CAFF")
    center_align = Alignment(horizontal="center", vertical="center")

    for ri, rec in enumerate(records, 2):
        fill = alt_fill if ri % 2 == 0 else normal_fill
        values = [ri - 1, rec.get("Title", ""), rec.get("Date", ""),
                  rec.get("Small Length", ""), rec.get("Small Width", ""),
                  rec.get("Small Height", ""), rec.get("Small Extra", ""),
                  rec.get("Small Result", ""),
                  rec.get("Big Length", ""),   rec.get("Big Width", ""),
                  rec.get("Big Height", ""),   rec.get("Big Extra", ""),
                  rec.get("Big Result", ""),   rec.get("Total", "")]
        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill, cell.font = fill, data_font
            cell.alignment, cell.border = center_align, border

    col_widths = [4, 18, 20, 12, 12, 12, 12, 13, 10, 10, 10, 10, 11, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 35
    ws.freeze_panes = "A2"
    wb.save(file_path)
    wb.close()


def _records_equal(r1, r2):
    keys = ["Title", "Small Length", "Small Width", "Small Height", "Small Extra",
            "Big Length",  "Big Width",  "Big Height",  "Big Extra"]
    return all(str(r1.get(k, "")) == str(r2.get(k, "")) for k in keys)


# ── DATE PICKER POPUP ─────────────────────────────────────────────────────────
class DatePickerPopup(Popup):
    def __init__(self, date_list, on_date_selected, title_text="Select Date", **kwargs):
        super().__init__(
            title=title_text,
            title_color=C_ACCENT,
            title_size=dp(17),
            separator_color=C_PRIMARY,
            background_color=(*C_BG[:3], 0.97),
            size_hint=(0.92, 0.80),
            **kwargs,
        )
        self.on_date_selected = on_date_selected

        root = BoxLayout(orientation='vertical', padding=dp(14), spacing=dp(10))

        root.add_widget(Label(
            text="Choose a saved date:",
            color=C_ACCENT, font_size=dp(15), bold=True,
            size_hint_y=None, height=dp(36), halign='center',
        ))

        scroll = ScrollView(size_hint=(1, 1))
        inner  = BoxLayout(orientation='vertical', spacing=dp(8), size_hint_y=None)
        inner.bind(minimum_height=inner.setter('height'))
        scroll.add_widget(inner)

        if not date_list:
            inner.add_widget(Label(
                text="No saved records found.",
                color=C_DANGER, font_size=dp(14),
                size_hint_y=None, height=dp(40),
            ))
        else:
            for ds in date_list:
                btn = make_button(ds, C_SURFACE, C_TEXT)
                btn.bind(on_press=lambda b, d=ds: self._pick(d))
                inner.add_widget(btn)

        root.add_widget(scroll)
        cancel_btn = make_button("Cancel", C_RESET)
        cancel_btn.bind(on_press=self.dismiss)
        root.add_widget(cancel_btn)
        self.content = root

    def _pick(self, date_str):
        self.dismiss()
        self.on_date_selected(date_str)


# ── RECORD LIST POPUP ─────────────────────────────────────────────────────────
class RecordListPopup(Popup):
    def __init__(self, date_str, records, indices,
                 on_record_selected, on_record_deleted=None, **kwargs):
        super().__init__(
            title=f"Records \u2014 {date_str}",
            title_color=C_ACCENT, title_size=dp(16),
            separator_color=C_PRIMARY,
            background_color=(*C_BG[:3], 0.97),
            size_hint=(0.95, 0.88),
            **kwargs,
        )
        self.on_record_selected = on_record_selected
        self.on_record_deleted  = on_record_deleted

        root = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(10))
        root.add_widget(Label(
            text="Tap a record to edit  \u2022  \U0001F5D1 to delete",
            color=C_MUTED, font_size=dp(13), bold=True,
            size_hint_y=None, height=dp(28), halign='center',
        ))

        scroll = ScrollView(size_hint=(1, 1))
        inner  = BoxLayout(orientation='vertical', spacing=dp(8), size_hint_y=None)
        inner.bind(minimum_height=inner.setter('height'))
        scroll.add_widget(inner)

        for pos, (idx, rec) in enumerate(zip(indices, records)):
            fl   = FloatLayout(size_hint_y=None, height=dp(88))
            card = BoxLayout(
                orientation='vertical', spacing=dp(2),
                size_hint=(1, None), height=dp(82),
                padding=[dp(12), dp(8)],
                pos_hint={'x': 0, 'y': 0},
            )

            def _draw_card(inst, val, c=card):
                c.canvas.before.clear()
                with c.canvas.before:
                    Color(*C_SURFACE)
                    RoundedRectangle(pos=c.pos, size=c.size, radius=[dp(14)])
                    Color(*C_PRIMARY[:3], 0.25)
                    Line(
                        rounded_rectangle=(c.pos[0]+1, c.pos[1]+1,
                                           c.size[0]-2, c.size[1]-2, dp(14)),
                        width=1.1,
                    )
            card.bind(pos=_draw_card, size=_draw_card)

            card.add_widget(Label(
                text=f"[b]{rec.get('Title', '(no title)')}[/b]",
                markup=True, color=C_TEXT, font_size=dp(15),
                size_hint_y=None, height=dp(28),
                halign='left',
                text_size=(max(Window.width, 100) * 0.68, None),
            ))
            card.add_widget(Label(
                text=(f"Time: {str(rec.get('Date',''))[-8:]}   "
                      f"Total: {rec.get('Total','')}   "
                      f"S: {rec.get('Small Result','')}   "
                      f"B: {rec.get('Big Result','')}"),
                color=C_MUTED, font_size=dp(12),
                size_hint_y=None, height=dp(22),
                halign='left',
                text_size=(max(Window.width, 100) * 0.68, None),
            ))

            edit_overlay = Button(
                background_normal='', background_color=(0, 0, 0, 0),
                size_hint=(0.78, None), height=dp(82),
                pos_hint={'x': 0, 'y': 0},
            )
            edit_overlay.bind(
                on_press=lambda b, p=pos, i=idx, r=rec: self._pick(p, i, r)
            )

            del_btn = Button(
                text='\U0001F5D1', font_size=dp(20),
                size_hint=(None, None), width=dp(50), height=dp(50),
                pos_hint={'right': 1, 'top': 1},
                background_normal='', background_color=(0, 0, 0, 0),
                color=C_DANGER,
            )

            def _del_draw(inst, val, b=del_btn):
                b.canvas.before.clear()
                with b.canvas.before:
                    Color(*C_DANGER[:3], 0.2)
                    RoundedRectangle(pos=b.pos, size=b.size, radius=[dp(14)])
            del_btn.bind(pos=_del_draw, size=_del_draw)
            del_btn.bind(
                on_press=lambda b, p=pos, i=idx, r=rec: self._confirm_delete(p, i, r)
            )

            fl.add_widget(card)
            fl.add_widget(edit_overlay)
            fl.add_widget(del_btn)
            inner.add_widget(fl)

        root.add_widget(scroll)
        cancel_btn = make_button("Close", C_RESET)
        cancel_btn.bind(on_press=self.dismiss)
        root.add_widget(cancel_btn)
        self.content = root

    def _pick(self, pos, idx, rec):
        self.dismiss()
        self.on_record_selected(pos, idx, rec)

    def _confirm_delete(self, pos, idx, rec):
        title = rec.get('Title', '(no title)')
        cp = Popup(
            title="Delete Record?",
            title_color=C_DANGER, title_size=dp(16),
            separator_color=C_DANGER,
            background_color=(*C_BG[:3], 0.97),
            size_hint=(0.85, 0.32),
        )
        box = BoxLayout(orientation='vertical', padding=dp(14), spacing=dp(10))
        box.add_widget(Label(
            text=f"Delete  [b]{title}[/b]?",
            markup=True, color=C_TEXT, font_size=dp(15),
            halign='center', size_hint_y=None, height=dp(40),
        ))
        btn_row = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=dp(54))
        yes_btn = make_button("Delete", C_DANGER)
        no_btn  = make_button("Cancel", C_MUTED)

        def do_delete(*a):
            cp.dismiss()
            self.dismiss()
            if self.on_record_deleted:
                self.on_record_deleted(idx)

        yes_btn.bind(on_press=do_delete)
        no_btn.bind(on_press=cp.dismiss)
        btn_row.add_widget(yes_btn)
        btn_row.add_widget(no_btn)
        box.add_widget(btn_row)
        cp.content = box
        cp.open()


# ── EDIT FIELDS POPUP ─────────────────────────────────────────────────────────
class EditFieldsPopup(Popup):
    def __init__(self, rec, on_confirm, **kwargs):
        super().__init__(
            title="Edit Record",
            title_color=C_ACCENT, title_size=dp(17),
            separator_color=C_PRIMARY,
            background_color=(*C_BG[:3], 0.97),
            size_hint=(0.95, 0.92),
            **kwargs,
        )
        self.rec        = rec
        self.on_confirm = on_confirm
        self.edit_inputs = {}

        root = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(10))

        root.add_widget(Label(
            text=(f"[b]{rec.get('Title','(no title)')}[/b]\n"
                  f"[color=9980D4]{rec.get('Date','')}[/color]"),
            markup=True, color=C_TEXT, font_size=dp(14),
            halign='center', size_hint_y=None, height=dp(48),
        ))

        root.add_widget(Label(
            text="Title", color=C_MUTED, font_size=dp(13),
            size_hint_y=None, height=dp(22), halign='left',
            text_size=(max(Window.width, 100) - dp(48), None),
        ))
        self.title_edit = TextInput(
            text=str(rec.get("Title", "")),
            multiline=False,
            font_size=dp(17),
            size_hint_y=None, height=dp(50),
            padding=[dp(12), dp(14)],
            background_color=(*C_SURFACE[:3], 1),
            foreground_color=C_TEXT,
            cursor_color=C_PRIMARY,
        )
        root.add_widget(self.title_edit)

        scroll = ScrollView(size_hint=(1, 1))
        inner  = BoxLayout(orientation='vertical', spacing=dp(10), size_hint_y=None)
        inner.bind(minimum_height=inner.setter('height'))
        scroll.add_widget(inner)

        section_inputs = {}
        for sec_name, prefix in [("Small Cone", "Small"), ("Big Cone", "Big")]:
            inner.add_widget(Label(
                text=sec_name, color=C_PRIMARY, font_size=dp(15), bold=True,
                size_hint_y=None, height=dp(28), halign='left',
                text_size=(max(Window.width, 100) - dp(48), None),
            ))
            section_inputs[sec_name] = []
            for field in ["Length", "Width", "Height", "Extra"]:
                key = f"{prefix} {field}"
                ti = TextInput(
                    text=str(rec.get(key, "") or ""),
                    multiline=False, input_filter='int',
                    font_size=dp(17), size_hint_y=None, height=dp(50),
                    padding=[dp(12), dp(14)],
                    background_color=(*C_SURFACE[:3], 1),
                    foreground_color=C_TEXT, cursor_color=C_PRIMARY,
                )
                self.edit_inputs[key] = ti
                section_inputs[sec_name].append((key, ti))

            grid = GridLayout(cols=2, spacing=[dp(10), dp(6)], size_hint_y=None)
            grid.bind(minimum_height=grid.setter('height'))
            for f, ti in section_inputs[sec_name]:
                short = f.split(" ")[1]
                cell  = BoxLayout(orientation='vertical', spacing=dp(2),
                                  size_hint_y=None, height=dp(74))
                cell.add_widget(Label(
                    text=short, color=C_MUTED, font_size=dp(13),
                    size_hint_y=None, height=dp(22), halign='left',
                    text_size=(max(Window.width, 100) * 0.40, None),
                ))
                cell.add_widget(ti)
                grid.add_widget(cell)
            inner.add_widget(grid)

        root.add_widget(scroll)

        btn_row = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=dp(58))
        confirm_btn = make_button("Confirm Edit", C_SUCCESS)
        cancel_btn  = make_button("Cancel", C_RESET)
        confirm_btn.bind(on_press=self._confirm)
        cancel_btn.bind(on_press=self.dismiss)
        btn_row.add_widget(confirm_btn)
        btn_row.add_widget(cancel_btn)
        root.add_widget(btn_row)
        self.content = root

    def _confirm(self, *args):
    
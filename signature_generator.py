"""Desktop version of the Atlas Copco e-mail signature generator.

The original project ships a web based generator in ``index.html``.  This
module recreates the same behaviour as a Tk based desktop application so that
it can be packaged as a Windows executable.  The graphical interface is built
with ``tkinter`` and ``tkhtmlview`` is used to preview the generated HTML
signature.
"""

from __future__ import annotations

import html
import os
import re
import textwrap
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk
from zipfile import ZipFile, ZipInfo

from tkhtmlview import HTMLLabel


BELGIAN_NATIONAL_PATTERN = re.compile(r"^04\d{8}$")
BELGIAN_INTERNATIONAL_PATTERN = re.compile(r"^\+32\d{9}$")
GERMAN_PATTERN = re.compile(r"^\+49(?:\(0\))?\s?\d{7,13}$")


def sanitize_tel(value: str) -> str:
    """Return a telephone value that only contains digits and ``+``.

    The web version normalises a couple of edge cases where users mix the
    international prefix with ``(0)``.  The same behaviour is reproduced here so
    that validation and tel-links work identically.
    """

    sanitized = re.sub(r"[\s()-]", "", value)
    sanitized = sanitized.replace("+320", "+32")
    sanitized = sanitized.replace("+490", "+49")
    sanitized = re.sub(r"[^+0-9]", "", sanitized)
    return sanitized


def sanitize_file_segment(value: str) -> str:
    """Return a filesystem friendly variant used for exported file names."""

    return re.sub(r"[^a-zA-Z0-9@._+-]", "_", value)


def escape_rtf(value: str) -> str:
    """Escape characters with special meaning in RTF documents."""

    return (
        value.replace("\\", "\\\\")
        .replace("{", "\\{")
        .replace("}", "\\}")
    )


def is_valid_mobile(number: str) -> bool:
    """Validate Belgian and German formatted mobile numbers."""

    trimmed = number.strip()
    clean = re.sub(r"[\s()-]", "", number)
    normalized = clean.replace("+320", "+32")

    if BELGIAN_NATIONAL_PATTERN.fullmatch(normalized):
        return True

    if BELGIAN_INTERNATIONAL_PATTERN.fullmatch(normalized):
        return True

    german_candidate = re.sub(r"\s+", " ", trimmed)
    return bool(GERMAN_PATTERN.fullmatch(german_candidate))


@dataclass
class SignatureData:
    name: str
    title: str
    email: str
    mobile: str
    phone: str

    @classmethod
    def from_entries(cls, entries: dict[str, tk.StringVar]) -> "SignatureData":
        return cls(
            name=entries["name"].get().strip(),
            title=entries["title"].get().strip(),
            email=entries["email"].get().strip(),
            mobile=entries["mobile"].get().strip(),
            phone=entries["phone"].get().strip(),
        )

    def validate(self) -> bool:
        """Return ``True`` when all fields contain valid information."""

        errors = []

        if not self.name:
            errors.append("Naam is verplicht.")

        if not self.title:
            errors.append("Functie is verplicht.")

        email_lower = self.email.lower()
        if not self.email or not email_lower.endswith("@atlascopco.com"):
            errors.append("E-mail is verplicht en moet eindigen op @atlascopco.com.")

        if not self.mobile or not is_valid_mobile(self.mobile):
            errors.append(
                "Mobiel nummer is verplicht. Gebruik 04xxxxxxxx, +32xxxxxxxxx, "
                "+32(0)4xxxxxxxx of +49(0) 123456789."
            )

        if errors:
            messagebox.showerror("Ongeldige invoer", "\n".join(errors))
            return False

        return True

    def contact_info(self) -> tuple[str, str]:
        """Return contact information as (HTML, plain text)."""

        safe_mobile = html.escape(self.mobile)
        safe_phone = html.escape(self.phone)
        safe_email = html.escape(self.email)

        tel_mobile = sanitize_tel(self.mobile)
        tel_phone = sanitize_tel(self.phone)

        contact_html = (
            f'<div style="font-size:9pt;color:#2F363A;">Mobile: '
            f'<a href="tel:{tel_mobile}" style="color:#2F363A;text-decoration:none;">'
            f"{safe_mobile}</a></div>"
        )
        contact_text = f"Mobile: {self.mobile}\n"

        if self.phone:
            contact_html += (
                f'<div style="font-size:9pt;color:#2F363A;">Phone: '
                f'<a href="tel:{tel_phone}" style="color:#2F363A;text-decoration:none;">'
                f"{safe_phone}</a></div>"
            )
            contact_text += f"Phone: {self.phone}\n"

        contact_html += (
            f'<div style="font-size:9pt;color:#2F363A;">E-mail: '
            f'<a href="mailto:{safe_email}" style="color:#0092BC;text-decoration:none;">'
            f"{safe_email}</a></div>"
        )
        contact_text += f"E-mail: {self.email}\n"

        return contact_html, contact_text

    def generate(self) -> "SignaturePayload":
        contact_html, contact_text = self.contact_info()

        safe_name = html.escape(self.name)
        safe_title = html.escape(self.title)
        safe_email = html.escape(self.email)

        html_signature = textwrap.dedent(
            f"""
            <html><head><meta charset="UTF-8"></head><body>
            <div style="font-family:Calibri,Arial,sans-serif; font-size:9pt; color:#2F363A; line-height:1.4; max-width:600px;">
              <div>Best regards,</div><br>
              <div style="font-weight:bold; color:#0092BC; font-size:11pt; margin-top:4px;">{safe_name}</div>
              <div style="font-weight:bold; font-size:9pt; margin-top:2px;">{safe_title}</div>

              <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

              <div style="font-weight:bold; font-size:10pt;">Atlas Copco Tools Central Europe GmbH</div>
              <div>Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen</div>
              <div style="color:red; font-weight:bold;">Head Office: <u>New Address as of Dec 1, 2025</u>: Wetterschacht 9, D-45139 Essen</div>
              <div>Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen</div>
              <div>Managing Director: Claus Schiedeck, Peter Edmonds</div><br>

              {contact_html}

              <div style="margin-top:6px;">VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)</div>
              <div>Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)</div>
              <div>Peppol ID (Belgium): 0208:0473470658</div>

              <div style="margin-top:6px;">Privacy Policy: <a href="https://www.atlascopco.com/nl-be/itba/privacy-policy" target="_blank" style="color:#0092BC;text-decoration:none;">https://www.atlascopco.com/nl-be/itba/privacy-policy</a></div>

              <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

              <div><b style="color:#0092BC;">Level up your experience at</b> <a href="https://www.atlascopco.com" target="_blank" style="color:#0092BC; font-style:italic; text-decoration:none;">atlascopco.com</a></div>

              <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

              <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td><img src="https://groupapp.atlascopco.com/mail-signature/Signature/AtlasCopco/ACBlueSquare.png" width="8" height="8" style="display:block;"></td>
                  <td style="padding-left:5px;">
                    <span style="font-size:8pt; color:#A1A9B4; font-weight:bold;">Part of Atlas Copco Group</span>
                  </td>
                </tr>
              </table>
            </div>
            </body></html>
            """
        ).strip()

        txt_signature = textwrap.dedent(
            f"""
            Best regards,
            {self.name}
            {self.title}

            ----------------------------------------
            Atlas Copco Tools Central Europe GmbH
            Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen
            Head Office: New Address as of Dec 1, 2025: Wetterschacht 9, D-45139 Essen
            Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen
            Managing Director: Claus Schiedeck, Peter Edmonds

            {contact_text}VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)
            Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)
            Peppol ID (Belgium): 0208:0473470658

            Privacy Policy: https://www.atlascopco.com/nl-be/itba/privacy-policy
            ----------------------------------------
            Level up your experience at atlascopco.com
            ----------------------------------------
            Part of Atlas Copco Group
            """
        ).strip()

        rtf_contact_text = contact_text.replace("\n", "\\line ")
        rtf_signature = textwrap.dedent(
            f"""
            {{\\rtf1\\ansi\\deff0
            {{\\fonttbl{{\\f0 Calibri;}}}}
            \\f0\\fs18 Best regards,\\line
            \\b\\cf1 {escape_rtf(self.name)}\\b0\\line
            {escape_rtf(self.title)}\\line
            \\line
            Atlas Copco Tools Central Europe GmbH\\line
            Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen\\line
            Head Office: New Address as of Dec 1, 2025: Wetterschacht 9, D-45139 Essen\\line
            Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen\\line
            Managing Director: Claus Schiedeck, Peter Edmonds\\line
            \\line
            {escape_rtf(rtf_contact_text)}\\line
            VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)\\line
            Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)\\line
            Peppol ID (Belgium): 0208:0473470658\\line
            \\line
            Privacy Policy: https://www.atlascopco.com/nl-be/itba/privacy-policy\\line
            Level up your experience at atlascopco.com\\line
            Part of Atlas Copco Group
            }}
            """
        ).strip()

        preview_html = html_signature.replace(
            "Managing Director: Claus Schiedeck, Peter Edmonds",
            "Managing Director: <span style=\"color:#6B7280;\">[Verborgen in preview]</span>",
        )

        return SignaturePayload(
            html=html_signature,
            txt=txt_signature,
            rtf=rtf_signature,
            preview_html=preview_html,
            safe_email=safe_email,
        )


@dataclass
class SignaturePayload:
    html: str
    txt: str
    rtf: str
    preview_html: str
    safe_email: str


class SignatureApp(tk.Tk):
    """Tkinter front-end for the signature generator."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Atlas Copco Handtekening Generator")
        self.geometry("920x640")
        self.minsize(880, 600)

        self.entries: dict[str, tk.StringVar] = {
            "name": tk.StringVar(),
            "title": tk.StringVar(),
            "email": tk.StringVar(),
            "mobile": tk.StringVar(),
            "phone": tk.StringVar(),
        }

        self.payload: SignaturePayload | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        form_frame = ttk.LabelFrame(container, text="Gegevens", padding=15)
        form_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        preview_frame = ttk.LabelFrame(container, text="Voorbeeld", padding=15)
        preview_frame.grid(row=0, column=1, sticky="nsew")

        for column in range(2):
            form_frame.columnconfigure(column, weight=1)

        labels = {
            "name": "Naam",
            "title": "Functie",
            "email": "E-mail",
            "mobile": "Mobiel (verplicht)",
            "phone": "Telefoon (optioneel)",
        }

        for index, key in enumerate(labels):
            row, column = divmod(index, 2)
            label = ttk.Label(form_frame, text=labels[key])
            label.grid(row=row * 2, column=column, sticky="w", pady=(0, 2))

            entry = ttk.Entry(form_frame, textvariable=self.entries[key])
            entry.grid(row=row * 2 + 1, column=column, sticky="ew", padx=(0, 10), pady=(0, 10))

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        button_frame.columnconfigure((0, 1, 2), weight=1)

        generate_btn = ttk.Button(button_frame, text="Genereer handtekening", command=self.on_generate)
        generate_btn.grid(row=0, column=0, padx=5)

        export_btn = ttk.Button(button_frame, text="Exporteer naar Outlook (ZIP)", command=self.on_export)
        export_btn.grid(row=0, column=1, padx=5)

        clear_btn = ttk.Button(button_frame, text="Velden leegmaken", command=self.on_clear)
        clear_btn.grid(row=0, column=2, padx=5)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.preview = HTMLLabel(
            preview_frame,
            html="""
            <div style='font-family:Inter, sans-serif; color:#4B5563;'>
              Vul je gegevens in en klik op <strong>Genereer handtekening</strong> om hier een voorbeeld te zien.
            </div>
            """,
            background="white",
        )
        self.preview.pack(fill=tk.BOTH, expand=True)

    def on_generate(self) -> None:
        data = SignatureData.from_entries(self.entries)
        if not data.validate():
            return

        self.payload = data.generate()
        self.preview.set_html(self.payload.preview_html)

    def on_export(self) -> None:
        if not self.payload:
            messagebox.showinfo("Geen handtekening", "Genereer eerst een handtekening voordat je exporteert.")
            return

        base_name = f"AtlasCopco({sanitize_file_segment(self.payload.safe_email)})"

        default_path = os.path.join(os.path.expanduser("~"), f"{base_name}.zip")
        file_path = filedialog.asksaveasfilename(
            title="Bewaar ZIP-bestand",
            defaultextension=".zip",
            filetypes=[("ZIP-bestand", "*.zip")],
            initialfile=f"{base_name}.zip",
            initialdir=os.path.dirname(default_path),
        )

        if not file_path:
            return

        with ZipFile(file_path, "w") as archive:
            archive.writestr(f"{base_name}.htm", self.payload.html)
            archive.writestr(f"{base_name}.txt", self.payload.txt)
            archive.writestr(f"{base_name}.rtf", self.payload.rtf)

            # Create an empty _files folder to mimic Outlook's expected structure.
            info = ZipInfo(f"{base_name}_files/")
            archive.writestr(info, b"")

        messagebox.showinfo("Export voltooid", f"ZIP-bestand opgeslagen als:\n{file_path}")

    def on_clear(self) -> None:
        for variable in self.entries.values():
            variable.set("")

        self.payload = None
        self.preview.set_html(
            """
            <div style='font-family:Inter, sans-serif; color:#4B5563;'>
              Vul je gegevens in en klik op <strong>Genereer handtekening</strong> om hier een voorbeeld te zien.
            </div>
            """
        )


def main() -> None:
    app = SignatureApp()
    app.mainloop()


if __name__ == "__main__":
    main()
